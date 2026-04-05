"""
server.py  –  Auto-PPT Generator · FastAPI Backend
====================================================
Wraps the existing MCP-based agent (agent_ppt.py) behind a clean REST API.

Endpoints
---------
POST /generate   – run the full MCP agent, return slide plan + file path
POST /update     – re-render pptx from an edited slide plan (no LLM)
GET  /download   – stream output.pptx to the browser
GET  /health     – liveness probe

Run
---
    uvicorn server:app --reload --port 8000
"""

# NOTE: do NOT use `from __future__ import annotations` here.
# Pydantic v2 + FastAPI OpenAPI schema generation fails to resolve forward
# references from PEP-563 lazy evaluation when nested models are used in
# endpoint signatures.

import json
import logging
import os
import re
import sys
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any, List, Optional

import uvicorn
from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

# ── Path bootstrap (ensure agent_ppt imports resolve) ─────────────────────────
HERE = Path(__file__).parent.resolve()
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

# ── Import ONLY from existing agent (DO NOT rewrite any logic) ────────────────
from agent_ppt import (          # noqa: E402
    extract_json,
    _llm_call,
    _text,
    SYSTEM_PROMPT,
    OUTPUT_FILE,
    get_tools,
    call_tool,
)
from mcp import ClientSession, StdioServerParameters   # noqa: E402
from mcp.client.stdio import stdio_client              # noqa: E402

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="[SERVER] %(asctime)s %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("server")


# ── MCP server parameters ─────────────────────────────────────────────────────
# Resolves MCP scripts from ./servers/<file>  OR  ./<file>  (whichever exists).

def _resolve(filename: str) -> str:
    """Return the first existing path for a MCP server script."""
    candidates = [
        HERE / "servers" / filename,   # your layout: servers/ subfolder
        HERE / filename,               # flat layout: same dir as server.py
    ]
    for p in candidates:
        if p.exists():
            log.info("MCP script found: %s", p)
            return str(p)
    fallback = str(HERE / "servers" / filename)
    log.warning("MCP script not found; tried %s — will attempt: %s",
                [str(c) for c in candidates], fallback)
    return fallback


def _fs_params() -> StdioServerParameters:
    return StdioServerParameters(
        command="python",
        args=[_resolve("filesystem_mcp_server.py")],
        env={**os.environ},
    )


def _ppt_params() -> StdioServerParameters:
    return StdioServerParameters(
        command="python",
        args=[_resolve("ppt_mcp_server.py")],
        env={**os.environ},
    )


# ── Core pipeline (owns MCP session lifecycle) ────────────────────────────────

async def _run_with_plan(user_request: str) -> dict:
    """
    Opens MCP sessions, calls the agent pipeline, returns the structured plan:
        {
            "presentation_title": str,
            "slides": [ {"title": str, "bullet_points": [str, ...]}, ... ]
        }
    Does NOT modify agent_ppt.py – reuses its helpers directly.
    """
    async with stdio_client(_fs_params()) as (r1, w1):
        async with ClientSession(r1, w1) as fs_sess:
            await fs_sess.initialize()
            log.info("FS MCP session ready")

            async with stdio_client(_ppt_params()) as (r2, w2):
                async with ClientSession(r2, w2) as ppt_sess:
                    await ppt_sess.initialize()
                    log.info("PPT MCP session ready")

                    # PHASE 1 – Plan
                    log.info("Phase 1: planning …")
                    plan_raw = _text(_llm_call(
                        messages=[{
                            "role": "user",
                            "content": (
                                f"{user_request}\n\n"
                                "Output ONLY the JSON plan. No other text."
                            ),
                        }],
                        system=SYSTEM_PROMPT,
                        max_tokens=2048,
                    ))

                    try:
                        plan: dict = json.loads(extract_json(plan_raw))
                    except Exception as exc:
                        raise RuntimeError(
                            f"LLM plan parse failed: {exc}\nRaw:\n{plan_raw}"
                        ) from exc

                    slides = plan.get("slides", [])
                    if not slides:
                        raise RuntimeError("LLM returned an empty slide list.")

                    log.info(
                        "Plan OK: '%s' | %d slides",
                        plan.get("presentation_title"), len(slides),
                    )

                    # PHASE 2 – Execute via MCP tools
                    log.info("Phase 2: executing MCP tools …")
                    await get_tools(fs_sess)
                    await get_tools(ppt_sess)

                    await call_tool(
                        ppt_sess, "create_presentation",
                        {"filename": OUTPUT_FILE},
                    )

                    for i, slide in enumerate(slides, 1):
                        stitle  = slide.get("title", f"Slide {i}")
                        bullets = slide.get("bullet_points", [])

                        if i > 1 and len(bullets) < 3:
                            log.info("Slide %d: generating bullets via LLM", i)
                            gen = _text(_llm_call(
                                messages=[{"role": "user", "content":
                                    f'Give 4 bullet points for slide "{stitle}" '
                                    f'in context: "{user_request}". '
                                    f'Output ONLY a JSON array of strings.'}],
                                max_tokens=512,
                            ))
                            try:
                                m = re.search(r"\[.*\]", gen, re.DOTALL)
                                bullets = json.loads(m.group(0)) if m else bullets
                            except Exception:
                                bullets = [
                                    f"Key point about {stitle}.",
                                    f"Why {stitle} matters.",
                                    f"Further detail on {stitle}.",
                                ]
                            slides[i - 1]["bullet_points"] = bullets

                        log.info("add_slide #%d: %s", i, stitle)
                        await call_tool(
                            ppt_sess, "add_slide",
                            {"title": stitle, "bullet_points": bullets},
                        )

                    # PHASE 3 – Save
                    log.info("Phase 3: saving …")
                    await call_tool(ppt_sess, "save_presentation", {})
                    log.info("Done → %s", OUTPUT_FILE)

                    return plan


async def _rebuild_from_plan(slides_data: list) -> None:
    """Re-render output.pptx from an edited slide list via PPT MCP only."""
    async with stdio_client(_ppt_params()) as (r2, w2):
        async with ClientSession(r2, w2) as ppt_sess:
            await ppt_sess.initialize()

            await call_tool(ppt_sess, "create_presentation", {"filename": OUTPUT_FILE})

            for i, slide in enumerate(slides_data, 1):
                title   = slide["title"]
                bullets = list(slide["bullet_points"])
                while i > 1 and len(bullets) < 3:
                    bullets.append("Additional detail to be added.")
                await call_tool(
                    ppt_sess, "add_slide",
                    {"title": title, "bullet_points": bullets},
                )

            await call_tool(ppt_sess, "save_presentation", {})


# ── Pydantic schemas (no forward refs – all types defined above use site) ──────

class SlideSchema(BaseModel):
    title:         str
    bullet_points: List[str]


class PlanSchema(BaseModel):
    presentation_title: str
    slides:             List[SlideSchema]


class GenerateRequest(BaseModel):
    prompt:     str = Field(..., min_length=5)
    num_slides: int = Field(5, ge=3, le=10)


class GenerateResponse(BaseModel):
    status:  str
    job_id:  str
    file:    str
    plan:    PlanSchema


class UpdateRequest(BaseModel):
    plan: PlanSchema


class UpdateResponse(BaseModel):
    status: str
    file:   str


# Force Pydantic v2 to fully resolve all models now (avoids lazy-eval issues)
for _m in (SlideSchema, PlanSchema, GenerateRequest, GenerateResponse,
           UpdateRequest, UpdateResponse):
    _m.model_rebuild()


# ── Lifespan ──────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("Auto-PPT API starting …")
    yield
    log.info("Auto-PPT API shutting down.")


# ── FastAPI app ───────────────────────────────────────────────────────────────

app = FastAPI(
    title="Auto-PPT Generator API",
    version="1.0.0",
    description=(
        "FastAPI wrapper around the MCP-powered PPT agent. "
        "POST /generate to build a deck; GET /download to fetch it."
    ),
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Endpoints ─────────────────────────────────────────────────────────────────

@app.get("/health", tags=["meta"])
async def health():
    """Liveness probe – returns 200 when the server is ready."""
    return {"status": "ok", "server": "auto-ppt-api"}


@app.post("/generate", response_model=GenerateResponse, tags=["ppt"])
async def generate(req: GenerateRequest):
    """
    Run the full MCP agent pipeline:
      1. LLM generates a structured slide plan
      2. Agent calls create_presentation / add_slide / save_presentation via MCP
      3. Returns the plan + file path for the frontend to preview & download.
    """
    full_prompt = (
        f"{req.prompt.strip()}  "
        f"The presentation must have exactly {req.num_slides} slides "
        f"(including the title slide)."
    )
    try:
        plan = await _run_with_plan(full_prompt)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))
    except Exception as exc:
        log.exception("Unexpected error in /generate")
        raise HTTPException(status_code=500, detail=f"Unexpected error: {exc}")

    if not (HERE / OUTPUT_FILE).exists():
        raise HTTPException(status_code=500, detail="output.pptx was not created.")

    return GenerateResponse(
        status="success",
        job_id="sync",
        file=OUTPUT_FILE,
        plan=PlanSchema(**plan),
    )


@app.get("/download", tags=["ppt"])
async def download():
    """Stream the generated output.pptx to the client."""
    output_path = HERE / OUTPUT_FILE
    if not output_path.exists():
        raise HTTPException(
            status_code=404,
            detail="No presentation found. Run POST /generate first.",
        )
    return FileResponse(
        path=str(output_path),
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation"
        ),
        filename="presentation.pptx",
        headers={"Content-Disposition": 'attachment; filename="presentation.pptx"'},
    )


@app.post("/update", response_model=UpdateResponse, tags=["ppt"])
async def update_presentation(body: UpdateRequest):
    """
    Re-render output.pptx from an edited slide plan.
    Calls PPT MCP tools directly – no LLM round-trip needed.
    """
    try:
        slides_data = [s.model_dump() for s in body.plan.slides]
        await _rebuild_from_plan(slides_data)
    except Exception as exc:
        log.exception("Error in /update")
        raise HTTPException(status_code=500, detail=f"Re-render failed: {exc}")

    if not (HERE / OUTPUT_FILE).exists():
        raise HTTPException(status_code=500, detail="Re-render produced no file.")

    return UpdateResponse(status="updated", file=OUTPUT_FILE)


# ── Dev entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    uvicorn.run("server:app", host="0.0.0.0", port=8000, reload=True)