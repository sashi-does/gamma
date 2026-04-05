"""
agent_ppt.py  –  Auto-PPT Agent
================================================
Usage:
    export ANTHROPIC_API_KEY=sk-ant-...
    python agent_ppt.py "Create a 5-slide presentation on black holes for beginners"
    python agent_ppt.py   # interactive prompt

Architecture:
    1. Spin up both MCP servers as subprocesses (stdio transport)
    2. Load tools from each server via ClientSession
    3. PLAN: one LLM call → JSON slide plan
    4. EXEC: for each slide → call add_slide via MCP (one at a time)
    5. SAVE: call save_presentation via MCP
"""

import asyncio
import json
import logging
import os
import re
import sys
import urllib.request
import urllib.error
# import google.generativeai as genai
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openai import OpenAI

client = OpenAI(
    api_key=os.getenv("OPENROUTER_API_KEY"),
    base_url="https://openrouter.ai/api/v1"
)


# genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
# model = genai.GenerativeModel("gemini-1.5-flash")
# model = genai.GenerativeModel("gemini-1.5-flash-latest")
# model = genai.GenerativeModel("gemini-1.5-pro-latest")
# model = genai.GenerativeModel("gemini-pro")

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="[AGENT] %(asctime)s %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("agent_ppt")

# ── Constants ──────────────────────────────────────────────────────────────────
OUTPUT_FILE = "output.pptx"

SYSTEM_PROMPT = """
You are a Presentation Planning Agent that produces PowerPoint decks via tools.

STRICT WORKFLOW - follow these phases IN ORDER:

PHASE 1 - PLAN
  Output a JSON object (and NOTHING else) with this exact structure:
  {
    "presentation_title": "<overall deck title>",
    "slides": [
      { "title": "<slide title>", "bullet_points": ["<point 1>", "<point 2>", "<point 3>"] },
      ...
    ]
  }
  Rules:
  - First slide entry is the TITLE slide - bullet_points may be [].
  - Remaining slides are CONTENT slides - each must have 3-5 bullet_points.
  - Number of slides = exactly what the user requested (including the title slide).
  - Each bullet point is a complete, informative sentence.
  - Never skip this planning phase.

Important:
- Never generate all slides in a single step.
- If content is unclear, make reasonable assumptions - never crash or refuse.
""".strip()


# ── LLM helper (pure stdlib HTTP) ─────────────────────────────────────────────

def _read_key_file(path: Path) -> str | None:
    try:
        return path.read_text().strip() or None
    except Exception:
        return None

def _llm_call(messages, system="", max_tokens=2048):
    response = client.chat.completions.create(
        model="openai/gpt-4o-mini", 
        messages=[
            {"role": "system", "content": system},
            *messages
        ],
        max_tokens=max_tokens,
        temperature=0.7
    )
    return response

def _text(response):
    return response.choices[0].message.content


# ── MCP helpers ────────────────────────────────────────────────────────────────

async def get_tools(session: ClientSession) -> list[dict]:
    result = await session.list_tools()
    tools = [
        {"name": t.name, "description": t.description or "",
         "input_schema": getattr(t, "inputSchema", {})}
        for t in result.tools
    ]
    log.info("Loaded %d tools", len(tools))
    return tools


async def call_tool(session: ClientSession, name: str, args: dict) -> str:
    result = await session.call_tool(name, args)
    return "\n".join(
        b.text for b in result.content if hasattr(b, "text")
    ) or "(no output)"


# ── UI helpers ─────────────────────────────────────────────────────────────────

def banner(label: str, w: int = 62) -> None:
    print(f"\n{'─'*w}\n  {label}\n{'─'*w}")


def extract_json(text: str) -> str:
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if m:
        return m.group(1)
    m = re.search(r"\{.*\}", text, re.DOTALL)
    return m.group(0) if m else text


# ── Core agent ─────────────────────────────────────────────────────────────────

async def run_agent(user_request: str, fs_sess: ClientSession, ppt_sess: ClientSession) -> None:

    await get_tools(fs_sess)    # load (logs count); tools available via MCP
    await get_tools(ppt_sess)

    # ── PHASE 1: Plan ─────────────────────────────────────────────────────────
    banner("PHASE 1 – PLANNING")
    print(f"  Request: {user_request}\n")

    plan_raw = _text(_llm_call(
        messages=[{"role": "user", "content":
            f"{user_request}\n\nOutput ONLY the JSON plan. No other text."}],
        system=SYSTEM_PROMPT,
        max_tokens=2048,
    ))

    print("  Slide Plan (raw LLM output):")
    print(plan_raw)

    try:
        plan = json.loads(extract_json(plan_raw))
    except Exception as exc:
        log.error("JSON parse failed: %s", exc)
        print("[ERROR] Could not parse plan. Aborting.")
        return

    slides  = plan.get("slides", [])
    title   = plan.get("presentation_title", "Presentation")
    n       = len(slides)
    log.info("Plan OK: '%s' | %d slides", title, n)
    print(f"\n  Plan accepted: '{title}'  ({n} slides)")

    if not slides:
        print("[ERROR] No slides in plan. Aborting.")
        return

    # ── PHASE 2: Execute ──────────────────────────────────────────────────────
    banner("PHASE 2 – EXECUTION")

    # create_presentation via PPT MCP
    res = await call_tool(ppt_sess, "create_presentation", {"filename": OUTPUT_FILE})
    log.info("create_presentation -> %s", res)
    print(f"  Creating file: {OUTPUT_FILE}")
    print(f"  {res}")

    # add_slide one at a time
    for i, slide in enumerate(slides, 1):
        stitle  = slide.get("title", f"Slide {i}")
        bullets = slide.get("bullet_points", [])

        # Content slides need 3-5 bullets; generate if missing
        if i > 1 and len(bullets) < 3:
            log.info("Slide %d: generating bullets via LLM", i)
            print(f"\n  Generating bullets for slide {i}: '{stitle}' ...")
            gen = _text(_llm_call(
                messages=[{"role": "user", "content":
                    f'Give 4 bullet points for slide "{stitle}" '
                    f'in context: "{user_request}". '
                    f'Output ONLY a JSON array, e.g. ["...", "...", "...", "..."]'}],
                max_tokens=512,
            ))
            try:
                m = re.search(r"\[.*\]", gen, re.DOTALL)
                bullets = json.loads(m.group(0)) if m else bullets
            except Exception:
                bullets = [f"Key point about {stitle}.",
                           f"Why {stitle} matters.", f"Further detail on {stitle}."]

        print(f"\n  Slide {i}/{n}: '{stitle}'")
        for b in bullets:
            print(f"    • {b}")

        res = await call_tool(ppt_sess, "add_slide", {"title": stitle, "bullet_points": bullets})
        log.info("add_slide #%d -> %s", i, res)
        print(f"  {res}")

    # ── PHASE 3: Save ─────────────────────────────────────────────────────────
    banner("PHASE 3 – SAVING")
    res = await call_tool(ppt_sess, "save_presentation", {})
    log.info("save_presentation -> %s", res)
    print(f"\n  {res}")

    banner("COMPLETE")
    print(f"\n  File   : {OUTPUT_FILE}")
    print(f"  Slides : {n}")
    print(f"  Topic  : {title}\n")


# ── Entry ──────────────────────────────────────────────────────────────────────

async def main() -> None:
    if len(sys.argv) > 1:
        user_request = " ".join(sys.argv[1:])
    else:
        print("\nAuto-PPT Agent")
        user_request = input("Enter request: ").strip()
        if not user_request:
            user_request = "Create a 5-slide presentation on black holes for beginners"

    print(f"\nRequest: {user_request}\n")

    here = Path(__file__).parent.resolve()
    fs_params = StdioServerParameters(
        command="python",
        args=[str(here / "servers/filesystem_mcp_server.py")],
        env={**os.environ}
    )

    ppt_params = StdioServerParameters(
        command="python",
        args=[str(here / "servers/ppt_mcp_server.py")],
        env={**os.environ}
    )

    async with stdio_client(fs_params) as (r1, w1):
        async with ClientSession(r1, w1) as fs_sess:
            await fs_sess.initialize()
            log.info("Filesystem MCP ready")
            async with stdio_client(ppt_params) as (r2, w2):
                async with ClientSession(r2, w2) as ppt_sess:
                    await ppt_sess.initialize()
                    log.info("PPT MCP ready")
                    await run_agent(user_request, fs_sess, ppt_sess)


if __name__ == "__main__":
    asyncio.run(main())