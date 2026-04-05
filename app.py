"""
app.py  –  Auto-PPT Generator · Streamlit Frontend
====================================================
Full-featured UI for the MCP-powered PPT agent.

Features
--------
  • Prompt input + slide-count slider
  • Animated 3-step progress bar (Planning → Generating → Saving)
  • Slide preview cards after generation
  • Inline per-slide editing (title + bullet points)
  • "Update Presentation" re-renders pptx via /update endpoint
  • Download button streams output.pptx from /download

Run
---
    streamlit run app.py
"""

import time
from typing import Any

import httpx
import streamlit as st

# ── Config ────────────────────────────────────────────────────────────────────
API_BASE = "http://127.0.0.1:8000"
TIMEOUT  = 300   # seconds – agent can be slow on large decks

# ── Page setup ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Auto-PPT Generator",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Global background ── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0d0b1e 0%, #1a1040 50%, #0d0b1e 100%);
    min-height: 100vh;
}
[data-testid="stSidebar"] {
    background: rgba(255,255,255,0.035);
    border-right: 1px solid rgba(255,255,255,0.07);
}
[data-testid="stSidebar"] * { color: rgba(255,255,255,0.75) !important; }

/* ── Hero ── */
.hero { text-align: center; padding: 2.8rem 1rem 0.5rem; }
.hero h1 {
    font-size: 2.9rem; font-weight: 800; margin: 0;
    background: linear-gradient(100deg, #7eb8f7, #a78bfa, #7eb8f7);
    background-size: 200%;
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    animation: shimmer 4s infinite linear;
}
.hero p { color: rgba(255,255,255,0.45); font-size: 1rem; margin-top: 0.5rem; }
@keyframes shimmer { 0%{background-position:0%} 100%{background-position:200%} }

/* ── Divider ── */
.divider {
    height: 1px;
    background: linear-gradient(90deg,transparent,rgba(126,184,247,0.4),transparent);
    margin: 1.6rem 0;
}

/* ── Slide preview card ── */
.slide-card {
    background: rgba(255,255,255,0.055);
    border: 1px solid rgba(126,184,247,0.22);
    border-radius: 14px;
    padding: 1.3rem 1.5rem 1.1rem;
    margin-bottom: 1rem;
    transition: border-color .2s, transform .15s;
    height: 100%;
}
.slide-card:hover { border-color: rgba(126,184,247,0.55); transform: translateY(-2px); }
.slide-label {
    font-size: 0.68rem; font-weight: 700; letter-spacing: .13em;
    color: #7eb8f7; text-transform: uppercase; margin-bottom: .35rem;
}
.slide-title {
    font-size: 1.18rem; font-weight: 700; color: #fff; margin-bottom: .75rem;
    line-height: 1.3;
}
.bullet { color: rgba(255,255,255,0.72); font-size: .88rem; margin: .28rem 0; }
.bullet::before { content: "• "; color: #a78bfa; font-weight: bold; }

/* ── Progress steps ── */
.step {
    display:flex; align-items:center; gap:.7rem;
    padding:.5rem .95rem; border-radius:10px;
    margin:.3rem 0; font-size:.9rem; font-weight:500;
}
.step-done    { background:rgba(74,222,128,.09); color:#4ade80; }
.step-active  { background:rgba(126,184,247,.13); color:#7eb8f7;
                animation: pulse 1.3s infinite ease-in-out; }
.step-pending { background:rgba(255,255,255,.04); color:rgba(255,255,255,.3); }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.5} }

/* ── Status pill ── */
.pill {
    display:inline-block; padding:.22rem .8rem; border-radius:999px;
    font-size:.78rem; font-weight:700; letter-spacing:.05em;
}
.pill-ok  { background:#0d2a1a; color:#4ade80; border:1px solid #4ade80; }
.pill-err { background:#2a0d0d; color:#f87171; border:1px solid #f87171; }
.pill-off { background:#1a1a2a; color:#94a3b8; border:1px solid #94a3b8; }

/* ── Button tweaks ── */
div[data-testid="stButton"] > button,
div[data-testid="stDownloadButton"] > button {
    border-radius: 10px !important;
    font-weight: 650 !important;
    transition: opacity .15s !important;
}
div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg,#1a3050,#1e2761) !important;
    border: 1px solid #7eb8f7 !important;
    color: white !important;
    width: 100%;
}

/* ── Expander & text-area ── */
[data-testid="stExpander"] {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid rgba(255,255,255,0.09) !important;
    border-radius: 10px !important;
}
textarea {
    background: rgba(255,255,255,0.06) !important;
    border-color: rgba(255,255,255,0.15) !important;
    color: white !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)


# ── Session state defaults ────────────────────────────────────────────────────
_DEFAULTS = {
    "plan":        None,   # dict: presentation_title + slides
    "generated":   False,
    "pptx_bytes":  None,   # bytes for st.download_button
    "last_prompt": "",
    "num_slides":  5,
    "edit_mode":   False,
    "error":       None,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ── API helpers ───────────────────────────────────────────────────────────────

def api_health() -> bool:
    try:
        r = httpx.get(f"{API_BASE}/health", timeout=3)
        return r.status_code == 200
    except Exception:
        return False


def api_generate(prompt: str, num_slides: int) -> dict:
    with httpx.Client(timeout=TIMEOUT) as c:
        r = c.post(f"{API_BASE}/generate",
                   json={"prompt": prompt, "num_slides": num_slides})
    r.raise_for_status()
    return r.json()


def api_download() -> bytes:
    with httpx.Client(timeout=60) as c:
        r = c.get(f"{API_BASE}/download")
    r.raise_for_status()
    return r.content


def api_update(plan: dict) -> dict:
    with httpx.Client(timeout=TIMEOUT) as c:
        r = c.post(f"{API_BASE}/update", json={"plan": plan})
    r.raise_for_status()
    return r.json()


# ── UI sub-components ─────────────────────────────────────────────────────────

def render_progress(active_step: int):
    """Show 3-step progress. active_step ∈ {1, 2, 3}."""
    steps = [
        ("🧠", "Planning slides…"),
        ("⚙️",  "Generating presentation…"),
        ("💾",  "Saving file…"),
    ]
    html = ""
    for i, (icon, label) in enumerate(steps, 1):
        if i < active_step:
            cls, icon = "step-done", "✅"
        elif i == active_step:
            cls = "step-active"
        else:
            cls = "step-pending"
        html += f'<div class="step {cls}">{icon}&nbsp; {label}</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_slide_card(idx: int, slide: dict):
    label   = "Title Slide" if idx == 1 else f"Slide {idx}"
    title   = slide.get("title", "Untitled")
    bullets = slide.get("bullet_points", [])
    bullets_html = "".join(
        f'<div class="bullet">{b}</div>' for b in bullets
    ) or "<div class='bullet' style='opacity:.35'>No bullet points</div>"
    st.markdown(f"""
    <div class="slide-card">
        <div class="slide-label">{label}</div>
        <div class="slide-title">{title}</div>
        {bullets_html}
    </div>
    """, unsafe_allow_html=True)


def render_edit_form(slides: list) -> list:
    """Editable form for all slides; returns updated slides list."""
    edited = []
    for i, slide in enumerate(slides):
        label  = "🎯 Title Slide" if i == 0 else f"📄 Slide {i + 1}"
        with st.expander(label, expanded=(i == 0)):
            new_title = st.text_input(
                "Title",
                value=slide.get("title", ""),
                key=f"edit_title_{i}",
            )
            bullets     = slide.get("bullet_points", [])
            new_bullets = []
            for j, bp in enumerate(bullets):
                edited_bp = st.text_area(
                    f"Bullet {j + 1}",
                    value=bp,
                    height=64,
                    key=f"edit_bp_{i}_{j}",
                )
                new_bullets.append(edited_bp)

            # Allow adding a bullet (max 5, only on content slides)
            if i > 0 and len(bullets) < 5:
                if st.button("➕ Add bullet point", key=f"add_bp_{i}"):
                    new_bullets.append("New bullet point.")

            edited.append({"title": new_title, "bullet_points": new_bullets})
    return edited


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## ⚙️ Settings")

    online = api_health()
    badge  = '<span class="pill pill-ok">● API Online</span>'  if online \
        else '<span class="pill pill-off">● API Offline</span>'
    st.markdown(badge, unsafe_allow_html=True)

    if not online:
        st.warning(
            "Backend is not running. Start it with:\n"
            "```bash\nuvicorn server:app --reload\n```"
        )

    st.markdown("---")
    st.markdown("""
### How it works
1. **Type** your topic and pick a slide count
2. **Generate** → the MCP agent plans and builds the deck
3. **Preview** all slides in the card view
4. **Edit** titles and bullets inline
5. **Update** to re-render the PPTX
6. **Download** your finished presentation
""")
    st.markdown("---")
    st.caption("FastAPI · FastMCP · python-pptx · Claude")


# ── Hero ──────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
  <h1>🎯 Auto-PPT Generator</h1>
  <p>Describe your topic — the AI agent plans, builds, and delivers your presentation.</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)


# ── Input panel ───────────────────────────────────────────────────────────────

col_prompt, col_opts = st.columns([3, 1], gap="large")

with col_prompt:
    prompt = st.text_area(
        "💬 Describe your presentation",
        placeholder=(
            "e.g.  Create a presentation on climate change for high-school students, "
            "covering causes, effects, and solutions."
        ),
        height=115,
        value=st.session_state.last_prompt,
    )

with col_opts:
    num_slides = st.slider(
        "🗂 Slide count",
        min_value=3, max_value=10,
        value=st.session_state.num_slides,
        help="Includes the title slide",
    )
    st.session_state.num_slides = num_slides
    st.markdown("<br>", unsafe_allow_html=True)
    generate_btn = st.button(
        "⚡ Generate PPT",
        use_container_width=True,
        type="primary",
        disabled=(not prompt.strip()) or (not online),
    )


# ── Generation flow ───────────────────────────────────────────────────────────

if generate_btn and prompt.strip():
    # Reset state for a fresh run
    st.session_state.last_prompt = prompt
    st.session_state.generated   = False
    st.session_state.plan        = None
    st.session_state.pptx_bytes  = None
    st.session_state.edit_mode   = False
    st.session_state.error       = None

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    progress_box = st.empty()

    try:
        # Step 1 – Planning (show while request is in flight)
        with progress_box.container():
            render_progress(1)

        result = api_generate(prompt.strip(), num_slides)

        # Step 2 – Generating
        with progress_box.container():
            render_progress(2)
        time.sleep(0.25)

        # Step 3 – Saving
        with progress_box.container():
            render_progress(3)
        time.sleep(0.25)

        # Fetch file bytes for the download button
        pptx_bytes = api_download()

        st.session_state.plan       = result.get("plan")
        st.session_state.pptx_bytes = pptx_bytes
        st.session_state.generated  = True

    except httpx.HTTPStatusError as exc:
        try:
            detail = exc.response.json().get("detail", str(exc))
        except Exception:
            detail = str(exc)
        st.session_state.error = f"API error {exc.response.status_code}: {detail}"
    except httpx.ConnectError:
        st.session_state.error = (
            "Cannot reach the backend at "
            f"{API_BASE}. "
            "Is `uvicorn server:app --reload` running?"
        )
    except Exception as exc:
        st.session_state.error = str(exc)

    progress_box.empty()


# ── Error banner ──────────────────────────────────────────────────────────────

if st.session_state.error:
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.error(f"❌  {st.session_state.error}")


# ── Results panel ─────────────────────────────────────────────────────────────

if st.session_state.generated and st.session_state.plan:
    plan: dict   = st.session_state.plan
    slides: list = plan.get("slides", [])
    prs_title    = plan.get("presentation_title", "Presentation")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ── Status bar ────────────────────────────────────────────────────────────
    st.markdown(
        f'<span class="pill pill-ok">✅ Generated</span>'
        f'&nbsp;&nbsp;<span style="color:rgba(255,255,255,.65);font-size:.95rem">'
        f'<strong style="color:#fff">{prs_title}</strong>'
        f' &nbsp;·&nbsp; {len(slides)} slides</span>',
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Action bar ────────────────────────────────────────────────────────────
    dl_col, edit_col, spacer = st.columns([2, 2, 4])

    with dl_col:
        st.download_button(
            label="📥  Download PPTX",
            data=st.session_state.pptx_bytes,
            file_name="presentation.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".presentationml.presentation"
            ),
            use_container_width=True,
        )

    with edit_col:
        toggle_label = "✏️ Edit Slides" if not st.session_state.edit_mode \
            else "👁 Preview Mode"
        if st.button(toggle_label, use_container_width=True):
            st.session_state.edit_mode = not st.session_state.edit_mode
            st.rerun()

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ── Preview mode ──────────────────────────────────────────────────────────
    if not st.session_state.edit_mode:
        st.markdown("### 📋 Slide Preview")

        # Two-column grid
        left_col, right_col = st.columns(2, gap="medium")
        for i, slide in enumerate(slides):
            with (left_col if i % 2 == 0 else right_col):
                render_slide_card(i + 1, slide)

    # ── Edit mode ─────────────────────────────────────────────────────────────
    else:
        st.markdown("### ✏️ Edit Your Slides")
        st.caption(
            "Modify any title or bullet point below, then click "
            "**Update Presentation** to regenerate the PPTX."
        )

        edited_slides = render_edit_form(slides)

        st.markdown("<br>", unsafe_allow_html=True)
        upd_col, _ = st.columns([2, 5])
        with upd_col:
            update_btn = st.button(
                "🔄 Update Presentation",
                type="primary",
                use_container_width=True,
            )

        if update_btn:
            updated_plan = {
                "presentation_title": prs_title,
                "slides": edited_slides,
            }
            with st.spinner("Re-rendering presentation…"):
                try:
                    api_update(updated_plan)
                    new_bytes = api_download()
                    # Persist edits into session state
                    st.session_state.plan["slides"] = edited_slides
                    st.session_state.pptx_bytes     = new_bytes
                    st.success(
                        "✅ Presentation updated! "
                        "Use the **Download PPTX** button to get the new file."
                    )
                except httpx.HTTPStatusError as exc:
                    try:
                        detail = exc.response.json().get("detail", str(exc))
                    except Exception:
                        detail = str(exc)
                    st.error(f"Update failed ({exc.response.status_code}): {detail}")
                except Exception as exc:
                    st.error(f"Update failed: {exc}")


# ── Empty state ───────────────────────────────────────────────────────────────

elif not st.session_state.error:
    st.markdown("""
    <div style="text-align:center;padding:3.5rem 0;color:rgba(255,255,255,.25);">
        <div style="font-size:4.5rem">🎯</div>
        <div style="font-size:1.05rem;margin-top:1rem;color:rgba(255,255,255,.4);">
            Enter a topic above and click
            <strong style="color:rgba(255,255,255,.55)">Generate PPT</strong>
            to get started
        </div>
        <div style="font-size:.82rem;margin-top:.45rem;">
            The AI agent will plan your slides, build the content,
            and assemble the PPTX — automatically.
        </div>
    </div>
    """, unsafe_allow_html=True)