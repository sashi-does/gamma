"""
app.py  –  SlideMind · AI Presentation Builder
===============================================
Sidebar: minimal left (collapse, theme, how-to-use)
Collapsed: logo + new-chat + help-popup + expand icon, API status centered above logo
Main: slide settings moved near textarea

Run
---
    streamlit run app.py
"""

import time
import httpx
import streamlit as st

# ── Config ────────────────────────────────────────────────────────────────────
API_BASE = "http://127.0.0.1:8000"
TIMEOUT  = 300

# ── Page setup ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SlideMind · AI Presentation Builder",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state defaults ────────────────────────────────────────────────────
_DEFAULTS = {
    "plan":              None,
    "generated":         False,
    "pptx_bytes":        None,
    "last_prompt":       "",
    "num_slides":        6,
    "error":             None,
    "theme":             "dark",
    "chat_history":      [],
    "active_tab":        "chat",
    "sidebar_collapsed": False,
    "prompt_value":      "",
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

IS_DARK      = st.session_state.theme == "dark"
SB_COLLAPSED = st.session_state.sidebar_collapsed

# ── Theme tokens ──────────────────────────────────────────────────────────────
DARK = {
    "bg":"#07070f","bg2":"#0d0d1a","bg3":"#131325","bg4":"#1a1a2e",
    "bd":"rgba(255,255,255,0.06)","bdh":"rgba(255,255,255,0.14)",
    "tx":"#e2e2ee","mu":"rgba(226,226,238,0.36)",
    "ac":"#7c6af7","ac2":"#a78bfa","glow":"rgba(124,106,247,0.16)",
    "ok":"#34d399","er":"#f87171","wa":"#fbbf24",
    "ub":"#12122a","bb":"#0d0d1a","inp":"#0d0d1a","sh":"rgba(0,0,0,0.6)",
    "sb":"#0a0a18","sb2":"#0f0f20","sb_bd":"rgba(255,255,255,0.05)",
}
LIGHT = {
    "bg":"#f4f4f8","bg2":"#ffffff","bg3":"#eeeef4","bg4":"#e4e4ee",
    "bd":"rgba(0,0,0,0.07)","bdh":"rgba(0,0,0,0.16)",
    "tx":"#0c0c18","mu":"rgba(12,12,24,0.38)",
    "ac":"#5b4de8","ac2":"#7c6af7","glow":"rgba(91,77,232,0.11)",
    "ok":"#059669","er":"#dc2626","wa":"#d97706",
    "ub":"#eceaff","bb":"#f4f4f8","inp":"#ffffff","sh":"rgba(0,0,0,0.06)",
    "sb":"#f0f0f7","sb2":"#ffffff","sb_bd":"rgba(0,0,0,0.06)",
}
T = DARK if IS_DARK else LIGHT

SB_WIDTH_EXP  = 200
SB_WIDTH_COL  = 56

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=Instrument+Sans:ital,wght@0,400;0,500;0,600;1,400&family=Fira+Code:wght@400;500&display=swap');

:root {{
  --bg:{T['bg']};--bg2:{T['bg2']};--bg3:{T['bg3']};--bg4:{T['bg4']};
  --bd:{T['bd']};--bdh:{T['bdh']};
  --tx:{T['tx']};--mu:{T['mu']};
  --ac:{T['ac']};--ac2:{T['ac2']};--glow:{T['glow']};
  --ok:{T['ok']};--er:{T['er']};--wa:{T['wa']};
  --ub:{T['ub']};--bb:{T['bb']};--inp:{T['inp']};--sh:{T['sh']};
  --sb:{T['sb']};--sb2:{T['sb2']};--sb-bd:{T['sb_bd']};
  --fd:'Syne',sans-serif;--fb:'Instrument Sans',sans-serif;--fm:'Fira Code',monospace;
  --sb-w:{SB_WIDTH_EXP}px;--sb-col:{SB_WIDTH_COL}px;
}}

*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
html,body,[data-testid="stAppViewContainer"]{{
  font-family:var(--fb)!important;background:var(--bg)!important;color:var(--tx)!important;
}}
[data-testid="stAppViewContainer"]>section>div{{background:transparent!important;}}
#MainMenu,footer,header{{visibility:hidden!important;}}
div.block-container, [data-testid="stMainBlockContainer"] {{padding-top:1.2rem!important;}}
[data-testid="stDeployButton"]{{display:none!important;}}
[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarCollapsedControl"],
button[data-testid="baseButton-headerNoPadding"]{{display:none!important;}}

::-webkit-scrollbar{{width:3px;height:3px;}}
::-webkit-scrollbar-track{{background:transparent;}}
::-webkit-scrollbar-thumb{{background:var(--bdh);border-radius:99px;}}

/* ════════════════════════════════════════
   SIDEBAR — BASE
════════════════════════════════════════ */
[data-testid="stSidebar"]{{
  background:var(--sb)!important;
  border-right:1px solid var(--sb-bd)!important;
  transition:min-width 0.2s cubic-bezier(0.4,0,0.2,1),
             max-width 0.2s cubic-bezier(0.4,0,0.2,1)!important;
}}
[data-testid="stSidebar"]>div:first-child, [data-testid="stSidebarUserContent"] {{
  padding-top: 0 !important;
  padding: 0 !important;
  overflow: hidden !important;
}}
[data-testid="stSidebar"] *{{color:var(--tx)!important;}}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{{gap:0!important;}}

/* ── Expanded sidebar ── */
[data-testid="stSidebar"]:not([class*="collapsed"]){{
  min-width:{SB_WIDTH_EXP}px!important;
  max-width:{SB_WIDTH_EXP}px!important;
}}

/* ── Collapsed sidebar ── */
[data-testid="stSidebar"].sb-collapsed,
[data-testid="stSidebar"].sb-collapsed > div:first-child {{
  min-width:{SB_WIDTH_COL}px!important;
  max-width:{SB_WIDTH_COL}px!important;
  width:{SB_WIDTH_COL}px!important;
  overflow:hidden!important;
}}

/* ════════════════════════════════════════
   SIDEBAR EXPANDED — minimal left panel
════════════════════════════════════════ */
.sb-wrap{{
  display:flex;flex-direction:column;height:100vh;
  width:{SB_WIDTH_EXP}px;overflow:hidden;
}}

/* API status bar at very top, centered */
.sb-status-top{{
  display:flex;align-items:center;justify-content:center;
  padding:0.55rem 0.75rem 0.55rem;
  border-bottom:1px solid var(--sb-bd);
  flex-shrink:0;
}}
.sb-pill{{
  display:inline-flex;align-items:center;gap:0.38rem;
  padding:0.22rem 0.7rem;border-radius:99px;
  font-family:var(--fm);font-size:0.62rem;font-weight:500;letter-spacing:0.04em;
}}
.sb-pill.on{{
  background:rgba(52,211,153,0.09);color:var(--ok);
  border:1px solid rgba(52,211,153,0.18);
}}
.sb-pill.off{{
  background:var(--bg3);color:var(--mu);border:1px solid var(--bd);
}}
.pill-dot{{width:5px;height:5px;border-radius:50%;flex-shrink:0;}}
.sb-pill.on .pill-dot{{background:var(--ok);animation:blink 2s infinite;}}
.sb-pill.off .pill-dot{{background:var(--mu);}}
@keyframes blink{{0%,100%{{opacity:1;}}50%{{opacity:0.35;}}}}

/* Logo header */
.sb-header{{
  display:flex;align-items:center;justify-content:space-between;
  padding:0.8rem 0.75rem 0.65rem;
  border-bottom:1px solid var(--sb-bd);
  flex-shrink:0;
}}
.sb-brand{{display:flex;align-items:center;gap:0.45rem;}}
.sb-gem{{
  width:26px;height:26px;border-radius:7px;flex-shrink:0;
  background:linear-gradient(135deg,var(--ac) 0%,var(--ac2) 100%);
  display:flex;align-items:center;justify-content:center;
  font-size:0.88rem;color:#fff;line-height:1;
  box-shadow:0 2px 10px rgba(124,106,247,0.35);
}}
.sb-name{{
  font-family:var(--fd);font-weight:700;font-size:0.88rem;letter-spacing:-0.02em;
}}
.sb-name em{{color:var(--ac);font-style:normal;}}

.sb-collapse-btn-wrap{{
  display:flex;align-items:center;
}}

/* Offline warning */
.sb-warn{{
  margin:0.45rem 0.75rem 0;
  background:rgba(251,191,36,.05);
  border:1px solid rgba(251,191,36,.16);
  border-radius:8px;padding:0.5rem 0.65rem;
  font-size:0.72rem;color:var(--wa);line-height:1.6;
}}
.sb-warn code{{
  font-family:var(--fm);font-size:0.65rem;
  background:rgba(251,191,36,.08);
  padding:0.1rem 0.3rem;border-radius:3px;
}}

/* Section label */
.sb-body{{flex:1;overflow-y:auto;padding:0 0.75rem 1rem;}}
.sb-sect{{
  font-family:var(--fm);font-size:0.57rem;font-weight:500;
  letter-spacing:0.12em;text-transform:uppercase;
  color:var(--mu)!important;display:block;
  margin:0.9rem 0 0.4rem;
}}
.sb-hr{{height:1px;background:var(--sb-bd);margin:0.55rem 0;}}

/* How to use box */
.how-to-use{{
  position:fixed;bottom:3rem;left:0.75rem;width:calc(var(--sb-w) - 1.5rem);
  background:var(--bg2);border:1px solid var(--bd);border-radius:10px;
  padding:0.75rem 0.85rem;margin-top:0.5rem;z-index:100;
}}
.how-to-use-title{{
  font-family:var(--fm);font-size:0.6rem;text-transform:uppercase;
  letter-spacing:0.1em;color:var(--mu);margin-bottom:0.55rem;
}}
.how-step{{
  display:flex;align-items:flex-start;gap:0.5rem;
  font-size:0.75rem;color:var(--mu);line-height:1.55;
  font-family:var(--fb);padding:0.1rem 0;
}}
.how-num{{
  width:16px;height:16px;border-radius:4px;
  background:var(--glow);border:1px solid rgba(124,106,247,0.2);
  display:flex;align-items:center;justify-content:center;
  font-family:var(--fm);font-size:0.55rem;font-weight:600;
  color:var(--ac);flex-shrink:0;margin-top:1px;
}}

/* Footer */
.sb-foot{{
  position:fixed;bottom:0;left:0;width:var(--sb-w);
  padding:0.65rem 0.75rem 0.85rem;border-top:1px solid var(--sb-bd);
  font-family:var(--fm);font-size:0.6rem;color:var(--mu);
  text-align:center;flex-shrink:0;letter-spacing:0.04em;
  background:var(--sb);z-index:100;
}}

/* ════════════════════════════════════════
   SIDEBAR COLLAPSED — minimal icon rail
════════════════════════════════════════ */
.rail{{
  display:flex;flex-direction:column;align-items:center;
  width:{SB_WIDTH_COL}px;height:100vh;
  padding:0;
  overflow:hidden;
}}

/* API status dot at very top center */
.rail-status-top{{
  display:flex;align-items:center;justify-content:center;
  width:100%;padding:0.55rem 0 0.55rem;
  border-bottom:1px solid var(--sb-bd);
  flex-shrink:0;
}}
.rail-status{{
  width:7px;height:7px;border-radius:50%;flex-shrink:0;
}}
.rail-status.on{{background:var(--ok);box-shadow:0 0 6px rgba(52,211,153,0.5);animation:blink 2s infinite;}}
.rail-status.off{{background:var(--mu);}}

/* Logo gem */
.rail-gem{{
  width:30px;height:30px;border-radius:8px;
  background:linear-gradient(135deg,var(--ac) 0%,var(--ac2) 100%);
  display:flex;align-items:center;justify-content:center;
  font-size:0.95rem;color:#fff;line-height:1;
  box-shadow:0 2px 10px rgba(124,106,247,0.3);
  margin:0.6rem 0 0.3rem;flex-shrink:0;
}}

.rail-divider{{
  width:28px;height:1px;
  background:var(--sb-bd);
  margin:0.35rem 0;flex-shrink:0;
}}

/* Icon buttons */
.rail-btn{{
  width:34px;height:34px;border-radius:8px;
  display:flex;align-items:center;justify-content:center;
  cursor:pointer;transition:background 0.15s,color 0.15s;
  color:var(--mu);margin:0.1rem 0;flex-shrink:0;
  border:1px solid transparent;
  position:relative;
}}
.rail-btn:hover{{
  background:rgba(255,255,255,0.07);
  color:var(--tx);
  border-color:var(--bd);
}}
.rail-btn svg{{display:block;}}

/* Tooltip for collapsed help */
.rail-btn .rail-tooltip{{
  position:absolute;left:calc(100% + 10px);top:50%;transform:translateY(-50%);
  background:var(--bg2);border:1px solid var(--bd);border-radius:10px;
  padding:0.6rem 0.75rem;width:180px;z-index:9999;
  box-shadow:0 8px 24px var(--sh);pointer-events:none;
  opacity:0;transition:opacity 0.15s;
  font-family:var(--fb);
}}
.rail-btn:hover .rail-tooltip{{opacity:1;}}
.rail-tooltip-title{{
  font-family:var(--fm);font-size:0.57rem;text-transform:uppercase;
  letter-spacing:0.1em;color:var(--mu);margin-bottom:0.4rem;
}}
.rail-tooltip-step{{
  display:flex;align-items:flex-start;gap:0.35rem;
  font-size:0.72rem;color:var(--mu);line-height:1.5;padding:0.05rem 0;
}}
.rail-tooltip-num{{
  width:14px;height:14px;border-radius:3px;
  background:var(--glow);border:1px solid rgba(124,106,247,0.2);
  display:flex;align-items:center;justify-content:center;
  font-family:var(--fm);font-size:0.52rem;font-weight:600;
  color:var(--ac);flex-shrink:0;margin-top:1px;
}}

/* Expand button at bottom */
.rail-expand{{
  position:fixed;bottom:1.5rem;left:11px;
  width:34px;height:28px;border-radius:7px;
  display:flex;align-items:center;justify-content:center;
  color:var(--mu);cursor:pointer;
  border:1px solid var(--sb-bd);
  background:var(--sb);transition:all 0.15s;z-index:99;
}}
.rail-expand:hover{{
  background:rgba(255,255,255,0.07);
  color:var(--tx);
}}

/* ════════════════════════════════════════
   STREAMLIT COMPONENT OVERRIDES
════════════════════════════════════════ */
*{{outline-color:var(--ac)!important;}}
button:focus,input:focus,textarea:focus,select:focus,[role="slider"]:focus{{
  outline:none!important;box-shadow:0 0 0 3px var(--glow)!important;
}}

div[data-testid="stButton"]>button{{
  border-radius:8px!important;font-family:var(--fb)!important;font-weight:500!important;
  font-size:0.84rem!important;transition:all 0.14s!important;
  border:1px solid var(--bd)!important;background:var(--bg2)!important;color:var(--tx)!important;
}}
div[data-testid="stButton"]>button:hover{{
  border-color:var(--bdh)!important;background:var(--bg3)!important;
}}
div[data-testid="stButton"]>button:focus{{
  box-shadow:0 0 0 3px var(--glow)!important;border-color:var(--ac)!important;outline:none!important;
}}
div[data-testid="stButton"]>button[kind="primary"]{{
  background:var(--ac)!important;color:#fff!important;border-color:transparent!important;
  font-family:var(--fd)!important;font-weight:600!important;
}}
div[data-testid="stButton"]>button[kind="primary"]:hover{{opacity:0.87!important;}}

div[data-testid="stDownloadButton"]>button{{
  border-radius:8px!important;font-family:var(--fb)!important;font-weight:500!important;
  background:var(--bg2)!important;color:var(--tx)!important;border:1px solid var(--bd)!important;width:100%!important;
}}

div[data-testid="stTextArea"] textarea{{
  background:var(--inp)!important;border:1px solid var(--bd)!important;
  color:var(--tx)!important;font-family:var(--fb)!important;font-size:0.87rem!important;
  border-radius:10px!important;padding:0.6rem 0.75rem!important;resize:none!important;
  transition:border-color 0.18s,box-shadow 0.18s!important;caret-color:var(--ac)!important;
}}
div[data-testid="stTextArea"] textarea:focus{{
  border-color:rgba(124,106,247,.5)!important;box-shadow:0 0 0 3px var(--glow)!important;outline:none!important;
}}
div[data-testid="stTextInput"] input{{
  background:var(--inp)!important;border:1px solid var(--bd)!important;
  color:var(--tx)!important;border-radius:8px!important;font-family:var(--fb)!important;
}}
div[data-testid="stTextInput"] input:focus{{
  border-color:rgba(124,106,247,.5)!important;box-shadow:0 0 0 3px var(--glow)!important;outline:none!important;
}}
textarea{{background:var(--bg3)!important;border:1px solid var(--bd)!important;color:var(--tx)!important;border-radius:8px!important;font-family:var(--fb)!important;}}
textarea:focus{{border-color:rgba(124,106,247,.5)!important;box-shadow:0 0 0 3px var(--glow)!important;outline:none!important;}}

div[data-testid="stSelectbox"]>div>div{{
  background:var(--inp)!important;border:1px solid var(--bd)!important;
  color:var(--tx)!important;border-radius:8px!important;
}}
div[data-testid="stSlider"] span{{color:var(--tx)!important;font-family:var(--fb)!important;}}
div[data-testid="stSlider"] div[data-baseweb="slider"] div[role="slider"]{{background:var(--ac)!important;border-color:var(--ac)!important;}}
div[data-testid="stSlider"] div[data-baseweb="slider"] div[role="slider"]:focus{{box-shadow:0 0 0 4px var(--glow)!important;}}
div[data-testid="stSlider"] div[data-baseweb="slider"] div[role="slider"] div {{color:var(--ac)!important;font-family:var(--fm)!important;}}
div[data-testid="stSlider"] div[data-baseweb="slider"] > div > div > div:first-child{{background:var(--ac)!important;}}

[data-testid="stExpander"]{{
  background:var(--bg2)!important;border:1px solid var(--bd)!important;
  border-radius:10px!important;overflow:hidden!important;
}}
[data-testid="stExpander"] summary{{
  color:var(--tx)!important;font-weight:500!important;font-family:var(--fb)!important;
}}

/* Hidden rail buttons */
div[data-testid="stButton"][data-key="sb_expand_btn"],
div[data-testid="stButton"][data-key="new_chat_btn_col"] {{
  position: absolute !important; opacity: 0 !important; pointer-events: none !important;
}}

/* Collapse button inside expanded sidebar */
div[data-testid="stButton"][data-key="sb_collapse_btn"] > button {{
  width:26px!important;min-height:22px!important;height:22px!important;
  padding:0!important;border-radius:6px!important;
  border-color:var(--sb-bd)!important;background-color:transparent!important;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='rgba(226,226,238,0.5)' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Crect x='3' y='3' width='18' height='18' rx='2' ry='2'%3E%3C/rect%3E%3Cline x1='9' y1='3' x2='9' y2='21'%3E%3C/line%3E%3C/svg%3E") !important;
  background-size: 14px 14px !important;
  background-position: center !important;
  background-repeat: no-repeat !important;
  color: transparent !important;
}}
div[data-testid="stButton"][data-key="sb_collapse_btn"] > button * {{
  display: none !important;
}}
div[data-testid="stButton"][data-key="sb_collapse_btn"] > button:hover {{
  background-color:rgba(255,255,255,0.07)!important;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23e2e2ee' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Crect x='3' y='3' width='18' height='18' rx='2' ry='2'%3E%3C/rect%3E%3Cline x1='9' y1='3' x2='9' y2='21'%3E%3C/line%3E%3C/svg%3E") !important;
}}

/* New chat button in sidebar */
div[data-testid="stButton"][data-key="new_chat_btn"] > button {{
  width:26px!important;min-height:22px!important;height:22px!important;
  padding:0!important;font-size:0.82rem!important;line-height:1!important;
  border-radius:6px!important;border-color:var(--sb-bd)!important;
  background:transparent!important;color:var(--mu)!important;
}}
div[data-testid="stButton"][data-key="new_chat_btn"] > button:hover {{
  background:rgba(255,255,255,0.07)!important;color:var(--tx)!important;
}}

/* Send button */
div[data-testid="stButton"][data-key="send_btn"] > button {{
  height:52px!important;font-size:1.3rem!important;
  padding:0!important;line-height:1!important;
}}

/* Chip buttons */
div[data-testid="stButton"] > button[kind="secondary"]{{
  border-radius:99px!important;font-size:0.76rem!important;
  padding:0.28rem 0.5rem!important;color:var(--mu)!important;
  background:var(--bg2)!important;border:1px solid var(--bd)!important;
}}
div[data-testid="stButton"] > button[kind="secondary"]:hover{{
  border-color:var(--ac)!important;color:var(--ac)!important;background:var(--glow)!important;
}}

/* ════════════════════════════════════════
   SETTINGS ROW near input
════════════════════════════════════════ */
.settings-row{{
  display:flex;align-items:center;gap:1rem;
  padding:0.45rem 0.1rem 0.3rem;
  flex-wrap:wrap;
}}
.settings-label{{
  font-family:var(--fm);font-size:0.6rem;text-transform:uppercase;
  letter-spacing:0.1em;color:var(--mu);margin-bottom:0.15rem;
}}
.slide-count-note{{font-size:0.72rem;color:var(--mu);margin-top:0.1rem;font-family:var(--fm);}}

/* ════════════════════════════════════════
   MAIN CONTENT STYLES
════════════════════════════════════════ */
.nav-bar{{display:flex;align-items:center;justify-content:space-between;padding:0.85rem 0 0.75rem;border-bottom:1px solid var(--bd);margin-bottom:1.5rem;}}
.nav-wm{{font-family:var(--fd);font-weight:800;font-size:1.05rem;letter-spacing:-0.03em;}}
.nav-wm em{{color:var(--ac);font-style:normal;}}
.nav-tags{{display:flex;gap:0.45rem;}}
.nav-tag{{font-family:var(--fm);font-size:0.59rem;letter-spacing:0.1em;text-transform:uppercase;
  padding:0.17rem 0.55rem;border-radius:4px;border:1px solid var(--bd);color:var(--mu);background:var(--bg2);}}

/* Chat */
.msg-row{{display:flex;gap:0.75rem;margin-bottom:1.25rem;animation:sUp 0.28s cubic-bezier(0.4,0,0.2,1) both;}}
.msg-row.user{{flex-direction:row-reverse;}}
@keyframes sUp{{from{{opacity:0;transform:translateY(8px);}}to{{opacity:1;transform:translateY(0);}}}}
.av{{width:28px;height:28px;border-radius:7px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-family:var(--fd);font-weight:700;font-size:0.85rem;}}
.av.bot{{background:linear-gradient(135deg,var(--ac),var(--ac2));color:white;}}
.av.usr{{background:var(--bg4);border:1px solid var(--bdh);color:var(--mu);}}
.bub{{padding:0.72rem 0.95rem;border-radius:12px;font-size:0.875rem;line-height:1.7;max-width:80%;border:1px solid var(--bd);font-family:var(--fb);}}
.bub.bot{{background:var(--bb);border-radius:2px 12px 12px 12px;}}
.bub.usr{{background:var(--ub);border-radius:12px 2px 12px 12px;}}
.msg-lbl{{font-family:var(--fm);font-size:0.59rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.28rem;opacity:0.35;}}

/* Progress */
.prog-card{{background:var(--bg2);border:1px solid var(--bd);border-radius:14px;padding:1.2rem 1.4rem;max-width:370px;margin:0.5rem auto;}}
.prog-head{{font-family:var(--fm);font-size:0.59rem;letter-spacing:0.12em;text-transform:uppercase;color:var(--mu);margin-bottom:0.9rem;}}
.prog-step{{display:flex;align-items:center;gap:0.65rem;padding:0.38rem 0;font-size:0.84rem;font-family:var(--fb);color:var(--mu);transition:color 0.25s;}}
.prog-step.active{{color:var(--tx);}}
.prog-step.done{{color:var(--ok);}}
.stic{{width:24px;height:24px;border-radius:7px;flex-shrink:0;display:flex;align-items:center;justify-content:center;transition:all 0.25s;}}
.stic svg{{width:13px;height:13px;}}
.stic.pending{{background:var(--bg4);border:1px solid var(--bd);}}
.stic.active{{background:var(--glow);border:1px solid rgba(124,106,247,.3);animation:spulse 1.6s infinite ease-in-out;}}
.stic.done{{background:rgba(52,211,153,.1);border:1px solid rgba(52,211,153,.25);}}
@keyframes spulse{{0%,100%{{box-shadow:0 0 0 0 var(--glow);}}50%{{box-shadow:0 0 0 5px transparent;}}}}
.prog-track{{height:2px;background:var(--bg4);border-radius:99px;margin-top:0.9rem;overflow:hidden;}}
.prog-fill{{height:100%;background:linear-gradient(90deg,var(--ac),var(--ac2));border-radius:99px;transition:width 0.5s ease;}}

/* Slide card */
.slide-card{{background:var(--bg2);border:1px solid var(--bd);border-radius:13px;padding:1.15rem 1.35rem;transition:border-color 0.18s,transform 0.16s,box-shadow 0.18s;position:relative;overflow:hidden;animation:sUp 0.32s ease both;}}
.slide-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--ac),var(--ac2));opacity:0;transition:opacity 0.18s;}}
.slide-card:hover{{border-color:var(--bdh);transform:translateY(-2px);box-shadow:0 10px 28px var(--sh);}}
.slide-card:hover::before{{opacity:1;}}
.snum{{font-family:var(--fm);font-size:0.59rem;font-weight:500;letter-spacing:0.12em;text-transform:uppercase;color:var(--ac);margin-bottom:0.42rem;}}
.stitle{{font-family:var(--fd);font-size:0.95rem;font-weight:600;color:var(--tx);margin-bottom:0.7rem;line-height:1.35;letter-spacing:-0.01em;}}
.bullet{{font-size:0.81rem;color:var(--mu);line-height:1.65;padding:0.1rem 0 0.1rem 1rem;position:relative;font-family:var(--fb);}}
.bullet::before{{content:'';position:absolute;left:0;top:0.65em;width:3px;height:3px;border-radius:50%;background:var(--ac2);}}

/* Empty state */
.empty{{text-align:center;padding:4rem 2rem 3rem;}}
.empty-mark{{width:48px;height:48px;border-radius:13px;background:var(--glow);border:1px solid rgba(124,106,247,.2);display:flex;align-items:center;justify-content:center;margin:0 auto 1rem;font-family:var(--fd);font-weight:800;font-size:1.45rem;color:var(--ac);line-height:1;}}
.empty-title{{font-family:var(--fd);font-size:1.02rem;font-weight:600;color:var(--tx);margin-bottom:0.42rem;letter-spacing:-0.02em;}}
.empty-sub{{font-size:0.83rem;color:var(--mu);max-width:320px;margin:0 auto;line-height:1.7;}}

/* Toast */
.toast{{display:inline-flex;align-items:center;gap:0.4rem;padding:0.4rem 0.82rem;border-radius:8px;font-size:0.79rem;font-weight:500;font-family:var(--fb);}}
.toast.info{{background:var(--glow);color:var(--ac);border:1px solid rgba(124,106,247,.25);}}
.toast.success{{background:rgba(52,211,153,.08);color:var(--ok);border:1px solid rgba(52,211,153,.2);}}

/* Stats */
.stat-row{{display:flex;gap:0.75rem;margin-bottom:1rem;}}
.stat-card{{flex:1;background:var(--bg2);border:1px solid var(--bd);border-radius:12px;padding:0.85rem 1rem;text-align:center;}}
.stat-val{{font-family:var(--fd);font-size:1.7rem;font-weight:800;color:var(--tx);line-height:1;letter-spacing:-0.04em;}}
.stat-lbl{{font-family:var(--fm);font-size:0.61rem;color:var(--mu);margin-top:0.28rem;text-transform:uppercase;letter-spacing:0.08em;}}
.div{{height:1px;background:var(--bd);margin:0.8rem 0;}}

@media(max-width:768px){{.stat-row{{flex-direction:column;}}.bub{{max-width:90%;}}}}
</style>
""", unsafe_allow_html=True)


# ── SVG Icon helper ────────────────────────────────────────────────────────────
def ico(name: str, size: int = 15) -> str:
    paths = {
        "chat":    '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
        "edit":    '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>',
        "preview": '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>',
        "sun":     '<circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>',
        "moon":    '<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>',
        "layers":  '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>',
        "trash":   '<polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>',
        "settings":'<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
        "check":   '<polyline points="20 6 9 17 4 12"/>',
        "cpu":     '<rect x="9" y="9" width="6" height="6"/><rect x="2" y="2" width="20" height="20" rx="2"/><line x1="9" y1="2" x2="9" y2="6"/><line x1="15" y1="2" x2="15" y2="6"/><line x1="9" y1="18" x2="9" y2="22"/><line x1="15" y1="18" x2="15" y2="22"/><line x1="2" y1="9" x2="6" y2="9"/><line x1="2" y1="15" x2="6" y2="15"/><line x1="18" y1="9" x2="22" y2="9"/><line x1="18" y1="15" x2="22" y2="15"/>',
        "zap":     '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
        "file":    '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>',
        "chevron-right": '<polyline points="9 18 15 12 9 6"/>',
        "chevron-left":  '<polyline points="15 18 9 12 15 6"/>',
        "palette": '<circle cx="12" cy="12" r="10"/><path d="M8 14s1.5 2 4 2 4-2 4-2"/><line x1="9" y1="9" x2="9.01" y2="9"/><line x1="15" y1="9" x2="15.01" y2="9"/>',
        "sliders": '<line x1="4" y1="21" x2="4" y2="14"/><line x1="4" y1="10" x2="4" y2="3"/><line x1="12" y1="21" x2="12" y2="12"/><line x1="12" y1="8" x2="12" y2="3"/><line x1="20" y1="21" x2="20" y2="16"/><line x1="20" y1="12" x2="20" y2="3"/><line x1="1" y1="14" x2="7" y2="14"/><line x1="9" y1="8" x2="15" y2="8"/><line x1="17" y1="16" x2="23" y2="16"/>',
        "help":    '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
        "plus":    '<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>',
        "message-square": '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    }
    sw = "2.5" if name == "check" else "1.75"
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" '
            f'fill="none" stroke="currentColor" stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round">'
            f'{paths.get(name,"")}</svg>')


# ── API helpers ────────────────────────────────────────────────────────────────
def api_health() -> bool:
    try:
        r = httpx.get(f"{API_BASE}/health", timeout=3)
        return r.status_code == 200
    except Exception:
        return False

def api_generate(prompt: str, num_slides: int) -> dict:
    with httpx.Client(timeout=TIMEOUT) as c:
        r = c.post(f"{API_BASE}/generate", json={"prompt": prompt, "num_slides": num_slides})
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

def generate_pdf(plan: dict) -> bytes:
    from fpdf import FPDF
    
    pdf = FPDF(unit="in", format=(13.33, 7.5))
    pdf.set_auto_page_break(auto=True, margin=0.5)
    pdf.set_font("Helvetica", "B", 44)
    
    title = plan.get("presentation_title", "Presentation")
    slides = plan.get("slides", [])
    
    # Title Slide
    pdf.add_page()
    pdf.set_fill_color(30, 39, 97)
    pdf.rect(0, 0, 13.33, 7.5, "F")
    
    pdf.set_fill_color(74, 144, 217)
    pdf.rect(0, 2.8, 0.12, 1.9, "F")
    
    pdf.set_text_color(255, 255, 255)
    pdf.set_y(2.5)
    pdf.set_x(0.4)
    pdf.multi_cell(12.5, 0.5, title, align="L")
    
    pdf.set_font("Helvetica", "I", 16)
    pdf.set_text_color(74, 144, 217)
    pdf.set_y(4.2)
    pdf.set_x(0.4)
    pdf.multi_cell(9, 0.3, "Auto-Generated by PPT Agent", align="L")
    
    # Content Slides
    for i, slide in enumerate(slides, 1):
        pdf.add_page()
        pdf.set_fill_color(240, 244, 255)
        pdf.rect(0, 0, 13.33, 7.5, "F")
        pdf.set_fill_color(30, 39, 97)
        pdf.rect(0, 0, 13.33, 1.4, "F")
        
        pdf.set_font("Helvetica", "B", 32)
        pdf.set_text_color(255, 255, 255)
        pdf.set_y(0.5)
        pdf.set_x(0.35)
        pdf.multi_cell(12.5, 0.4, slide.get("title", f"Slide {i}"), align="L")
        
        pdf.set_fill_color(74, 144, 217)
        pdf.rect(0.35, 1.55, 0.06, 5.5, "F")
        
        bullet_top = 1.65
        bps = slide.get("bullet_points", [])
        bullet_spacing = min(0.95, 5.6 / max(len(bps), 1))
        
        pdf.set_font("Helvetica", "", 18)
        pdf.set_text_color(30, 39, 97)
        
        for j, bp in enumerate(bps[:5]):
            y_pos = bullet_top + j * bullet_spacing
            pdf.set_fill_color(74, 144, 217)
            pdf.rect(0.55, y_pos + 0.12, 0.18, 0.18, "F")
            
            pdf.set_y(y_pos)
            pdf.set_x(0.9)
            pdf.multi_cell(12.0, 0.3, bp, align="L")
            
    return pdf.output()


# ── Render helpers ─────────────────────────────────────────────────────────────
def render_progress(active_step: int):
    steps = [("cpu","Planning slide structure"),("zap","Generating content"),("file","Building PPTX file")]
    pct   = {1:15, 2:55, 3:90}.get(active_step, 15)
    rows  = ""
    for i, (ik, label) in enumerate(steps, 1):
        if i < active_step:    cls, svg = "done",    ico("check")
        elif i == active_step: cls, svg = "active",  ico(ik)
        else:                  cls, svg = "pending", ico(ik)
        rows += f'<div class="prog-step {cls}"><div class="stic {cls}">{svg}</div><span>{label}</span></div>'
    st.markdown(f"""
    <div class="prog-card">
      <div class="prog-head">Processing your request</div>
      {rows}
      <div class="prog-track"><div class="prog-fill" style="width:{pct}%"></div></div>
    </div>""", unsafe_allow_html=True)


def render_slide_card(idx: int, slide: dict):
    label   = "Title Slide" if idx == 1 else f"Slide {idx:02d}"
    title   = slide.get("title", "Untitled")
    bullets = slide.get("bullet_points", [])
    b_html  = "".join(f'<div class="bullet">{b}</div>' for b in bullets)
    if not b_html:
        b_html = '<div class="bullet" style="opacity:.3;font-style:italic">No bullet points</div>'
    st.markdown(f"""
    <div class="slide-card">
      <div class="snum">{label}</div>
      <div class="stitle">{title}</div>
      {b_html}
    </div>""", unsafe_allow_html=True)


def render_chat_message(role: str, content: str, plan: dict | None = None):
    if role == "user":
        st.markdown(f"""
        <div class="msg-row user">
          <div class="av usr">U</div>
          <div class="bub usr"><div class="msg-lbl">You</div>{content}</div>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="msg-row">
          <div class="av bot">◈</div>
          <div class="bub bot"><div>SlideMind</div>{content}</div>
        </div>""", unsafe_allow_html=True)
        if plan and plan.get("slides"):
            prs_title = plan.get("presentation_title", "Presentation")
            n = len(plan["slides"])
            st.markdown(f"""
            <div style="margin-left:2.55rem;margin-top:0.18rem;margin-bottom:0.4rem;">
              <div class="toast info">{ico("layers")} <strong>{prs_title}</strong> &nbsp;&middot;&nbsp; {n} slides generated</div>
            </div>""", unsafe_allow_html=True)


def render_edit_form(slides: list) -> list:
    edited = []
    for i, slide in enumerate(slides):
        label = "Title Slide" if i == 0 else f"Slide {i + 1}"
        with st.expander(label, expanded=(i == 0)):
            new_title = st.text_input("Title", value=slide.get("title", ""), key=f"et_{i}")
            bullets   = slide.get("bullet_points", [])
            new_bullets = []
            for j, bp in enumerate(bullets):
                new_bullets.append(st.text_area(f"Point {j + 1}", value=bp, height=60, key=f"eb_{i}_{j}"))
            if i > 0 and len(bullets) < 6:
                if st.button("+ Add point", key=f"add_{i}"):
                    new_bullets.append("New point.")
            edited.append({"title": new_title, "bullet_points": new_bullets})
    return edited


# ── Chip buttons ───────────────────────────────────────────────────────────────
def chip_row(topics: list, prefix: str):
    cols = st.columns(len(topics))
    for i, (col, topic) in enumerate(zip(cols, topics)):
        with col:
            short = topic.split(" — ")[0].split(" for ")[0]
            short = short[:28] + ("…" if len(short) > 28 else "")
            if st.button(short, key=f"{prefix}_{i}", use_container_width=True):
                st.session_state.prompt_value = topic
                if "main_input" in st.session_state:
                    del st.session_state["main_input"]
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
online = api_health()
active_tab = st.session_state.active_tab

import streamlit.components.v1 as components

with st.sidebar:

    # ── COLLAPSED STATE ───────────────────────────────────────────────────────
    if SB_COLLAPSED:

        components.html(f"""
        <script>
          (function() {{
            var sid = window.parent.document.querySelector('[data-testid="stSidebar"]');
            if (sid) {{
              sid.style.cssText += 'min-width:{SB_WIDTH_COL}px!important;max-width:{SB_WIDTH_COL}px!important;width:{SB_WIDTH_COL}px!important;flex:0 0 {SB_WIDTH_COL}px!important;overflow:hidden!important;';
              var inner = sid.querySelector('section > div');
              if (inner) inner.style.cssText += 'min-width:{SB_WIDTH_COL}px!important;max-width:{SB_WIDTH_COL}px!important;overflow:hidden!important;';
            }}
          }})();
        </script>
        """, height=0)

        status_cls = "on" if online else "off"
        status_title = "API Connected" if online else "API Offline"

        

        # How-to tooltip HTML
        how_tooltip = """
        <div class="rail-tooltip">
          <div class="rail-tooltip-step"><div class="rail-tooltip-num">1</div><span>Describe your topic in chat</span></div>
          <div class="rail-tooltip-step"><div class="rail-tooltip-num">2</div><span>Set slide count near input</span></div>
          <div class="rail-tooltip-step"><div class="rail-tooltip-num">3</div><span>Press Generate</span></div>
          <div class="rail-tooltip-step"><div class="rail-tooltip-num">4</div><span>Preview slide cards inline</span></div>
          <div class="rail-tooltip-step"><div class="rail-tooltip-num">5</div><span>Switch to Editor to refine</span></div>
          <div class="rail-tooltip-step"><div class="rail-tooltip-num">6</div><span>Download your PPTX</span></div>
        </div>
        """

        st.markdown(f"""
        <div class="rail">
          <!-- Logo gem -->
          <div class="rail-status-top">
            <div class="rail-status {status_cls}" title="{status_title}"></div>
          </div>
          <div class="rail-gem" title="SlideMind" style="margin-top:0.2rem;">◈</div>


          <div class="rail-divider"></div>

          <!-- New Chat icon -->
          <div class="rail-btn" id="rail-new-chat-btn" title="New Chat">
            {ico("message-square", 15)}
          </div>

          <!-- Help icon with hover tooltip -->
          <div class="rail-btn" title="How to Use">
            {ico("help", 15)}
            {how_tooltip}
          </div>

          <!-- Expand button at bottom -->
          <div class="rail-expand" id="rail-expand-btn" title="Expand sidebar">
            {ico("chevron-right", 13)}
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Hidden Streamlit buttons
        if st.button("›", key="sb_expand_btn", help="Expand sidebar"):
            st.session_state.sidebar_collapsed = False
            st.rerun()

        if st.button("⊕", key="new_chat_btn_col", help="New chat"):
            st.session_state.chat_history = []
            st.session_state.plan = None
            st.session_state.generated = False
            st.session_state.pptx_bytes = None
            st.session_state.error = None
            st.session_state.active_tab = "chat"
            st.rerun()

        # Wire HTML buttons → Streamlit buttons
        components.html("""
        <script>
          (function wire() {
            var parent = window.parent.document;

            // Expand button
            var expandHtml = parent.getElementById('rail-expand-btn');
            var stBtns = parent.querySelectorAll('[data-testid="stButton"] button');
            var expandSt = null, newChatSt = null;
            stBtns.forEach(function(b) {
              if (b.textContent.trim() === '›') expandSt = b;
              if (b.textContent.trim() === '⊕') newChatSt = b;
            });

            if (expandHtml && expandSt) {
              expandHtml.addEventListener('click', function() { expandSt.click(); });
            }

            // New chat button
            var newChatHtml = parent.getElementById('rail-new-chat-btn');
            if (newChatHtml && newChatSt) {
              newChatHtml.addEventListener('click', function() { newChatSt.click(); });
            }

            if (!expandSt || !newChatSt) setTimeout(wire, 200);
          })();
        </script>
        """, height=0)

    # ── EXPANDED STATE ────────────────────────────────────────────────────────
    else:

        components.html("""
        <script>
          (function() {
            var sid = window.parent.document.querySelector('[data-testid="stSidebar"]');
            if (sid) {
              sid.style.removeProperty('min-width');
              sid.style.removeProperty('max-width');
              sid.style.removeProperty('width');
              sid.style.removeProperty('flex');
              var inner = sid.querySelector('section > div');
              if (inner) {
                inner.style.removeProperty('min-width');
                inner.style.removeProperty('max-width');
                inner.style.removeProperty('width');
                inner.style.removeProperty('flex');
              }
            }
          })();
        </script>
        """, height=0)

        # Logo + collapse button
        hcol1, hcol2 = st.columns([5, 1])
        with hcol1:
            st.markdown("""
            <div class="sb-header" style="border:none;padding:0.2rem 0.75rem 0.5rem;">
              <div class="sb-brand">
                <div class="sb-gem">◈</div>
                <div class="sb-name">Slide<em>Mind</em></div>
              </div>
            </div>""", unsafe_allow_html=True)
        with hcol2:
            st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
            if st.button("‹", key="sb_collapse_btn", help="Collapse sidebar"):
                st.session_state.sidebar_collapsed = True
                st.rerun()

        # API status
        pill_lbl = "API Connected" if online else "API Offline"
        pill_cls = "on" if online else "off"
        st.markdown(
            f'<div class="sb-status-top"><span class="sb-pill {pill_cls}"><span class="pill-dot"></span>{pill_lbl}</span></div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="sb-hr" style="margin:0;"></div>', unsafe_allow_html=True)

        if not online:
            st.markdown(f"""
            <div class="sb-warn">
              Start backend:<br>
              <code>uvicorn server:app --reload</code>
            </div>""", unsafe_allow_html=True)

        # Body
        st.markdown('<div class="sb-body">', unsafe_allow_html=True)

        # Appearance
        st.markdown('<span class="sb-sect">Theme</span>', unsafe_allow_html=True)
        t1, t2 = st.columns(2)
        with t1:
            if st.button("Light", use_container_width=True,
                         type="primary" if not IS_DARK else "secondary", key="th_light"):
                st.session_state.theme = "light"; st.rerun()
        with t2:
            if st.button("Dark", use_container_width=True,
                         type="primary" if IS_DARK else "secondary", key="th_dark"):
                st.session_state.theme = "dark"; st.rerun()

        st.markdown('<div class="sb-hr"></div>', unsafe_allow_html=True)

        # How to use

        st.markdown("""
<div class="how-to-use">
  <div class="how-step"><div class="how-num">1</div><span>Describe your topic in chat</span></div>
  <div class="how-step"><div class="how-num">2</div><span>Set slide count near input</span></div>
  <div class="how-step"><div class="how-num">3</div><span>Press Generate ↑</span></div>
  <div class="how-step"><div class="how-num">4</div><span>Preview cards inline</span></div>
  <div class="how-step"><div class="how-num">5</div><span>Refine in Editor tab</span></div>
  <div class="how-step"><div class="how-num">6</div><span>Download your PPTX</span></div>
</div>""", unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

        # Footer
        st.markdown(
            '<div class="sb-foot">SlideMind · FastAPI · Claude</div>',
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
# MAIN AREA
# ══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="nav-bar">
  <div class="nav-wm">Slide<em>Mind</em></div>
  <div class="nav-tags"><div class="nav-tag">Beta</div><div class="nav-tag">AI Powered</div></div>
</div>""", unsafe_allow_html=True)

c1, c2, c3, _ = st.columns([1.1, 1.1, 1.1, 5])
with c1:
    if st.button("Chat", use_container_width=True,
                 type="primary" if st.session_state.active_tab == "chat" else "secondary", key="tab_chat"):
        st.session_state.active_tab = "chat"; st.rerun()
with c2:
    if st.button("Editor", use_container_width=True,
                 type="primary" if st.session_state.active_tab == "editor" else "secondary",
                 disabled=not st.session_state.generated, key="tab_editor"):
        st.session_state.active_tab = "editor"; st.rerun()
with c3:
    if st.button("Preview", use_container_width=True,
                 type="primary" if st.session_state.active_tab == "preview" else "secondary",
                 disabled=not st.session_state.generated, key="tab_preview"):
        st.session_state.active_tab = "preview"; st.rerun()

st.markdown('<div class="div" style="margin-top:0.3rem;"></div>', unsafe_allow_html=True)

QUICK_TOPICS = [
    "Climate change overview for high school students",
    "Startup pitch deck for a SaaS product",
    "Introduction to Machine Learning for beginners",
    "Q4 business review and next quarter goals",
    "CRISPR gene editing — science and ethics",
]

FOLLOWUP_TOPICS = [
    "Make it more visual",
    "Add an executive summary slide",
    "Shorten the bullet points",
    "Add a conclusion slide",
    "Make the tone more formal",
]


# ══════════════════════════════════════════════════════════════════════════════
# TAB: CHAT
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.active_tab == "chat":

    if not st.session_state.chat_history:
        st.markdown("""
        <div class="empty">
          <div class="empty-mark">◈</div>
          <div class="empty-title">What would you like to present?</div>
          <div class="empty-sub">Describe your topic and I'll build a complete, structured deck — ready to download.</div>
        </div>""", unsafe_allow_html=True)
        st.markdown(
            '<p style="text-align:center;font-size:0.73rem;color:var(--mu);margin-bottom:0.35rem;font-family:var(--fm);">Quick start</p>',
            unsafe_allow_html=True,
        )
        chip_row(QUICK_TOPICS, "chip_empty")
    else:
        for msg in st.session_state.chat_history:
            render_chat_message(msg["role"], msg["content"], msg.get("plan"))

        if st.session_state.generated and st.session_state.plan:
            slides = st.session_state.plan.get("slides", [])
            if slides:
                st.markdown('<div style="margin-left:2.55rem;margin-top:0.1rem;">', unsafe_allow_html=True)
                for row_start in range(0, len(slides), 2):
                    row = slides[row_start:row_start + 2]
                    cols = st.columns(len(row), gap="small")
                    for ci, slide in enumerate(row):
                        with cols[ci]:
                            render_slide_card(row_start + ci + 1, slide)
                st.markdown('</div>', unsafe_allow_html=True)

                dl_c, _, _ = st.columns([2, 2, 4])
                with dl_c:
                    st.download_button(
                        label="Download PPTX",
                        data=st.session_state.pptx_bytes,
                        file_name="presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )

    if st.session_state.error:
        st.markdown(f"""
        <div class="msg-row">
          <div class="av bot">◈</div>
          <div class="bub bot" style="border-color:rgba(248,113,113,.25);">
            <div class="msg-lbl" style="color:var(--er)">Error</div>
            {st.session_state.error}
          </div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="div"></div>', unsafe_allow_html=True)

    if st.session_state.chat_history:
        st.markdown(
            '<p style="font-size:0.72rem;color:var(--mu);margin-bottom:0.3rem;font-family:var(--fm);">Suggestions</p>',
            unsafe_allow_html=True,
        )
        chip_row(FOLLOWUP_TOPICS, "chip_fu")
        st.markdown("<div style='height:0.15rem'></div>", unsafe_allow_html=True)

    # ── Settings row above input ───────────────────────────────────────────────
    scol1, scol2, _ = st.columns([2, 2, 4])
    with scol1:
        st.markdown('<div class="settings-label">Slide Count</div>', unsafe_allow_html=True)
        num_slides = st.slider(
            "", min_value=3, max_value=12,
            value=st.session_state.num_slides,
            label_visibility="collapsed",
            key="slide_count_main",
        )
        st.session_state.num_slides = num_slides

    with scol2:
        st.markdown('<div class="settings-label" style="margin-top:0.05rem;">Export Format</div>', unsafe_allow_html=True)
        st.selectbox(
            "", ["PowerPoint (.pptx)", "PDF"],
            label_visibility="collapsed",
            key="exp_fmt_main",
        )

    st.markdown("<div style='height:0.2rem'></div>", unsafe_allow_html=True)

    # ── Input row ─────────────────────────────────────────────────────────────
    inp_col, btn_col = st.columns([8, 1], gap="small")

    with inp_col:
        prompt = st.text_area(
            "Message",
            value=st.session_state.prompt_value,
            placeholder="Describe your presentation topic…  e.g. 'A deck on the future of renewable energy for investors'",
            height=80,
            label_visibility="collapsed",
            key="main_input",
        )

    with btn_col:
        st.markdown("<div style='height:0.9rem'></div>", unsafe_allow_html=True)
        send_btn = st.button(
            "↑",
            type="primary",
            use_container_width=True,
            disabled=not prompt.strip() or not online,
            key="send_btn",
        )

    if send_btn and prompt.strip():
        st.session_state.prompt_value = ""
        if "main_input" in st.session_state:
            del st.session_state["main_input"]

        st.session_state.chat_history.append({"role": "user", "content": prompt.strip()})
        st.session_state.last_prompt = prompt.strip()
        st.session_state.error = None

        prog = st.empty()
        try:
            with prog.container(): render_progress(1)
            result = api_generate(prompt.strip(), st.session_state.num_slides)
            with prog.container(): render_progress(2)
            time.sleep(0.2)
            with prog.container(): render_progress(3)
            time.sleep(0.2)

            pptx_bytes = api_download()
            plan       = result.get("plan", {})
            prs_title  = plan.get("presentation_title", "Your Presentation")
            n_slides   = len(plan.get("slides", []))

            st.session_state.chat_history.append({
                "role": "assistant",
                "content": (f"I've built <strong>{prs_title}</strong> — {n_slides} slides covering your topic. "
                            "Preview the cards below or switch to the <strong>Editor</strong> tab to refine."),
                "plan": plan,
            })
            st.session_state.plan       = plan
            st.session_state.pptx_bytes = pptx_bytes
            st.session_state.generated  = True

        except httpx.HTTPStatusError as exc:
            try:    detail = exc.response.json().get("detail", str(exc))
            except: detail = str(exc)
            st.session_state.error = f"API error {exc.response.status_code}: {detail}"
            st.session_state.chat_history.append({"role":"assistant","content":f"Something went wrong: <em>{detail}</em>"})
        except httpx.ConnectError:
            st.session_state.error = f"Cannot reach backend at {API_BASE}."
            st.session_state.chat_history.append({"role":"assistant","content":"Cannot reach the backend. Run <code>uvicorn server:app --reload</code>."})
        except Exception as exc:
            st.session_state.error = str(exc)
            st.session_state.chat_history.append({"role":"assistant","content":f"Unexpected error: <em>{exc}</em>"})

        prog.empty()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB: EDITOR
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "editor":
    if not st.session_state.generated:
        st.markdown("""
        <div class="empty">
          <div class="empty-mark" style="font-size:1.1rem;">✏</div>
          <div class="empty-title">Nothing to edit yet</div>
          <div class="empty-sub">Generate a presentation from the Chat tab first.</div>
        </div>""", unsafe_allow_html=True)
    else:
        plan      = st.session_state.plan
        slides    = plan.get("slides", [])
        prs_title = plan.get("presentation_title", "Presentation")

        col_h, col_open, col_dl = st.columns([3, 1, 1])
        with col_h:
            st.markdown(f"""
            <div style="margin-bottom:1rem;">
              <div style="font-family:var(--fm);font-size:0.59rem;color:var(--mu);text-transform:uppercase;letter-spacing:0.1em;margin-bottom:0.28rem;">Editing</div>
              <div style="font-family:var(--fd);font-size:1.12rem;font-weight:700;color:var(--tx);letter-spacing:-0.02em;">{prs_title}</div>
              <div style="font-size:0.78rem;color:var(--mu);font-family:var(--fb);">{len(slides)} slides</div>
            </div>""", unsafe_allow_html=True)
        with col_open:
            if st.button("Open Native", key="open_ppt_editor", use_container_width=True):
                import tempfile
                import subprocess
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                tmp.write(st.session_state.pptx_bytes)
                tmp.close()
                subprocess.Popen(["open", tmp.name])
        with col_dl:
            is_pdf = (st.session_state.get("exp_fmt_main") == "PDF")
            st.download_button(
                "Download",
                data=generate_pdf(st.session_state.plan) if is_pdf else st.session_state.pptx_bytes,
                file_name="presentation.pdf" if is_pdf else "presentation.pptx",
                mime="application/pdf" if is_pdf else "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

        st.markdown('<div class="div"></div>', unsafe_allow_html=True)
        edited_slides = render_edit_form(slides)
        st.markdown("<br>", unsafe_allow_html=True)

        upd_col, _ = st.columns([2, 5])
        with upd_col:
            if st.button("Update Presentation", type="primary", use_container_width=True, key="update_btn"):
                updated_plan = {"presentation_title": prs_title, "slides": edited_slides}
                with st.spinner("Re-rendering…"):
                    try:
                        api_update(updated_plan)
                        new_bytes = api_download()
                        st.session_state.plan["slides"] = edited_slides
                        st.session_state.pptx_bytes     = new_bytes
                        st.markdown(
                            f'<div class="toast success" style="margin-top:0.65rem;">'
                            f'{ico("check")} Presentation updated successfully.</div>',
                            unsafe_allow_html=True,
                        )
                    except Exception as exc:
                        st.error(f"Update failed: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB: PREVIEW
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "preview":
    if not st.session_state.generated:
        st.markdown("""
        <div class="empty">
          <div class="empty-mark" style="font-size:1.1rem;">▦</div>
          <div class="empty-title">No slides to preview</div>
          <div class="empty-sub">Generate a presentation from the Chat tab first.</div>
        </div>""", unsafe_allow_html=True)
    else:
        plan      = st.session_state.plan
        slides    = plan.get("slides", [])
        prs_title = plan.get("presentation_title", "Presentation")

        total = sum(len(s.get("bullet_points", [])) for s in slides)
        avg   = round(total / max(len(slides), 1), 1)

        st.markdown(f"""
        <div class="stat-row">
          <div class="stat-card"><div class="stat-val">{len(slides)}</div><div class="stat-lbl">Slides</div></div>
          <div class="stat-card"><div class="stat-val">{total}</div><div class="stat-lbl">Total Points</div></div>
          <div class="stat-card"><div class="stat-val">{avg}</div><div class="stat-lbl">Avg / Slide</div></div>
        </div>""", unsafe_allow_html=True)

        hc, oc, dc = st.columns([3, 1, 1])
        with hc:
            st.markdown(
                f'<div style="font-family:var(--fd);font-size:1.08rem;font-weight:700;letter-spacing:-0.02em;margin-bottom:0.85rem;">{prs_title}</div>',
                unsafe_allow_html=True,
            )
        with oc:
            if st.button("Open Native", key="open_ppt_preview", use_container_width=True):
                import tempfile
                import subprocess
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                tmp.write(st.session_state.pptx_bytes)
                tmp.close()
                subprocess.Popen(["open", tmp.name])
        with dc:
            is_pdf = (st.session_state.get("exp_fmt_main") == "PDF")
            st.download_button(
                "Download",
                data=generate_pdf(st.session_state.plan) if is_pdf else st.session_state.pptx_bytes,
                file_name="presentation.pdf" if is_pdf else "presentation.pptx",
                mime="application/pdf" if is_pdf else "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

        st.markdown('<div class="div"></div>', unsafe_allow_html=True)

        for row_start in range(0, len(slides), 2):
            row = slides[row_start:row_start + 2]
            cols = st.columns(len(row), gap="medium")
            for ci, slide in enumerate(row):
                with cols[ci]:
                    render_slide_card(row_start + ci + 1, slide)
            st.markdown("<div style='height:0.2rem'></div>", unsafe_allow_html=True)