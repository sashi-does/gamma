**SlideMind**  
AI-Powered Presentation Builder  
FastAPI · MCP · OpenRouter · Streamlit · python-pptx

**Python 3.10+**  
**FastAPI Backend**  
**Streamlit UI**  
**OpenRouter**

**Project Overview**  
SlideMind is a full-stack, AI-powered presentation generator that turns a plain-text topic or prompt into a fully formatted, themed PowerPoint deck (.pptx) — without touching any slide software. Users interact through a modern Streamlit chat interface. A FastAPI backend orchestrates an autonomous agent that plans the slide structure using an LLM and assembles the PPTX programmatically through MCP (Model Context Protocol) tool servers.

The architecture deliberately separates concerns: the LLM is responsible only for content planning, while all file operations (creating slides, writing to disk) happen through sandboxed MCP subprocess servers. This ensures the AI can never perform arbitrary filesystem operations — it can only call explicitly defined tools.

**Repository Structure & File Reference**

File / Path                        | Type          | Size       | Purpose & Detailed Description
-----------------------------------|---------------|------------|--------------------------------
app.py                             | Frontend      | 1310 ln    | The Streamlit application — the entire user-facing experience. Renders the dark glassmorphic UI, manages session state, handles chat history, dispatches requests to the FastAPI backend via httpx, and renders slide preview cards inline. Includes full custom CSS/JS theming.
server.py                          | Backend       | 370 ln     | The FastAPI REST server that acts as the orchestration layer. Exposes /generate, /update, /download, and /health endpoints.
agent_ppt.py                       | AI Agent      | 269 ln     | The core autonomous AI agent implementing the 3-phase pipeline (PLAN → EXEC → SAVE).
requirements.txt                   | Config        | 8 ln       | Python dependency manifest.
servers/ppt_mcp_server.py          | MCP Server    | ~150 ln    | PowerPoint MCP tool server exposing create_presentation, add_slide, and save_presentation.
servers/filesystem_mcp_server.py   | MCP Server    | ~80 ln     | Filesystem MCP tool server for read/list access.
.env (not tracked)                 | Secret        | —          | Contains OPENROUTER_API_KEY.

**System Architecture**

Layer                | Tech                  | Responsibilities
---------------------|-----------------------|-----------------------------------
Frontend (UI)        | Streamlit             | Chat interface, session state, slide previews, editor, PDF export
API Gateway          | FastAPI               | REST endpoints, validation, MCP lifecycle, file streaming
AI Agent Core        | agent_ppt.py          | 3-phase orchestration: planning → tool dispatch → save
LLM Provider         | OpenRouter            | Routes to gpt-4o-mini
PPT MCP Server       | FastMCP + python-pptx | Slide creation and .pptx assembly with Midnight Executive theme
FS MCP Server        | FastMCP               | Filesystem verification

**Agent Pipeline Detail**
- **Phase 1 — PLAN**: LLM call via OpenRouter to generate structured JSON slide plan.
- **Phase 2 — EXEC**: Calls PPT MCP tools to create slides (with optional extra LLM call for bullet points if needed).
- **Phase 3 — SAVE**: Saves the presentation as output.pptx.

**Key Features**
- Conversational slide plan generation from any topic
- Configurable slide count (3–10)
- Live HTML slide preview cards
- Built-in Editor tab for editing titles and bullets
- One-click PPTX download with Midnight Executive dark theme
- Instant PDF export (generated in frontend)
- Dark/Light theme toggle
- Sidebar with API health indicator
- Secure MCP-based sandboxing for file operations

**Getting Started**

1. Clone the repository:  
   `git clone https://github.com/sashi-does/gamma.git && cd gamma`

2. Create and activate virtual environment

3. Install dependencies:  
   `pip install -r requirements.txt`

4. Create `.env` file with your key:  
   `OPENROUTER_API_KEY="sk-or-v1-..."`

5. Start backend:  
   `uvicorn server:app --reload`

6. Start frontend (in new terminal):  
   `streamlit run app.py`

**Environment Variables**  
Create a `.env` file containing:  
`OPENROUTER_API_KEY="sk-or-v1-your-key-here"`

**Future Improvements**
- Multiple theme options
- Pre-built slide templates
- Real-time collaboration
- Voice input support
- Image generation per slide
- Direct export to Google Slides
- Streaming generation progress

**License & Credits**  
This project is open source. Built by sashi-does.  
Repository: https://github.com/sashi-does/gamma

