# SlideMind – AI Presentation Builder ◈

SlideMind is a modern, full-stack AI-powered presentation generator. It allows you to rapidly generate completely structured slide decks by simply providing a text prompt. Start with a topic, generate full slides with AI, preview them as cards, edit if needed, and download instantly as a beautiful native PowerPoint (`.pptx`) file.

## Demo Video

Watch SlideMind in action:

[![Watch Demo](https://img.youtube.com/vi/QqBGYNpWWhI/0.jpg)](https://youtu.be/QqBGYNpWWhI)

The project splits the workload cleanly between a stunning, lightning-fast **Streamlit** frontend and a robust **FastAPI + MCP** (Model Context Protocol) backend capable of assembling raw PowerPoint files procedurally.

## Technical Stack

### Frontend
- **Streamlit**: Powers the entire real-time conversational UI, slide preview grid, and editor.
- **Custom CSS / JS**: Implements a polished dark glassmorphic design with custom fonts, tooltips, and responsive layout.
- **FPDF2**: Used directly in the frontend to generate replica PDF exports on-the-fly from the slide plan.

### Backend
- **FastAPI**: Provides the REST API layer with endpoints for generation, updates, and downloads.
- **MCP (Model Context Protocol)**: Modular tool architecture via FastMCP that restricts LLM permissions to only defined tools.
- **Python-PPTX**: Used inside the MCP server to programmatically create slides, text boxes, bullet lists, and apply the 'Midnight Executive' dark theme.
- **OpenRouter**: Powers the intelligent planning and content generation phase using gpt-4o-mini.

## System Architecture

| Layer                | Technology                  | Responsibilities |
|----------------------|-----------------------------|------------------|
| Frontend (UI)        | Streamlit                   | Chat interface, session state management, live slide preview cards, interactive editor, theme toggle, PDF export |
| API Gateway          | FastAPI                     | REST endpoints (/generate, /update, /download, /health), request validation, CORS, MCP session management, file streaming |
| AI Agent Core        | agent_ppt.py                | 3-phase pipeline: Plan (LLM), Execute (MCP tools), Save |
| LLM Provider         | OpenRouter                  | Routes calls to gpt-4o-mini for slide structure planning and bullet point generation |
| PPT MCP Server       | FastMCP + python-pptx       | Creates and assembles PowerPoint slides with enforced Midnight Executive theme |
| FS MCP Server        | FastMCP                     | Provides safe filesystem read/list access for verification |

```ascii
┌─────────────────────────────────────────────────────┐
│                Streamlit Frontend                   │
│                   (port 8501)                       │
│                                                     │
│  • Chat Interface                                   │
│  • Generate Button                                  │
│  • Live Slide Preview Cards                         │
│  • Interactive Editor Tab                           │
│  • Theme Toggle (Dark/Light)                        │
│  • PDF Export (fpdf2)                               │
│  • Download PPTX Button                             │
└──────────────────────┬──────────────────────────────┘
                       │ HTTP (httpx)
                       │ POST /generate
                       │ POST /update
                       │ GET /download
                       │ GET /health
                       ▼
┌─────────────────────────────────────────────────────┐
│               FastAPI Backend                       │
│                 server.py (port 8000)               │
│                                                     │
│  • /generate     → Full 3-phase agent pipeline      │
│  • /update       → Re-render PPTX from edited plan  │
│  • /download     → Stream output.pptx               │
│  • /health       → API liveness probe               │
│                                                     │
│  Manages MCP ClientSessions via asynccontextmanager │
└──────────────────────┬──────────────────────────────┘
                       │ stdio (MCP Protocol)
                       ▼
┌──────────────────────┬──────────────────────────────┐
│   ppt_mcp_server.py  │  filesystem_mcp_server.py    │
│     (MCP Server 1)   │       (MCP Server 2)         │
│                      │                              │
│  • create_presentation │  • read_file               │
│  • add_slide         │  • list_dir                  │
│  • save_presentation │                              │
│                      │                              │
│  Uses python-pptx    │  Safe filesystem access      │
│  Midnight Executive  │  for post-save verification  │
│  dark theme          │                              │
└──────────────────────┴──────────────────────────────┘

                  ▲
                  │ LLM Calls via OpenRouter (gpt-4o-mini)
                  │
             agent_ppt.py
       (3-Phase Pipeline: PLAN → EXEC → SAVE)

```
## Agent Pipeline

The AI agent follows a strict 3-phase process:

1. **Phase 1 – PLAN**: Sends the user prompt to OpenRouter (gpt-4o-mini) and receives a structured JSON slide plan containing presentation title and list of slides with bullet points.
2. **Phase 2 – EXEC**: Initializes the presentation and calls the PPT MCP server to add each slide. If a slide has fewer than 3 bullet points, an additional LLM call expands the content.
3. **Phase 3 – SAVE**: Calls the save tool to write the final `output.pptx` file to disk.

## Local Development & Setup

Both the backend and frontend must run simultaneously.

### 1. Requirements & Environment
```bash
# Activate virtual environment (if not already active)
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Environment Variables
Create a `.env` file in the project root with:
```bash
OPENROUTER_API_KEY="sk-or-v1-..."
```

### 3. Run the Backend API
```bash
uvicorn server:app --reload
```
This starts the FastAPI server on http://127.0.0.1:8000

### 4. Run the Frontend
Open a new terminal and run:
```bash
streamlit run app.py
```
The application will open automatically in your browser.

**Note**: The sidebar will show a green status indicator when the backend is connected.

## Key Features
- Conversational Builder – Describe your topic and get a complete slide deck
- Configurable number of slides (3 to 10)
- Live slide preview cards rendered directly in the UI
- Interactive Editor tab to modify titles and bullet points
- One-click download of native `.pptx` file with Midnight Executive theme
- Instant PDF export generated in the frontend
- Clean dark glassmorphic UI with theme toggle
- Secure sandboxed architecture using MCP tools

## Repository Structure (Key Files)

- `app.py` – Streamlit frontend application
- `server.py` – FastAPI backend server
- `agent_ppt.py` – Core AI agent logic
- `servers/ppt_mcp_server.py` – PowerPoint tool server
- `servers/filesystem_mcp_server.py` – Filesystem tool server
- `requirements.txt` – Project dependencies

## Future Improvements
- Theme customization engine
- Pre-built slide template marketplace
- Real-time collaboration support
- Voice-to-presentation input
- Automatic image generation per slide
- Direct export to Google Slides & compatible to canva
