# SlideMind – AI Presentation Builder ◈

SlideMind is a modern, full-stack AI-powered presentation generator. It allows you to rapidly generate completely structured slide decks by simply talking to it like a chatbot. Start with a prompt, generate full slides with AI, directly preview the cards, and download them instantly as beautiful native PowerPoint (`.pptx`).

---

## 🎥 Demo Video

Watch SlideMind in action:

[![Watch Demo](https://img.youtube.com/vi/QqBGYNpWWhI/0.jpg)](https://youtu.be/QqBGYNpWWhI)

---

The project splits the workload cleanly between a stunning, lightning-fast **Streamlit** frontend and a robust **FastAPI / MCP** (Model Context Protocol) backend capable of assembling raw PowerPoint files procedurally.

---

## Technical Stack

### Frontend
- **Streamlit**: Powers the entire real-time conversational UI and slide editor grid.
- **Custom CSS / JS**: Implements a highly polished vibrantly dark glassmorphic design theme, custom tooltips, slide carousels, and responsive sidebars.
- **FPDF2**: A pure-Python implementation integrated into the frontend to generate replica PDF exports on-the-fly directly from the AI's generated structural plan.

### Backend
- **FastAPI**: Provides the REST API layer handling jobs (`/generate`, `/update`, `/download`).
- **MCP (Model Context Protocol)**: Uses a modular tool architecture via FastMCP (`agent_ppt.py` and `servers/ppt_mcp_server.py`) limiting direct LLM permissions.
- **Python-PPTX**: The heavy lifter script within the tool server used to programmatically render PowerPoint shapes, text grids, and themed colors (The slide builder enforces a strict 'Midnight Executive' theme standard).
- **OpenRouter API**: Intelligent model backing the prompt expansion, planning, and structured content routing phase.

---

## Local Development & Setup

Both the backend generation server and the frontend UI daemon need to be running simultaneously to use the application.

### 1. Requirements & Environment

Initialize your local Python environment and install the required dependencies:

```bash
# Create and activate your virtual environment (venv is already configured in this directory)
source venv/bin/activate

# Install all primary dependencies
pip install -r requirements.txt
```

---

### 2. Environment Variables

You must have a valid API key exposed to the background runtime so the agent can communicate with the LLM API.

Ensure your `.env` file exists in the directory containing:

```bash
OPENROUTER_API_KEY="sk-or-v1-..."
```

---

### 3. Run the Backend API 

This spins up the FastAPI layer on localized port 8000. It manages the long-running generation tasks, rendering pipelines, and the internal MCP server sub-processes.

```bash
uvicorn server:app --reload
```

*Note: The application requires the backend to be online to unlock the primary interface.*

---

### 4. Run the Frontend 

Open a new terminal tab, ensure your virtual environment is active, and launch the Streamlit dashboard:

```bash
streamlit run app.py
```

This will automatically bridge with your browser and connect to the local backend.

---

## Key Features

- **Conversational Builder**: Chat interface routing naturally to presentation parameters.
- **Live Markdown Preview**: Native visually rendered card arrays corresponding to structural slides built inside Streamlit.
- **Editor Mode**: Interactive form expansion modifying text blocks/bullets dynamically across all slides.
- **Multi-Format Export**: Quick download bridging direct `.pptx` assembly or procedural `.pdf` drafting.
- **Vibrant Styling**: Custom native UI theming (No default generic component bounds).

---

##  Future Improvements (Optional Ideas)

- Theme customization engine
- Slide templates marketplace
- Real-time collaboration
- Voice-to-presentation input
