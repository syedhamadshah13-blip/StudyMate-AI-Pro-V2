# StudyMate AI Pro V2
## Autonomous Multi-Agent Learning System

StudyMate AI Pro V2 upgrades the original RAG study assistant into an observable
multi-agent workflow for the Pak Angels Cohort 11 — 2nd Hackathon.

### Core hackathon flow
1. Student enters a study objective.
2. ChromaDB retrieves relevant indexed notes.
3. Architect Agent creates a prioritized 3-day syllabus.
4. Examiner Agent receives the exact syllabus and builds five MCQs.
5. A Python pre-check detects obvious structure problems.
6. QA Critic receives the syllabus, exam and pre-check issues.
7. QA Critic corrects scope/quality problems and certifies the module.
8. The app exports Markdown, PDF and an audit log.

### Other capabilities
- PDF, TXT, DOCX and PPTX study-file ingestion
- Audio/video lecture processing
- Live lecture audio recording
- RAG Socratic chat with source labels
- Existing quick exam/flashcard tools
- Adaptive revision coach
- Agent audit trail
- Environment/security status page
- Multi-agent architecture page

### Windows 10 quick start

```powershell
cd PATH_TO_PROJECT
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
copy .env.example .env
notepad .env
python -m streamlit run app.py
```

Put your real Gemini API key in `.env`:

```text
GEMINI_API_KEY=YOUR_REAL_KEY
```

Never commit `.env` to GitHub.

### Hackathon mapping
- Generative AI: Gemini-generated learning artifacts.
- Agentic AI: specialized goal-driven agent roles.
- Multi-Agent Systems: Architect → Examiner → QA Critic.
- AI Workflows: automated artifact handoffs and validation.
- Business Process Automation: one-click learning-content production and QA pipeline.

### Demo path
Upload a real study PDF → enter a study objective → launch workflow →
show the RAG sources → show three agent handoffs → show QA corrections →
download certified module + audit log.


## Ultimate V2 judge-impact upgrades

- Visual Learning Mind Map: Architect syllabus -> Mermaid flowchart rendered in Streamlit.
- Agent 4 Challenger: Socratic reasoning defense and follow-up evaluation.
- Vision-Enabled RAG: optional Gemini Vision analysis for PDF pages and PPTX images; useful insights are indexed in ChromaDB.
- Voice-to-Voice Socratic Tutor: Streamlit audio input + Gemini response + browser Text-to-Speech.

### Recommended live demo
1. Upload a PDF with Vision Analysis enabled.
2. Launch the autonomous workflow.
3. Show RAG sources and the Architect Mind Map.
4. Show Examiner -> Python Pre-Check -> QA Critic.
5. Download the certified PDF and audit JSON.
6. Enter a quiz result and show Adaptive Revision.
7. Use the Challenger Agent once.
8. Ask one short voice question and play the spoken answer.

Mermaid visualization loads a browser-side module from a CDN, so internet access is required for the visual rendering.
