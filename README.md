# AI PPT Generator (Gamma-style)

Describe a topic, pick one of 5 templates, and get a full presentation —
previewed live as an interactive HTML deck, downloadable as a dynamic HTML
file or a real, editable PowerPoint (.pptx).

## Features

- **LangChain agent + Gemini** — same `create_agent` pattern as the resume
  generator, calling `gemini-3.5-flash-lite` to write a structured slide
  outline (title, subtitle, bullets, speaker notes) as strict JSON.
- **Optional live research** — if a Tavily key is provided, the agent can
  pull a few current facts about your topic to ground the content (used as
  inspiration for the AI, not quoted directly).
- **5 built-in templates**, selectable via clickable preview cards:
  1. **Minimal Light** — clean, spacious, editorial
  2. **Dark Professional** — bold, high-contrast, executive
  3. **Gradient Modern** — vibrant gradient background, startup style
  4. **Corporate Blue** — structured, formal, trustworthy
  5. **Creative Pastel** — playful, soft, approachable
- **One data source, two outputs** — the same generated slide JSON drives
  both the HTML deck and the PPTX file, so they always stay visually
  consistent and switching templates updates both instantly (no re-generation
  needed).
- **6 slide layouts** used automatically by the AI: title, content,
  two-column, quote, section divider, and closing — each styled per template.
- **Live HTML preview** with prev/next buttons and keyboard arrow navigation,
  right inside the app.
- **Speaker notes** are included in the downloaded PPTX (visible in
  PowerPoint's Notes pane).

## Requirements

- Python 3.10+
- **Google API key** (required) — powers the content-writing agent.
- **Groq / Tavily API keys** (optional) — Tavily enables the "enrich with
  live web research" option.

## Setup

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r ppt_requirements.txt
```

## Running the app

```bash
streamlit run ppt_generator.py
```

Then:

1. Enter your **Google API key** in the sidebar (Tavily is optional).
2. Pick a **template** from the 5 preview cards.
3. Describe your presentation topic and choose a **slide count** (5–15).
4. Optionally enable **"Enrich content with live web research"**.
5. Click **✨ Generate Presentation**.
6. Preview the deck, then download as **HTML** (instant) or click
   **Prepare PPTX** to generate and download the PowerPoint file.

## Notes

- Changing the template after generation re-renders the HTML preview
  immediately from the same slide content — no new AI call is made.
- The PPTX is built with `python-pptx` using manually-positioned text boxes
  and shapes (not PowerPoint's built-in themes), so the 5 designs render
  identically to their HTML counterparts, including gradient backgrounds.
- If the AI's output isn't valid JSON (rare), the app shows an error and
  asks you to regenerate rather than guessing at broken content.
