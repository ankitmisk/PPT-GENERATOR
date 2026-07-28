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
- **7 slide layouts** used automatically by the AI: title, content,
  two-column, quote, section divider, closing, and image-focus — each styled per template.
- **AI-generated images (optional)** — toggle "🖼️ Generate AI images for
  slides" to have the agent decide which slides need a visual, write a short
  image prompt for each, and generate it with **Gemini 2.5 Flash Image**
  ("Nano Banana") using the *same* Google API key already entered above (no
  separate image API/key needed). Pick from 5 image styles (flat vector,
  photorealistic, 3D render, line art, watercolor) applied consistently
  across the deck. Images are embedded directly in both the HTML preview and
  the PPTX (as real embedded pictures, not links).
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
- **Image generation** uses Google's `gemini-2.5-flash-image` model via a
  direct REST call (`generativelanguage.googleapis.com`). The free tier is
  roughly 500 requests/day at the time of writing, but Google's limits can
  change — check https://ai.google.dev for current quotas. If an image fails
  to generate (quota, safety filter, etc.), that one slide just falls back
  to a text-only layout and the rest of the deck still completes.
- `gemini-3-pro-image-preview` ("Nano Banana Pro") is a separate, higher-end
  model with **no free tier** — this app intentionally uses the Flash Image
  model instead to stay within the free quota.
