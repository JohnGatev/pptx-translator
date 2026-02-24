# PPTX Translator

Translate PowerPoint decks and Word documents using a LiteLLM/OpenAI-compatible endpoint while keeping the original layout intact.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

## Notes

- PPTX and DOCX paragraphs are translated in batches, then written back to runs to preserve formatting.
- PPTX charts and SmartArt text are translated by editing OOXML after save.
- Some embedded objects (OLE/math) remain unchanged.
- Provide your API key in the sidebar before translating.
- Image text translation requires a vision-capable model; text-only models skip image processing.

## Finance lexicon RAG (required)

Place the PDF lexicon file `Finance 1 Lexicon ENG NL.pdf` in the project root. The app builds a lightweight RAG index from this PDF and injects the most relevant entries into each translation call as context, so the model can apply the terms contextually (including variations and grammar).
The matcher scans the full lexicon per batch and uses fuzzy/token matching plus a repair pass to enforce glossary usage.

## Deploy (Streamlit Community Cloud)

1. Push this repo to GitHub.
2. Go to Streamlit Community Cloud and click **New app**.
3. Select your repo/branch and set the main file to `app.py`.
4. Deploy. Dependencies are installed from `requirements.txt` and Python is pinned in `runtime.txt`.
