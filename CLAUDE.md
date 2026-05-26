# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a Udemy course repository for the Azure AI-102 exam, containing standalone Python samples for each Azure AI service category. There is no shared application framework — each subdirectory is a self-contained script with its own `requirements.txt` and `.env` file.

## Running Scripts

Each script must be run from its own directory, since file paths (images, audio, docs) are relative:

```bash
cd 5_AI_Language/sentiment_analysis
pip install -r requirements.txt
python sentiment_analysis.py
```

Scripts write results to `result.json` (or a `.md` file for Document Intelligence) in the same directory.

## Environment Variables

Every script reads credentials from a `.env` file (loaded via `python-dotenv`) or directly from environment variables. The `.env` file is gitignored in each directory. Variable names by service:

| Service | Variables |
|---|---|
| AI Vision | `VISION_ENDPOINT`, `VISION_KEY` |
| AI Language (TextAnalytics) | `LANGUAGE_ENDPOINT`, `LANGUAGE_KEY` |
| Speech | `SPEECH_KEY`, `SPEECH_REGION` |
| Azure AI Search | `SEARCH_ENDPOINT`, `SEARCH_KEY` |
| Azure OpenAI (Embeddings) | `AOAI_ENDPOINT`, `AOAI_KEY` |
| Document Intelligence | `DI_ENDPOINT`, `DI_KEY` |
| Translator | `TRANSLATOR_KEY`, `TRANSLATOR_REGION` |

## Architecture by Module

### 3_AIVision
- `computer_vision/analyze_image/` — uses `ImageAnalysisClient` with `VisualFeatures` flags (CAPTION, READ, etc.)
- `computer_vision/delete_background/` — calls the REST API directly (`computervision/imageanalysis:segment`) with `mode=backgroundRemoval`; does not use the SDK
- `ocr_use_aivision/` — uses `ImageAnalysisClient` with `VisualFeatures.READ`, saves result to `result.json`

### 4_CustomVision
Training images only (ball classification dataset). No Python scripts; intended for upload to the Azure Custom Vision portal.

### 5_AI_Language
All text analysis scripts (`sentiment_analysis`, `key_phrase_extraction`, `language-detection`, `named-entity-recognition`, `extract-linked-entities`) follow the same pattern: read `docs.md` → call `TextAnalyticsClient` → write `result.json`.

Speech scripts require **ffmpeg** and **pydub** to convert `.m4a` → `.wav` before calling `speechsdk`. On Linux:
```bash
sudo apt-get install ffmpeg
```

`text_translator/` calls the Translator REST API directly (not an SDK) using `TRANSLATOR_KEY` and `TRANSLATOR_REGION`.

### 6_AISearch
Three progressive search implementations sharing the same document schema (`DocumentId`, `DocumentName`, `Content`, `Tags`). All three scripts run destructively on startup: delete all existing documents → re-upload from `docs/` → run a sample query.

- `full_text_search/` — creates the index schema programmatically using `SearchIndexClient`; uses `ja.lucene` analyzer for Japanese content
- `vector_search/` — adds a `ContentVector` field; calls Azure OpenAI `text-embedding-3-large` to vectorize documents and queries
- `hybrid_search/` — combines full-text `search_text` with `VectorizedQuery` in a single `search()` call

The `infra/` folders contain screenshots of the Azure portal index setup (no Bicep/Terraform).

### 7_DocumentIntelligence
`Layout/document-intelligence.py` — uses `DocumentIntelligenceClient` with `prebuilt-layout` model and `output_content_format="markdown"`. Writes the extracted markdown to `<input_file>.md`.

### 8_OpenAI
No Python code — contains a sample PDF (`2024年7-9月GDP.pdf`) and a `readme.md` with expected system message and QA pairs for testing Azure OpenAI file-based Q&A (file upload / retrieval scenario).

## Dev Container

The `.devcontainer/devcontainer.json` uses `mcr.microsoft.com/devcontainers/python:1-3.11-bookworm` (Python 3.11). No post-create commands are configured; install dependencies per-module as needed.
