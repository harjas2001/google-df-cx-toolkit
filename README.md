# Dialogflow CX Toolkit

A collection of diagnostic and analysis tools for Dialogflow CX agent development and quality assurance. Built from production NLU work on enterprise-scale conversational AI deployments.

Four focused tools — covering training phrase extraction, NLU overlap analysis, casing quality checks, and flow page auditing.

---

## Background

Built as supporting tooling during an NLU quality initiative across multiple production Dialogflow CX agents handling millions of chats via LLM playbooks annually across two national telco brands.

The core problem: as conversational AI agents grow, intent libraries accumulate training phrase overlap, inconsistent casing, and underpopulated intents — all of which quietly degrade NLU confidence without any obvious signal in aggregate metrics. These tools were built to surface those issues systematically and produce prioritised, actionable output rather than raw data dumps. LLM playbooks need context of misalignment in NLU and match as well

The analysis pipeline (`phrase_extraction.py` → `analyse_training_phrases.py`) was used to produce remediation reports across agent deployments, identifying intent pairs with critical Jaccard overlap, intents with insufficient phrase coverage, and shared tokens that provided no discriminative signal to the model.

---

## Tools

### 1. `phrase_extraction.py`

Extracts training phrases from a filtered set of intents in a Dialogflow CX agent export and writes them to a CSV file. The output is the input for the analysis script.

```bash
python phrase_extraction.py
```

Reads intent names from `config/include_intents.txt` (gitignored). Only intents listed there are extracted.

**Output:** `training_phrases.csv` (or as configured in `.env`)

---

### 2. `analyse_training_phrases.py`

Deep NLU analysis of training phrases. Takes the CSV from the extraction step and produces a fully formatted 6-sheet Excel report.

```bash
python analyse_training_phrases.py
```

| Sheet | What it shows |
|---|---|
| Summary | Per-intent phrase count, word count stats, health flags |
| Token Overlap Matrix | Pairwise Jaccard similarity heatmap across all intents |
| High Overlap Pairs | Ranked intent pairs above the overlap threshold |
| Shared Tokens Detail | Every shared token and which intents use it |
| Intent Word Stats | Full word-count distribution per intent |
| Action Items | Prioritised remediation tasks (CRITICAL / HIGH / MEDIUM / LOW) |

**Output:** `training_phrase_analysis.xlsx` (or as configured in `.env`)

---

### 3. `check_for_upper_lower_case.py`

Scans all intents in a Dialogflow CX agent export and flags training phrases written entirely in CAPITAL LETTERS. All-caps phrases are a common data quality issue — they cause the NLU model to overfit to casing and typically originate from CRM exports or internal documentation.

```bash
python check_for_upper_lower_case.py
```

Expects the standard Dialogflow CX export folder structure (`df_cx_agent/intents/`).

---

### 4. `lastPage_list.py`

Iterates the exported agent config and prints all pages grouped by flow. Useful for auditing flow structure, mapping terminal pages, and identifying orphaned or unreachable pages before a release.

```bash
python lastPage_list.py
```

---

## Setup

```bash
git clone https://github.com/your-username/dialogflow-cx-toolkit.git
cd dialogflow-cx-toolkit

python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

pip install -r requirements.txt

cp .env.example .env
# Edit .env with your agent paths and filenames

cp config/include_intents.example.txt config/include_intents.txt
# Edit include_intents.txt with your intent names
```

---

## Configuration

All paths and filenames are set in `.env` — see `.env.example` for all available options. No hardcoded values in any script.

The intent whitelist for extraction is managed separately in `config/include_intents.txt` (gitignored). This keeps internal intent names out of the repository while making the tool fully reusable across different agents.

---

## Expected Agent Export Structure

```
df_cx_agent/
├── flows/
│   └── <flow-name>/
│       └── pages/
│           └── <page-name>.json
└── intents/
    └── <intent-name>/
        └── trainingPhrases/
            └── en.json
```

This matches the standard folder structure produced by the Dialogflow CX export function. See `sample_data/training_phrases_schema.json` for the expected JSON schema of a training phrases file.

---

## Stack

Python · openpyxl · python-dotenv
