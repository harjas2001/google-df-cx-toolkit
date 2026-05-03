"""
phrase_extraction.py
─────────────────────────────────────────────────────────────────────────────
Extracts training phrases from a filtered set of Dialogflow CX intents and
writes them to a CSV file (columns: utterance, intent).

The output CSV is the input for analyse_training_phrases.py.

Configuration:
  AGENT_INTENTS_PATH  — path to the intents folder in your agent export
  OUTPUT_CSV          — output filename
  Both are loaded from .env (see .env.example).

Intent filter:
  Only intents listed in config/include_intents.txt are extracted.
  Copy config/include_intents.example.txt to get started.
─────────────────────────────────────────────────────────────────────────────
"""

import os
import json
import csv
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
AGENT_INTENTS_PATH = os.getenv("AGENT_INTENTS_PATH", "agent/intents")
OUTPUT_CSV         = os.getenv("OUTPUT_CSV", "training_phrases.csv")

# ── Intent whitelist — loaded from private config file ────────────────────────
_INTENTS_FILE = Path("config/include_intents.txt")

if not _INTENTS_FILE.exists():
    raise FileNotFoundError(
        f"\n[phrase_extraction] Intent filter file not found: {_INTENTS_FILE}\n"
        f"  Copy config/include_intents.example.txt → config/include_intents.txt\n"
        f"  and populate it with the intent names you want to extract."
    )

INCLUDE_INTENTS = {
    line.strip()
    for line in _INTENTS_FILE.read_text(encoding="utf-8").splitlines()
    if line.strip() and not line.strip().startswith("#")
}

# ── Extraction ────────────────────────────────────────────────────────────────
data = {}

for intent_folder in os.listdir(AGENT_INTENTS_PATH):
    intent_dir = os.path.join(AGENT_INTENTS_PATH, intent_folder)

    if not os.path.isdir(intent_dir):
        continue

    intent_name = intent_folder

    if intent_name not in INCLUDE_INTENTS:
        continue

    training_phrases_dir = os.path.join(intent_dir, "trainingPhrases")

    if not os.path.isdir(training_phrases_dir):
        continue

    for filename in os.listdir(training_phrases_dir):
        if not filename.endswith(".json"):
            continue

        file_path = os.path.join(training_phrases_dir, filename)

        with open(file_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # Dialogflow CX training phrase schema:
        # { "trainingPhrases": [ { "parts": [ { "text": "..." } ], ... } ] }
        for phrase_obj in json_data.get("trainingPhrases", []):
            phrase = "".join(
                part.get("text", "")
                for part in phrase_obj.get("parts", [])
            )
            if phrase:
                if intent_name not in data:
                    data[intent_name] = []
                data[intent_name].append(phrase)

# ── Write CSV ─────────────────────────────────────────────────────────────────
with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["utterance", "intent"])
    for intent, phrases in data.items():
        for phrase in phrases:
            writer.writerow([phrase, intent])

print(f"Training phrases extracted → {OUTPUT_CSV}")

# Flag intents with fewer than 4 phrases
intents_with_few_phrases = {
    intent: len(phrases)
    for intent, phrases in data.items()
    if len(phrases) < 4
}

for intent, count in intents_with_few_phrases.items():
    print(f'  ⚠ Intent "{intent}" has fewer than 5 phrases: {count}')

print(f"Total intents with fewer than 5 phrases: {len(intents_with_few_phrases)}")
