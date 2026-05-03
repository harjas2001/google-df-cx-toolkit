"""
check_for_upper_lower_case.py
─────────────────────────────────────────────────────────────────────────────
Scans all intents in a Dialogflow CX exported agent and flags training phrases
that are written entirely in CAPITAL LETTERS.

All-caps training phrases are a common data quality issue — they cause the NLU
model to overfit to casing patterns rather than semantic content, and usually
indicate phrases that were copy-pasted from internal documentation or CRM exports.

Configuration:
  AGENT_BASE_PATH — base path of your exported agent.
  Loaded from .env (see .env.example). Expects df_cx_agent/intents inside.
─────────────────────────────────────────────────────────────────────────────
"""

import json
import os
from collections import defaultdict
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
AGENT_BASE_PATH = os.getenv("AGENT_BASE_PATH", ".")


def extract_training_phrases_from_json(json_file_path):
    """
    Extract training phrases from a Dialogflow CX training phrases JSON file.
    Joins all 'text' parts from each phrase's 'parts' array.
    """
    training_phrases = []

    try:
        with open(json_file_path, "r", encoding="utf-8") as file:
            data = json.load(file)

        if "trainingPhrases" in data:
            for phrase_group in data["trainingPhrases"]:
                if "parts" in phrase_group:
                    complete_phrase = "".join(
                        part["text"]
                        for part in phrase_group["parts"]
                        if "text" in part and part["text"]
                    )
                    if complete_phrase.strip():
                        training_phrases.append(complete_phrase.strip())

    except (json.JSONDecodeError, FileNotFoundError, KeyError) as e:
        print(f"Error processing {json_file_path}: {e}")

    return training_phrases


def is_all_caps(text):
    """
    Returns True if all alphabetic characters in the text are uppercase.
    Ignores numbers, punctuation, and whitespace.
    """
    alphabetic = [c for c in text if c.isalpha()]
    return len(alphabetic) > 0 and all(c.isupper() for c in alphabetic)


def extract_all_training_phrases(base_folder_path):
    """
    Walk all intent folders in the agent export and extract training phrases.
    Returns {intent_name: [phrase, ...]}
    """
    intents_folder = os.path.join(base_folder_path, "df_cx_agent", "intents")
    all_training_phrases = {}

    if not os.path.exists(intents_folder):
        print(f"Intents folder not found: {intents_folder}")
        return all_training_phrases

    for intent_folder in os.listdir(intents_folder):
        intent_path = os.path.join(intents_folder, intent_folder)

        if os.path.isdir(intent_path):
            training_phrases_folder = os.path.join(intent_path, "trainingPhrases")

            if os.path.exists(training_phrases_folder):
                en_json_path = os.path.join(training_phrases_folder, "en.json")

                if os.path.exists(en_json_path):
                    phrases = extract_training_phrases_from_json(en_json_path)
                    if phrases:
                        all_training_phrases[intent_folder] = phrases
                        print(f"  {intent_folder}: {len(phrases)} phrases")
                else:
                    print(f"  en.json not found in: {training_phrases_folder}")
            else:
                print(f"  trainingPhrases folder not found in: {intent_path}")

    return all_training_phrases


def find_capitalized_training_phrases(all_training_phrases):
    """Return all-caps training phrases grouped by intent."""
    capitalized = defaultdict(list)

    for intent_name, phrases in all_training_phrases.items():
        for phrase in phrases:
            if is_all_caps(phrase):
                capitalized[intent_name].append(phrase)

    return capitalized


def main():
    print(f"Scanning agent at: {AGENT_BASE_PATH}\n")
    all_training_phrases = extract_all_training_phrases(AGENT_BASE_PATH)

    if not all_training_phrases:
        print("No training phrases found. Check the agent folder structure.")
        return

    print(f"\nFound training phrases in {len(all_training_phrases)} intents.")

    capitalized = find_capitalized_training_phrases(all_training_phrases)

    if capitalized:
        print("\n" + "=" * 60)
        print("TRAINING PHRASES WITH ALL CAPITAL LETTERS")
        print("=" * 60)

        for intent_name, phrases in capitalized.items():
            print(f"\nIntent: {intent_name}")
            print("-" * 40)
            for phrase in phrases:
                print(f"  • {phrase}")

        total = sum(len(p) for p in capitalized.values())
        print(f"\n{'='*60}")
        print(f"SUMMARY: {total} all-caps phrases across {len(capitalized)} intents.")
        print("=" * 60)
    else:
        print("\nNo all-caps training phrases found.")


if __name__ == "__main__":
    main()
