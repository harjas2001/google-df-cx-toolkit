"""
lastPage_list.py
─────────────────────────────────────────────────────────────────────────────
Iterates through a Dialogflow CX exported agent config and prints a complete
list of all pages grouped by flow.

Useful for auditing flow structure, mapping last pages, and spotting orphaned
or unreachable pages before a deployment.

Configuration:
  AGENT_FLOWS_PATH — path to the flows folder in your agent export.
  Loaded from .env (see .env.example).
─────────────────────────────────────────────────────────────────────────────
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
base_dir = os.getenv("AGENT_FLOWS_PATH", "df_cx_agent/flows")

# ── Page extraction ───────────────────────────────────────────────────────────
flow_pages = {}

for flow_name in os.listdir(base_dir):
    flow_path = os.path.join(base_dir, flow_name)

    if os.path.isdir(flow_path):
        pages_path = os.path.join(flow_path, "pages")

        if os.path.isdir(pages_path):
            json_files = [f for f in os.listdir(pages_path) if f.endswith(".json")]
            page_names = [os.path.splitext(f)[0] for f in json_files]
            flow_pages[flow_name] = page_names

# ── Output ────────────────────────────────────────────────────────────────────
for flow, pages in flow_pages.items():
    print(f"\n{flow}")
    print("-" * len(flow))
    for page in pages:
        print(page)
