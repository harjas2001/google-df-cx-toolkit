# Private config — never commit
config/include_intents.txt

# Environment
.env
*.env.local

# Agent exports — real agent data should never be committed
df_cx_agent/
*.zip

# Output files
*.xlsx
*.csv
!sample_data/

# Python
__pycache__/
*.py[cod]
*.pyo
*.pyd
*.egg-info/
dist/
build/

# Virtual environments
venv/
.venv/
env/

# IDE
.vscode/
.idea/
*.swp
*.swo
.DS_Store
Thumbs.db

# Logs
*.log
logs/
