import json
from pathlib import Path

SCHEMA_FILE = Path("schemas.json")

# Agar yo‘q bo‘lsa — yaratamiz
if not SCHEMA_FILE.exists():
    SCHEMA_FILE.write_text("{}", encoding="utf-8")


def load_schemas() -> dict:
    return json.loads(SCHEMA_FILE.read_text(encoding="utf-8"))


def save_schemas(schemas: dict):
    SCHEMA_FILE.write_text(
        json.dumps(schemas, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )


# 🔁 Server ishga tushganda yuklanadi
schemas: dict = load_schemas()
