
import argparse
import json
from pathlib import Path
from datetime import datetime

def create_project_structure(base_path: Path, project_name: str, meta: dict):
    (base_path / "src" / project_name).mkdir(parents=True, exist_ok=True)
    (base_path / "tests").mkdir(parents=True, exist_ok=True)
    (base_path / ".venv").mkdir(exist_ok=True)
    (base_path / "README.md").write_text(f"# {project_name}\n\nProject auto-generated.\n")
    (base_path / ".gitignore").write_text(".venv/\n__pycache__/\n.env\n")
    (base_path / "requirements.txt").write_text("fastapi\nuvicorn\n")
    (base_path / ".projectmeta.json").write_text(json.dumps(meta, indent=4))

def main():
    parser = argparse.ArgumentParser(description="Initialize a new dual-instance project.")
    parser.add_argument("name", help="Project name (use kebab-case)")
    parser.add_argument("--type", default="backend", help="Project type (backend/gui/cli)")
    parser.add_argument("--lang", default="python", help="Programming language")
    parser.add_argument("--status", default="Planning", help="Project status (e.g., Planning, Active, etc.)")
    parser.add_argument("--tags", default="", help="Comma-separated list of tags")
    parser.add_argument("--archive", default="C:/Users/damian/projects/tests/simulate_d/", help="Path to archive root")
    parser.add_argument("--dev", required=True, help="Path to dev project root")
    args = parser.parse_args()

    tags = [tag.strip() for tag in args.tags.split(",") if tag.strip()]
    archive_path = Path(args.archive) / args.name
    dev_path = Path(args.dev) / args.name

    meta = {
        "project_name": args.name,
        "created": datetime.now().isoformat(),
        "archive_path": str(archive_path),
        "dev_path": str(dev_path),
        "type": args.type,
        "lang": args.lang,
        "status": args.status,
        "tags": tags,
        "synced": datetime.now().isoformat()
    }

    create_project_structure(archive_path, args.name, meta)
    create_project_structure(dev_path, args.name, meta)
    print(f"Project '{args.name}' created in archive and dev locations.")


if __name__ == "__main__":
    main()
