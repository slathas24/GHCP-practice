from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Dict, List, Tuple

import openpyxl


HIGH_KEYWORDS = {"down", "outage", "critical", "failure", "unavailable", "emergency", "panic"}
MEDIUM_KEYWORDS = {"slow", "latency", "intermittent", "degraded", "issue", "performance", "delay"}
LOW_KEYWORDS = {"typo", "cosmetic", "minor", "info", "warning", "notification"}


def categorize_incident(description: str) -> str:
    text = (description or "").lower()
    if any(keyword in text for keyword in HIGH_KEYWORDS):
        return "high"
    if any(keyword in text for keyword in MEDIUM_KEYWORDS):
        return "medium"
    if any(keyword in text for keyword in LOW_KEYWORDS):
        return "low"
    if "security" in text or "data" in text:
        return "high"
    return "medium"


def extract_keywords(description: str) -> List[str]:
    text = (description or "").lower()
    words = [word.strip(" ,.-") for word in text.split()]
    return [word for word in words if len(word) > 3]


def read_incidents(input_dir: Path) -> List[Dict[str, str]]:
    for candidate in ["incidents.csv", "incidents.json", "incidents.txt"]:
        path = input_dir / candidate
        if path.exists():
            if path.suffix == ".csv":
                with path.open("r", encoding="utf-8", newline="") as handle:
                    reader = csv.DictReader(handle)
                    return [
                        {
                            "incident_id": row.get("incident_id", ""),
                            "title": row.get("title", ""),
                            "description": row.get("description", ""),
                        }
                        for row in reader
                    ]
            if path.suffix == ".json":
                data = json.loads(path.read_text(encoding="utf-8"))
                return data if isinstance(data, list) else list(data.values())
            with path.open("r", encoding="utf-8") as handle:
                return [
                    {
                        "incident_id": line.strip(),
                        "title": line.strip(),
                        "description": line.strip(),
                    }
                    for line in handle if line.strip()
                ]
    return []


def read_engineers(input_dir: Path) -> List[Dict[str, object]]:
    workbook_path = input_dir / "Service_Engg.xlsx"
    if not workbook_path.exists():
        return []

    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []

    header = [str(cell).strip().lower() if cell is not None else "" for cell in rows[0]]
    engineers: List[Dict[str, object]] = []
    for row in rows[1:]:
        if not any(cell is not None and str(cell).strip() for cell in row):
            continue
        values = {header[index]: row[index] if index < len(row) else "" for index in range(len(header))}
        engineers.append(
            {
                "engineer_name": values.get("engineer_name", ""),
                "experience_years": int(values.get("experience_years", 0) or 0),
                "specialties": [str(item).strip().lower() for item in str(values.get("specialties", "")).split(",") if str(item).strip()],
                "current_queue_count": int(values.get("current_queue_count", 0) or 0),
            }
        )
    return engineers


def select_engineer(incident: Dict[str, str], engineers: List[Dict[str, object]]) -> Tuple[Dict[str, object], str]:
    category = categorize_incident(incident.get("description", ""))
    keywords = extract_keywords(incident.get("description", ""))
    ranked: List[Tuple[Tuple[int, int, int, str], Dict[str, object]]] = []
    for engineer in engineers:
        specialties = set(engineer.get("specialties", []))
        keyword_matches = sum(1 for keyword in keywords if keyword in specialties)
        queue_count = int(engineer.get("current_queue_count", 0))
        experience_years = int(engineer.get("experience_years", 0))
        score = (queue_count, -experience_years, -keyword_matches, str(engineer.get("engineer_name", "")))
        ranked.append((score, engineer))

    ranked.sort(key=lambda item: item[0])
    if not ranked:
        return {}, "No engineer available"

    chosen = ranked[0][1]
    reason = (
        f"Category {category}; queue {chosen.get('current_queue_count')}; experience {chosen.get('experience_years')} years"
    )
    return chosen, reason


def run_agent(input_dir: Path | str, output_dir: Path | str) -> Path:
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    incidents = read_incidents(input_path)
    engineers = read_engineers(input_path)

    output_file = output_path / "allocated_incidents.csv"
    with output_file.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "incident_id",
            "title",
            "description",
            "category",
            "assigned_engineer",
            "assigned_engineer_experience",
            "assignment_reason",
        ])
        for incident in incidents:
            category = categorize_incident(incident.get("description", ""))
            engineer, reason = select_engineer(incident, engineers)
            writer.writerow(
                [
                    incident.get("incident_id", ""),
                    incident.get("title", ""),
                    incident.get("description", ""),
                    category,
                    engineer.get("engineer_name", ""),
                    engineer.get("experience_years", ""),
                    reason,
                ]
            )

    return output_file


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Assign incidents to service engineers")
    parser.add_argument("--input", default="input", help="Folder containing incidents and the engineer workbook")
    parser.add_argument("--output", default="output", help="Folder where the allocation report is written")
    args = parser.parse_args()
    output_path = run_agent(args.input, args.output)
    print(f"Wrote allocation report to {output_path}")
