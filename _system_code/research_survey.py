from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path

try:  # Use a bundled copy first when this folder is distributed with a QMD.
    from . import research_survey_workbook as workbook
except ImportError:  # pragma: no cover - repo-local fallback
    try:
        from analytics import research_survey_workbook as workbook
    except ImportError:  # pragma: no cover - assignment folder fallback
        import research_survey_workbook as workbook


DEFAULT_ASSIGNMENT_CODE = "M01_A_emplo"
DEFAULT_OUTPUT_DIR = "research_submissions"


def assignment_code() -> str:
    workbook.load_dotenv_file()
    return os.environ.get("RESEARCH_ASSIGNMENT_CODE", DEFAULT_ASSIGNMENT_CODE).strip() or DEFAULT_ASSIGNMENT_CODE


def research_dir() -> Path:
    workbook.load_dotenv_file()
    return Path(os.environ.get("RESEARCH_SUBMISSION_DIR", DEFAULT_OUTPUT_DIR))


def print_runtime_summary() -> dict:
    report = workbook.runtime_report()
    print("Operating System:", report.get("operating_system", ""))
    print("Python Version:", report.get("python_version", ""))
    print("Machine:", report.get("machine", ""))
    print("Processor:", report.get("processor", ""))
    if "memory_total_mb" in report:
        print("Total Memory (MB):", report["memory_total_mb"])
    if "memory_available_mb" in report:
        print("Available Memory (MB):", report["memory_available_mb"])
    return report


def prepare_workbook() -> Path:
    code = assignment_code()
    output_dir = research_dir()
    output_dir.mkdir(parents=True, exist_ok=True)

    existing = existing_workbook_path(code, output_dir)
    if existing:
        print(f"Using existing workbook: {existing}")
        return existing

    survey, survey_url = workbook.fetch_survey(code)
    workbook_path = timestamped_workbook_path(code, output_dir)
    workbook.create_workbook(survey, workbook_path, force=False)
    pointer_path(code, output_dir).write_text(str(workbook_path), encoding="utf-8")
    print(f"Created workbook: {workbook_path}")
    print(f"Survey source: {survey_url}")
    return workbook_path


def validate_workbook(workbook_path: str | Path | None = None) -> Path:
    code = assignment_code()
    output_dir = research_dir()
    resolved_workbook = Path(workbook_path) if workbook_path else existing_workbook_path(code, output_dir)
    if not resolved_workbook:
        raise FileNotFoundError("No research survey workbook found. Run prepare_workbook() first.")
    if not resolved_workbook.exists():
        raise FileNotFoundError(f"Research survey workbook not found: {resolved_workbook}")

    output_json = resolved_workbook.with_name(f"{resolved_workbook.stem}_validation.json")
    notebook_run_id = os.environ.get("NOTEBOOK_RUN_ID", resolved_workbook.stem)
    workbook.validate_workbook(
        resolved_workbook,
        assignment_code=code,
        notebook_run_id=notebook_run_id,
        output_json=output_json,
    )
    print("Workbook validation passed. No data was submitted to the course API or Django database.")
    print(f"Local validation report: {output_json}")
    return output_json


def existing_workbook_path(code: str | None = None, output_dir: str | Path | None = None) -> Path | None:
    code = code or assignment_code()
    output_dir = Path(output_dir) if output_dir else research_dir()
    pointer = pointer_path(code, output_dir)
    if pointer.exists():
        candidate = Path(pointer.read_text(encoding="utf-8").strip())
        if candidate.exists():
            return candidate

    candidates = sorted(output_dir.glob(f"{safe_slug(code)}_response_workbook_*.xlsx"))
    if candidates:
        latest = candidates[-1]
        pointer.write_text(str(latest), encoding="utf-8")
        return latest
    return None


def timestamped_workbook_path(code: str | None = None, output_dir: str | Path | None = None) -> Path:
    code = code or assignment_code()
    output_dir = Path(output_dir) if output_dir else research_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_dir / f"{safe_slug(code)}_response_workbook_{timestamp}.xlsx"


def pointer_path(code: str, output_dir: Path) -> Path:
    return output_dir / f"{safe_slug(code)}_current_workbook.txt"


def safe_slug(value: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip())
    return slug.strip("._-") or DEFAULT_ASSIGNMENT_CODE
