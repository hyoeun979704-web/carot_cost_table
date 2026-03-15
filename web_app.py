#!/usr/bin/env python3
"""
통신사 단가표 웹 서비스

사용법:
    uvicorn web_app:app --reload --port 8000
    # 또는
    python web_app.py
"""

import os
import uuid
import shutil
import tempfile
import threading
from pathlib import Path

from dotenv import load_dotenv
load_dotenv(dotenv_path=Path(__file__).parent / ".env")

import openpyxl
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from update_carrier_data import (
    analyze_image,
    update_excel,
    load_excel_model_rows,
    CARRIER_SHEET_MAP,
)

app = FastAPI(title="단가표 업데이트 서비스")

# ─────────────────────────────────────────────
# 작업 저장소 (in-memory)
# ─────────────────────────────────────────────

_jobs: dict[str, dict] = {}
# 구조: { job_id: { status, log: [], output_path, tmp_dir } }
# status: "pending" | "running" | "done" | "error"


# ─────────────────────────────────────────────
# 백그라운드 분석 작업
# ─────────────────────────────────────────────

def _run(job_id: str, image_paths: list[str], template_path: str) -> None:
    job = _jobs[job_id]
    job["status"] = "running"

    def log(msg: str) -> None:
        job["log"].append(msg)

    try:
        api_key = os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            raise RuntimeError(
                "GOOGLE_API_KEY가 설정되지 않았습니다.\n"
                ".env 파일에 GOOGLE_API_KEY=... 를 입력하세요."
            )

        os.environ["GOOGLE_API_KEY"] = api_key

        log(f"📂 템플릿 로드: {Path(template_path).name}")
        wb = openpyxl.load_workbook(template_path)
        model_rows = load_excel_model_rows(wb)
        log(f"  ▸ 입력용: {len(model_rows.get('입력용', {}))}개 모델")
        for sheet in CARRIER_SHEET_MAP.values():
            log(f"  ▸ {sheet}: {len(model_rows.get(sheet, {}))}개 모델")

        log(f"\n🔍 이미지 분석 ({len(image_paths)}개)...")
        all_extracted = []
        for path in image_paths:
            log(f"  → {Path(path).name}")
            data = analyze_image(path)
            all_extracted.append(data)
            carriers = list(data.get("carriers", {}).keys())
            log(f"    ✓ 통신사 감지: {carriers if carriers else '없음'}")

        log("\n📝 엑셀 업데이트 중...")
        change_log: list[str] = []
        total = update_excel(wb, all_extracted, model_rows, change_log)

        for entry in change_log:
            marker = "  ✓" if "→" in entry else "  ⚠"
            log(f"{marker} {entry}")

        output_path = str(Path(job["tmp_dir"]) / "updated.xlsx")
        wb.save(output_path)

        log(f"\n✅ 완료 — {total}개 셀 업데이트")
        job["output_path"] = output_path
        job["status"] = "done"

    except Exception as exc:
        log(f"\n❌ 오류: {exc}")
        job["status"] = "error"


# ─────────────────────────────────────────────
# API 엔드포인트
# ─────────────────────────────────────────────

@app.post("/api/analyze")
async def start_analysis(
    images: list[UploadFile] = File(..., description="단가표 이미지 (여러 장 가능)"),
    template: UploadFile = File(..., description="엑셀 템플릿 (.xlsx)"),
):
    """이미지와 템플릿을 업로드하면 백그라운드에서 분석을 시작합니다."""
    if not template.filename.endswith(".xlsx"):
        raise HTTPException(400, "템플릿은 .xlsx 파일이어야 합니다.")

    tmp_dir = tempfile.mkdtemp(prefix="carot_")
    job_id = str(uuid.uuid4())
    _jobs[job_id] = {"status": "pending", "log": [], "output_path": None, "tmp_dir": tmp_dir}

    # 파일 저장
    image_paths = []
    for img in images:
        dest = Path(tmp_dir) / img.filename
        dest.write_bytes(await img.read())
        image_paths.append(str(dest))

    tpl_path = str(Path(tmp_dir) / template.filename)
    Path(tpl_path).write_bytes(await template.read())

    threading.Thread(target=_run, args=(job_id, image_paths, tpl_path), daemon=True).start()
    return {"job_id": job_id}


@app.get("/api/status/{job_id}")
def get_status(job_id: str):
    """작업 상태 및 진행 로그를 반환합니다."""
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")
    return {"status": job["status"], "log": job["log"]}


@app.get("/api/download/{job_id}")
def download(job_id: str):
    """완료된 작업의 엑셀 파일을 다운로드합니다."""
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")
    if job["status"] != "done" or not job["output_path"]:
        raise HTTPException(400, "아직 완료되지 않은 작업입니다.")
    return FileResponse(
        job["output_path"],
        filename="updated.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.delete("/api/job/{job_id}")
def cleanup_job(job_id: str):
    """작업 임시 파일을 정리합니다."""
    job = _jobs.pop(job_id, None)
    if job and job.get("tmp_dir"):
        shutil.rmtree(job["tmp_dir"], ignore_errors=True)
    return {"ok": True}


# static 파일은 마지막에 마운트 (API 라우트보다 후순위)
app.mount("/", StaticFiles(directory="static", html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_app:app", host="0.0.0.0", port=8000, reload=True)
