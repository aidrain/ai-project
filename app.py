from __future__ import annotations

import io
import logging
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file

import db
import tools

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)

LOGO_PATH = Path(__file__).parent / "assets" / "suwon_logo.png"

# 업로드/편집 중인 데이터셋을 메모리에만 보관 (디스크에 임시 저장하지 않음)
STATE: dict[str, Any] = {
    "df_2026": None,
    "meta": {"college": "경영공학대학", "major": "경제금융"},
    "filename": "",
    "changes": [],
}


def _persist_state() -> None:
    """현재 작업 상태를 SQLite에 자동 저장합니다."""
    if STATE["df_2026"] is not None and STATE["filename"]:
        db.save_workspace(STATE["filename"], STATE["meta"], STATE["df_2026"], STATE["changes"])


def _init_default_dataset() -> None:
    """DB에 저장된 최근 작업이 있으면 복원하고, 없으면 기본 샘플 파일을 불러옵니다."""
    db.init_db()
    latest = db.load_latest_workspace()
    if latest is not None:
        STATE["df_2026"] = latest["df_2026"]
        STATE["meta"] = latest["meta"]
        STATE["filename"] = latest["filename"]
        STATE["changes"] = latest["changes"]
        logger.info("DB에서 최근 작업 복원 완료: %s (변경 %d건)", STATE["filename"], len(STATE["changes"]))
        return

    if tools.DEFAULT_EXCEL_PATH.exists():
        df, meta = tools.parse_curriculum_excel(tools.DEFAULT_EXCEL_PATH)
        STATE["df_2026"] = df
        STATE["meta"] = meta
        STATE["filename"] = tools.DEFAULT_EXCEL_PATH.name
        STATE["changes"] = []
        _persist_state()
        logger.info("기본 샘플 파일 로드 완료: %s (%d개 과목)", STATE["filename"], len(df))


def _build_state_payload() -> dict[str, Any]:
    """프론트엔드 화면 갱신에 필요한 현재 상태 JSON을 만듭니다."""
    df_2026 = STATE["df_2026"]
    if df_2026 is None:
        return {"loaded": False, "filename": "", "meta": STATE["meta"]}

    df_2027, revision_df = tools.apply_changes(df_2026, STATE["changes"])
    metrics = tools.get_summary_metrics(df_2026, df_2027, revision_df)

    categories = []
    for cat in df_2027["과목구분"].unique():
        sub_df = df_2027[df_2027["과목구분"] == cat]
        categories.append(
            {
                "name": cat,
                "count": len(sub_df),
                "subtotal": tools.sum_credits(sub_df),
                "rows": sub_df.to_dict(orient="records"),
            }
        )

    return {
        "loaded": True,
        "filename": STATE["filename"],
        "meta": STATE["meta"],
        "metrics": metrics,
        "courses_2027": df_2027.to_dict(orient="records"),
        "revision_rows": revision_df.to_dict(orient="records"),
        "categories": categories,
        "saved_workspaces": db.list_workspaces(),
    }


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.get("/logo")
def logo() -> Any:
    if LOGO_PATH.exists():
        return send_file(LOGO_PATH, mimetype="image/png")
    return ("", 404)


@app.get("/api/state")
def get_state() -> Any:
    return jsonify(_build_state_payload())


@app.post("/api/upload")
def upload_excel() -> Any:
    file = request.files.get("file")
    if file is None or not file.filename:
        return jsonify({"error": "업로드할 엑셀 파일(.xlsx)을 선택해 주세요."}), 400
    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({"error": ".xlsx 형식의 엑셀 파일만 업로드할 수 있습니다."}), 400

    try:
        buf = io.BytesIO(file.read())
        df, meta = tools.parse_curriculum_excel(buf)
        if df.empty:
            return jsonify({"error": "표에서 교과목 데이터를 찾지 못했습니다."}), 400
    except Exception as exc:
        logger.exception("엑셀 파싱 실패")
        return jsonify({"error": f"엑셀 파일을 읽는 중 오류가 발생했습니다: {exc}"}), 400

    STATE["df_2026"] = df
    STATE["meta"] = meta
    STATE["filename"] = file.filename
    STATE["changes"] = []
    _persist_state()
    logger.info("업로드 완료: %s (%d개 과목)", file.filename, len(df))
    return jsonify(_build_state_payload())


@app.post("/api/change/add")
def add_course() -> Any:
    if STATE["df_2026"] is None:
        return jsonify({"error": "먼저 2026학년도 엑셀 파일을 불러와 주세요."}), 400

    data = request.get_json(silent=True) or {}
    name = str(data.get("교과목명", "")).strip()
    term = str(data.get("학년-학기", "")).strip()
    if not name:
        return jsonify({"error": "교과목명을 입력해 주세요."}), 400
    if not term:
        return jsonify({"error": "학년-학기를 입력해 주세요."}), 400

    credit = int(data.get("학점", 3))
    theory = int(data.get("이론", 3))
    practice = int(data.get("실습", 0))

    STATE["changes"].append(
        {
            "유형": "신설",
            "과목구분": str(data.get("과목구분", "전공선택")).strip(),
            "전공": STATE["meta"]["major"],
            "학수번호": "",
            "교과목명": name,
            "학년-학기": term,
            "학점-이론-실습": f"{credit}-{theory}-{practice}",
        }
    )
    _persist_state()
    return jsonify(_build_state_payload())


@app.post("/api/change/edit")
def edit_course() -> Any:
    if STATE["df_2026"] is None:
        return jsonify({"error": "먼저 2026학년도 엑셀 파일을 불러와 주세요."}), 400

    data = request.get_json(silent=True) or {}
    target = str(data.get("대상과목명", "")).strip()
    df_2027, _ = tools.apply_changes(STATE["df_2026"], STATE["changes"])
    matched = df_2027[df_2027["교과목명"] == target]
    if matched.empty:
        return jsonify({"error": "변경할 과목을 찾을 수 없습니다."}), 404

    cur = matched.iloc[0]
    new_cat = str(data.get("새_과목구분", cur["과목구분"])).strip()
    new_name = str(data.get("새_교과목명", cur["교과목명"])).strip()
    new_term = str(data.get("새_학년학기", cur["학년-학기"])).strip()
    new_ctp = f"{int(data.get('학점', 0))}-{int(data.get('이론', 0))}-{int(data.get('실습', 0))}"

    if not new_name:
        return jsonify({"error": "교과목명을 입력해 주세요."}), 400

    if (
        new_cat == cur["과목구분"]
        and new_name == cur["교과목명"]
        and new_term == cur["학년-학기"]
        and new_ctp == cur["학점-이론-실습"]
    ):
        return jsonify({"error": "변경된 내용이 없습니다."}), 400

    STATE["changes"].append(
        {
            "유형": "변경",
            "과목구분": cur["과목구분"],
            "기존_과목구분": cur["과목구분"],
            "새_과목구분": new_cat,
            "대상과목명": target,
            "기존_교과목명": cur["교과목명"],
            "기존_학년학기": cur["학년-학기"],
            "기존_학점이론실습": cur["학점-이론-실습"],
            "새_교과목명": new_name,
            "새_학년학기": new_term,
            "새_학점이론실습": new_ctp,
        }
    )
    _persist_state()
    return jsonify(_build_state_payload())


@app.post("/api/change/delete")
def delete_course() -> Any:
    if STATE["df_2026"] is None:
        return jsonify({"error": "먼저 2026학년도 엑셀 파일을 불러와 주세요."}), 400

    data = request.get_json(silent=True) or {}
    target = str(data.get("대상과목명", "")).strip()
    df_2027, _ = tools.apply_changes(STATE["df_2026"], STATE["changes"])
    matched = df_2027[df_2027["교과목명"] == target]
    if matched.empty:
        return jsonify({"error": "폐설할 과목을 찾을 수 없습니다."}), 404

    cur = matched.iloc[0]
    STATE["changes"].append(
        {
            "유형": "폐지",
            "과목구분": cur["과목구분"],
            "대상과목명": target,
            "기존_학년학기": cur["학년-학기"],
            "기존_학점이론실습": cur["학점-이론-실습"],
        }
    )
    _persist_state()
    return jsonify(_build_state_payload())


@app.post("/api/change/undo")
def undo_change() -> Any:
    if STATE["changes"]:
        STATE["changes"].pop()
        _persist_state()
    return jsonify(_build_state_payload())


@app.post("/api/change/reset")
def reset_changes() -> Any:
    STATE["changes"] = []
    _persist_state()
    return jsonify(_build_state_payload())


@app.post("/api/workspace/load")
def load_saved_workspace() -> Any:
    data = request.get_json(silent=True) or {}
    ws_id = data.get("id")
    if ws_id is None:
        return jsonify({"error": "불러올 작업을 선택해 주세요."}), 400
    ws_data = db.load_workspace(int(ws_id))
    if ws_data is None:
        return jsonify({"error": "저장된 작업을 찾을 수 없습니다."}), 404

    STATE["df_2026"] = ws_data["df_2026"]
    STATE["meta"] = ws_data["meta"]
    STATE["filename"] = ws_data["filename"]
    STATE["changes"] = ws_data["changes"]
    logger.info("저장된 작업 불러오기: %s", STATE["filename"])
    return jsonify(_build_state_payload())


@app.post("/api/workspace/delete")
def delete_saved_workspace() -> Any:
    data = request.get_json(silent=True) or {}
    ws_id = data.get("id")
    if ws_id is None:
        return jsonify({"error": "삭제할 작업을 선택해 주세요."}), 400
    db.delete_workspace(int(ws_id))
    return jsonify(_build_state_payload())


@app.get("/api/download")
def download_excel() -> Any:
    if STATE["df_2026"] is None:
        return jsonify({"error": "먼저 2026학년도 엑셀 파일을 불러와 주세요."}), 400

    df_2027, revision_df = tools.apply_changes(STATE["df_2026"], STATE["changes"])
    excel_bytes = tools.build_2027_excel_bytes(df_2027, revision_df, STATE["meta"])
    major = STATE["meta"].get("major", "전공")
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"2027_전공교육과정_{major}.xlsx",
    )


_init_default_dataset()

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
