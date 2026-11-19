from __future__ import annotations

import io
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

DEFAULT_EXCEL_PATH = Path(__file__).parent / "sample-data" / "2026_전공교육과정_경제금융.xlsx"


def format_course_code(value: Any) -> str:
    """학수번호 셀 값을 깔끔한 문자열로 변환합니다."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_credit_triplet(text: str) -> tuple[int, int, int]:
    """'3-3-0' 형태의 문자열을 (학점, 이론, 실습) 정수로 분리합니다."""
    parts = [p.strip() for p in str(text).split("-")]
    if len(parts) != 3:
        return (0, 0, 0)
    try:
        return (int(parts[0]), int(parts[1]), int(parts[2]))
    except ValueError:
        return (0, 0, 0)


def sum_credits(df: pd.DataFrame) -> str:
    """DataFrame의 '학점-이론-실습' 열을 합산해 'N-N-N' 문자열로 반환합니다."""
    if df.empty:
        return "0-0-0"
    c_sum = t_sum = p_sum = 0
    for val in df["학점-이론-실습"]:
        c, t, p = parse_credit_triplet(val)
        c_sum += c
        t_sum += t
        p_sum += p
    return f"{c_sum}-{t_sum}-{p_sum}"


def parse_curriculum_excel(file_source: Any) -> tuple[pd.DataFrame, dict[str, str]]:
    """2026학년도 전공 교육과정 엑셀 파일의 6~41행 편성표를 읽어옵니다."""
    wb = openpyxl.load_workbook(file_source, data_only=True)
    ws = wb[wb.sheetnames[0]]

    college = str(ws.cell(4, 2).value or "경영공학대학").strip()
    major = str(ws.cell(7, 5).value or "경제금융").strip()

    rows: list[dict[str, str]] = []
    current_category = ""

    for r in range(7, ws.max_row + 1):
        cat_val = ws.cell(r, 2).value
        if cat_val is not None and str(cat_val).strip():
            current_category = str(cat_val).strip()

        col_e = str(ws.cell(r, 5).value or "").strip()
        course_name = ws.cell(r, 8).value

        # 42행 이후 빈 줄이나 교과목 해설 영역에 도달하면 편성표 읽기 종료
        if r > 41 and course_name is None and col_e != "소 계":
            break

        if col_e == "소 계" or course_name is None:
            continue

        rows.append(
            {
                "과목구분": current_category,
                "전공": major,
                "학수번호": format_course_code(ws.cell(r, 6).value),
                "교과목명": str(course_name).strip(),
                "학년-학기": str(ws.cell(r, 9).value or "").strip(),
                "학점-이론-실습": str(ws.cell(r, 10).value or "0-0-0").strip(),
            }
        )

    df = pd.DataFrame(rows)
    meta = {"college": college, "major": major}
    return df, meta


def apply_changes(
    df_base: pd.DataFrame, changes: list[dict[str, Any]]
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """개정 내역(신설/변경/폐지)을 26학년도 표에 반영해 27학년도 표와 개정표를 만듭니다."""
    df_2027 = df_base.copy()
    revision_rows: list[dict[str, str]] = []

    for item in changes:
        action = item["유형"]
        if action == "신설":
            new_row = {
                "과목구분": item["과목구분"],
                "전공": item["전공"],
                "학수번호": item["학수번호"],
                "교과목명": item["교과목명"],
                "학년-학기": item["학년-학기"],
                "학점-이론-실습": item["학점-이론-실습"],
            }
            df_2027 = pd.concat([df_2027, pd.DataFrame([new_row])], ignore_index=True)
            revision_rows.append(
                {
                    "구분": "신설",
                    "과목구분": item["과목구분"],
                    "변경 전 (2026)": "-",
                    "변경 후 (2027)": f"[{item['과목구분']}] {item['교과목명']} ({item['학년-학기']}, {item['학점-이론-실습']})",
                }
            )
        elif action == "변경":
            target_name = item["대상과목명"]
            old_cat = item.get("기존_과목구분", item["과목구분"])
            new_cat = item.get("새_과목구분", item["과목구분"])
            mask = df_2027["교과목명"] == target_name
            if mask.any():
                df_2027.loc[mask, "과목구분"] = new_cat
                df_2027.loc[mask, "교과목명"] = item["새_교과목명"]
                df_2027.loc[mask, "학년-학기"] = item["새_학년학기"]
                df_2027.loc[mask, "학점-이론-실습"] = item["새_학점이론실습"]

            cat_summary = new_cat if old_cat == new_cat else f"{old_cat} → {new_cat}"
            revision_rows.append(
                {
                    "구분": "변경",
                    "과목구분": cat_summary,
                    "변경 전 (2026)": f"[{old_cat}] {item['기존_교과목명']} ({item['기존_학년학기']}, {item['기존_학점이론실습']})",
                    "변경 후 (2027)": f"[{new_cat}] {item['새_교과목명']} ({item['새_학년학기']}, {item['새_학점이론실습']})",
                }
            )
        elif action == "폐지":
            target_name = item["대상과목명"]
            df_2027 = df_2027[df_2027["교과목명"] != target_name].reset_index(drop=True)
            revision_rows.append(
                {
                    "구분": "폐지",
                    "과목구분": item["과목구분"],
                    "변경 전 (2026)": f"[{item['과목구분']}] {item['대상과목명']} ({item['기존_학년학기']}, {item['기존_학점이론실습']})",
                    "변경 후 (2027)": "-",
                }
            )

    if not df_2027.empty:
        category_order = {"전공핵심": 0, "전공선택": 1}
        df_2027["_cat_order"] = df_2027["과목구분"].map(lambda x: category_order.get(x, 9))
        df_2027 = (
            df_2027.sort_values(by=["_cat_order", "학년-학기"], kind="stable")
            .drop(columns=["_cat_order"])
            .reset_index(drop=True)
        )

    revision_df = pd.DataFrame(
        revision_rows, columns=["구분", "과목구분", "변경 전 (2026)", "변경 후 (2027)"]
    )
    return df_2027, revision_df


def get_summary_metrics(
    df_2026: pd.DataFrame, df_2027: pd.DataFrame, revision_df: pd.DataFrame
) -> dict[str, Any]:
    """2026년 대비 2027년 과목 수, 변동 내역, 학점-이론-실습 합계를 요약합니다."""
    rev_counts = revision_df["구분"].value_counts().to_dict() if not revision_df.empty else {}
    add_cnt = int(rev_counts.get("신설", 0))
    edit_cnt = int(rev_counts.get("변경", 0))
    del_cnt = int(rev_counts.get("폐지", 0))
    diff_cnt = len(df_2027) - len(df_2026)
    diff_sign = f"+{diff_cnt}" if diff_cnt > 0 else str(diff_cnt)

    delta_text = (
        "2026년 대비 변동 없음"
        if revision_df.empty
        else f"{diff_sign}개 (신설 {add_cnt} · 변경 {edit_cnt} · 폐지 {del_cnt})"
    )

    return {
        "count_2026": len(df_2026),
        "count_2027": len(df_2027),
        "diff_count": diff_cnt,
        "delta_text": delta_text,
        "add_count": add_cnt,
        "edit_count": edit_cnt,
        "delete_count": del_cnt,
        "credits_2026": sum_credits(df_2026),
        "credits_2027": sum_credits(df_2027),
    }


def build_2027_excel_bytes(
    df_2027: pd.DataFrame, revision_df: pd.DataFrame, meta: dict[str, str]
) -> bytes:
    """2026년 원본과 동일한 표 형태(병합/소계 포함)로 2027년 엑셀 파일을 생성합니다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet_01"

    thin = Side(style="thin", color="808080")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    font_title = Font(name="맑은 고딕", size=14, bold=True)
    font_header = Font(name="맑은 고딕", size=10, bold=True)
    font_body = Font(name="맑은 고딕", size=10)
    fill_header = PatternFill("solid", fgColor="E9EFF7")
    fill_subtotal = PatternFill("solid", fgColor="F5F7FA")

    ws.merge_cells("B2:J2")
    ws["B2"] = "2027학년도 전공 교육과정"
    ws["B2"].font = font_title
    ws["B2"].alignment = align_center

    ws.merge_cells("B4:F4")
    ws["B4"] = meta.get("college", "경영공학대학")
    ws["B4"].font = font_header
    ws["B4"].alignment = align_left

    ws.merge_cells("B6:E6")
    ws["B6"] = "과목구분"
    ws.merge_cells("F6:G6")
    ws["F6"] = "학수번호"
    ws["H6"] = "교과목명"
    ws["I6"] = "학년-학기"
    ws["J6"] = "학점-이론-실습"

    for col in range(2, 11):
        cell = ws.cell(6, col)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all

    row_idx = 7
    major = meta.get("major", "경제금융")
    categories = [c for c in ["전공핵심", "전공선택"] if c in df_2027["과목구분"].values]
    extra_cats = [c for c in df_2027["과목구분"].unique() if c not in categories]

    for cat in categories + extra_cats:
        sub_df = df_2027[df_2027["과목구분"] == cat]
        start_row = row_idx

        for _, row in sub_df.iterrows():
            ws.merge_cells(start_row=row_idx, start_column=6, end_row=row_idx, end_column=7)
            ws.cell(row_idx, 6, row["학수번호"])
            ws.cell(row_idx, 8, row["교과목명"])
            ws.cell(row_idx, 9, row["학년-학기"])
            ws.cell(row_idx, 10, row["학점-이론-실습"])

            for col in range(2, 11):
                c = ws.cell(row_idx, col)
                c.font = font_body
                c.border = border_all
                c.alignment = align_left if col == 8 else align_center
            row_idx += 1

        subtotal_row = row_idx
        ws.merge_cells(start_row=subtotal_row, start_column=5, end_row=subtotal_row, end_column=9)
        ws.cell(subtotal_row, 5, "소 계")
        ws.cell(subtotal_row, 10, sum_credits(sub_df))
        for col in range(2, 11):
            c = ws.cell(subtotal_row, col)
            c.font = font_header
            c.fill = fill_subtotal
            c.border = border_all
            c.alignment = align_center

        ws.merge_cells(start_row=start_row, start_column=2, end_row=subtotal_row, end_column=4)
        ws.cell(start_row, 2, cat).alignment = align_center
        ws.cell(start_row, 2).font = font_header

        if subtotal_row - 1 >= start_row:
            ws.merge_cells(start_row=start_row, start_column=5, end_row=subtotal_row - 1, end_column=5)
            ws.cell(start_row, 5, major).alignment = align_center

        row_idx += 1

    col_widths = {2: 5, 3: 5, 4: 5, 5: 12, 6: 7, 7: 7, 8: 28, 9: 12, 10: 16}
    for col, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    ws_rev = wb.create_sheet(title="교과목개정표")
    headers = ["구분", "과목구분", "변경 전 (2026)", "변경 후 (2027)"]
    for col_i, h in enumerate(headers, start=1):
        c = ws_rev.cell(1, col_i, h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_all

    for r_i, (_, rev_row) in enumerate(revision_df.iterrows(), start=2):
        for col_i, h in enumerate(headers, start=1):
            c = ws_rev.cell(r_i, col_i, rev_row[h])
            c.font = font_body
            c.border = border_all
            c.alignment = align_center if col_i <= 2 else align_left

    for col_i, width in enumerate([10, 18, 38, 38], start=1):
        ws_rev.column_dimensions[openpyxl.utils.get_column_letter(col_i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
