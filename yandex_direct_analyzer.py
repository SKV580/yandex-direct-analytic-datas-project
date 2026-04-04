#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import math
import datetime
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ==================== НАСТРОЙКИ ====================
# Укажите здесь путь к вашему файлу с данными или папке
DEFAULT_INPUT_FILE = "data/BD_RSY_MK_BFL_OPORA_01.03-04.04 2026.xlsx"
# Укажите здесь имя для файла с результатом
DEFAULT_OUTPUT_FILE = "out/analysis_result.xlsx"
# ===================================================


CANONICAL_COLUMNS = {
    "campaign": [
        "Название кампании", "Кампания", "campaign", "campaign name"
    ],
    "ad_group": [
        "Название группы", "Группа", "Группа объявлений", "ad group"
    ],
    "placement": [
        "Название площадки",
        "Название площадок yandex",
        "Название площадки yandex",
        "Площадка",
        "Site",
        "Placement"
    ],
    "device": [
        "Тип устройства", "Устройство", "device"
    ],
    "gender": [
        "Пол", "gender"
    ],
    "age": [
        "Возраст", "age"
    ],
    "targeting_type": [
        "Тип условия показа", "Условие показа", "targeting type"
    ],
    "title": [
        "Заголовок", "Заголовок 1", "title"
    ],
    "image": [
        "Изображение",
        "image",
        "Название файла изображения"
    ],
    "spend": [
        "Расход, ₽", "Расход", "Cost", "Spend"
    ],
    "clicks": [
        "Клики", "Clicks"
    ],
    "conversions": [
        "Конверсии", "Conversions"
    ],
    "impressions": [
        "Показы", "Impressions"
    ],
    "bounce_rate": [
        "Отказы, %", "Показатель отказов, %", "Bounce rate"
    ],
    "depth": [
        "Глубина просмотра", "Depth"
    ],
}

DIMENSIONS = [
    ("campaign", "Кампания"),
    ("device", "Устройство"),
    ("gender", "Пол"),
    ("age", "Возраст"),
    ("targeting_type", "Тип условия показа"),
    ("placement", "Площадка"),
    ("ad_group", "Группа"),
    ("title", "Заголовок"),
    ("image", "Изображение"),
]


class AnalyzerError(Exception):
    pass


def normalize_text(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("ё", "е")
    s = re.sub(r"\s+", " ", s)
    return s


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    normalized_map = {normalize_text(c): c for c in df.columns}
    for c in candidates:
        if normalize_text(c) in normalized_map:
            return normalized_map[normalize_text(c)]
    return None


def clean_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace("\u00a0", " ", regex=False)
    s = s.str.replace("₽", "", regex=False)
    s = s.str.replace("%", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    s = s.replace({"nan": np.nan, "None": np.nan, "": np.nan, "-": np.nan})
    return pd.to_numeric(s, errors="coerce")


def resolve_input(path_str: str) -> Path:
    p = Path(path_str)
    if p.is_file():
        return p
    if p.is_dir():
        files = sorted([x for x in p.iterdir() if x.suffix.lower() in {".xlsx", ".xls"} and not x.name.startswith("~$")])
        if not files:
            raise AnalyzerError(f"В папке {p} не найдено Excel-файлов")
        return files[0]
    raise AnalyzerError(f"Путь не найден: {p}")


def load_data(path: Path) -> pd.DataFrame:
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        raise AnalyzerError(f"Не удалось открыть Excel-файл: {e}")
    frames = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            if not df.empty:
                df["__sheet__"] = sheet
                frames.append(df)
        except Exception:
            continue
    if not frames:
        raise AnalyzerError("Не удалось прочитать ни одного непустого листа")
    return pd.concat(frames, ignore_index=True)


def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    result = {}
    for key, candidates in CANONICAL_COLUMNS.items():
        col = find_column(df, candidates)
        if col:
            result[key] = col
    required = ["campaign", "spend", "clicks", "conversions"]
    missing = [k for k in required if k not in result]
    if missing:
        raise AnalyzerError(f"Не найдены обязательные колонки: {', '.join(missing)}")
    return result


def prepare_dataframe(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    work = pd.DataFrame()
    for key, col in mapping.items():
        work[key] = df[col]

    for col in ["spend", "clicks", "conversions", "impressions", "bounce_rate", "depth"]:
        if col in work.columns:
            work[col] = clean_numeric(work[col])
        else:
            work[col] = np.nan

    # "Отказы, %" обычно попадает как число типа 12.3 (без символа %).
    # В Excel формат '0.00%' ожидает долю (0.123), поэтому приводим.
    if "bounce_rate" in work.columns:
        s = work["bounce_rate"].dropna()
        if not s.empty and s.max() > 1.0:
            work["bounce_rate"] = work["bounce_rate"] / 100.0

    for col in ["campaign", "ad_group", "placement", "device", "gender", "age", "targeting_type", "title", "image"]:
        if col not in work.columns:
            work[col] = "Не указано"
        work[col] = work[col].fillna("Не указано").astype(str).str.strip().replace({"": "Не указано"})

    work["spend"] = work["spend"].fillna(0.0)
    work["clicks"] = work["clicks"].fillna(0.0)
    work["conversions"] = work["conversions"].fillna(0.0)
    work["impressions"] = work["impressions"].fillna(0.0)

    return work


def add_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["cpc"] = np.where(df["clicks"] > 0, df["spend"] / df["clicks"], np.nan)
    df["cpa"] = np.where(df["conversions"] > 0, df["spend"] / df["conversions"], np.nan)
    df["cr"] = np.where(df["clicks"] > 0, df["conversions"] / df["clicks"], np.nan)
    df["ctr"] = np.where(df["impressions"] > 0, df["clicks"] / df["impressions"], np.nan)
    return df


def summarize(df: pd.DataFrame, dim: str) -> pd.DataFrame:
    agg = df.groupby(dim, dropna=False).agg(
        spend=("spend", "sum"),
        clicks=("clicks", "sum"),
        conversions=("conversions", "sum"),
        impressions=("impressions", "sum"),
        bounce_rate=("bounce_rate", "mean"),
        depth=("depth", "mean"),
    ).reset_index()
    agg["cpc"] = np.where(agg["clicks"] > 0, agg["spend"] / agg["clicks"], np.nan)
    agg["cpa"] = np.where(agg["conversions"] > 0, agg["spend"] / agg["conversions"], np.nan)
    agg["cr"] = np.where(agg["clicks"] > 0, agg["conversions"] / agg["clicks"], np.nan)
    agg["ctr"] = np.where(agg["impressions"] > 0, agg["clicks"] / agg["impressions"], np.nan)
    agg = agg.sort_values(["conversions", "spend"], ascending=[False, False]).reset_index(drop=True)
    return agg


def overall_metrics(df: pd.DataFrame) -> Dict[str, float]:
    spend = df["spend"].sum()
    clicks = df["clicks"].sum()
    conv = df["conversions"].sum()
    impr = df["impressions"].sum()
    return {
        "spend": spend,
        "clicks": clicks,
        "conversions": conv,
        "impressions": impr,
        "cpc": spend / clicks if clicks else np.nan,
        "cpa": spend / conv if conv else np.nan,
        "cr": conv / clicks if clicks else np.nan,
        "ctr": clicks / impr if impr else np.nan,
    }


def fmt_money(x) -> str:
    if pd.isna(x):
        return "—"
    return f"{x:,.2f} ₽".replace(",", " ")


def fmt_num(x) -> str:
    if pd.isna(x):
        return "—"
    return f"{int(round(x)):,}".replace(",", " ")


def fmt_pct(x) -> str:
    if pd.isna(x):
        return "—"
    return f"{x * 100:.2f}%"


def best_and_worst(summary: pd.DataFrame, min_clicks: int, min_spend: float) -> Tuple[pd.DataFrame, pd.DataFrame]:
    eligible = summary[(summary["clicks"] >= min_clicks) | (summary["spend"] >= min_spend)].copy()
    if eligible.empty:
        return eligible, eligible
    good = eligible[eligible["conversions"] > 0].sort_values(["cpa", "conversions"], ascending=[True, False]).head(5)
    bad = eligible[(eligible["conversions"] == 0) | (eligible["cpa"] > eligible["cpa"].median())].sort_values(["conversions", "spend"], ascending=[True, False]).head(5)
    return good, bad


def campaign_recommendations(
    df: pd.DataFrame,
    campaign_summary: pd.DataFrame,
    min_clicks: int,
    min_spend: float,
    available_dims: set,
) -> pd.DataFrame:
    account_cpa = overall_metrics(df)["cpa"]
    rows = []
    for _, row in campaign_summary.iterrows():
        camp = row["campaign"]
        part = df[df["campaign"] == camp].copy()
        recs = []

        for dim, label in [
            ("device", "устройствам"),
            ("gender", "полу"),
            ("age", "возрасту"),
            ("targeting_type", "типу условия показа"),
            ("placement", "площадкам"),
        ]:
            # Строим рекомендации только по тем измерениям, которые реально присутствуют в отчёте.
            if dim not in available_dims:
                continue

            s = summarize(part, dim)
            good, bad = best_and_worst(s, min_clicks=min_clicks, min_spend=min_spend)
            if not good.empty:
                good_names = ", ".join(good[dim].astype(str).head(3).tolist())
                recs.append(f"Усилить по {label}: {good_names}")
            if not bad.empty:
                bad_names = ", ".join(bad[dim].astype(str).head(3).tolist())
                recs.append(f"Сократить по {label}: {bad_names}")

        status = "усилить"
        if pd.notna(row["cpa"]) and pd.notna(account_cpa):
            if row["cpa"] > account_cpa * 1.5:
                status = "сократить и пересобрать"
            elif row["cpa"] > account_cpa * 1.15:
                status = "оптимизировать"

        rows.append({
            "campaign": camp,
            "spend": row["spend"],
            "clicks": row["clicks"],
            "conversions": row["conversions"],
            "cpc": row["cpc"],
            "cpa": row["cpa"],
            "cr": row["cr"],
            "status": status,
            "recommendations": " | ".join(recs[:10]) if recs else "Недостаточно данных для рекомендаций",
        })
    return pd.DataFrame(rows).sort_values(["conversions", "spend"], ascending=[False, False])


def print_terminal_report(df: pd.DataFrame, campaign_summary: pd.DataFrame, recs: pd.DataFrame) -> None:
    totals = overall_metrics(df)
    print("=" * 90)
    print("АНАЛИЗ ОТЧЕТА ЯНДЕКС ДИРЕКТ")
    print("=" * 90)
    print(f"Расход: {fmt_money(totals['spend'])}")
    print(f"Клики: {fmt_num(totals['clicks'])}")
    print(f"Конверсии: {fmt_num(totals['conversions'])}")
    print(f"CPC: {fmt_money(totals['cpc'])}")
    print(f"CPA: {fmt_money(totals['cpa'])}")
    print(f"CR: {fmt_pct(totals['cr'])}")
    print()
    print("КАМПАНИИ")
    print("-" * 90)
    cols = ["campaign", "spend", "clicks", "conversions", "cpc", "cpa", "cr"]
    temp = campaign_summary[cols].copy()
    temp["spend"] = temp["spend"].map(fmt_money)
    temp["clicks"] = temp["clicks"].map(fmt_num)
    temp["conversions"] = temp["conversions"].map(fmt_num)
    temp["cpc"] = temp["cpc"].map(fmt_money)
    temp["cpa"] = temp["cpa"].map(fmt_money)
    temp["cr"] = temp["cr"].map(fmt_pct)
    print(temp.to_string(index=False))
    print()
    print("РЕКОМЕНДАЦИИ ПО КАМПАНИЯМ")
    print("-" * 90)
    for _, row in recs.iterrows():
        print(f"\nКампания: {row['campaign']}")
        print(f"Статус: {row['status']}")
        for part in str(row["recommendations"]).split(" | "):
            print(f"- {part}")
    print()


def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", str(name))
    return name[:31]


def write_df_sheet(ws, df: pd.DataFrame, table_name: str) -> None:
    if df.empty:
        ws["B2"] = "Нет данных"
        return
    start_row, start_col = 2, 2
    headers = list(df.columns)
    for i, h in enumerate(headers, start=start_col):
        c = ws.cell(row=start_row, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for r_idx, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for c_idx, val in enumerate(row, start=start_col):
            ws.cell(row=r_idx, column=c_idx, value=None if pd.isna(val) else val)

    end_row = start_row + len(df)
    end_col = start_col + len(headers) - 1
    ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
    tab = Table(displayName=table_name[:25], ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)
    ws.freeze_panes = f"{get_column_letter(start_col)}{start_row + 1}"

    money_cols = {"spend", "cpc", "cpa"}
    pct_cols = {"cr", "ctr", "bounce_rate"}
    int_cols = {"clicks", "conversions", "impressions"}

    for idx, h in enumerate(headers, start=start_col):
        col_letter = get_column_letter(idx)
        # Исправлена ошибка TypeError при вычислении ширины колонок
        max_len = df.iloc[:, idx - start_col].apply(lambda x: len(str(x)) if pd.notna(x) else 0).max() if len(df) else 0
        width = max(12, min(45, max(len(str(h)), max_len) + 2))
        ws.column_dimensions[col_letter].width = width
        if h in money_cols:
            for cell in ws[col_letter][start_row: end_row]:
                cell.number_format = '#,##0.00 [$₽-419]'
        elif h in pct_cols:
            for cell in ws[col_letter][start_row: end_row]:
                cell.number_format = '0.00%'
        elif h in int_cols:
            for cell in ws[col_letter][start_row: end_row]:
                cell.number_format = '#,##0'

    if "spend" in headers:
        col_letter = get_column_letter(start_col + headers.index("spend"))
        ws.conditional_formatting.add(
            f"{col_letter}{start_row+1}:{col_letter}{end_row}",
            ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='BDD7EE')
        )
    if "cpa" in headers:
        col_letter = get_column_letter(start_col + headers.index("cpa"))
        ws.conditional_formatting.add(
            f"{col_letter}{start_row+1}:{col_letter}{end_row}",
            ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B')
        )


def build_excel(output_path: Path, totals: Dict[str, float], campaign_summary: pd.DataFrame, recs: pd.DataFrame, slices: Dict[str, pd.DataFrame]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Обзор"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 18

    ws["B2"] = "Анализ отчета Яндекс Директ"
    ws["B2"].font = Font(size=14, bold=True)

    overview = [
        ("Общий расход", totals["spend"]),
        ("Клики", totals["clicks"]),
        ("Конверсии", totals["conversions"]),
        ("Показы", totals["impressions"]),
        ("Средний CPC", totals["cpc"]),
        ("Средний CPA", totals["cpa"]),
        ("CR", totals["cr"]),
        ("CTR", totals["ctr"]),
    ]
    r = 4
    for name, value in overview:
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=None if pd.isna(value) else float(value))
        r += 1
    for rr in range(4, r):
        label = ws.cell(row=rr, column=2).value
        cell = ws.cell(row=rr, column=3)
        if label in {"Общий расход", "Средний CPC", "Средний CPA"}:
            cell.number_format = '#,##0.00 [$₽-419]'
        elif label in {"CR", "CTR"}:
            cell.number_format = '0.00%'
        else:
            cell.number_format = '#,##0'

    camp_sheet = wb.create_sheet("Кампании")
    write_df_sheet(camp_sheet, campaign_summary, "CampaignSummary")

    rec_sheet = wb.create_sheet("Рекомендации")
    write_df_sheet(rec_sheet, recs, "CampaignRecs")

    for key, label in DIMENSIONS:
        if key in slices:
            sh = wb.create_sheet(safe_sheet_name(label))
            write_df_sheet(sh, slices[key], f"T_{key[:20]}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Анализ Excel-отчета Яндекс Директ")
    parser.add_argument("--input", default=DEFAULT_INPUT_FILE, help=f"Путь к Excel-файлу или к папке. По умолчанию: '{DEFAULT_INPUT_FILE}'")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_FILE, help=f"Путь к итоговому Excel-файлу. По умолчанию: '{DEFAULT_OUTPUT_FILE}'")
    parser.add_argument("--min-spend", type=float, default=500.0, help="Минимальный расход для рекомендаций")
    parser.add_argument("--min-clicks", type=int, default=20, help="Минимум кликов для рекомендаций")
    args = parser.parse_args()

    source = resolve_input(args.input)
    raw = load_data(source)
    mapping = map_columns(raw)
    df = prepare_dataframe(raw, mapping)
    df = add_metrics(df)

    totals = overall_metrics(df)
    campaign_summary = summarize(df, "campaign")

    # available_dims: измерения, которые были найдены по фактическим колонкам отчёта.
    available_dims = set(mapping.keys())

    slices = {}
    for key, _ in DIMENSIONS:
        if key in available_dims:
            try:
                slices[key] = summarize(df, key)
            except Exception:
                pass

    recs = campaign_recommendations(
        df,
        campaign_summary,
        args.min_clicks,
        args.min_spend,
        available_dims=available_dims,
    )
    print_terminal_report(df, campaign_summary, recs)
    build_excel(Path(args.output), totals, campaign_summary, recs, slices)
    print(f"Готово. Excel-отчет сохранен: {args.output}")


if __name__ == "__main__":
    main()
