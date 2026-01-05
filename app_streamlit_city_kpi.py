#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv
import io
import posixpath
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import streamlit as st

try:
    import pandas as pd  # Optional; used if available for nicer tables.
except Exception:  # pragma: no cover
    pd = None

EXCEL_NS_MAIN = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXCEL_NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

SUMMARY_COLUMNS = [
    "城市",
    "期间",
    "产成率(%)",
    "产值(元/kg)",
    "生肉产量(吨)",
    "销量(吨)",
    "产销率(%)",
    "宰鸡量(千只)",
    "均重(kg/只)",
]

PART_ORDER = [
    "腿类",
    "胸部",
    "大胸",
    "胸皮",
    "里肌",
    "翅类",
    "骨架",
    "爪类",
    "鸡肝",
    "鸡心",
    "脖类",
    "鸡胗",
    "鸡头",
    "油类",
    "骨架+鸡脖+鸡头",
    "骨架+鸡脖",
    "主产品",
    "副产品",
    "总计",
]

PART_ALIASES = {
    "胸部": ["胸类", "胸部"],
    "大胸": ["胸类-胸", "大胸"],
    "胸皮": ["胸类-胸皮", "胸皮"],
    "里肌": ["里肌类", "里肌"],
    "骨架": ["骨架类", "骨架"],
    "爪类": ["爪类"],
    "鸡肝": ["鸡肝类", "鸡肝"],
    "鸡心": ["鸡心类", "鸡心"],
    "鸡胗": ["鸡胗类", "鸡胗"],
    "鸡头": ["鸡头类", "鸡头"],
    "脖类": ["脖类", "鸡脖"],
    "油类": ["油类"],
    "骨架+鸡脖+鸡头": ["鸡头+鸡脖+骨架", "骨架+鸡脖+鸡头"],
    "骨架+鸡脖": ["骨架+鸡脖"],
    "腿类": ["腿类"],
}

DERIVED_PARTS = {
    "骨架+鸡脖+鸡头": ["骨架", "脖类", "鸡头"],
    "骨架+鸡脖": ["骨架", "脖类"],
    "主产品": ["胸部", "大胸", "里肌", "翅类", "爪类"],
    "副产品": ["鸡肝", "鸡心", "脖类", "鸡胗", "鸡头", "油类", "骨架"],
}

HEADER_ALIASES = {
    "日期": "期间",
    "产成率": "产成率(%)",
    "产成率(%)": "产成率(%)",
    "产值": "产值(元/kg)",
    "产值(元/kg)": "产值(元/kg)",
    "生肉产量": "生肉产量(吨)",
    "生肉产量(吨)": "生肉产量(吨)",
    "产量": "生肉产量(吨)",
    "销量": "销量(吨)",
    "销量(吨)": "销量(吨)",
    "产销率": "产销率(%)",
    "产销率(%)": "产销率(%)",
    "宰鸡量": "宰鸡量(千只)",
    "宰鸡量(千只)": "宰鸡量(千只)",
    "均重": "均重(kg/只)",
    "均重(kg/只)": "均重(kg/只)",
}

CHART_PATH = Path("/home/fan/8f573a0c22d5ad3933140a75507eacb8.png")


def read_shared_strings(zip_file: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for si in root.findall("m:si", EXCEL_NS_MAIN):
        text = "".join(t.text or "" for t in si.findall(".//m:t", EXCEL_NS_MAIN))
        strings.append(text)
    return strings


def column_to_index(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def read_first_sheet_bytes(data: bytes) -> Tuple[List[str], List[List[str]]]:
    """Return header row + dense rows from the first worksheet in bytes."""
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        shared = read_shared_strings(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall("rel:Relationship", EXCEL_NS_REL)}
        first_sheet = workbook.find("m:sheets/m:sheet", EXCEL_NS_MAIN)
        if first_sheet is None:
            return [], []
        rel_id = first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rels.get(rel_id, "worksheets/sheet1.xml")
        sheet_path = posixpath.normpath(posixpath.join("xl", target))
        root = ET.fromstring(zf.read(sheet_path))

        rows: List[Dict[int, str]] = []
        max_col = -1
        for row in root.findall("m:sheetData/m:row", EXCEL_NS_MAIN):
            values: Dict[int, str] = {}
            for cell in row.findall("m:c", EXCEL_NS_MAIN):
                ref = cell.attrib.get("r", "")
                col_letters = "".join(ch for ch in ref if ch.isalpha())
                if not col_letters:
                    continue
                col_idx = column_to_index(col_letters)
                max_col = max(max_col, col_idx)
                cell_type = cell.attrib.get("t")
                value = ""
                v = cell.find("m:v", EXCEL_NS_MAIN)
                inline = cell.find("m:is", EXCEL_NS_MAIN)
                if cell_type == "s" and v is not None and v.text is not None:
                    s_idx = int(v.text)
                    if 0 <= s_idx < len(shared):
                        value = shared[s_idx]
                elif cell_type == "inlineStr" and inline is not None:
                    value = "".join(t.text or "" for t in inline.findall(".//m:t", EXCEL_NS_MAIN))
                elif v is not None and v.text is not None:
                    value = v.text
                if value != "":
                    values[col_idx] = value
            rows.append(values)
        width = max_col + 1 if max_col >= 0 else 0
        dense_rows: List[List[str]] = []
        for row in rows:
            dense_rows.append([row.get(i, "") for i in range(width)])
        header = dense_rows[0] if dense_rows else []
        return header, dense_rows


def normalize_date(text: str) -> str:
    cleaned = (text or "").strip()
    if not cleaned:
        return ""
    try:
        as_float = float(cleaned.replace(",", ""))
        if as_float.is_integer():
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=int(as_float))).date().isoformat()
    except ValueError:
        pass
    return cleaned


def build_column_map(header_row: List[str]) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for idx, name in enumerate(header_row):
        normalized = HEADER_ALIASES.get((name or "").strip())
        if normalized:
            mapping[normalized] = idx
    return mapping


def extract_records_from_bytes(data: bytes, filename: str) -> List[Dict[str, str]]:
    _, rows = read_first_sheet_bytes(data)
    header_idx = None
    for idx, row in enumerate(rows):
        if any(cell.strip() == "日期" for cell in row):
            header_idx = idx
            break
    if header_idx is None:
        return []

    header_row = rows[header_idx]
    col_map = build_column_map(header_row)
    data_rows = rows[header_idx + 1 :]
    records: List[Dict[str, str]] = []
    city = Path(filename).stem.split("_", 1)[0]

    for row in data_rows:
        if not any(cell.strip() for cell in row):
            continue
        period_idx = col_map.get("期间")
        period = normalize_date(row[period_idx]) if period_idx is not None and period_idx < len(row) else ""
        values: Dict[str, str] = {"城市": city, "期间": period}
        empty_value = True
        for col in SUMMARY_COLUMNS[2:]:
            idx = col_map.get(col)
            value = row[idx].strip() if idx is not None and idx < len(row) else ""
            values[col] = value
            if value:
                empty_value = False
        if period or not empty_value:
            records.append(values)
    return records


def to_csv_bytes(records: List[Dict[str, str]]) -> bytes:
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=SUMMARY_COLUMNS)
    writer.writeheader()
    writer.writerows(records)
    return buf.getvalue().encode("utf-8")


def to_csv_bytes_generic(records: List[Dict[str, str]]) -> bytes:
    if not records:
        return "".encode("utf-8")
    fields = sorted({k for r in records for k in r.keys()})
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fields)
    writer.writeheader()
    writer.writerows(records)
    return buf.getvalue().encode("utf-8")


def extract_monthly_value_from_bytes(data: bytes, filename: str) -> List[Dict[str, str]]:
    """优先读取“组合还原后产成率总览（月累计）”，缺失则回退“本月至今累计”，取含税单价/金额/销量。"""
    _, rows = read_first_sheet_bytes(data)

    def _read_block(marker: str) -> List[Dict[str, str]]:
        start_idx = None
        for idx, row in enumerate(rows):
            if row and str(row[0]).strip() == marker:
                start_idx = idx
                break
        if start_idx is None:
            return []
        header_idx = None
        for j in range(start_idx + 1, min(start_idx + 5, len(rows))):
            if rows[j] and str(rows[j][0]).strip() == "项目":
                header_idx = j
                break
        if header_idx is None:
            return []

        header = rows[header_idx]
        name_idx = 0
        sales_idx = _find_col_idx(header, ["销量(kg)", "销量", "销售量", "销量kg", "销量KG"])
        amount_idx = _find_col_idx(header, ["含税金额", "含税金额(元)", "金额"])
        unit_idx = _find_col_idx(header, ["含税单价", "单价"])
        if sales_idx is None and amount_idx is None and unit_idx is None:
            return []

        stop_words = {"部位还原后的产成率", "组合还原后产成率总览（月累计）", "组合还原后产成率总览"}
        data_rows = rows[header_idx + 1 :]
        plant = Path(filename).stem.split("_", 1)[0]
        out: List[Dict[str, str]] = []
        for row in data_rows:
            if not any(str(c).strip() for c in row):
                break
            first = str(row[name_idx]).strip() if len(row) > name_idx else ""
            if first in stop_words:
                break
            if not first:
                continue
            part = normalize_part_name(first)
            sales = _parse_number(row[sales_idx]) if (sales_idx is not None and sales_idx < len(row)) else None
            amount = _parse_number(row[amount_idx]) if (amount_idx is not None and amount_idx < len(row)) else None
            unit = _parse_number(row[unit_idx]) if (unit_idx is not None and unit_idx < len(row)) else None
            if unit is None and amount is not None and sales is not None and sales > 0:
                unit = amount / sales
            if unit is None and amount is None and sales is None:
                continue
            out.append({"城市": plant, "部位": part, "销量": sales, "含税金额": amount, "含税单价": unit})
        return out

    return _read_block("组合还原后产成率总览（月累计）") or _read_block("本月至今累计")


def _find_col_idx(header: List[str], name_cands: List[str]) -> int | None:
    clean = [str(c or "").strip() for c in header]
    for cand in name_cands:
        if cand in clean:
            return clean.index(cand)
    # loose contains
    for i, col in enumerate(clean):
        for cand in name_cands:
            if cand and cand in col:
                return i
    return None


def extract_monthly_parts_from_bytes(data: bytes, filename: str) -> List[Dict[str, str]]:
    """优先读取“组合还原后产成率总览（月累计）”，若缺失回退“本月至今累计”."""
    header_row, rows = read_first_sheet_bytes(data)

    def _read_block(marker: str) -> List[Dict[str, str]]:
        start_idx = None
        for idx, row in enumerate(rows):
            if row and str(row[0]).strip() == marker:
                start_idx = idx
                break
        if start_idx is None:
            return []
        header_idx = None
        for j in range(start_idx + 1, min(start_idx + 5, len(rows))):
            if rows[j] and str(rows[j][0]).strip() == "项目":
                header_idx = j
                break
        if header_idx is None:
            return []
        header = rows[header_idx]
        rate_idx = _find_col_idx(header, ["产成率%", "产成率", "调整后产成率(%)", "产成率(%)"])
        name_idx = 0
        stop_words = {"组合还原后产成率总览（月累计）", "组合还原后产成率总览"}
        data_rows = rows[header_idx + 1 :]
        out: List[Dict[str, str]] = []
        plant = Path(filename).stem.split("_", 1)[0]
        for row in data_rows:
            if not any(str(c).strip() for c in row):
                break
            first = str(row[name_idx]).strip() if len(row) > name_idx else ""
            if first in stop_words:
                break
            if not first:
                continue
            rate_val = row[rate_idx] if (rate_idx is not None and rate_idx < len(row)) else ""
            out.append({"城市": plant, "部位": normalize_part_name(first), "产成率": rate_val})
        return out

    return _read_block("组合还原后产成率总览（月累计）") or _read_block("本月至今累计")


def _parse_number(val: str) -> float | None:
    text = str(val or "").replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _parse_percent(val: str) -> float | None:
    text = str(val or "").strip()
    if not text:
        return None
    has_percent = "%" in text
    text = text.replace("%", "").replace(",", "").strip()
    if not text:
        return None
    try:
        num = float(text)
    except ValueError:
        return None
    if has_percent:
        return num / 100
    if abs(num) > 1:
        return num / 100
    return num


def normalize_part_name(name: str) -> str:
    cleaned = str(name or "").strip()
    for target, aliases in PART_ALIASES.items():
        for a in aliases:
            if cleaned == a:
                return target
    return cleaned


def _format_thousands0(val: str) -> str:
    num = _parse_number(val)
    return f"{num:,.0f}" if num is not None else ""


def _format_percent2(val: str) -> str:
    num = _parse_percent(val)
    return f"{num * 100:.2f}%" if num is not None else ""


def _format_percent0(val: str) -> str:
    num = _parse_percent(val)
    return f"{num * 100:.0f}%" if num is not None else ""


def _format_decimal2(val: float | str | None) -> str:
    num = _parse_number(val)
    return f"{num:.2f}" if num is not None else ""


def main() -> None:
    st.set_page_config(page_title="城市 KPI 汇总", layout="wide")
    st.title("城市 KPI 汇总（多文件上传）")
    st.caption("上传多个 Excel（每个文件首个工作表，包含“日期”表头），自动汇总核心指标。")

    if CHART_PATH.exists():
        st.subheader("图表预览")
        st.image(str(CHART_PATH), use_column_width=True)

    uploaded_files = st.file_uploader("选择多个 .xlsx 文件", type=["xlsx"], accept_multiple_files=True)
    if not uploaded_files:
        st.info("等待上传文件。")
        return

    import re
    month_total_re = re.compile(r"^\s*(\d{1,2})月累计\s*$")

    def _is_month_total(text: str) -> bool:
        return bool(month_total_re.match(str(text or "")))

    records: List[Dict[str, str]] = []
    part_records: List[Dict[str, str]] = []
    value_records: List[Dict[str, str]] = []
    for f in uploaded_files:
        try:
            data = f.getvalue()
            recs = extract_records_from_bytes(data, f.name)
            records.extend(recs)
            part_records.extend(extract_monthly_parts_from_bytes(data, f.name))
            value_records.extend(extract_monthly_value_from_bytes(data, f.name))
        except Exception as exc:  # pragma: no cover - surfaced to user
            st.error(f"{f.name} 解析失败: {exc}")

    if not records:
        st.warning("未解析到任何数据，请检查表头是否包含“日期”。")
        return

    if pd:
        df = pd.DataFrame(records, columns=SUMMARY_COLUMNS)
        core_df = df[df["期间"].apply(_is_month_total)].copy()
        if core_df.empty:
            st.info("未找到“X月累计”行。")
            return

        periods = sorted(core_df["期间"].dropna().unique().tolist())
        sel = st.multiselect("选择期间（可多选）", options=periods, default=periods)
        core_view = core_df if not sel else core_df[core_df["期间"].isin(sel)]
        if core_view.empty:
            st.warning("当前选择下无数据。")
            return

        display_df = core_view.copy()
        for col in ["生肉产量(吨)", "销量(吨)", "宰鸡量(千只)"]:
            if col in display_df:
                display_df[col] = display_df[col].apply(_format_thousands0)
        if "产成率(%)" in display_df:
            display_df["产成率(%)"] = display_df["产成率(%)"].apply(_format_percent2)
        if "产销率(%)" in display_df:
            display_df["产销率(%)"] = display_df["产销率(%)"].apply(_format_percent0)

        st.subheader("月累计汇总（核心指标概览）")
        st.dataframe(display_df.sort_values(["期间", "城市"]), use_container_width=True)
        csv_bytes = display_df.to_csv(index=False).encode("utf-8")
    else:
        core_records = [r for r in records if _is_month_total(r.get("期间", ""))]
        if not core_records:
            st.info("未找到“X月累计”行。")
            return
        periods = sorted({r.get("期间", "") for r in core_records if r.get("期间", "")})
        sel = st.multiselect("选择期间（可多选）", options=periods, default=periods)
        core_view = [r for r in core_records if (not sel) or (r.get("期间", "") in sel)]
        if not core_view:
            st.warning("当前选择下无数据。")
            return
        formatted_records = []
        for r in core_view:
            item = dict(r)
            item["生肉产量(吨)"] = _format_thousands0(r.get("生肉产量(吨)", ""))
            item["销量(吨)"] = _format_thousands0(r.get("销量(吨)", ""))
            item["宰鸡量(千只)"] = _format_thousands0(r.get("宰鸡量(千只)", ""))
            item["产成率(%)"] = _format_percent2(r.get("产成率(%)", ""))
            item["产销率(%)"] = _format_percent0(r.get("产销率(%)", ""))
            formatted_records.append(item)
        st.subheader("月累计汇总（核心指标概览）")
        st.dataframe(formatted_records, use_container_width=True, height=300)
        csv_bytes = to_csv_bytes(formatted_records)

    st.download_button(
        "下载汇总 CSV",
        data=csv_bytes,
        file_name="city_kpi_summary.csv",
        mime="text/csv",
        use_container_width=True,
    )

    # === 月产成细项对比 ===
    st.divider()
    st.subheader("月产成细项对比（本月至今累计）")
    if not part_records:
        st.info("未找到“本月至今累计”明细。")
        return
    # 归一化部位名称
    for r in part_records:
        r["部位"] = normalize_part_name(r.get("部位", ""))

    def _pivot_parts(records: List[Dict[str, str]]):
        if pd:
            dfp = pd.DataFrame(records)
            # 尝试解析小数
            dfp["产成率"] = dfp["产成率"].apply(_parse_percent)
            dfp = dfp.dropna(subset=["产成率"])
            dfp["部位"] = dfp["部位"].apply(normalize_part_name)
            pivot = dfp.pivot_table(index="部位", columns="城市", values="产成率", aggfunc="first")
            # 补充派生行
            for new_part, parts in DERIVED_PARTS.items():
                vals = {}
                for city in pivot.columns:
                    vals[city] = sum(pivot.at[p, city] if p in pivot.index and pd.notna(pivot.at[p, city]) else 0 for p in parts)
                pivot.loc[new_part] = vals
            # 若已有“总计”则保留，否则求和
            if "总计" not in pivot.index:
                pivot.loc["总计"] = pivot.sum(min_count=1)
            # 重新排序
            pivot = pivot.reindex(PART_ORDER, fill_value=pd.NA)
            # 格式化百分比（两位小数）
            return pivot.applymap(lambda x: f"{x*100:.2f}%" if pd.notna(x) else "")
        else:
            # 手工 pivot
            parts = {}
            for r in records:
                part = r.get("部位")
                city = r.get("城市")
                val = _parse_percent(r.get("产成率"))
                if part and city and val is not None:
                    part_norm = normalize_part_name(part)
                    parts.setdefault(part_norm, {})[city] = val

            # 派生行
            all_cities = {c for v in parts.values() for c in v}
            for new_part, comp in DERIVED_PARTS.items():
                city_vals = {}
                for city in all_cities:
                    city_vals[city] = sum(parts.get(p, {}).get(city, 0) for p in comp)
                parts[new_part] = city_vals

            # 总计
            if "总计" not in parts:
                parts["总计"] = {c: sum(parts.get(p, {}).get(c, 0) for p in parts if p != "总计") for c in all_cities}

            rows = []
            all_cities = sorted({c for v in parts.values() for c in v})
            for part in PART_ORDER:
                row = {"部位": part}
                for c in all_cities:
                    val = parts.get(part, {}).get(c)
                    row[c] = f"{val*100:.2f}%" if val is not None else ""
                rows.append(row)
            return rows

    pivoted = _pivot_parts(part_records)
    if pd and isinstance(pivoted, pd.DataFrame):
        st.dataframe(pivoted, use_container_width=True)
        part_csv = pivoted.to_csv().encode("utf-8")
    else:
        st.dataframe(pivoted, use_container_width=True)
        part_csv = to_csv_bytes_generic(pivoted if isinstance(pivoted, list) else [])

    st.download_button(
        "下载月产成细项对比 CSV",
        data=part_csv,
        file_name="city_kpi_parts.csv",
        mime="text/csv",
        use_container_width=True,
    )

    # === 月产值细项对比（含税单价） ===
    st.divider()
    st.subheader("月产值细项对比（本月至今累计，含税单价）")
    if not value_records:
        st.info("未找到“本月至今累计”明细。")
        return

    def _pivot_value(records: List[Dict[str, str]]):
        if pd:
            dfv = pd.DataFrame(records)
            dfv["含税单价"] = dfv.apply(
                lambda r: r["含税单价"]
                if pd.notna(r.get("含税单价"))
                else (r["含税金额"] / r["销量"] if pd.notna(r.get("含税金额")) and pd.notna(r.get("销量")) and r["销量"] > 0 else pd.NA),
                axis=1,
            )
            dfv = dfv.dropna(subset=["含税单价"])
            dfv["部位"] = dfv["部位"].apply(normalize_part_name)
            base = dfv[["部位", "城市", "含税单价", "销量", "含税金额"]]

            derived_rows = []
            for new_part, parts in DERIVED_PARTS.items():
                sub = base[base["部位"].isin(parts)]
                if sub.empty:
                    continue
                # Prefer weighted average from unit price + sales to avoid amount unit drift.
                unit_sub = sub.dropna(subset=["含税单价", "销量"]).copy()
                if not unit_sub.empty:
                    unit_sub["加权金额"] = unit_sub["含税单价"] * unit_sub["销量"]
                    agg = unit_sub.groupby("城市").agg({"加权金额": "sum", "销量": "sum"})
                    agg["含税单价"] = agg.apply(
                        lambda r: r["加权金额"] / r["销量"] if pd.notna(r["销量"]) and r["销量"] > 0 else pd.NA,
                        axis=1,
                    )
                    for city, row in agg.iterrows():
                        derived_rows.append(
                            {
                                "部位": new_part,
                                "城市": city,
                                "含税单价": row["含税单价"],
                                "销量": row["销量"],
                                "含税金额": row["加权金额"],
                            }
                        )
                    continue
                # Fallback to amount/sales if unit price is unavailable.
                amt_sub = sub.dropna(subset=["含税金额", "销量"])
                if amt_sub.empty:
                    continue
                agg = amt_sub.groupby("城市").agg({"含税金额": "sum", "销量": "sum"})
                agg["含税单价"] = agg.apply(
                    lambda r: r["含税金额"] / r["销量"] if pd.notna(r["销量"]) and r["销量"] > 0 else pd.NA,
                    axis=1,
                )
                for city, row in agg.iterrows():
                    derived_rows.append(
                        {
                            "部位": new_part,
                            "城市": city,
                            "含税单价": row["含税单价"],
                            "销量": row["销量"],
                            "含税金额": row["含税金额"],
                        }
                    )

            if derived_rows:
                derived_df = pd.DataFrame(derived_rows)
                base = base.merge(derived_df[["部位", "城市"]], on=["部位", "城市"], how="left", indicator=True)
                base = base[base["_merge"] == "left_only"].drop(columns=["_merge"])
                combined = pd.concat([base, derived_df], ignore_index=True)
            else:
                combined = base

            # 总计（仅 base 计算）
            if "总计" not in combined["部位"].values:
                total = base.groupby("城市").agg({"含税金额": "sum", "销量": "sum"})
                total["含税单价"] = total.apply(
                    lambda r: r["含税金额"] / r["销量"] if pd.notna(r["销量"]) and r["销量"] > 0 else pd.NA,
                    axis=1,
                )
                for city, row in total.iterrows():
                    combined = pd.concat(
                        [
                            combined,
                            pd.DataFrame(
                                [
                                    {
                                        "部位": "总计",
                                        "城市": city,
                                        "含税单价": row["含税单价"],
                                        "销量": row["销量"],
                                        "含税金额": row["含税金额"],
                                    }
                                ]
                            ),
                        ],
                        ignore_index=True,
                    )

            pivot = combined.pivot_table(index="部位", columns="城市", values="含税单价", aggfunc="first")
            pivot = pivot.reindex(PART_ORDER, fill_value=pd.NA)
            return pivot.applymap(lambda x: _format_decimal2(x) if pd.notna(x) else "")
        else:
            # 手工
            parts_map: Dict[str, Dict[str, Dict[str, float]]] = {}
            for r in records:
                part = normalize_part_name(r.get("部位"))
                city = r.get("城市")
                if not part or not city:
                    continue
                sales = _parse_number(r.get("销量"))
                amount = _parse_number(r.get("含税金额"))
                unit = _parse_number(r.get("含税单价"))
                if unit is None and amount is not None and sales is not None and sales > 0:
                    unit = amount / sales
                if unit is None:
                    continue
                parts_map.setdefault(part, {}).setdefault(city, {"amount": 0.0, "sales": 0.0})
                if amount is not None:
                    parts_map[part][city]["amount"] += amount
                if sales is not None:
                    parts_map[part][city]["sales"] += sales
                # store unit as fallback if no amount/sales
                parts_map[part][city]["unit"] = unit

            # 派生行
            all_cities = {c for v in parts_map.values() for c in v}
            for new_part, comp in DERIVED_PARTS.items():
                city_vals = {}
                for city in all_cities:
                    amt = sum(parts_map.get(p, {}).get(city, {}).get("amount", 0) for p in comp)
                    sal = sum(parts_map.get(p, {}).get(city, {}).get("sales", 0) for p in comp)
                    unit = amt / sal if sal > 0 else None
                    if unit is not None:
                        city_vals[city] = {"amount": amt, "sales": sal, "unit": unit}
                if city_vals:
                    parts_map[new_part] = city_vals

            # 总计
            if "总计" not in parts_map:
                total_city = {}
                for city in all_cities:
                    amt = sum(parts_map.get(p, {}).get(city, {}).get("amount", 0) for p in parts_map)
                    sal = sum(parts_map.get(p, {}).get(city, {}).get("sales", 0) for p in parts_map)
                    unit = amt / sal if sal > 0 else None
                    if unit is not None:
                        total_city[city] = {"amount": amt, "sales": sal, "unit": unit}
                if total_city:
                    parts_map["总计"] = total_city

            rows = []
            all_cities = sorted({c for v in parts_map.values() for c in v})
            for part in PART_ORDER:
                row = {"部位": part}
                for c in all_cities:
                    unit = parts_map.get(part, {}).get(c, {}).get("unit")
                    row[c] = _format_decimal2(unit)
                rows.append(row)
            return rows

    pivot_v = _pivot_value(value_records)
    if pd and isinstance(pivot_v, pd.DataFrame):
        st.dataframe(pivot_v, use_container_width=True)
        value_csv = pivot_v.to_csv().encode("utf-8")
    else:
        st.dataframe(pivot_v, use_container_width=True)
        value_csv = to_csv_bytes_generic(pivot_v if isinstance(pivot_v, list) else [])

    st.download_button(
        "下载月产值细项对比 CSV",
        data=value_csv,
        file_name="city_kpi_value_parts.csv",
        mime="text/csv",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
