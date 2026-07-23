# XML-like TXT <reference> Diff Tool - Core Python Version
# No UI. Same calculation/diff workflow as the R-Shiny app.
#
# Usage:
#   python xml_reference_diff_core.py old.txt new.txt
#   python xml_reference_diff_core.py old.txt new.txt --out diff.xlsx
#   python xml_reference_diff_core.py old.txt new.txt --out diff.csv
#
# Output columns:
#   group_id, group_title, diff_status, line_key, file0_summary_line, file1_summary_line, color_diff_memo

import argparse
import html
import os
import re
from typing import Dict, List

import pandas as pd


# ============================================================
# Core logic converted from R make_diff_table()
# ============================================================

def xml_unescape(x: str) -> str:
    if x is None:
        return ""
    return html.unescape(str(x)).replace("\u00a0", " ")


def str_squish(x: str) -> str:
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x)).strip()


def clean_text(x) -> str:
    if x is None:
        return ""
    if isinstance(x, (list, tuple)):
        x = " ".join(str(v) for v in x)
    else:
        x = str(x)

    x = re.sub(r"<!--.*?-->", " ", x, flags=re.S)
    # Same as R: self-closing xref is treated as blank.
    x = re.sub(r"<xref\b[^>]*/>", " ", x, flags=re.I)
    x = re.sub(r"<[^>]+>", " ", x, flags=re.S)
    x = xml_unescape(x)
    return str_squish(x)


def get_attr(tag_txt: str, attr_name: str) -> str:
    if tag_txt is None or str(tag_txt) == "":
        return ""
    pattern = rf"\b{re.escape(attr_name)}\s*=\s*(['\"])(.*?)\1"
    m = re.search(pattern, str(tag_txt), flags=re.S)
    return m.group(2) if m else ""


def read_txt_file(path: str) -> str:
    # Mirrors R readLines(..., encoding='UTF-8') but is tolerant of odd bytes.
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def extract_references(txt: str) -> pd.DataFrame:
    ref_blocks = re.findall(r"<reference\b[^>]*>.*?</reference>", txt, flags=re.S | re.I)
    if not ref_blocks:
        return pd.DataFrame(columns=["group_id", "group_title", "block"])

    rows = []
    for i, block in enumerate(ref_blocks, start=1):
        m_start = re.search(r"<reference\b[^>]*>", block, flags=re.S | re.I)
        start_tag = m_start.group(0) if m_start else ""
        rid = get_attr(start_tag, "id") or f"NO_ID_{i}"

        m_title = re.search(r"<title\b[^>]*>(.*?)</title>", block, flags=re.S | re.I)
        title = clean_text(m_title.group(1)) if m_title else ""
        rows.append({"group_id": rid, "group_title": title, "block": block})

    return pd.DataFrame(rows, columns=["group_id", "group_title", "block"])


def parse_row_to_line(row_block: str) -> str:
    entry_pattern = re.compile(
        r"<entry\b([^>]*)>(.*?)</entry>|<entry\b([^>]*)/>",
        flags=re.S | re.I,
    )
    entries = list(entry_pattern.finditer(row_block))
    if not entries:
        return ""

    tmp = []
    for j, m in enumerate(entries, start=1):
        attrs = m.group(1) if m.group(1) is not None else m.group(3)
        val_raw = m.group(2) if m.group(2) is not None else ""
        cn = get_attr(f"<x {attrs or ''}>", "colname")
        if not cn:
            cn = f"col{900 + j}"

        val = clean_text(val_raw)
        col_num_match = re.search(r"\d+", cn)
        col_order = int(col_num_match.group(0)) if col_num_match else 999
        tmp.append((col_order, cn, val))

    tmp.sort(key=lambda z: z[0])
    vals = [val for _, _, val in tmp if val != ""]
    return str_squish(" ".join(vals))


def unique_keep_order(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def parse_group_lines(block: str) -> List[str]:
    lines = []

    # Table-like content: one <row> becomes one summarized line.
    row_blocks = re.findall(r"<row\b[^>]*>.*?</row>", block, flags=re.S | re.I)
    for row in row_blocks:
        one_line = parse_row_to_line(row)
        if one_line:
            lines.append(one_line)

    # Paragraph/list/bold content outside tables.
    block_no_table = re.sub(r"<table\b[^>]*>.*?</table>", " ", block, flags=re.S | re.I)
    for tg in ["p", "li", "b"]:
        pattern = re.compile(rf"<{tg}\b[^>]*>(.*?)</{tg}>", flags=re.S | re.I)
        for m in pattern.finditer(block_no_table):
            val = clean_text(m.group(1))
            if val:
                lines.append(val)

    lines = [str_squish(x) for x in lines if str_squish(x)]
    return unique_keep_order(lines)


def line_key(x: str) -> str:
    x = str_squish(x)
    return x.split()[0] if x else ""


def first_ref_row(refs: pd.DataFrame, rid: str) -> Dict[str, str] | None:
    hit = refs[refs["group_id"] == rid]
    if hit.empty:
        return None
    return hit.iloc[0].to_dict()


def split_by_line_key(lines: List[str]) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {}
    for ln in lines:
        k = line_key(ln)
        out.setdefault(k, []).append(ln)
    return out


def make_diff_table(file0: str, file1: str) -> pd.DataFrame:
    txt0 = read_txt_file(file0)
    txt1 = read_txt_file(file1)

    refs0 = extract_references(txt0)
    refs1 = extract_references(txt1)

    ref_ids = sorted(set(refs0["group_id"].tolist()) | set(refs1["group_id"].tolist()))
    result = []

    for rid in ref_ids:
        r0 = first_ref_row(refs0, rid)
        r1 = first_ref_row(refs1, rid)

        if r1 is not None and r1.get("group_title", ""):
            title = r1["group_title"]
        elif r0 is not None and r0.get("group_title", ""):
            title = r0["group_title"]
        else:
            title = ""

        if r0 is None:
            l1 = parse_group_lines(r1["block"])
            if not l1:
                l1 = ["[GROUP EXISTS BUT NO LOWEST TAG TEXT PARSED]"]
            for ln in l1:
                result.append({
                    "group_id": rid,
                    "group_title": title,
                    "diff_status": "GROUP_ADDED",
                    "line_key": line_key(ln),
                    "file0_summary_line": "",
                    "file1_summary_line": ln,
                })
            continue

        if r1 is None:
            l0 = parse_group_lines(r0["block"])
            if not l0:
                l0 = ["[GROUP EXISTS BUT NO LOWEST TAG TEXT PARSED]"]
            for ln in l0:
                result.append({
                    "group_id": rid,
                    "group_title": title,
                    "diff_status": "GROUP_REMOVED",
                    "line_key": line_key(ln),
                    "file0_summary_line": ln,
                    "file1_summary_line": "",
                })
            continue

        l0 = parse_group_lines(r0["block"])
        l1 = parse_group_lines(r1["block"])

        d0 = split_by_line_key(l0)
        d1 = split_by_line_key(l1)
        keys = sorted(set(d0.keys()) | set(d1.keys()))

        for k in keys:
            a = d0.get(k, [])
            b = d1.get(k, [])
            if a == b:
                continue

            maxn = max(len(a), len(b))
            for i in range(maxn):
                va = a[i] if i < len(a) else ""
                vb = b[i] if i < len(b) else ""
                if va == vb:
                    continue

                if va and vb:
                    st = "MODIFIED"
                elif va and not vb:
                    st = "REMOVED"
                else:
                    st = "ADDED"

                result.append({
                    "group_id": rid,
                    "group_title": title,
                    "diff_status": st,
                    "line_key": k,
                    "file0_summary_line": va,
                    "file1_summary_line": vb,
                })

    cols = ["group_id", "group_title", "diff_status", "line_key", "file0_summary_line", "file1_summary_line"]
    df = pd.DataFrame(result, columns=cols)
    if df.empty:
        return df
    return df.sort_values(["group_id", "diff_status", "line_key"], kind="stable").reset_index(drop=True)


# ============================================================
# Preview/export helper: same as right-side Tab 1 preview table
# ============================================================

def split_words_for_display(x: str) -> List[str]:
    x = str_squish(x or "")
    return x.split() if x else []


def make_color_diff_memo(status: str, old_line: str, new_line: str) -> str:
    status = str(status or "")
    old_line = str_squish(old_line or "")
    new_line = str_squish(new_line or "")

    if status in ["ADDED", "GROUP_ADDED"]:
        return "New row/text added in file1; green words show added content."
    if status in ["REMOVED", "GROUP_REMOVED"]:
        return "Old row/text removed from file0; red words show removed content."
    if status == "MODIFIED":
        old_tokens = split_words_for_display(old_line)
        new_tokens = split_words_for_display(new_line)
        removed = []
        added = []
        for tok in old_tokens:
            if tok not in new_tokens and tok not in removed:
                removed.append(tok)
        for tok in new_tokens:
            if tok not in old_tokens and tok not in added:
                added.append(tok)
        removed_txt = " ".join(removed[:8])
        added_txt = " ".join(added[:8])
        if removed_txt and added_txt:
            return f"Modified: red removed/changed from file0 [{removed_txt}]; green added/changed in file1 [{added_txt}]."
        if removed_txt:
            return f"Modified: red words removed/changed from file0 [{removed_txt}]."
        if added_txt:
            return f"Modified: green words added/changed in file1 [{added_txt}]."
        return "Modified: line text changed; color tab highlights the different words."
    return "Color tab: red = file0 removed/old difference; green = file1 added/new difference."


def make_plain_preview_table(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["group_id", "group_title", "diff_status", "line_key", "file0_summary_line", "file1_summary_line"]
    if df.empty:
        out = pd.DataFrame(columns=cols + ["color_diff_memo"])
        return out

    out = df.loc[:, cols].copy()
    out["color_diff_memo"] = [
        make_color_diff_memo(st, old, new)
        for st, old, new in zip(out["diff_status"], out["file0_summary_line"], out["file1_summary_line"])
    ]
    return out



def save_output(df: pd.DataFrame, out_path: str) -> None:
    ext = os.path.splitext(out_path)[1].lower()
    if ext in [".xlsx", ".xlsm"]:
        df.to_excel(out_path, index=False, engine="openpyxl")
    elif ext == ".csv":
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:
        raise ValueError("Output file must be .xlsx or .csv")


# ============================================================
# Color HTML export helper
# - Same diff result, only display color markup for easy review.
# - Red = old/file0 removed or changed token.
# - Green = new/file1 added or changed token.
# ============================================================

def escape_html(x: str) -> str:
    return html.escape("" if x is None else str(x), quote=True)


def highlight_line_diff(old_line: str, new_line: str, side: str) -> str:
    old_tokens = split_words_for_display(old_line)
    new_tokens = split_words_for_display(new_line)

    if side == "old" and not old_tokens:
        return ""
    if side == "new" and not new_tokens:
        return ""

    if side == "old":
        keep_counts = {}
        for t in new_tokens:
            keep_counts[t] = keep_counts.get(t, 0) + 1
        used = {}
        out = []
        for tok in old_tokens:
            used[tok] = used.get(tok, 0) + 1
            safe = escape_html(tok)
            if used[tok] > keep_counts.get(tok, 0):
                out.append(f'<span class="diff-del">{safe}</span>')
            else:
                out.append(safe)
        return " ".join(out)

    keep_counts = {}
    for t in old_tokens:
        keep_counts[t] = keep_counts.get(t, 0) + 1
    used = {}
    out = []
    for tok in new_tokens:
        used[tok] = used.get(tok, 0) + 1
        safe = escape_html(tok)
        if used[tok] > keep_counts.get(tok, 0):
            out.append(f'<span class="diff-add">{safe}</span>')
        else:
            out.append(safe)
    return " ".join(out)


def make_color_preview_table(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["group_id", "group_title", "diff_status", "line_key", "file0_summary_line", "file1_summary_line"]
    if df.empty:
        return pd.DataFrame(columns=cols)

    out = pd.DataFrame()
    out["group_id"] = df["group_id"].map(escape_html)
    out["group_title"] = df["group_title"].map(escape_html)
    out["diff_status"] = df["diff_status"].map(escape_html)
    out["line_key"] = df["line_key"].map(escape_html)
    out["file0_summary_line"] = [
        highlight_line_diff(old, new, "old")
        for old, new in zip(df["file0_summary_line"], df["file1_summary_line"])
    ]
    out["file1_summary_line"] = [
        highlight_line_diff(old, new, "new")
        for old, new in zip(df["file0_summary_line"], df["file1_summary_line"])
    ]
    return out


def write_color_diff_html(df: pd.DataFrame, out_path: str) -> None:
    color_df = make_color_preview_table(df)
    css = """
    body { font-family: Arial, sans-serif; margin: 18px; background: white; }
    h2 { margin: 0 0 10px 0; font-size: 22px; }
    .memo { margin: 0 0 14px 0; font-size: 13px; color: #333; }
    table { border-collapse: collapse; width: 100%; table-layout: auto; font-size: 12px; }
    th { background: #D9EAF7; font-weight: bold; }
    th, td { border: 1px solid #999; padding: 5px 7px; vertical-align: top; white-space: nowrap; }
    td.wrap-cell { white-space: normal; word-break: break-word; overflow-wrap: anywhere; min-width: 420px; max-width: 760px; line-height: 1.35; }
    .diff-del { background-color: #ffd6d6; text-decoration: line-through; padding: 1px 2px; border-radius: 3px; }
    .diff-add { background-color: #d8f8d8; padding: 1px 2px; border-radius: 3px; }
    tr.added td, tr.group_added td { background: #F4FFF4; }
    tr.removed td, tr.group_removed td { background: #FFF4F4; }
    tr.modified td { background: #FFFDF0; }
    """

    if color_df.empty:
        body = "<h2>XML Reference Color Diff</h2><p>No diff rows found.</p>"
    else:
        headers = "".join(f"<th>{escape_html(c)}</th>" for c in color_df.columns)
        rows = []
        for _, r in color_df.iterrows():
            cls = re.sub(r"[^a-z0-9_]+", "_", str(r["diff_status"]).lower())
            cells_list = []
            for c in color_df.columns:
                if c in ["file0_summary_line", "file1_summary_line"]:
                    cells_list.append(f'<td class="wrap-cell">{r[c]}</td>')
                else:
                    cells_list.append(f"<td>{r[c]}</td>")
            cells = "".join(cells_list)
            rows.append(f'<tr class="{cls}">{cells}</tr>')
        body = (
            "<h2>XML Reference Color Diff Preview</h2>"
            "<p class='memo'>Red = file0 old/removed difference; green = file1 new/added difference.</p>"
            f"<table><thead><tr>{headers}</tr></thead><tbody>{''.join(rows)}</tbody></table>"
        )

    html_text = f"<!DOCTYPE html><html><head><meta charset='UTF-8'><style>{css}</style></head><body>{body}</body></html>"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_text)


def find_default_input_files(folder: str) -> tuple[str, str]:
    old_path = os.path.join(folder, "old.txt")
    new_path = os.path.join(folder, "new.txt")
    if not os.path.exists(old_path) or not os.path.exists(new_path):
        raise FileNotFoundError(
            "Cannot find old.txt and new.txt in this folder.\n"
            f"Current folder: {folder}\n\n"
            "Please put these 3 files in the same folder:\n"
            "  1) this .py file\n"
            "  2) old.txt\n"
            "  3) new.txt\n\n"
            "Then run: python xml_reference_diff_auto_folder.py"
        )
    return old_path, new_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Auto diff old.txt and new.txt in the same folder. UI not required."
    )
    parser.add_argument("file0", nargs="?", default="", help="Optional old TXT/XML/DITA path. Default: ./old.txt")
    parser.add_argument("file1", nargs="?", default="", help="Optional new TXT/XML/DITA path. Default: ./new.txt")
    parser.add_argument("--out-prefix", default="xml_reference_diff", help="Output prefix. Default: xml_reference_diff")
    parser.add_argument("--raw", action="store_true", help="Export raw diff table without color_diff_memo column")
    parser.add_argument("--max-rows", type=int, default=20, help="Rows to print in terminal. Default: 20")
    args = parser.parse_args()

    # User workflow: cd to the folder and run python. Use current working folder.
    work_dir = os.getcwd()

    if args.file0 and args.file1:
        file0 = os.path.abspath(args.file0)
        file1 = os.path.abspath(args.file1)
    else:
        file0, file1 = find_default_input_files(work_dir)

    print("XML Reference Diff Tool")
    print("Working folder:", work_dir)
    print("Old file:", file0)
    print("New file:", file1)
    print("Running diff...\n")

    raw_df = make_diff_table(file0, file1)
    preview_df = raw_df if args.raw else make_plain_preview_table(raw_df)

    out_xlsx = os.path.join(work_dir, f"{args.out_prefix}_preview.xlsx")
    out_csv = os.path.join(work_dir, f"{args.out_prefix}_preview.csv")
    out_html = os.path.join(work_dir, f"{args.out_prefix}_color_diff.html")

    save_output(preview_df, out_xlsx)
    save_output(preview_df, out_csv)
    write_color_diff_html(raw_df, out_html)

    print(f"Total diff rows: {len(raw_df)}")
    if not raw_df.empty:
        print(raw_df["diff_status"].value_counts().to_string())
        print("\nPreview:")
        with pd.option_context("display.max_columns", None, "display.width", 220, "display.max_colwidth", 90):
            print(preview_df.head(args.max_rows).to_string(index=False))
    else:
        print("No diff rows found.")

    print("\nGenerated files:")
    print(" -", out_xlsx)
    print(" -", out_csv)
    print(" -", out_html)


if __name__ == "__main__":
    main()
