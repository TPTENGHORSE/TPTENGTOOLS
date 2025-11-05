import os
import sys
import json
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required. Please install dependencies and re-run.")
    sys.exit(1)


EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


def find_target_file(path_hint: str) -> str | None:
    """
    If path_hint is a file, return it. If it's a directory, look for a file
    starting with 'QUOTATION TOOL DATA' and Excel extensions, pick the most recent.
    """
    if os.path.isfile(path_hint):
        return path_hint
    if not os.path.isdir(path_hint):
        return None
    candidates = []
    for name in os.listdir(path_hint):
        lower = name.lower()
        if lower.startswith("quotation tool data") and lower.endswith(EXCEL_EXTS):
            full = os.path.join(path_hint, name)
            try:
                mtime = os.path.getmtime(full)
            except Exception:
                mtime = 0
            candidates.append((mtime, full))
    if not candidates:
        # fallback: any excel file containing the phrase
        for name in os.listdir(path_hint):
            lower = name.lower()
            if "quotation tool data" in lower and lower.endswith(EXCEL_EXTS):
                full = os.path.join(path_hint, name)
                try:
                    mtime = os.path.getmtime(full)
                except Exception:
                    mtime = 0
                candidates.append((mtime, full))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


def summarize_excel(path: str) -> dict:
    summary: dict = {
        "file": path,
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "sheets": []
    }
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        summary["error"] = f"Failed to open Excel: {e}"
        return summary
    for sheet in xls.sheet_names:
        try:
            df_head = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            cols = list(df_head.columns.astype(str))
            sample = df_head.fillna("").astype(str).to_dict(orient="records")
            info = {
                "name": sheet,
                "columns": cols,
                "sample_rows": sample,
                "row_count_estimate": None,
            }
            # try to estimate row count cheaply
            try:
                df_small = pd.read_excel(xls, sheet_name=sheet, usecols=[0], engine=None)
                info["row_count_estimate"] = int(df_small.shape[0])
            except Exception:
                pass
            summary["sheets"].append(info)
        except Exception as e:
            summary["sheets"].append({
                "name": sheet,
                "error": str(e)
            })
    return summary


def write_outputs(summary: dict, out_dir: str):
    os.makedirs(out_dir, exist_ok=True)
    json_path = os.path.join(out_dir, "qtool_schema.json")
    md_path = os.path.join(out_dir, "qtool_schema.md")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    # Simple Markdown summary
    lines = []
    lines.append(f"# QUOTATION TOOL DATA summary\n")
    lines.append(f"File: `{summary.get('file','')}`  ")
    lines.append(f"Generated: {summary.get('generated_at','')}\n")
    if "error" in summary:
        lines.append(f"ERROR: {summary['error']}\n")
    for sheet in summary.get("sheets", []):
        lines.append(f"## Sheet: {sheet.get('name','')}\n")
        if "error" in sheet:
            lines.append(f"- Error: {sheet['error']}\n")
            continue
        cols = sheet.get("columns", [])
        lines.append(f"- Columns ({len(cols)}): {', '.join(cols)}")
        rcount = sheet.get("row_count_estimate")
        if rcount is not None:
            lines.append(f"- Approx. rows: {rcount}")
        sample = sheet.get("sample_rows", [])
        if sample:
            lines.append(f"- Sample (first 5 rows):\n")
            # render a tiny table-like block
            header = " | ".join(cols[:6])
            lines.append(f"  | {header} |\n  | {' | '.join(['---']*min(6,len(cols)))} |")
            for row in sample:
                vals = [str(row.get(c, "")) for c in cols[:6]]
                lines.append(f"  | {' | '.join(vals)} |")
        lines.append("")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Written: {json_path}")
    print(f"Written: {md_path}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python Quotations/inspect_qtool_data.py <path-to-file-or-folder>")
        sys.exit(2)
    hint = sys.argv[1]
    target = find_target_file(hint)
    if not target:
        print(f"ERROR: Could not find target Excel in '{hint}'.")
        sys.exit(3)
    summary = summarize_excel(target)
    write_outputs(summary, out_dir=os.path.join(os.path.dirname(__file__), "_out"))


if __name__ == "__main__":
    main()
