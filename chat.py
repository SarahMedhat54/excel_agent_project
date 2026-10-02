"""Excel chat without any LLM: a command parser on top of the tools in server.py."""
import json
import re

from openpyxl.utils import get_column_letter

from server import FILES_DIR, create_chart, create_excel, describe_excel, edit_excel, run_eda

HELP = """Commands:
  create Month,Sales,Profit | Jan,100,20 | Feb,130,30   -> new file (rows after each |)
  files                                                  -> list saved files
  open <file_id>                                         -> switch to a file
  show                                                   -> structure + first rows
  eda                                                    -> analysis
  add column Total = Sales*Profit                        -> column with formula (use column names)
  add column Notes                                       -> empty column
  delete column Notes
  rename column Notes to Comments
  add row Mar,90,15
  delete rows 3 2                                        -> from row 3, count 2
  set B2 = 10
  chart bar Month vs Sales,Profit title My Chart         -> bar | line | pie
  help / exit"""

state = {"file_id": None}


# ---------- helpers ----------
def val(s: str):
    """Convert text to int / float / formula / text."""
    s = s.strip()
    if s.startswith("="):
        return s
    for cast in (int, float):
        try:
            return cast(s)
        except ValueError:
            pass
    return s


def split_vals(s: str) -> list:
    return [val(x) for x in s.split(",")]


def need_file() -> str:
    if not state["file_id"]:
        raise ValueError("No active file. Use 'create ...' or 'open <file_id>' first.")
    return state["file_id"]


def to_formula(expr: str, headers: list) -> str:
    """'Sales*Profit' -> '=B{row}*C{row}' using the real column positions."""
    for i, h in sorted(enumerate(headers), key=lambda t: -len(str(t[1]))):
        expr = re.sub(re.escape(str(h)), f"{get_column_letter(i + 1)}{{row}}", expr, flags=re.I)
    return "=" + expr.strip()


def pretty(obj) -> str:
    return json.dumps(obj, indent=2, ensure_ascii=False)


# ---------- command handler ----------
def handle(line: str) -> str:
    low = line.lower().strip()

    if low in ("help", "?"):
        return HELP

    if low == "files":
        return ", ".join(p.stem for p in FILES_DIR.glob("*.xlsx")) or "No files yet."

    if m := re.match(r"open\s+(\w+)$", line, re.I):
        describe_excel(m.group(1))  # raises an error if the file doesn't exist
        state["file_id"] = m.group(1)
        return f"Opened {m.group(1)}"

    if m := re.match(r"create\s+(.+)", line, re.I):
        parts = [p.strip() for p in m.group(1).split("|")]
        columns = [c.strip() for c in parts[0].split(",")]
        rows = [split_vals(p) for p in parts[1:]]
        r = create_excel(columns, rows)
        state["file_id"] = r["file_id"]
        return f"Created {r['path']}  (file_id = {r['file_id']})"

    if low == "show":
        return pretty(describe_excel(need_file()))

    if low == "eda":
        return pretty(run_eda(need_file()))

    if m := re.match(r"add column\s+(.+?)\s*=\s*(.+)", line, re.I):
        fid = need_file()
        headers = describe_excel(fid)["sheets"]
        headers = list(headers.values())[0]["headers"]
        op = {"op": "add_column", "header": m.group(1).strip(), "formula": to_formula(m.group(2), headers)}
        return pretty(edit_excel(fid, [op]))

    if m := re.match(r"add column\s+(.+)", line, re.I):
        return pretty(edit_excel(need_file(), [{"op": "add_column", "header": m.group(1).strip()}]))

    if m := re.match(r"delete column\s+(.+)", line, re.I):
        return pretty(edit_excel(need_file(), [{"op": "delete_column", "header": m.group(1).strip()}]))

    if m := re.match(r"rename column\s+(.+?)\s+to\s+(.+)", line, re.I):
        op = {"op": "rename_column", "old": m.group(1).strip(), "new": m.group(2).strip()}
        return pretty(edit_excel(need_file(), [op]))

    if m := re.match(r"add row\s+(.+)", line, re.I):
        return pretty(edit_excel(need_file(), [{"op": "add_rows", "rows": [split_vals(m.group(1))]}]))

    if m := re.match(r"delete rows?\s+(\d+)(?:\s+(\d+))?", line, re.I):
        op = {"op": "delete_rows", "start": int(m.group(1)), "count": int(m.group(2) or 1)}
        return pretty(edit_excel(need_file(), [op]))

    if m := re.match(r"set\s+([A-Za-z]+\d+)\s*=\s*(.+)", line, re.I):
        op = {"op": "set_cell", "cell": m.group(1).upper(), "value": val(m.group(2))}
        return pretty(edit_excel(need_file(), [op]))

    if m := re.match(r"chart\s+(bar|line|pie)\s+(.+?)\s+vs\s+(.+?)(?:\s+title\s+(.+))?$", line, re.I):
        y_cols = [c.strip() for c in m.group(3).split(",")]
        r = create_chart(need_file(), m.group(1).lower(), m.group(2).strip(), y_cols, title=m.group(4) or "")
        return pretty(r)

    return "I didn't understand that. Type 'help' to see the commands."


def main():
    print("Excel chat (no LLM). Type 'help' for commands, 'exit' to quit.")
    while True:
        line = input("\nYou: ").strip()
        if not line:
            continue
        if line.lower() in ("exit", "quit"):
            break
        try:
            print("\nBot:", handle(line))
        except Exception as e:
            print("\n[error]", e)


if __name__ == "__main__":
    main()