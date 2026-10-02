"""Tiny Excel generator: type a request, the file downloads. Uses agent.py + server.py."""
import asyncio
import re

from flask import Flask, jsonify, request, send_file

from agent import run_agent
from server import FILES_DIR

app = Flask(__name__)
histories: dict = {}


def clean_error(e: Exception) -> str:
    while hasattr(e, "exceptions") and e.exceptions:
        e = e.exceptions[0]
    msg = str(e)
    if "credit balance" in msg:
        return "Your Anthropic API credit balance is too low."
    return msg


@app.get("/")
def index():
    return PAGE


@app.post("/api/chat")
def chat():
    d = request.get_json()
    sid, text, fid = d["sid"], d["message"].strip(), d.get("file_id")
    if not text:
        return jsonify(error="Write something first."), 400
    msgs = histories.setdefault(sid, [])
    msgs.append({"role": "user", "content": f"[Current file_id: {fid}]\n{text}" if fid else text})
    mark = len(msgs) - 1
    try:
        reply, file_ids = asyncio.run(run_agent(msgs))
    except Exception as e:
        del msgs[mark:]
        return jsonify(error=clean_error(e)), 500
    return jsonify(reply=reply, file_id=file_ids[-1] if file_ids else None)


@app.get("/download/<file_id>")
def download(file_id):
    path = FILES_DIR / f"{file_id}.xlsx"
    if not re.fullmatch(r"\w+", file_id) or not path.exists():
        return "not found", 404
    return send_file(path, as_attachment=True, download_name=f"{file_id}.xlsx")


PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>Excel Maker</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{--bg:#f4f6f9;--card:#fff;--ink:#1f2937;--muted:#6b7280;--line:#d1d5db;--brand:#1f4e78;--bad:#b91c1c}
@media (prefers-color-scheme:dark){:root{--bg:#0f172a;--card:#1e293b;--ink:#e5e7eb;--muted:#94a3b8;--line:#334155;--brand:#4f8fc9;--bad:#f87171}}
body{margin:0;min-height:100vh;display:grid;place-items:center;background:var(--bg);color:var(--ink);font:15px/1.5 "Segoe UI",system-ui,Arial,sans-serif}
.card{width:min(520px,92vw);background:var(--card);border:1px solid var(--line);border-radius:16px;padding:22px}
h1{margin:0 0 4px;font-size:20px;color:var(--brand)}
p{margin:0 0 14px;color:var(--muted);font-size:13.5px}
textarea{width:100%;height:110px;resize:vertical;padding:10px;border:1px solid var(--line);border-radius:10px;background:var(--card);color:var(--ink);font:inherit}
textarea:focus{outline:2px solid var(--brand);outline-offset:-1px}
.row{display:flex;gap:8px;margin-top:10px}
button{border:0;border-radius:10px;padding:10px 16px;font:inherit;font-weight:600;cursor:pointer}
#go{flex:1;background:var(--brand);color:#fff}
#go:disabled{opacity:.5;cursor:default}
#new{background:transparent;color:var(--brand);border:1px solid var(--line)}
#out{margin-top:14px;white-space:pre-wrap;font-size:14px}
#out.err{color:var(--bad)}
#out a{display:inline-block;margin-top:8px;padding:7px 12px;border-radius:8px;background:var(--brand);color:#fff;text-decoration:none;font-weight:600}
#cur{font-size:12.5px;color:var(--muted);margin-top:8px}
</style></head>
<body><div class="card">
  <h1>📊 Excel Maker</h1>
  <p>Describe the Excel file you want. To change it afterwards, just write the change.</p>
  <textarea id="box" placeholder="e.g. Create a monthly expenses sheet with date, category and amount, with 6 sample rows"></textarea>
  <div class="row"><button id="go">Generate</button><button id="new" title="Start a new file">New file</button></div>
  <div id="cur"></div>
  <div id="out"></div>
</div>
<script>
const $ = id => document.getElementById(id);
let sid = crypto.randomUUID ? crypto.randomUUID() : String(Math.random()), fid = null;

function setFile(id) { fid = id; $('cur').textContent = id ? 'Editing the current file. Click "New file" to start over.' : ''; }

$('new').onclick = () => { sid = crypto.randomUUID ? crypto.randomUUID() : String(Math.random()); setFile(null); $('out').textContent = ''; $('box').value = ''; $('box').focus(); };

$('go').onclick = async () => {
  const text = $('box').value.trim(); if (!text) return;
  $('go').disabled = true; $('out').className = ''; $('out').textContent = 'Working on it...';
  try {
    const r = await fetch('/api/chat', {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({sid, message: text, file_id: fid})});
    const d = await r.json();
    if (!r.ok) { $('out').className = 'err'; $('out').textContent = d.error; }
    else {
      $('out').textContent = d.reply || 'Done.';
      if (d.file_id) {
        setFile(d.file_id);
        const a = document.createElement('a'); a.href = '/download/' + d.file_id; a.textContent = '⬇ Download Excel';
        $('out').appendChild(document.createElement('br')); $('out').appendChild(a);
        a.click();  // download starts automatically
      }
      $('box').value = '';
    }
  } catch (e) { $('out').className = 'err'; $('out').textContent = 'Could not reach the server.'; }
  $('go').disabled = false;
};
</script></body></html>"""


if __name__ == "__main__":
    app.run(port=5000, debug=False)