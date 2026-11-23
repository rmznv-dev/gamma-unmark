#!/usr/bin/env python3
"""
gamma-unmark — remove the "Made with Gamma" badge from exported .pptx and .pdf files.

Local web UI.  PPTX uses only the standard library.
PDF uses PyMuPDF (pip install pymupdf).

    python gu.py

For educational purposes only. See the Disclaimer in README.md.
"""

import argparse
import io
import json
import os
import re
import secrets
import socket
import sys
import threading
import urllib.parse
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import pymupdf
except ImportError:  # older PyMuPDF releases
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None

HERE = os.path.dirname(os.path.abspath(__file__))
MAX_UPLOAD = 1024 * 1024 * 1024  # 1 GiB

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
PDF_MIME = "application/pdf"

# Static files the UI may request (nothing else is ever served from disk).
ASSETS = {
    "/logo.png": ("logo-dark.png", "image/png"),
    "/favicon.png": ("favicon.png", "image/png"),
}


def _asset_path(filename):
    """assets/<file> first, then next to gu.py."""
    for base in (os.path.join(HERE, "assets"), HERE):
        p = os.path.join(base, filename)
        if os.path.isfile(p):
            return p
    return None


# =========================================================================== #
#  PPTX cleaning (standard library)
# =========================================================================== #

XML_PARTS = re.compile(
    r"^ppt/(slides|slideLayouts|slideMasters|notesSlides|notesMasters)/[^/]+\.xml$"
)
REL_TAG = re.compile(r"<Relationship\b[^>]*>", re.I)
GAMMA_URL = re.compile(r"https?://([\w-]+\.)*gamma\.app", re.I)
SHAPE = re.compile(r"<p:(sp|pic)\b[^>]*>.*?</p:\1>", re.S)
TEXT_RUN = re.compile(r"<a:t\b[^>]*>(.*?)</a:t>", re.S)
RID_ATTR = re.compile(r'r:(?:id|embed|link)="([^"]+)"')
CNV_PR = re.compile(r"<p:cNvPr\b[^>]*>")
MADE_WITH = re.compile(r"made\s+with\s+gamma", re.I)


def _attr(tag, name):
    m = re.search(r'\b%s="([^"]*)"' % name, tag)
    return m.group(1) if m else ""


def _gamma_rids(rels_xml):
    """Relationship ids that point at gamma.app."""
    out = set()
    for tag in REL_TAG.findall(rels_xml):
        if GAMMA_URL.search(_attr(tag, "Target")):
            out.add(_attr(tag, "Id"))
    return out


def _is_badge(kind, xml, gamma_rids):
    text = " ".join(TEXT_RUN.findall(xml)).strip()
    meta = " ".join(
        _attr(t, "name") + " " + _attr(t, "descr") for t in CNV_PR.findall(xml)
    )
    linked = any(r in gamma_rids for r in RID_ATTR.findall(xml))

    if MADE_WITH.search(text) or MADE_WITH.search(meta):
        return True
    if linked and (kind == "pic" or len(text) <= 40):
        return True
    if kind == "pic" and "gamma" in meta.lower():
        return True
    return False


def clean_pptx(data):
    """Return (cleaned_bytes, removed_count). Raises ValueError on bad input."""
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise ValueError("Not a valid .pptx file")

    names = set(zin.namelist())
    if "[Content_Types].xml" not in names or not any(
        n.startswith("ppt/") for n in names
    ):
        raise ValueError("Not a PowerPoint file")

    out = io.BytesIO()
    removed = 0
    with zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if XML_PARTS.match(item.filename):
                folder, base = item.filename.rsplit("/", 1)
                rels_name = folder + "/_rels/" + base + ".rels"
                rels = (
                    zin.read(rels_name).decode("utf-8", "replace")
                    if rels_name in names
                    else ""
                )
                rids = _gamma_rids(rels)
                xml = blob.decode("utf-8")
                count = [0]

                def repl(m):
                    if _is_badge(m.group(1), m.group(0), rids):
                        count[0] += 1
                        return ""
                    return m.group(0)

                new_xml = SHAPE.sub(repl, xml)
                if count[0]:
                    removed += count[0]
                    blob = new_xml.encode("utf-8")
            zout.writestr(item, blob)
    return out.getvalue(), removed


# =========================================================================== #
#  PDF cleaning (PyMuPDF)
# =========================================================================== #
#
#  1. Find the badge area on every page: the clickable gamma.app link
#     (fallback: the text "Made with Gamma").
#  2. Delete small images inside that area, then redact everything else that
#     is fully inside it: text, vector shapes, also inside nested objects.
#  3. Remove the link itself.
#  4. Re-render the area to verify that something really changed.

PAD = 3.0  # points of breathing room around the badge area


def _need_pymupdf():
    if pymupdf is None:
        raise ValueError("PDF support needs PyMuPDF:  pip install pymupdf")


def _badge_rects(page):
    rects, links = [], []
    for ln in page.get_links():
        uri = ln.get("uri") or ""
        if "gamma.app" in uri.lower() and ln.get("from") is not None:
            rects.append(pymupdf.Rect(ln["from"]))
            links.append(ln)
    if not rects:  # no link: look for the text itself
        for hit in page.search_for("made with gamma"):
            rects.append(pymupdf.Rect(hit) + (-12, -6, 12, 6))
    return rects, links


def _snapshot(page, rects):
    out = []
    for r in rects:
        pix = page.get_pixmap(clip=r, dpi=72, alpha=False)
        out.append(bytes(pix.samples))
    return out


def clean_pdf(data):
    """Return (cleaned_bytes_or_None, status, changed_visually)."""
    _need_pymupdf()
    try:
        doc = pymupdf.open(stream=data, filetype="pdf")
    except Exception:
        raise ValueError("Not a valid .pdf file")
    try:
        if doc.needs_pass or doc.is_encrypted:
            raise ValueError("Encrypted PDFs are not supported")
        if not doc.is_pdf:
            raise ValueError("Not a valid .pdf file")

        pages_hit = pages_changed = links_removed = 0
        line_art = getattr(pymupdf, "PDF_REDACT_LINE_ART_REMOVE_IF_COVERED", None)

        for page in doc:
            rects, links = _badge_rects(page)
            if not rects:
                continue
            pages_hit += 1
            pads = [r + (-PAD, -PAD, PAD, PAD) for r in rects]
            before = _snapshot(page, pads)

            # small images (logo-like) that sit inside the badge area
            try:
                for info in page.get_image_info(xrefs=True):
                    box = pymupdf.Rect(info["bbox"])
                    if not info.get("xref") or box.is_empty:
                        continue
                    for pad in pads:
                        inter = box & pad
                        if inter.is_empty:
                            continue
                        center = (box.tl + box.br) / 2
                        if pad.contains(center) and box.get_area() <= pad.get_area() * 1.6:
                            page.delete_image(info["xref"])
                            break
            except Exception:
                pass

            for pad in pads:
                page.add_redact_annot(pad)
            kw = {"images": pymupdf.PDF_REDACT_IMAGE_NONE}
            if line_art is not None:
                kw["graphics"] = line_art
            page.apply_redactions(**kw)

            for ln in links:
                try:
                    page.delete_link(ln)
                    links_removed += 1
                except Exception:
                    pass

            after = _snapshot(page, pads)
            if before != after:
                pages_changed += 1

        if not pages_hit:
            return None, "No watermark found", False

        out = doc.tobytes(garbage=3, deflate=True)
    finally:
        doc.close()

    if pages_changed:
        return out, "Badge removed on %d page%s" % (pages_changed, "" if pages_changed == 1 else "s"), True
    return out, "Link removed, badge graphic not detected", False


# =========================================================================== #
#  Unified entry point
# =========================================================================== #

def sniff(data):
    if data[:5] == b"%PDF-":
        return "pdf"
    if data[:2] == b"PK":
        return "pptx"
    return None


def clean_any(data):
    """Return dict(kind, data|None, status, ok). Raises ValueError."""
    kind = sniff(data)
    if kind == "pptx":
        out, n = clean_pptx(data)
        if not n:
            return {"kind": kind, "data": None, "status": "No watermark found", "ok": False}
        return {
            "kind": kind, "data": out, "ok": True,
            "status": "%d element%s removed" % (n, "" if n == 1 else "s"),
        }
    if kind == "pdf":
        out, status, ok = clean_pdf(data)
        return {"kind": kind, "data": out, "status": status, "ok": ok}
    raise ValueError("Only .pptx and .pdf files are supported")


# =========================================================================== #
#  Web UI
# =========================================================================== #

HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>gamma-unmark</title>
<link rel="icon" href="/favicon.png">
<style>
:root{--bg:#0a0a0a;--card:#141414;--line:#242424;--text:#f2f2f2;--muted:#8a8a8a;--lime:#c6ff00;--lime-dim:#c6ff0022}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:"Space Grotesk",Inter,system-ui,-apple-system,"Segoe UI",sans-serif;
  -webkit-font-smoothing:antialiased;min-height:100vh;display:flex;flex-direction:column}
body::before{content:"";position:fixed;left:50%;top:-220px;width:700px;height:700px;transform:translateX(-50%);
  background:radial-gradient(circle,#c6ff0016,transparent 65%);pointer-events:none}
.mono{font-family:ui-monospace,"JetBrains Mono",SFMono-Regular,Menlo,Consolas,monospace}
.wrap{width:100%;max-width:720px;margin:0 auto;padding:0 20px;position:relative}
.top{display:flex;justify-content:center;padding:22px 0 0}
.top img{width:min(300px,70vw);height:auto;display:block}
main{flex:1;padding:6px 0 30px;text-align:center}
h1{font-size:clamp(36px,7vw,60px);line-height:1;letter-spacing:-.045em}
h1 em{font-style:normal;color:var(--lime)}

.drop{margin-top:38px;border:1.5px dashed #333;border-radius:22px;background:var(--card);padding:50px 20px;cursor:pointer;transition:.2s}
.drop:hover,.drop.over{border-color:var(--lime);background:#c6ff000a;box-shadow:0 0 0 5px var(--lime-dim)}
.drop b{display:block;font-size:20px;font-weight:600}
.drop span{display:block;margin-top:8px;color:var(--muted);font-size:14px}
.plus{width:54px;height:54px;margin:0 auto 18px;border-radius:16px;background:var(--lime);color:#000;display:grid;place-items:center;font-size:30px;font-weight:700}

#list{margin-top:22px;display:flex;flex-direction:column;gap:10px;text-align:left}
.item{display:flex;align-items:center;gap:14px;background:var(--card);border:1px solid var(--line);border-radius:16px;padding:14px 16px}
.item .info{flex:1;min-width:0}
.item .nm{font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.item .st{font-size:13px;color:var(--muted);margin-top:2px}
.item .st.ok{color:var(--lime)}
.item .st.warn{color:#ffb74d}
.item .st.err{color:#ff6b6b}
.item.done{border-color:#c6ff0055}
.btn{background:var(--lime);color:#000;border:0;border-radius:11px;font:700 14px inherit;font-family:inherit;padding:11px 20px;cursor:pointer;transition:.15s;white-space:nowrap}
.btn:hover{transform:translateY(-1px);box-shadow:0 6px 24px #c6ff0044}
.spin{width:18px;height:18px;border:2px solid #333;border-top-color:var(--lime);border-radius:50%;animation:s .7s linear infinite}
@keyframes s{to{transform:rotate(360deg)}}
footer{padding:26px 0 36px;text-align:center;color:#555;font-size:12px;line-height:1.7}
</style>
</head>
<body>
<div class="wrap top"><img src="/logo.png" alt="gamma-unmark"></div>
<main class="wrap">
  <h1>Drop. <em>Clean.</em> Save.</h1>

  <div class="drop" id="drop">
    <div class="plus">+</div>
    <b>Drop .pptx or .pdf files here</b>
    <span>or click to browse</span>
  </div>
  <input type="file" id="file" accept=".pptx,.pdf" multiple hidden>

  <div id="list"></div>
</main>
<footer class="mono">runs locally · nothing leaves your machine<br>for educational purposes · use at your own responsibility</footer>

<script>
const $ = s => document.querySelector(s);
const drop = $('#drop'), input = $('#file'), list = $('#list');

drop.onclick = () => input.click();
input.onchange = () => { handle(input.files); input.value = ''; };
['dragenter','dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', ev => handle(ev.dataTransfer.files));
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());

function handle(files){ [...files].forEach(processFile); }

function row(name){
  const el = document.createElement('div');
  el.className = 'item';
  el.innerHTML = '<div class="info"><div class="nm"></div><div class="st">Cleaning…</div></div><div class="spin"></div>';
  el.querySelector('.nm').textContent = name;
  list.prepend(el);
  return el;
}

function finish(el, r, j){
  const st = el.querySelector('.st');
  el.querySelector('.spin').remove();
  if(!r.ok){ st.textContent = j.error || 'Failed'; st.className = 'st err'; return; }
  st.textContent = j.status;
  if(!j.id) return;
  st.className = 'st ' + (j.clean ? 'ok' : 'warn');
  if(j.clean) el.classList.add('done');
  const b = document.createElement('button');
  b.className = 'btn'; b.textContent = 'Save as…';
  b.onclick = () => save(j);
  el.appendChild(b);
}

async function processFile(file){
  const el = row(file.name);
  if(!/\.(pptx|pdf)$/i.test(file.name)){
    el.querySelector('.spin').remove();
    const st = el.querySelector('.st'); st.textContent = 'Only .pptx and .pdf are supported'; st.className = 'st err'; return;
  }
  try{
    const r = await fetch('/api/clean', {method:'POST', headers:{'X-Filename': encodeURIComponent(file.name)}, body:file});
    finish(el, r, await r.json());
  }catch(e){ finish(el, {ok:false}, {error:'Request failed'}); }
}

async function save(j){
  const url = '/api/download/' + j.id;
  const ext = j.name.split('.').pop().toLowerCase();
  const types = ext === 'pdf'
    ? [{description:'PDF', accept:{'application/pdf':['.pdf']}}]
    : [{description:'PowerPoint', accept:{'application/vnd.openxmlformats-officedocument.presentationml.presentation':['.pptx']}}];
  if(window.showSaveFilePicker){
    try{
      const h = await showSaveFilePicker({suggestedName: j.name, types});
      const blob = await (await fetch(url)).blob();
      const w = await h.createWritable(); await w.write(blob); await w.close();
    }catch(e){ if(e.name !== 'AbortError') alert(e.message); }
  }else{
    const a = document.createElement('a'); a.href = url; a.download = j.name; a.click();
  }
}
</script>
</body>
</html>
"""


# =========================================================================== #
#  HTTP server
# =========================================================================== #

RESULTS = {}  # id -> (filename, bytes, mime)


def _safe_name(name, default="presentation"):
    return re.sub(r"[\\/:*?\"<>|\x00-\x1f]", "_", name).strip() or default


def _store_result(name, kind, res):
    base = re.sub(r"\.(pptx|pdf)$", "", _safe_name(name), flags=re.I)
    out_name = "%s_clean.%s" % (base, kind)
    payload = {"status": res["status"], "clean": res["ok"]}
    if res["data"] is not None:
        rid = secrets.token_urlsafe(12)
        RESULTS[rid] = (out_name, res["data"], PPTX_MIME if kind == "pptx" else PDF_MIME)
        if len(RESULTS) > 20:
            RESULTS.pop(next(iter(RESULTS)))
        payload.update({"id": rid, "name": out_name})
    return payload


class Handler(BaseHTTPRequestHandler):
    server_version = "gamma-unmark"

    def log_message(self, *a):  # quiet
        pass

    def _local_only(self):
        """Accept only loopback Host headers and same-origin browser requests
        (blocks DNS-rebinding and cross-site requests to the local server)."""
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]")
        if host not in ("127.0.0.1", "localhost", "::1"):
            return False
        origin = self.headers.get("Origin")
        if origin:
            oh = urllib.parse.urlparse(origin).hostname
            if oh not in ("127.0.0.1", "localhost", "::1"):
                return False
        return True

    def _send(self, code, body, ctype="application/json", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body).encode()
        elif isinstance(body, str):
            body = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if not self._local_only():
            return self._send(403, {"error": "Forbidden"})
        path = urllib.parse.urlparse(self.path).path
        if path == "/":
            return self._send(200, HTML, "text/html; charset=utf-8")
        if path in ASSETS:
            name, mime = ASSETS[path]
            p = _asset_path(name)
            if not p:
                return self._send(404, b"", mime)
            with open(p, "rb") as f:
                return self._send(200, f.read(), mime)
        if path.startswith("/api/download/"):
            item = RESULTS.get(path.rsplit("/", 1)[1])
            if not item:
                return self._send(404, {"error": "Not found"})
            name, blob, mime = item
            return self._send(
                200, blob, mime,
                {"Content-Disposition": "attachment; filename*=UTF-8''" + urllib.parse.quote(name)},
            )
        self._send(404, {"error": "Not found"})

    def do_POST(self):
        if not self._local_only():
            return self._send(403, {"error": "Forbidden"})
        path = urllib.parse.urlparse(self.path).path
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        if length <= 0 or length > MAX_UPLOAD:
            return self._send(400, {"error": "Invalid request size"})

        try:
            if path == "/api/clean":
                name = urllib.parse.unquote(self.headers.get("X-Filename", "presentation"))
                data = self.rfile.read(length)
            else:
                return self._send(404, {"error": "Not found"})

            res = clean_any(data)
            self._send(200, _store_result(name, res["kind"], res))
        except ValueError as e:
            self._send(400, {"error": str(e)})
        except Exception:
            self._send(500, {"error": "Could not process this file"})


def free_port(preferred):
    for p in (preferred, 0):
        s = socket.socket()
        try:
            s.bind(("127.0.0.1", p))
            return s.getsockname()[1]
        except OSError:
            continue
        finally:
            s.close()
    return 0


def main():
    ap = argparse.ArgumentParser(description="Remove the Gamma badge from .pptx / .pdf files.")
    ap.add_argument("--port", type=int, default=8765, help="port (default 8765)")
    ap.add_argument("--no-browser", action="store_true", help="don't open the browser")
    args = ap.parse_args()

    if pymupdf is None:
        print("Note: PyMuPDF is not installed, so PDF files are disabled (PPTX works).")
        print("      Enable PDF with:  pip install pymupdf")

    if not _asset_path("logo-dark.png"):
        print("Note: logo not found. Put logo-dark.png into the 'assets' folder next to gu.py.")

    port = free_port(args.port)
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = "http://127.0.0.1:%d" % port
    print("gamma-unmark is running at %s  (Ctrl+C to stop)" % url)
    print("For educational purposes only. Use at your own responsibility.")
    if not args.no_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nbye")
        server.server_close()
        sys.exit(0)


if __name__ == "__main__":
    main()