import datetime, hashlib, json, os, re, shutil, subprocess, sys, time, zipfile
import xml.etree.ElementTree as ET

MIN_CHARS = 80
DRAFT = os.path.expanduser("~/.pandect-markitdown-drafts")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CYR = re.compile(r"[Ѐ-ӿ]")
LAT = re.compile(r"[A-Za-z]")
TOKEN = re.compile(r"[Ѐ-ӿA-Za-z]{2,}")

def sha(p):
    """Хеш на съдържанието; при I/O грешка от моста опитва повторно."""
    for attempt in range(3):
        try:
            h = hashlib.sha256()
            with open(p, "rb") as f:
                for c in iter(lambda: f.read(1 << 20), b""):
                    h.update(c)
            return h.hexdigest()[:16]
        except OSError:
            if attempt == 2:
                raise
            time.sleep(0.5)

def md_path(p):
    return os.path.splitext(p)[0] + ".md"

def skip(p):
    """Временни файлове на Word (~$....docx) и скрити — не са документи."""
    b = os.path.basename(p)
    return b.startswith("~$") or b.startswith(".")

def suspect_ocr(text):
    """Разпознава боклучав текстов слой от чужд OCR."""
    toks = TOKEN.findall(text)
    if len(toks) < 40:
        return None
    mixed = sum(1 for t in toks if CYR.search(t) and LAT.search(t))
    cyr = len(CYR.findall(text)); lat = len(LAT.findall(text))
    if cyr + lat < 200:
        return None
    if mixed / len(toks) > 0.12:
        return "смесена кирилица/латиница в %d%% от думите" % (100 * mixed // len(toks))
    if cyr / (cyr + lat) < 0.45:
        return "само %d%% кирилица" % (100 * cyr // (cyr + lat))
    return None

def header(p, s, method, pages, note=None):
    h = ("---\nsource: %s\nsource_sha256: %s\nextracted: %s\nmethod: %s\npages: %s\n"
         % (os.path.basename(p), s,
            datetime.datetime.now().isoformat(timespec="seconds"), method, pages))
    if note:
        h += "ocr_note: %s\n" % note
    return h + "---\n"

def docx_md(p):
    if shutil.which("pandoc"):
        r = subprocess.run(["pandoc", "-f", "docx", "-t", "gfm", "--wrap=none", p],
                           capture_output=True)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.decode("utf-8", "replace"), "docx-pandoc"
    with zipfile.ZipFile(p) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    paras = []
    for para in root.iter(W + "p"):
        buf = []
        for n in para.iter():
            if n.tag == W + "t": buf.append(n.text or "")
            elif n.tag == W + "tab": buf.append("\t")
            elif n.tag == W + "br": buf.append("\n")
        paras.append("".join(buf).rstrip())
    return "\n\n".join(paras), "docx-xml"

def pdftotext(p):
    """Три опита — I/O грешката от моста към машината е преходна."""
    err = ""
    for attempt in range(3):
        r = subprocess.run(["pdftotext", "-layout", p, "-"], capture_output=True)
        if r.returncode == 0:
            return r, None
        err = r.stderr.decode("utf-8", "replace")
        if "Input/output error" not in err and "I/O" not in err:
            return r, err[:200]
        time.sleep(0.5)
    return r, err[:200]

def handle(p, force, dry):
    md, s = md_path(p), sha(p)
    rec = {"file": p, "md": md}
    if not force and os.path.exists(md):
        if "source_sha256: " + s in open(md, encoding="utf-8", errors="replace").read(800):
            rec["status"] = "fresh"; return rec
    ext = os.path.splitext(p)[1].lower()
    if ext == ".docx":
        rec.update(status="written", method="docx")
        if not dry:
            body, method = docx_md(p)
            rec["method"] = method
            open(md, "w", encoding="utf-8").write(header(p, s, method, "n/a") + "\n" + body + "\n")
        return rec
    r, err = pdftotext(p)
    if err:
        rec.update(status="error", detail=err); return rec
    txt = r.stdout.decode("utf-8", "replace")
    pg = txt.split("\f")
    if pg and pg[-1] == "": pg.pop()
    empty = [i + 1 for i, t in enumerate(pg) if len(t.strip()) < MIN_CHARS]
    note = suspect_ocr(txt)
    if note:
        empty = list(range(1, len(pg) + 1))          # целият текстов слой е боклук
    method = "text-layer" if not empty else ("vision" if len(empty) == len(pg) else "mixed")
    rec.update(method=method, pages=len(pg))
    if note: rec["ocr_note"] = note
    if not empty:
        rec["status"] = "written"
        if not dry:
            body = "".join("\n<!-- стр. %d -->\n\n%s\n" % (i + 1, t.strip())
                           for i, t in enumerate(pg))
            open(md, "w", encoding="utf-8").write(header(p, s, method, len(pg)) + body)
        return rec
    rec.update(status="needs-vision", sha=s, empty_pages=len(empty))
    if not dry:
        os.makedirs(DRAFT, exist_ok=True)
        body = "".join("\n<!-- стр. %d -->\n\n%s\n"
                       % (i + 1, "@@СКАН@@" if (i + 1) in empty else t.strip())
                       for i, t in enumerate(pg))
        draft = os.path.join(DRAFT, s + ".md")
        open(draft, "w", encoding="utf-8").write(header(p, s, method, len(pg), note) + body)
        rec["draft"] = draft
    return rec

if __name__ == "__main__":
    a = sys.argv[1:]
    force = "--force" in a; dry = "--scan" in a
    out = []
    for p in [x for x in a if not x.startswith("--")]:
        if skip(p):
            continue
        try:
            out.append(handle(p, force, dry))
        except Exception as e:
            out.append({"file": p, "status": "error", "detail": "%s: %s" % (type(e).__name__, e)})
    print(json.dumps(out, ensure_ascii=False))
