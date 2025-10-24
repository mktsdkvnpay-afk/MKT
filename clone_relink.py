# ============================================================
# clone_relink.py — Public v1.1.14 (2025-10)
# Sheets ↔ Slides: Clone + Relink + Refresh (+ optional GAS WebApp)
# Không dùng Apps Script API. Chỉ POST đến Web App URL nếu có.
# Placeholders trên Slide:
#   IMG='Sheet'!A1        | ảnh từ ô
#   IMGURL=<https|http>   | ảnh từ URL
#   IMGID=<driveId>       | ảnh từ fileId
#   EXPORTPNG='Sheet'!A1:D20
#   RANGE='Sheet'!A1:D20  | TABLEIMG (GVIZ)
#   {{A1}} hoặc {{'Sheet'!B2}} trong ô Table
# ============================================================

import os, re, io, json, time, uuid, argparse, sys, requests
from typing import List, Dict, Any, Tuple
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import AuthorizedSession

# --- OPTIONAL GAS WEB APP URL ---
DEFAULT_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzrlBBWNjJIu6vJKqYSJzI9lnT6tpf_qDNRWIGMkjKDiBicJZ5xN4VLfw0mxTWBCr48/exec"
# Ưu tiên tham số CLI/UI; nếu không có thì lấy từ biến môi trường GAS_WEBAPP_URL
DEFAULT_WEBAPP_URL = os.getenv("GAS_WEBAPP_URL","").strip()

# ---------------- AUTH ----------------
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/spreadsheets",
]

def _build_services():
    try:
        from google.colab import auth as colab_auth
        colab_auth.authenticate_user()
        import google.auth
        creds, _ = google.auth.default(scopes=SCOPES)
    except Exception:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", SCOPES)
            creds = flow.run_local_server(port=0)
            with open("token.json", "w") as f: f.write(creds.to_json())
    from googleapiclient.discovery import build
    drive  = build("drive","v3",credentials=creds)
    slides = build("slides","v1",credentials=creds)
    sheets = build("sheets","v4",credentials=creds)
    return drive, slides, sheets, creds

drive, slides, sheets, creds = _build_services()

# ---------------- UTILS ----------------
ID_RX = re.compile(r"/d/([A-Za-z0-9_-]+)")
RANGE_RX = re.compile(r"RANGE\s*=\s*([^\r\n]+)")
IMG_ANY_RX = re.compile(r"\b(IMGID|IMGURL|IMG)(?:URL|ID)?\s*=\s*(?:\"([^\"]+)\"|'([^']+)'|([^\r\n]+))")
EXPORTPNG_RX = re.compile(r"\bEXPORTPNG\s*=\s*([^\r\n]+)")
CELL_ANY_TAG_RX = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")

def _extract_id(link_or_id:str)->str:
    if re.fullmatch(r"[A-Za-z0-9_-]+", link_or_id or ""): return link_or_id
    m = ID_RX.search(link_or_id or "")
    if not m: raise ValueError("Không tìm thấy ID trong link.")
    return m.group(1)

def _retry(fn, tries=5, base=0.6, fact=1.7, what="call"):
    for i in range(tries):
        try: return fn()
        except Exception:
            if i == tries-1: raise
            time.sleep(base*(fact**i))

def _drive_get(fid, fields): return _retry(lambda: drive.files().get(fileId=fid, fields=fields).execute(), what="drive.get")
def _drive_meta(fid, fields="id,name,mimeType"): return _drive_get(fid, fields)
def _is_sheet(m): return m.get("mimeType") == "application/vnd.google-apps.spreadsheet"
def _is_slide(m): return m.get("mimeType") == "application/vnd.google-apps.presentation"

def _slides_get(pid, fields=None):
    if not fields:
        return _retry(lambda: slides.presentations().get(presentationId=pid).execute(), what="slides.get")
    return _retry(lambda: slides.presentations().get(presentationId=pid, fields=fields).execute(), what="slides.get")

def _slides_batch(pid, reqs, chunk=80):
    if not reqs: return
    for i in range(0, len(reqs), chunk):
        _retry(lambda: slides.presentations().batchUpdate(presentationId=pid, body={"requests": reqs[i:i+chunk]}).execute(), what="slides.batchUpdate")

def _sheets_get(sid, fields): return _retry(lambda: sheets.spreadsheets().get(spreadsheetId=sid, fields=fields).execute(), what="sheets.get")
def _sheets_values_batch_get(sid, ranges):
    return _retry(lambda: sheets.spreadsheets().values().batchGet(spreadsheetId=sid, ranges=ranges, valueRenderOption="FORMATTED_VALUE", dateTimeRenderOption="FORMATTED_STRING").execute(), what="sheets.values.batchGet")

def verify_access(fid, label): _drive_get(fid, "id,name,mimeType")
def parent_folder(fid): return _drive_get(fid, "parents").get("parents", [None])[0]

def copy_same_folder(fid, suffix):
    meta = _drive_get(fid, "name")
    body = {"name": f"{meta['name']} {suffix}".strip()}
    p = parent_folder(fid)
    if p: body["parents"] = [p]
    newf = _retry(lambda: drive.files().copy(fileId=fid, body=body).execute(), what="drive.copy")
    return newf["id"], body["name"]

# ---------------- A1 + helpers ----------------
def normalize_a1(rng:str)->str:
    rng=(rng or "").strip()
    if "!" in rng:
        sh, r = rng.split("!",1)
        sh = sh.strip().strip("'").strip('"')
        r  = re.sub(r"\s+","",r)
        return f"'{sh}'!{r}"
    return rng

def _valid_rng(r:str)->bool:
    return bool(re.match(r"^'?.+?'?![A-Za-z]+\d+(:[A-Za-z]+\d+)?$", r))

# ---------------- Clean text ----------------
def _clean_note(s: str) -> str:
    s = (s or "").replace("“","'").replace("”","'").replace("‘","'").replace("’","'")
    return " ".join(s.split())

def _element_note_with_text(el: dict) -> str:
    desc = el.get("description", "") or ""
    title = el.get("title", "") or ""
    texts = []
    shp = el.get("shape", {})
    if shp:
        for te in shp.get("text", {}).get("textElements", []):
            c = te.get("textRun", {}).get("content", "")
            if c: texts.append(c)
    tbl = el.get("table", {})
    if tbl:
        for row in tbl.get("tableRows", []):
            for cell in row.get("tableCells", []):
                for te in cell.get("text", {}).get("textElements", []):
                    c = te.get("textRun", {}).get("content", "")
                    if c: texts.append(c)
    return _clean_note(" ".join([desc, title, " ".join(texts)]))

# ---------------- PNG makers ----------------
def _gviz_url(sid, rng):
    rng = normalize_a1(rng).replace("'","")
    clean = rng.split("!",1)[1] if "!" in rng else rng
    return f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:png&range={clean}&cb={uuid.uuid4().hex}"

def _fetch_png_gviz(sid, rng):
    ses = AuthorizedSession(creds)
    resp = _retry(lambda: ses.get(_gviz_url(sid, rng)), what="gviz.get")
    if resp.status_code != 200: raise RuntimeError(f"GVIZ {resp.status_code}: {resp.text[:200]}")
    return resp.content

def _sheet_gid_by_title(sid:str, title:str)->str:
    meta = _sheets_get(sid, "sheets(properties(sheetId,title))")
    for s in meta.get("sheets",[]):
        p=s["properties"]
        if p["title"]==title: return str(p["sheetId"])
    return "0"

def _export_png_url(sid:str, a1:str)->str:
    a1 = normalize_a1(a1)
    sh = a1.split("!",1)[0].strip().strip("'")
    rng = a1.split("!",1)[1]
    gid = _sheet_gid_by_title(sid, sh)
    return f"https://docs.google.com/spreadsheets/d/{sid}/export?format=png&gid={gid}&range={rng}&cb={uuid.uuid4().hex}"

def _fetch_png_export(sid:str, a1:str)->bytes:
    ses = AuthorizedSession(creds)
    resp = _retry(lambda: ses.get(_export_png_url(sid, a1)), what="sheet.export")
    if resp.status_code != 200: raise RuntimeError(f"EXPORTPNG {resp.status_code}: {resp.text[:200]}")
    return resp.content

# ---------------- Upload/rehost ----------------
_MAX_BYTES = 15 * 1024 * 1024

def _upload_bytes_public(b: bytes, filename: str, mime: str) -> str:
    media = MediaIoBaseUpload(io.BytesIO(b), mimetype=mime or "application/octet-stream", resumable=False)
    f = _retry(lambda: drive.files().create(body={"name": filename}, media_body=media, fields="id,webContentLink,mimeType").execute(), what="drive.upload")
    try:
        _retry(lambda: drive.permissions().create(fileId=f["id"], body={"role":"reader","type":"anyone"}).execute(), what="perm.anyone")
    except Exception:
        pass
    url = f"https://drive.google.com/uc?export=view&id={f['id']}&cb={uuid.uuid4().hex}"
    def _ok():
        try:
            r = requests.get(url, timeout=10)
            ct = r.headers.get("Content-Type","").lower()
            return r.status_code == 200 and ct.startswith("image/")
        except Exception:
            return False
    for _ in range(8):
        if _ok(): break
        time.sleep(0.7)
    return url

def _upload_png(png_bytes, name="tbl.png"):
    return _upload_bytes_public(png_bytes, name, "image/png")

def _fetch_bytes_http(url:str)->Tuple[bytes,str]:
    r = _retry(lambda: requests.get(url, timeout=30), what="http.get")
    if r.status_code != 200: raise RuntimeError(f"HTTP {r.status_code}")
    b = r.content
    if len(b) > _MAX_BYTES: raise RuntimeError("Image too large")
    ctype = r.headers.get("Content-Type","").split(";")[0].strip().lower()
    if not ctype.startswith("image/"):
        ext = os.path.splitext(url.split("?")[0])[1].lower()
        ctype = {"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png","gif":"image/gif","webp":"image/webp"}.get(ext.lstrip("."),"image/png")
    return b, ctype

def _fetch_bytes_drive(file_id:str)->Tuple[bytes,str]:
    if "drive.google.com" in file_id:
        m = re.search(r"/d/([A-Za-z0-9_-]+)", file_id) or re.search(r"id=([A-Za-z0-9_-]+)", file_id)
        if m: file_id = m.group(1)
    ses = AuthorizedSession(creds)
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    r = _retry(lambda: ses.get(url), what="drive.media")
    if r.status_code != 200: raise RuntimeError(f"Drive media {r.status_code}")
    b = r.content
    if len(b) > _MAX_BYTES: raise RuntimeError("Image too large")
    ctype = r.headers.get("Content-Type","").split(";")[0].strip().lower() or "image/png"
    return b, ctype

def _rehost_from_src(src:str, hint:str="img")->str:
    s = src.strip()
    drv_id = None
    if "drive.google.com" in s:
        m = re.search(r"/d/([A-Za-z0-9_-]+)", s) or re.search(r"[?&]id=([A-Za-z0-9_-]+)", s)
        if m: drv_id = m.group(1)
    if drv_id:
        public_url = f"https://drive.google.com/uc?export=view&id={drv_id}&cb={uuid.uuid4().hex}"
        try:
            r = _retry(lambda: requests.get(public_url, timeout=12), what="drive.uc")
            if r.status_code == 200 and r.headers.get("Content-Type","").lower().startswith("image/"):
                try: requests.head(public_url, timeout=5)
                except Exception: pass
                return public_url
        except Exception:
            pass
        b, ctype = _fetch_bytes_drive(drv_id)
        ext = {"image/jpeg":".jpg","image/png":".png","image/gif":".gif","image/webp":".webp"}.get(ctype, ".png")
        return _upload_bytes_public(b, f"{hint}{ext}", ctype)
    if s.startswith("http://") or s.startswith("https://"):
        b, ctype = _fetch_bytes_http(s)
        ext = {"image/jpeg":".jpg","image/png":".png","image/gif":".gif","image/webp":".webp"}.get(ctype, ".png")
        return _upload_bytes_public(b, f"{hint}{ext}", ctype)
    b, ctype = _fetch_bytes_drive(s)
    ext = {"image/jpeg":".jpg","image/png":".png","image/gif":".gif","image/webp":".webp"}.get(ctype, ".png")
    return _upload_bytes_public(b, f"{hint}{ext}", ctype)

# ---------------- GAS Web App (HTTP) ----------------
def run_gas_webapp(slide_id:str, sheet_id:str, webapp_url:str|None)->dict|None:
    url = (webapp_url or "").strip()
    if not url: return None
    try:
        r = requests.post(url, json={"slideId": slide_id, "sheetId": sheet_id}, timeout=45)
        return r.json() if r.headers.get("content-type","").lower().startswith("application/json") else {"ok": r.ok, "text": r.text}
    except Exception as e:
        print("GAS WebApp lỗi:", e); return None

# ---------------- Sheet values + image URL resolve ----------------
def _first_sheet_title(sid:str)->str:
    meta = _sheets_get(sid, "sheets(properties(title,index))")
    arr = sorted([s["properties"] for s in meta.get("sheets",[])], key=lambda x:x.get("index",0))
    return arr[0]["title"] if arr else "Sheet1"

FORMULA_URL_RX = re.compile(
    r'''(?ix)
    ^\s*=
    (?:IMAGE|HYPERLINK)
    \s*\(\s*
    ["'](?P<url>.+?)["']
    (?:\s*[,;].*)?\)\s*$
    '''
)

def _read_cell(sid:str, a1:str)->str:
    a1 = normalize_a1(a1)
    if "!" not in a1: a1 = f"'{_first_sheet_title(sid)}'!{a1}"
    resp = _retry(lambda: sheets.spreadsheets().values().get(
        spreadsheetId=sid, range=a1,
        valueRenderOption="FORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING"
    ).execute(), what="values.get")
    vals = resp.get("values",[[]])
    val  = (vals[0][0] if vals and vals[0] else "").strip()
    if val: return val
    resp2 = _retry(lambda: sheets.spreadsheets().values().get(
        spreadsheetId=sid, range=a1,
        valueRenderOption="FORMULA",
        dateTimeRenderOption="FORMATTED_STRING"
    ).execute(), what="values.get(formula)")
    vals2 = resp2.get("values",[[]])
    f = (vals2[0][0] if vals2 and vals2[0] else "").strip()
    m = FORMULA_URL_RX.match(f)
    if m: return m.group("url").strip()
    if f.startswith("=\"") and f.endswith("\""): return f[2:-1]
    return ""

def _image_url_from_cell(sid:str, ref:str)->str:
    v = _read_cell(sid, ref)
    if not v: return ""
    return _rehost_from_src(v.strip(), hint=f"cell_{ref.replace('!','_')}")

def _resolve_image_url_from_token(ttype:str, raw_ref:str, sheet_id:str)->str:
    raw_ref = raw_ref.strip()
    if ttype == "IMGID":  return _rehost_from_src(raw_ref, hint=f"imgid_{raw_ref[:8]}")
    if ttype == "IMGURL": return _rehost_from_src(raw_ref, hint="imgurl")
    return _image_url_from_cell(sheet_id, normalize_a1(raw_ref))  # IMG

def _create_img_req(page_id,size,transform,url,desc=""):
    try: requests.head(url, timeout=5)
    except Exception: pass
    img_id=f"img_{uuid.uuid4().hex}"
    return [
        {"createImage":{"url":url,"elementProperties":{"pageObjectId":page_id,"size":size,"transform":transform},"objectId":img_id}},
        {"updatePageElementAltText":{"objectId":img_id,"description":desc}}
    ]

# ---------------- Fields mask ----------------
FIELDS_ALL = (
  "slides("
    "objectId,"
    "pageElements("
      "objectId,description,title,size,transform,"
      "shape(text(textElements(textRun(content)))),"
      "table(tableRows(tableCells(text(textElements(textRun(content))))))"
    ")"
  ")"
)

# ---------------- TABLEIMG (GVIZ) ----------------
def convert_table_placeholders(pres_id, sheet_id)->int:
    pres = _slides_get(pres_id, FIELDS_ALL)
    reqs=[]; n=0
    for sl in pres.get("slides",[]):
        pid=sl["objectId"]
        for el in sl.get("pageElements",[]):
            note = _element_note_with_text(el)
            m = RANGE_RX.search(note)
            if not m: continue
            rng = normalize_a1(m.group(1))
            if not _valid_rng(rng): continue
            png = _fetch_png_gviz(sheet_id, rng)
            url = _upload_png(png, f"{rng}.png")
            size = el.get("size") or {"width":{"magnitude":3000000,"unit":"EMU"},"height":{"magnitude":2000000,"unit":"EMU"}}
            transform = el.get("transform") or {"scaleX":1,"scaleY":1,"translateX":0,"translateY":0,"unit":"EMU"}
            reqs += _create_img_req(pid, size, transform, url, f"TABLEIMG|RANGE={rng}")
            reqs.append({"deleteObject":{"objectId":el["objectId"]}})
            n+=1
    _slides_batch(pres_id, reqs); return n

# ---------------- EXPORTPNG ----------------
def convert_exportpng_placeholders(pres_id, sheet_id)->int:
    pres = _slides_get(pres_id, FIELDS_ALL)
    reqs=[]; n=0
    for sl in pres.get("slides",[]):
        pid=sl["objectId"]
        for el in sl.get("pageElements",[]):
            note = _element_note_with_text(el)
            m = EXPORTPNG_RX.search(note)
            if not m: continue
            a1 = normalize_a1(m.group(1))
            png = _fetch_png_export(sheet_id, a1)
            url = _upload_png(png, f"{a1}.png")
            size = el.get("size") or {"width":{"magnitude":3000000,"unit":"EMU"},"height":{"magnitude":2000000,"unit":"EMU"}}
            transform = el.get("transform") or {"scaleX":1,"scaleY":1,"translateX":0,"translateY":0,"unit":"EMU"}
            reqs += _create_img_req(pid, size, transform, url, f"EXPORTPNG|RANGE={a1}")
            reqs.append({"deleteObject":{"objectId":el["objectId"]}})
            n+=1
    _slides_batch(pres_id, reqs); return n

# ---------------- IMG / IMGID / IMGURL ----------------
def convert_img_placeholders(pres_id, sheet_id) -> int:
    pres = _slides_get(pres_id, FIELDS_ALL)
    reqs, total_new, to_delete = [], 0, []
    for sl in pres.get("slides", []):
        pid = sl["objectId"]
        for el in sl.get("pageElements", []):
            note = _element_note_with_text(el)
            matches = []
            for m in IMG_ANY_RX.finditer(note):
                ttype = m.group(1)
                raw = m.group(2) or m.group(3) or m.group(4) or ""
                matches.append((ttype, raw))
            if not matches: continue
            size = el.get("size") or {"width":{"magnitude":3000000,"unit":"EMU"},"height":{"magnitude":2000000,"unit":"EMU"}}
            transform = el.get("transform") or {"scaleX":1,"scaleY":1,"translateX":0,"translateY":0,"unit":"EMU"}
            ok=False
            for ttype, raw in matches:
                raw = (raw or "").strip()
                if "http" in raw: raw = re.sub(r"\s+", "", raw)
                try: url = _resolve_image_url_from_token(ttype, raw, sheet_id)
                except Exception: continue
                if not url: continue
                tag = f"{ttype}|VALUE={normalize_a1(raw) if ttype=='IMG' else raw}"
                reqs += _create_img_req(pid, size, transform, url, tag)
                total_new += 1; ok=True
            if ok: to_delete.append(el["objectId"])
    for oid in to_delete:
        reqs.append({"deleteObject":{"objectId":oid}})
    _slides_batch(pres_id, reqs); return total_new

# ---------------- REFRESH (ảnh & bảng) ----------------
def refresh_linked_images(pres_id, sheet_id)->Dict[str,int]:
    pres = _slides_get(pres_id, "slides(pageElements(objectId,description,image))")
    reqs=[]; n_img=0; n_export=0; n_tbl=0
    for sl in pres.get("slides",[]):
        for el in sl.get("pageElements",[]):
            d = el.get("description") or ""
            m = re.search(r"(IMG|IMGID|IMGURL)\|VALUE\s*=\s*([^\r\n]+)", d)
            if m:
                ttype, val = m.group(1), m.group(2).strip()
                ref = normalize_a1(val) if ttype=="IMG" else re.sub(r"\s+","",val)
                try: url = _resolve_image_url_from_token(ttype, ref, sheet_id)
                except Exception: continue
                if url:
                    try: requests.head(url, timeout=5)
                    except Exception: pass
                    reqs.append({"replaceImage":{"imageObjectId":el["objectId"],"imageReplaceMethod":"CENTER_INSIDE","url":url}})
                    n_img+=1
                continue
            m = re.search(r"EXPORTPNG\|RANGE\s*=\s*([^\r\n]+)", d)
            if m:
                a1 = normalize_a1(m.group(1))
                png = _fetch_png_export(sheet_id, a1)
                url = _upload_png(png, f"{a1}.png")
                try: requests.head(url, timeout=5)
                except Exception: pass
                reqs.append({"replaceImage":{"imageObjectId":el["objectId"],"imageReplaceMethod":"CENTER_INSIDE","url":url}})
                n_export+=1; continue
            if d.startswith("TABLEIMG|"):
                m = RANGE_RX.search(d)
                if m:
                    rng = normalize_a1(m.group(1))
                    png = _fetch_png_gviz(sheet_id, rng)
                    url = _upload_png(png, f"{rng}.png")
                    try: requests.head(url, timeout=5)
                    except Exception: pass
                    reqs.append({"replaceImage":{"imageObjectId":el["objectId"],"imageReplaceMethod":"CENTER_INSIDE","url":url}})
                    n_tbl+=1
    _slides_batch(pres_id, reqs)
    return {"images_refreshed": n_img, "exportpng_refreshed": n_export, "tableimg_refreshed": n_tbl}

# ---------------- {{A1}} in TABLE ----------------
def _norm_cell_ref(ref:str, default_sheet:str)->str:
    ref=ref.strip()
    if "!" in ref:
        sh,a1=ref.split("!",1); sh=sh.strip().strip("'").strip('"')
        return f"{sh}!{a1.strip().upper()}"
    return f"{default_sheet}!{ref.upper()}"

def _batch_vals(sid, refs:List[str])->Dict[str,str]:
    by_sh={}
    for r in refs:
        sh,a1=r.split("!",1); by_sh.setdefault(sh,[]).append(a1)
    out={}; rngs=[]
    for sh, arr in by_sh.items():
        rngs += [f"'{sh}'!{a}" for a in arr]
    resp=_sheets_values_batch_get(sid, rngs)
    for vr in resp.get("valueRanges",[]):
        rng = vr.get("range","")
        if "!" in rng:
            shn,a = rng.split("!",1); shn=shn.strip().strip("'")
            a1 = a.split(":")[0]
            key=f"{shn}!{a1}"
            vals=vr.get("values",[[]]); val = vals[0][0] if vals and vals[0] else ""
            out[key]=str(val)
    return out

def update_table_cell_links(pres_id, sheet_id)->int:
    pres = _slides_get(pres_id, "slides(objectId,pageElements(objectId,table(tableRows(tableCells(text(textElements(textRun(content))))))))")
    default_sheet=_first_sheet_title(sheet_id)
    cells=[]; uniq=[]
    for sl in pres.get("slides",[]):
        for el in sl.get("pageElements", []):
            tbl = el.get("table")
            if not tbl: continue
            for r_idx,row in enumerate(tbl.get("tableRows",[])):
                for c_idx,cell in enumerate(row.get("tableCells",[])):
                    texts = cell.get("text",{}).get("textElements",[])
                    s="".join([(t.get("textRun",{}) or {}).get("content","") for t in texts])
                    ms=list(CELL_ANY_TAG_RX.finditer(s))
                    if not ms: continue
                    raw=[m.group(1) for m in ms]; norm=[]
                    for rf in raw:
                        n=_norm_cell_ref(rf, default_sheet); norm.append(n)
                        if n not in uniq: uniq.append(n)
                    cells.append((el["objectId"], r_idx, c_idx, s, norm, raw))
    if not cells: return 0
    valmap=_batch_vals(sheet_id, uniq)
    reqs=[]; n=0
    for oid, r, c, orig, norms, raws in cells:
        new=orig
        for raw, nm in zip(raws, norms):
            new=re.sub(r"\{\{\s*"+re.escape(raw)+r"\s*\}\}", str(valmap.get(nm,"")), new, count=1)
        if new!=orig:
            reqs.append({"deleteText":{"objectId":oid,"cellLocation":{"rowIndex":r,"columnIndex":c},"textRange":{"type":"ALL"}}})
            reqs.append({"insertText":{"objectId":oid,"cellLocation":{"rowIndex":r,"columnIndex":c},"insertionIndex":0,"text":new}})
            n+=1
    _slides_batch(pres_id, reqs); return n

# ---------------- Charts relink (order-based) ----------------
def clone_and_relink_charts(new_pid, old_sid, new_sid)->int:
    def _charts(sid): return _sheets_get(sid, "sheets(charts(chartId))")
    old = _charts(old_sid); new = _charts(new_sid)
    map_chart={}
    try:
        old_ids=[c["chartId"] for s in old.get("sheets",[]) for c in s.get("charts",[])]
        new_ids=[c["chartId"] for s in new.get("sheets",[]) for c in s.get("charts",[])]
        for so, sn in zip(old_ids, new_ids): map_chart[so]=sn
    except: pass
    pres = _slides_get(new_pid, "slides(objectId,pageElements(objectId,size,transform,sheetsChart(chartId,spreadsheetId)))")
    reqs=[]; done=0
    for sl in pres.get("slides",[]):
        pid2=sl["objectId"]
        for el in sl.get("pageElements",[]):
            sc=el.get("sheetsChart")
            if not sc: continue
            oc=sc["chartId"]; nc=map_chart.get(oc)
            if not nc: continue
            reqs += [
                {"createSheetsChart":{"spreadsheetId":new_sid,"chartId":nc,"linkingMode":"LINKED","elementProperties":{"pageObjectId":pid2,"size":el.get("size"),"transform":el.get("transform")}}},
                {"deleteObject":{"objectId":el["objectId"]}}
            ]; done+=1
    _slides_batch(new_pid, reqs); return done

# ---------------- CLONE & RELINK ----------------
def clone_and_relink(sheet_link, slide_link, suffix="Clone", webapp_url:str|None=None)->Dict[str,Any]:
    sid=_extract_id(sheet_link); pid=_extract_id(slide_link)
    sm = _drive_meta(sid); pm = _drive_meta(pid)
    if _is_slide(sm) and _is_sheet(pm):
        sid, pid = pid, sid
        sm, pm = pm, sm
    if not _is_sheet(sm):  raise ValueError("--sheet phải là Google Sheet.")
    if not _is_slide(pm):  raise ValueError("--slide phải là Google Slide.")
    verify_access(sid,"sheet"); verify_access(pid,"slide")

    new_sid,new_sname = copy_same_folder(sid, suffix)
    new_pid,new_pname = copy_same_folder(pid, suffix)

    # Gọi Web App (nếu có), để GAS xử lý placeholder tùy biến
    run_gas_webapp(new_pid, new_sid, webapp_url or DEFAULT_WEBAPP_URL)

    # Python native
    charts_done = clone_and_relink_charts(new_pid, sid, new_sid)
    tbl = convert_table_placeholders(new_pid, new_sid)
    exp = convert_exportpng_placeholders(new_pid, new_sid)
    imgs = convert_img_placeholders(new_pid, new_sid)
    return {
        "new_sheet_id": new_sid, "new_slide_id": new_pid,
        "new_sheet_name": new_sname, "new_slide_name": new_pname,
        "charts_processed": charts_done, "tables_converted": tbl,
        "exportpng_created": exp, "images_created": imgs
    }

# ---------------- REFRESH ----------------
def refresh_all(pres_id, sheet_id, webapp_url:str|None=None)->Dict[str,int]:
    # Web App trước (nếu có)
    run_gas_webapp(pres_id, sheet_id, webapp_url or DEFAULT_WEBAPP_URL)

    # Python refresh
    created_tbl  = convert_table_placeholders(pres_id, sheet_id)
    created_exp  = convert_exportpng_placeholders(pres_id, sheet_id)
    created_imgs = convert_img_placeholders(pres_id, sheet_id)
    img_stats = refresh_linked_images(pres_id, sheet_id)
    cells = update_table_cell_links(pres_id, sheet_id)
    return {
        "tables_converted_now": created_tbl,
        "exportpng_created_now": created_exp,
        "images_created_now": created_imgs,
        **img_stats,
        "cells_updated": cells,
    }

# ---------------- UI ----------------
def render_widgets(default_sheet="", default_slide="", default_suffix="Clone 1", default_webapp=DEFAULT_WEBAPP_URL):
    try:
        from IPython.display import display, HTML, clear_output, Javascript
        import ipywidgets as w
    except Exception:
        print("Widgets không khả dụng ngoài Jupyter/Colab."); return
    sheet = w.Text(value=default_sheet, description='Sheet A:', layout=w.Layout(width='80%'))
    slide = w.Text(value=default_slide, description='Slide A:', layout=w.Layout(width='80%'))
    suff  = w.Text(value=default_suffix, description='Suffix:',  layout=w.Layout(width='50%'))
    web   = w.Text(value=default_webapp, description='WebApp URL:', layout=w.Layout(width='90%'))
    b1 = w.Button(description="Bước 1: Clone", button_style='primary')
    b2 = w.Button(description="Bước 2: Mở file")
    b3 = w.Button(description="Bước 3: Refresh", button_style='success')
    out = w.Output()
    STATE="last_clone.json"
    def save_state(d):
        with open(STATE,"w",encoding="utf-8") as f: json.dump(d,f,ensure_ascii=False,indent=2)
    def load_state():
        if not os.path.exists(STATE): return None
        with open(STATE,"r",encoding="utf-8") as f: return json.load(f)
    def on_clone(_):
        with out:
            clear_output(); print("Đang clone")
            try:
                r = clone_and_relink(sheet.value, slide.value, suff.value, webapp_url=web.value or None)
                save_state({"new_sheet_id":r["new_sheet_id"],"new_slide_id":r["new_slide_id"],"webapp_url":(web.value or "")})
                display(HTML(
                    f"<b>OK</b><br>"
                    f"<a href='https://docs.google.com/spreadsheets/d/{r['new_sheet_id']}/edit' target='_blank'>Mở Sheet</a> · "
                    f"<a href='https://docs.google.com/presentation/d/{r['new_slide_id']}/edit' target='_blank'>Mở Slide</a><br>"
                    f"Charts: {r['charts_processed']} | TABLEIMG: {r['tables_converted']} | EXPORTPNG: {r['exportpng_created']} | IMG: {r['images_created']}"
                ))
            except Exception as e: print("Lỗi:", e)
    def on_open(_):
        with out:
            clear_output()
            s = load_state()
            if not s: print("Chưa có file clone. Chạy Bước 1."); return
            su = f"https://docs.google.com/spreadsheets/d/{s['new_sheet_id']}/edit"
            pu = f"https://docs.google.com/presentation/d/{s['new_slide_id']}/edit"
            js = f"""
            (function(){
              try{
                const A='{su}', B='{pu}';
                if(window._ndOpenTabsLock) return; window._ndOpenTabsLock=true;
                function openOnce(u,n){let w=window.open('',n); if(!w||w.closed){w=window.open(u,n,'noopener');} else {try{w.location.href=u;}catch(e){window.open(u,n,'noopener');}}}
                openOnce(A,'nd_sheet_tab'); openOnce(B,'nd_slide_tab');
                setTimeout(()=>{window._ndOpenTabsLock=false;},1200);
              }catch(e){}
            })();"""
            try: display(Javascript(js))
            except Exception: pass
            display(HTML(f"Mở 2 tab. <a target='nd_sheet_tab' href='{su}'>Sheet</a> · <a target='nd_slide_tab' href='{pu}'>Slide</a>"))
    def on_refresh(_):
        with out:
            clear_output()
            s = load_state()
            if not s: print("Chưa có file clone. Chạy Bước 1."); return
            print("Đang refresh")
            try:
                rs = refresh_all(
                    s["new_slide_id"],
                    s["new_sheet_id"],
                    webapp_url=(s.get("webapp_url") or web.value or DEFAULT_WEBAPP_URL)
                )
                display(HTML(
                    f"OK. IMG: {rs.get('images_refreshed',0)} | EXPORTPNG: {rs.get('exportpng_refreshed',0)} | TABLEIMG: {rs.get('tableimg_refreshed',0)} | "
                    f"Created → IMG:{rs.get('images_created_now',0)} EXPORTPNG:{rs.get('exportpng_created_now',0)} TABLE:{rs.get('tables_converted_now',0)} | "
                    f"Ô bảng: {rs.get('cells_updated',0)} · "
                    f"<a href='https://docs.google.com/presentation/d/{s['new_slide_id']}/edit' target='_blank'>Mở Slide</a>"
                ))
            except Exception as e: print("Lỗi refresh:", e)
    b1.on_click(on_clone); b2.on_click(on_open); b3.on_click(on_refresh)
    from IPython.display import display as _d
    _d(w.VBox([w.HBox([sheet]), w.HBox([slide]), w.HBox([suff]), web, w.HBox([b1,b2,b3]), out]))

# ---------------- CLI ----------------
def main(parsed=None):
    p=argparse.ArgumentParser(description="Clone & Relink Sheets↔Slides (WebApp optional)")
    p.add_argument("--sheet", help="Link hoặc ID Google Sheet nguồn")
    p.add_argument("--slide", help="Link hoặc ID Google Slide nguồn")
    p.add_argument("--suffix", default="Clone 1")
    p.add_argument("--webapp-url", default=DEFAULT_WEBAPP_URL, help="URL Web App GAS (tùy chọn)")
    p.add_argument("--refresh", action="store_true", help="Refresh tài nguyên trên bản clone")
    p.add_argument("--state", default="last_clone.json", help="File lưu trạng thái clone")
    p.add_argument("--open-tabs", action="store_true", help="In link sau khi clone")
    p.add_argument("--widgets", action="store_true", help="UI 3 nút (Jupyter/Colab)")
    args,_=p.parse_known_args(parsed)

    if args.widgets:
        render_widgets(args.sheet or "", args.slide or "", args.suffix, default_webapp=args.webapp_url or DEFAULT_WEBAPP_URL); return

    if args.refresh:
        if not os.path.exists(args.state): print("Thiếu --state."); sys.exit(2)
        s=json.load(open(args.state,"r",encoding="utf-8"))
        rs=refresh_all(
            s["new_slide_id"],
            s["new_sheet_id"],
            webapp_url=(args.webapp_url or s.get("webapp_url") or DEFAULT_WEBAPP_URL)
        )
        print(json.dumps(rs,ensure_ascii=False,indent=2)); return

    if not args.sheet or not args.slide:
        print("Cần --sheet và --slide"); sys.exit(2)

    sid=_extract_id(args.sheet); pid=_extract_id(args.slide)
    sm=_drive_meta(sid); pm=_drive_meta(pid)
    if _is_slide(sm) and _is_sheet(pm):
        sid, pid = pid, sid
    r=clone_and_relink(sid, pid, args.suffix, webapp_url=args.webapp_url)
    with open(args.state,"w",encoding="utf-8") as f:
        json.dump({
            "new_sheet_id":r["new_sheet_id"],
            "new_slide_id":r["new_slide_id"],
            "webapp_url": (args.webapp_url or DEFAULT_WEBAPP_URL)
        },f)
    print(json.dumps(r,ensure_ascii=False,indent=2))
    if args.open-tabs:
        print(f"Sheet: https://docs.google.com/spreadsheets/d/{r['new_sheet_id']}/edit")
        print(f"Slide: https://docs.google.com/presentation/d/{r['new_slide_id']}/edit")

if __name__=="__main__":
    IN_NOTEBOOK = bool(getattr(sys, "ps1", None)) or "ipykernel" in sys.modules
    if IN_NOTEBOOK and not any(a in sys.argv for a in ["--sheet","--slide","--refresh","--widgets"]):
        render_widgets("", "", "Clone 1", DEFAULT_WEBAPP_URL)
    else:
        main()
