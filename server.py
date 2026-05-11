"""
CEPS API server - Railway / cloud deployment
Lokalne: python server.py
Railway: automaticky pres Procfile
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import json, requests, xml.etree.ElementTree as ET, os, time
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta

# Nacti .env soubor (pokud existuje)
def _load_env():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"): continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
_load_env()

BASE_URL  = "https://www.ceps.cz/_layouts/CepsData.asmx"
NAMESPACE = "https://www.ceps.cz/CepsData/"
DATA_NS   = "https://www.ceps.cz/CepsData/StructuredData/1.0"

# ENTSO-E
ENTSOE_URL   = "https://web-api.tp.entsoe.eu/api"
ENTSOE_TOKEN = os.environ.get("ENTSOE_TOKEN", "")
CZ_DOMAIN    = "10YCZ-CEPS-----N"

# Anthropic Claude API
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL      = "claude-sonnet-4-5"

def _request_with_retry(func, *args, retries=3, delay=1.5, **kwargs):
    """Retry wrapper kvuli obcasnym SSL EOF chybam."""
    last_err = None
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except (requests.exceptions.SSLError,
                requests.exceptions.ConnectionError,
                requests.exceptions.Timeout) as e:
            last_err = e
            print(f"  [retry {attempt+1}/{retries}] {type(e).__name__}", flush=True)
            if attempt < retries - 1:
                time.sleep(delay * (attempt + 1))
    raise last_err

def call_ceps(method, params):
    px = "".join(f"  <{k}>{v}</{k}>\n" for k, v in params.items() if v != "")
    body = f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <{method} xmlns="{NAMESPACE}">
{px}    </{method}>
  </soap:Body>
</soap:Envelope>"""
    r = _request_with_retry(
        requests.post, BASE_URL,
        data=body.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8",
                 "SOAPAction": f'"{NAMESPACE}{method}"'},
        timeout=25
    )
    return r.text, r.status_code

def parse_ceps(xml_text):
    """Parser pro SOAP XML odpovedi."""
    root = ET.fromstring(xml_text)
    ns = DATA_NS
    col_map = {}
    for s in root.iter(f"{{{ns}}}serie"):
        if s.get("id") and s.get("name"):
            col_map[s.get("id")] = s.get("name")
    if not col_map:
        for s in root.iter("serie"):
            if s.get("id") and s.get("name"):
                col_map[s.get("id")] = s.get("name")
    items = list(root.iter(f"{{{ns}}}item")) or list(root.iter("item"))
    rows = []
    for item in items:
        row = {"date": item.get("date")}
        for vid, name in col_map.items():
            row[name] = item.get(vid)
        if not col_map:
            for k, v in item.attrib.items():
                if k != "date": row[k] = v
        rows.append(row)
    return {"columns": list(col_map.values()), "rows": rows}

# ============================================================
# ENTSO-E
# ============================================================

def call_entsoe(params):
    """Vola ENTSO-E REST API. Vraci (xml_text, status_code)."""
    if not ENTSOE_TOKEN:
        raise RuntimeError("ENTSOE_TOKEN neni nastaveny - vytvor .env soubor s 'ENTSOE_TOKEN=...'")
    p = dict(params)
    p["securityToken"] = ENTSOE_TOKEN
    r = _request_with_retry(requests.get, ENTSOE_URL, params=p, timeout=25)
    return r.text, r.status_code

def parse_entsoe_xml(xml_text, period_start_utc):
    """Parser pro ENTSO-E XML s TimeSeries. Vraci list bodu se sloupci:
    [{ts: "2026-05-06T09:30Z", value: 1234.5}, ...]
    period_start_utc: datetime pro pripad chybejicich timestampu na pointech.
    """
    # Strip namespace pro snadny xpath
    body = xml_text
    # ENTSO-E XML pouziva default namespace, odstrani ho ze stringu
    import re
    body_no_ns = re.sub(r'\sxmlns="[^"]+"', '', body, count=1)

    try:
        root = ET.fromstring(body_no_ns)
    except ET.ParseError as e:
        return []

    points = []
    for ts_node in root.findall(".//TimeSeries"):
        period = ts_node.find("Period")
        if period is None: continue
        time_interval = period.find("timeInterval")
        if time_interval is None: continue
        start_text = time_interval.find("start").text  # "2026-05-06T00:00Z"
        # Resolution: PT15M, PT30M, PT60M
        res_text = period.find("resolution").text
        res_min = 15
        m = re.match(r"PT(\d+)M", res_text)
        if m: res_min = int(m.group(1))
        # Start timestamp
        try:
            start_dt = datetime.strptime(start_text.replace("Z", ""), "%Y-%m-%dT%H:%M")
        except Exception:
            continue
        for pt in period.findall("Point"):
            pos_el = pt.find("position")
            qty_el = pt.find("quantity")
            if pos_el is None or qty_el is None: continue
            try:
                pos = int(pos_el.text)
                qty = float(qty_el.text)
            except Exception:
                continue
            # ENTSO-E pozice je 1-based, kazdy bod = (start + (pos-1)*resolution)
            ts = start_dt + timedelta(minutes=(pos-1) * res_min)
            points.append({"ts": ts.strftime("%Y-%m-%dT%H:%MZ"), "value": qty})
    return points

def fmt_entsoe_period(d):
    """Format datetime -> 'YYYYMMDDhhmm' pro ENTSO-E."""
    return d.strftime("%Y%m%d%H%M")

# ============================================================
# XLSX PARSER (bez externích knihoven, pro OTE VDT)
# ============================================================

import zipfile, io as _io

def _xlsx_iter_rows(xlsx_bytes):
    """Vraci [(col_letter, value), ...] pro kazdy radek z prvniho sheetu XLSX."""
    with zipfile.ZipFile(_io.BytesIO(xlsx_bytes)) as z:
        # Shared strings
        shared = []
        try:
            with z.open("xl/sharedStrings.xml") as f:
                tree = ET.parse(f)
                ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                for si in tree.getroot().findall(f"{ns}si"):
                    t = si.find(f"{ns}t")
                    if t is not None and t.text:
                        shared.append(t.text)
                    else:
                        text = "".join(r.text or "" for r in si.findall(f".//{ns}t"))
                        shared.append(text)
        except KeyError:
            pass
        
        # Sheet 1
        sheet_path = "xl/worksheets/sheet1.xml"
        try:
            with z.open(sheet_path) as f:
                tree = ET.parse(f)
        except KeyError:
            return
        
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        for row in tree.getroot().findall(f".//{ns}row"):
            cells = []
            for c in row.findall(f"{ns}c"):
                ref = c.get("r", "")
                col_letter = "".join(ch for ch in ref if ch.isalpha())
                t = c.get("t", "")
                v = c.find(f"{ns}v")
                value = None
                if v is not None and v.text is not None:
                    if t == "s":
                        try: value = shared[int(v.text)]
                        except: value = v.text
                    elif t == "inlineStr":
                        is_el = c.find(f"{ns}is/{ns}t")
                        value = is_el.text if is_el is not None else None
                    else:
                        try: value = float(v.text)
                        except: value = v.text
                cells.append((col_letter, value))
            yield cells

# Cache pro VDT data (per den)
_VDT_CACHE = {}  # {YYYY-MM-DD: (fetched_at, data_obj)}
_VDT_CACHE_TTL = 600  # 10 minut

def call_claude(snapshot_text):
    """Pošle snapshot do Claude API a vrátí text odpovědi."""
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY neni nastaveny")
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": CLAUDE_MODEL,
        "max_tokens": 600,
        "messages": [{"role": "user", "content": snapshot_text}],
    }
    r = _request_with_retry(
        requests.post, ANTHROPIC_API_URL,
        headers=headers, json=body, timeout=45
    )
    if r.status_code != 200:
        raise RuntimeError(f"Claude API {r.status_code}: {r.text[:300]}")
    data = r.json()
    # Extrahovat text z prvniho content bloku
    parts = data.get("content", [])
    text = "\n".join(p.get("text", "") for p in parts if p.get("type") == "text")
    return text.strip()

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} {args[1]}", flush=True)

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors(); self.end_headers()

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/ai/analyze":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                raw = self.rfile.read(length).decode("utf-8") if length else "{}"
                payload = json.loads(raw)
                snapshot = payload.get("snapshot", "")
                if not snapshot:
                    self._json({"error": "missing snapshot"}, 400); return
                print(f"  -> Claude analyze: {len(snapshot)} chars input", flush=True)
                analysis = call_claude(snapshot)
                print(f"  -> Claude analyze: {len(analysis)} chars output", flush=True)
                self._json({"analysis": analysis})
            except Exception as e:
                print(f"  -> Claude analyze ERROR: {e}", flush=True)
                self._json({"error": str(e)}, 502)
            return
        self._json({"error": "use POST /ai/analyze"}, 404)

    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        g = lambda k, d="": qs.get(k, [d])[0]

        if parsed.path == "/health":
            self._json({"status": "ok", "time": datetime.now().isoformat()}); return

        if parsed.path in ("/", "/index.html", "/live_odchylky.html"):
            self._html(); return

        # ENTSO-E Solar Forecast (Day-ahead + Intraday) pro CR
        if parsed.path == "/entsoe/solar":
            self._entsoe_solar(qs); return
        
        # OTE VDT - jeden den
        if parsed.path == "/ote/vdt":
            self._ote_vdt(qs); return
        
        # OTE VDT - rozsah dní
        if parsed.path == "/ote/vdt/range":
            self._ote_vdt_range(qs); return
        
        # ENTSO-E Residual Load Forecast (DE)
        if parsed.path == "/entsoe/residual":
            self._entsoe_residual(qs); return
        
        # SMARD actual residual load (DE)
        if parsed.path == "/smard/residual-load":
            self._smard_residual_load(qs); return
        
        # ENTSO-E Cross-border Physical Flows pro CZ
        if parsed.path == "/entsoe/cross-border":
            self._entsoe_cross_border(qs); return
        
        # HTML stranky
        if parsed.path == "/hruska.html":
            self._html_hruska(); return
        
        if parsed.path == "/kapacity.html":
            self._html_kapacity(); return

        if parsed.path != "/api":
            self._json({"error": "use /api"}, 404); return

        method = g("method", "AktualniSystemovaOdchylkaCR")
        df     = g("dateFrom")
        dt_    = g("dateTo")
        agr    = g("agregation", "MI")
        fn     = g("function",   "AVG")
        ver    = g("version",    "RT")
        para1  = g("para1", "")

        if not df:
            now = datetime.now()
            df  = now.strftime("%Y-%m-%dT00:00:00")
            dt_ = now.strftime("%Y-%m-%dT%H:%M:%S")

        params = {"dateFrom": df, "dateTo": dt_}

        if method == "AktualniSystemovaOdchylkaCR":
            params.update({"agregation": agr, "function": fn})
        elif method == "AktivaceSVRvCR":
            params.update({"aggregation": agr, "function": fn})
            if para1: params["para1"] = para1
        elif method in ["Load","Generation","GenerationRES","CrossborderPowerFlows"]:
            params.update({"agregation": agr, "function": fn, "version": ver})
            if para1: params["para1"] = para1
        elif method == "ExportImportSVR":
            params.update({"agregation": agr, "function": fn})
            if para1: params["para1"] = para1

        try:
            xml_text, status = call_ceps(method, params)
        except Exception as e:
            print(f"  -> {method} REQUEST FAIL: {e}", flush=True)
            self._json({"error": f"CEPS request failed: {e}"}, 502); return

        if status != 200:
            try:
                root = ET.fromstring(xml_text)
                fs = root.find(".//{http://schemas.xmlsoap.org/soap/envelope/}faultstring")
                msg = fs.text if fs is not None else xml_text[:300]
            except Exception:
                msg = xml_text[:300]
            self._json({"error": f"CEPS {status}: {msg}"}, 502); return

        data = parse_ceps(xml_text)
        print(f"  -> {method}: {len(data['rows'])} radku, cols={data['columns']}", flush=True)
        self._json(data)

    def _entsoe_solar(self, qs):
        """Solar Forecast pro CR - vraci Day-ahead a Intraday serie pro dnesek.
        Optional ?day=YYYY-MM-DD pro jiny den.
        Vystup: {"day_ahead": [{ts, value}], "intraday": [{ts, value}]}
        """
        try:
            day_str = qs.get("day", [None])[0]
            if day_str:
                day = datetime.strptime(day_str, "%Y-%m-%d")
            else:
                day = datetime.utcnow()
            # Cely den 00:00-23:00 UTC
            ps = day.replace(hour=0, minute=0, second=0, microsecond=0)
            pe = ps + timedelta(hours=23)

            base = {
                "documentType": "A69",   # Wind and Solar Forecast
                "psrType":      "B16",   # Solar
                "in_Domain":    CZ_DOMAIN,
                "periodStart":  fmt_entsoe_period(ps),
                "periodEnd":    fmt_entsoe_period(pe),
            }

            # Day-ahead: processType=A01
            da_xml, da_st = call_entsoe({**base, "processType": "A01"})
            da_points = parse_entsoe_xml(da_xml, ps) if da_st == 200 else []
            print(f"  -> ENTSO-E Solar DA: status={da_st}, points={len(da_points)}", flush=True)

            # Intraday: processType=A40 (Intraday Process) - obnovuje se pres den
            id_xml, id_st = call_entsoe({**base, "processType": "A40"})
            id_points = parse_entsoe_xml(id_xml, ps) if id_st == 200 else []
            print(f"  -> ENTSO-E Solar ID: status={id_st}, points={len(id_points)}", flush=True)

            self._json({
                "day": ps.strftime("%Y-%m-%d"),
                "day_ahead": da_points,
                "intraday":  id_points,
            })
        except Exception as e:
            print(f"  -> ENTSO-E Solar ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    # ============================================================
    # OTE VDT (Vnitrodenni trh - intraday spot ceny CR)
    # ============================================================
    
    def _ote_vdt_fetch_day(self, day_dt):
        """Stahne VDT XLSX pro jeden den a vrati seznam bodu.
        Returns: {"date": "YYYY-MM-DD", "data": [{ts, last}, ...]} nebo None pri 404.
        Pouziva cache _VDT_CACHE.
        """
        import urllib.request, urllib.error
        cache_key = day_dt.strftime("%Y-%m-%d")
        now_ts = time.time()
        
        if cache_key in _VDT_CACHE:
            fetched_at, cached = _VDT_CACHE[cache_key]
            if now_ts - fetched_at < _VDT_CACHE_TTL:
                return cached
        
        url = (f"https://www.ote-cr.cz/pubweb/attachments/27/{day_dt.year}/"
               f"month{day_dt.month:02d}/day{day_dt.day:02d}/"
               f"VDT_15MIN_{day_dt.day:02d}_{day_dt.month:02d}_{day_dt.year}_CZ.xlsx")
        
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            # Retry logic
            xlsx_bytes = None
            for attempt in range(2):
                try:
                    with urllib.request.urlopen(req, timeout=15) as r:
                        xlsx_bytes = r.read()
                    break
                except (urllib.error.URLError, TimeoutError) as e:
                    if attempt == 0:
                        time.sleep(0.5)
                        continue
                    raise
            if xlsx_bytes is None:
                return None
            
            # Parse XLSX - hledame sloupec H = "Posledni cena"
            points = []
            last_col_idx = None
            header_row_idx = None
            
            for row_idx, cells in enumerate(_xlsx_iter_rows(xlsx_bytes)):
                cell_dict = {col: val for col, val in cells}
                
                # Najit header
                if header_row_idx is None:
                    for col, val in cells:
                        if isinstance(val, str) and ("Posled" in val or "Last" in val):
                            last_col_idx = col
                            header_row_idx = row_idx
                            break
                    continue
                
                if last_col_idx and last_col_idx in cell_dict:
                    interval = cell_dict.get("A") or cell_dict.get("B")
                    last_val = cell_dict.get(last_col_idx)
                    
                    if interval and last_val is not None:
                        # Parse interval "00:00 - 00:15"
                        if isinstance(interval, str) and ":" in interval:
                            parts = interval.split("-")
                            time_start = parts[0].strip()
                            try:
                                hh, mm = time_start.split(":")
                                ts = day_dt.replace(hour=int(hh), minute=int(mm), second=0, microsecond=0)
                                try:
                                    price = float(last_val)
                                    points.append({
                                        "ts": ts.strftime("%Y-%m-%dT%H:%MZ"),
                                        "timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"),
                                        "last": price,
                                    })
                                except (ValueError, TypeError):
                                    pass
                            except ValueError:
                                pass
            
            result = {"date": cache_key, "data": points}
            _VDT_CACHE[cache_key] = (now_ts, result)
            return result
            
        except urllib.error.HTTPError as e:
            if e.code == 404:
                # Den jeste nema data, cache jako None
                _VDT_CACHE[cache_key] = (now_ts, None)
                return None
            print(f"  -> VDT {cache_key}: HTTP {e.code}", flush=True)
            return None
        except Exception as e:
            print(f"  -> VDT {cache_key}: {type(e).__name__}: {e}", flush=True)
            return None
    
    def _ote_vdt(self, qs):
        """VDT pro jeden den. Default = dnes. ?day=YYYY-MM-DD"""
        try:
            day_str = qs.get("day", [None])[0]
            if day_str:
                day = datetime.strptime(day_str, "%Y-%m-%d")
            else:
                day = datetime.now()
            
            result = self._ote_vdt_fetch_day(day)
            if result is None:
                self._json({"date": day.strftime("%Y-%m-%d"), "data": [], "error": "no data"})
                return
            
            self._json(result)
        except Exception as e:
            print(f"  -> /ote/vdt ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)
    
    def _ote_vdt_range(self, qs):
        """VDT za rozsah dni dozadu. ?days_back=30 (default 30)"""
        from concurrent.futures import ThreadPoolExecutor, as_completed
        try:
            days_back = int(qs.get("days_back", ["30"])[0])
            days_back = max(1, min(days_back, 60))
            
            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            target_dates = [today - timedelta(days=i) for i in range(days_back, -1, -1)]
            
            results = {}
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = {executor.submit(self._ote_vdt_fetch_day, d): d for d in target_dates}
                for future in as_completed(futures, timeout=60):
                    d = futures[future]
                    try:
                        result = future.result()
                        if result and result.get("data"):
                            results[d.strftime("%Y-%m-%d")] = result["data"]
                    except Exception as e:
                        print(f"  -> VDT range {d}: {e}", flush=True)
            
            # Slouceni vsech bodu chronologicky
            all_points = []
            for d_key in sorted(results.keys()):
                all_points.extend(results[d_key])
            
            self._json({
                "days_back": days_back,
                "data": all_points,
                "loaded_days": len(results),
                "failed_days": len(target_dates) - len(results),
                "fetched_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            })
        except Exception as e:
            print(f"  -> /ote/vdt/range ERROR: {e}", flush=True)
            import traceback; traceback.print_exc()
            self._json({"error": str(e)}, 502)

    # ============================================================
    # ENTSO-E Residual Load Forecast (DE)
    # ============================================================
    
    def _entsoe_residual(self, qs):
        """Residual Load Forecast pro DE (Total Load - Wind - Solar)."""
        try:
            now_dt = datetime.utcnow()
            ps = now_dt - timedelta(hours=24)
            pe = now_dt + timedelta(hours=24)
            
            DE_DOMAIN = "10Y1001A1001A82H"  # DE-LU
            
            # Total Load Forecast (A65, processType A01)
            load_xml, load_st = call_entsoe({
                "documentType": "A65",
                "processType": "A01",
                "outBiddingZone_Domain": DE_DOMAIN,
                "periodStart": fmt_entsoe_period(ps),
                "periodEnd": fmt_entsoe_period(pe),
            })
            load_pts = parse_entsoe_xml(load_xml, ps) if load_st == 200 else []
            
            # Wind+Solar Forecast (A69)
            wind_xml, wind_st = call_entsoe({
                "documentType": "A69",
                "processType": "A01",
                "psrType": "B19",  # Wind onshore
                "in_Domain": DE_DOMAIN,
                "periodStart": fmt_entsoe_period(ps),
                "periodEnd": fmt_entsoe_period(pe),
            })
            wind_pts = parse_entsoe_xml(wind_xml, ps) if wind_st == 200 else []
            
            solar_xml, solar_st = call_entsoe({
                "documentType": "A69",
                "processType": "A01",
                "psrType": "B16",  # Solar
                "in_Domain": DE_DOMAIN,
                "periodStart": fmt_entsoe_period(ps),
                "periodEnd": fmt_entsoe_period(pe),
            })
            solar_pts = parse_entsoe_xml(solar_xml, ps) if solar_st == 200 else []
            
            # Residual = Load - Wind - Solar (po timestampech)
            wind_map = {p["ts"]: p["value"] for p in wind_pts}
            solar_map = {p["ts"]: p["value"] for p in solar_pts}
            
            data = []
            for lp in load_pts:
                ts = lp["ts"]
                w = wind_map.get(ts, 0)
                s = solar_map.get(ts, 0)
                residual = lp["value"] - w - s
                data.append({
                    "ts": ts,
                    "berlin_time": ts,
                    "residual_load_mw": residual,
                    "demand_mw": lp["value"],
                    "wind_mw": w,
                    "solar_mw": s,
                })
            
            print(f"  -> /entsoe/residual: {len(data)} bodu", flush=True)
            self._json({"data": data, "country": "DE", "source": "ENTSO-E"})
        except Exception as e:
            print(f"  -> /entsoe/residual ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    # ============================================================
    # SMARD Residual Load (DE actual)
    # ============================================================
    
    _SMARD_CACHE = {"ts": 0, "data": None}
    
    def _smard_residual_load(self, qs):
        """Actual Residual Load z bundesnetzagentur.de/SMARD (DE)."""
        try:
            now_ts = time.time()
            if self._SMARD_CACHE["data"] and (now_ts - self._SMARD_CACHE["ts"]) < 900:
                out = dict(self._SMARD_CACHE["data"])
                out["_cache"] = "hit"
                self._json(out); return
            
            # SMARD filter 4359 = Residuallast 15-min
            now_dt = datetime.utcnow()
            # SMARD používá týdenní soubory - timestamp začátku týdne (pondělí 00:00)
            days_back = now_dt.weekday()
            week_start = (now_dt - timedelta(days=days_back)).replace(hour=0, minute=0, second=0, microsecond=0)
            ts_ms = int(week_start.timestamp() * 1000)
            
            url = f"https://www.smard.de/app/chart_data/4359/DE/4359_DE_quarterhour_{ts_ms}.json"
            r = requests.get(url, timeout=15)
            r.raise_for_status()
            j = r.json()
            
            data = []
            for row in j.get("series", []):
                if not row or len(row) < 2: continue
                ts_ms_pt, val = row[0], row[1]
                if val is None: continue
                dt = datetime.utcfromtimestamp(ts_ms_pt / 1000)
                data.append({
                    "ts": dt.strftime("%Y-%m-%dT%H:%MZ"),
                    "berlin_time": dt.strftime("%Y-%m-%dT%H:%M"),
                    "residual_load_actual_mw": float(val),
                })
            
            out = {"data": data, "country": "DE", "source": "SMARD", "_cache": "miss"}
            self._SMARD_CACHE = {"ts": now_ts, "data": out}
            print(f"  -> /smard/residual-load: {len(data)} bodu", flush=True)
            self._json(out)
        except Exception as e:
            print(f"  -> /smard/residual-load ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    # ============================================================
    # ENTSO-E Cross-border Physical Flows pro CZ
    # ============================================================
    
    _CROSSBORDER_CACHE = {"ts": 0, "data": None}
    
    def _entsoe_cross_border(self, qs):
        """Implicit IDA (Intraday Auction) allocated capacity pro CZ borders.
        Vraci kolik kapacity (MW) bylo prideleno v IDA1/2/3 pro kazdy smer.
        
        DocumentType A09 = Cross-zonal capacity allocation
        contract_MarketAgreement.Type A13 = Daily
        auction.Type A02 = Implicit
        classificationSequence_Position = 1/2/3 = IDA1/IDA2/IDA3
        """
        from concurrent.futures import ThreadPoolExecutor
        try:
            if not ENTSOE_TOKEN:
                self._json({"error": "ENTSOE_TOKEN neni nastaveny"}, 500); return
            
            now_ts = time.time()
            if self._CROSSBORDER_CACHE["data"] and (now_ts - self._CROSSBORDER_CACHE["ts"]) < 900:
                out = dict(self._CROSSBORDER_CACHE["data"])
                out["_cache"] = "hit"
                self._json(out); return
            
            # EIC kódy
            CZ = "10YCZ-CEPS-----N"
            DE = "10Y1001A1001A82H"
            AT = "10YAT-APG------L"
            SK = "10YSK-SEPS-----K"
            PL = "10YPL-AREA-----S"
            
            # Den D-1 (vcera 00:00) az D+1 (zitra 24:00)
            now_dt = datetime.utcnow()
            today_00 = now_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            ps = today_00 - timedelta(days=1)
            pe = today_00 + timedelta(days=2)
            period_start = fmt_entsoe_period(ps)
            period_end = fmt_entsoe_period(pe)
            
            def fetch_capacity(name, in_domain, out_domain):
                """Implicit allocated capacity pro jeden smer.
                Slouci IDA1+IDA2+IDA3 do jedne serie (vezme nejnovejsi pro kazdou hodinu).
                """
                try:
                    # Zkusi vsechny tri IDA - IDA3 je nejnovejsi takze prepise IDA1/2
                    merged = {}
                    for ida_pos in ["1", "2", "3"]:
                        xml_text, st = call_entsoe({
                            "documentType": "A09",
                            "contract_MarketAgreement.Type": "A13",
                            "auction.Type": "A02",
                            "classificationSequence_AttributeInstanceComponent.Position": ida_pos,
                            "in_Domain": in_domain,
                            "out_Domain": out_domain,
                            "periodStart": period_start,
                            "periodEnd": period_end,
                        })
                        if st != 200:
                            continue
                        pts = parse_entsoe_xml(xml_text, ps)
                        for p in pts:
                            # Pozdejsi IDA prepisuje drivejsi (IDA3 > IDA2 > IDA1)
                            merged[p["ts"]] = {"mw": p["value"], "ida": ida_pos}
                    
                    return (name, [{"ts": ts, "mw": v["mw"], "ida": v["ida"]} 
                                   for ts, v in sorted(merged.items())])
                except Exception as e:
                    print(f"  -> IDA capacity {name}: {e}", flush=True)
                    return (name, [])
            
            borders = [
                ("CZ_DE_in",  CZ, DE),  # DE → CZ (import kapacita)
                ("CZ_DE_out", DE, CZ),  # CZ → DE (export kapacita)
                ("CZ_AT_in",  CZ, AT),
                ("CZ_AT_out", AT, CZ),
                ("CZ_SK_in",  CZ, SK),
                ("CZ_SK_out", SK, CZ),
                ("CZ_PL_in",  CZ, PL),
                ("CZ_PL_out", PL, CZ),
            ]
            
            results = {}
            with ThreadPoolExecutor(max_workers=8) as executor:
                futures = [executor.submit(fetch_capacity, n, i, o) for n, i, o in borders]
                for f in futures:
                    name, points = f.result(timeout=30)
                    results[name] = points
            
            def merge_directions(in_key, out_key):
                """Slouci kapacity v obou smerech do bodu per ts."""
                in_map = {p["ts"]: p["mw"] for p in results.get(in_key, [])}
                out_map = {p["ts"]: p["mw"] for p in results.get(out_key, [])}
                all_ts = sorted(set(in_map.keys()) | set(out_map.keys()))
                return [{
                    "ts": ts,
                    "in_mw": in_map.get(ts, 0),     # kapacita pro IMPORT do CZ
                    "out_mw": out_map.get(ts, 0),   # kapacita pro EXPORT z CZ
                    "net_mw": in_map.get(ts, 0) - out_map.get(ts, 0),
                } for ts in all_ts]
            
            out = {
                "borders": {
                    "CZ_DE": merge_directions("CZ_DE_in", "CZ_DE_out"),
                    "CZ_AT": merge_directions("CZ_AT_in", "CZ_AT_out"),
                    "CZ_SK": merge_directions("CZ_SK_in", "CZ_SK_out"),
                    "CZ_PL": merge_directions("CZ_PL_in", "CZ_PL_out"),
                },
                "country": "CZ",
                "source": "ENTSO-E IDA Implicit Capacity (A09/A02)",
                "fetched_at": now_dt.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "_cache": "miss",
            }
            self._CROSSBORDER_CACHE = {"ts": now_ts, "data": out}
            
            total_pts = sum(len(b) for b in out["borders"].values())
            print(f"  -> /entsoe/cross-border (IDA): 4 borders, {total_pts} body", flush=True)
            self._json(out)
        except Exception as e:
            print(f"  -> /entsoe/cross-border ERROR: {e}", flush=True)
            import traceback; traceback.print_exc()
            self._json({"error": str(e)}, 502)

    # ============================================================
    # HTML stranky
    # ============================================================
    
    def _html_hruska(self):
        html_path = os.path.join(os.path.dirname(__file__), "hruska.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            inject = """<script>
window.addEventListener('DOMContentLoaded', () => {
  API_URL = window.location.origin;
  localStorage.setItem('ceps_api_url', API_URL);
  const modal = document.getElementById('setupModal');
  if (modal) modal.style.display = 'none';
  if (typeof loadVdtWind === 'function') loadVdtWind();
});
</script>"""
            html = html.replace("</body>", inject + "\n</body>")
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self._cors(); self.end_headers()
            self.wfile.write(body)
        except FileNotFoundError:
            self._json({"error": "hruska.html not found"}, 404)
    
    def _html_kapacity(self):
        html_path = os.path.join(os.path.dirname(__file__), "kapacity.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            inject = """<script>
window.addEventListener('DOMContentLoaded', () => {
  window.API_URL = window.location.origin;
  if (typeof loadCapacity === 'function') loadCapacity();
});
</script>"""
            html = html.replace("</body>", inject + "\n</body>")
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self._cors(); self.end_headers()
            self.wfile.write(body)
        except FileNotFoundError:
            self._json({"error": "kapacity.html not found"}, 404)

    def _html(self):
        html_path = os.path.join(os.path.dirname(__file__), "live_odchylky.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            inject = """<script>
window.addEventListener('DOMContentLoaded', () => {
  const apiUrl = window.location.origin;
  API_URL = apiUrl;
  document.getElementById('setupModal').style.display = 'none';
  loadAll();
  setRefresh(30);
});
</script>"""
            html = html.replace("</body>", inject + "\n</body>")
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self._cors(); self.end_headers()
            self.wfile.write(body)
        except FileNotFoundError:
            self._json({"error": "live_odchylky.html not found"}, 404)

    def _json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors(); self.end_headers()
        self.wfile.write(body)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    print(f"CEPS API server -> port {port}", flush=True)
    HTTPServer(("0.0.0.0", port), Handler).serve_forever()
