"""
CEPS API server - Railway / cloud deployment
Lokalne: python server.py
Render: automaticky pres Procfile

POZADAVKY (requirements.txt):
    requests
    (zadne extra knihovny - pouziva jen Python stdlib)
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import json, requests, xml.etree.ElementTree as ET, os, time, io, threading
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, timezone
# Pouzivame vlastni rychly XLSX streamer pres zipfile + iterparse
# (openpyxl je 30-50x pomalejsi - parsoval 120s misto 3s)

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

# Regelleistung.net (DE) aFRR ENERGY market
RL_BASE_URL = "https://www.regelleistung.net/apps/crds/api/v2/tenders/results/anonymous"
RL_CACHE_TTL_SEC = 30 * 60   # 30 minut
RL_REQUEST_TIMEOUT = 60      # XLSX muze byt vetsi soubor
_RL_CACHE = {}               # {(productType, market, date): (timestamp, data)}
_RL_REFRESH_LOCK = threading.Lock()
_RL_REFRESH_INFLIGHT = set()


def _request_with_retry(func, *args, retries=3, delay=1.5, **kwargs):
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
    if not ENTSOE_TOKEN:
        raise RuntimeError("ENTSOE_TOKEN neni nastaveny")
    p = dict(params)
    p["securityToken"] = ENTSOE_TOKEN
    r = _request_with_retry(requests.get, ENTSOE_URL, params=p, timeout=25)
    return r.text, r.status_code

def parse_entsoe_xml(xml_text, period_start_utc):
    body = xml_text
    import re
    body_no_ns = re.sub(r'\sxmlns="[^"]+"', '', body, count=1)
    try:
        root = ET.fromstring(body_no_ns)
    except ET.ParseError:
        return []
    points = []
    for ts_node in root.findall(".//TimeSeries"):
        period = ts_node.find("Period")
        if period is None: continue
        time_interval = period.find("timeInterval")
        if time_interval is None: continue
        start_text = time_interval.find("start").text
        res_text = period.find("resolution").text
        res_min = 15
        m = re.match(r"PT(\d+)M", res_text)
        if m: res_min = int(m.group(1))
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
            ts = start_dt + timedelta(minutes=(pos-1) * res_min)
            points.append({"ts": ts.strftime("%Y-%m-%dT%H:%MZ"), "value": qty})
    return points

def fmt_entsoe_period(d):
    return d.strftime("%Y%m%d%H%M")


# ============================================================
# Regelleistung.net (DE) aFRR ENERGY bidy
# ============================================================

def _rl_cache_get(key, allow_stale=False):
    entry = _RL_CACHE.get(key)
    if not entry: return None
    ts, data = entry
    age = time.time() - ts
    if age > RL_CACHE_TTL_SEC and not allow_stale:
        return None
    return data

def _rl_cache_age(key):
    entry = _RL_CACHE.get(key)
    if not entry: return None
    return time.time() - entry[0]

def _rl_cache_set(key, data):
    _RL_CACHE[key] = (time.time(), data)


def fetch_regelleistung_xlsx(product_type, market, delivery_date):
    """Stahne XLSX z regelleistung.net cpp-publisher API. Vraci bytes."""
    params = {
        "productType":  product_type,
        "market":       market,
        "exportFormat": "xlsx",
        "deliveryDate": delivery_date,
    }
    r = _request_with_retry(
        requests.get, RL_BASE_URL,
        params=params, timeout=RL_REQUEST_TIMEOUT,
        headers={"User-Agent": "Mozilla/5.0 (compatible; afrr-dashboard)"}
    )
    if r.status_code != 200:
        raise RuntimeError(f"Regelleistung HTTP {r.status_code}: {r.text[:200]}")
    if not r.content or len(r.content) < 100:
        raise RuntimeError(f"Regelleistung returned empty response ({len(r.content)} bytes)")
    return r.content


def _xlsx_iter_rows(xlsx_bytes):
    """Streamuje radky XLSX jako list stringu, BEZ openpyxl.
    Pouziva primy ZIP + iterparse - 30-50x rychlejsi nez openpyxl read_only.

    XLSX je v podstate ZIP s XML soubory. Potrebujeme:
    - xl/sharedStrings.xml: tabulka stringu (ulozena oddelene od bunek)
    - xl/worksheets/sheet1.xml: vlastni data
    """
    import zipfile
    import re as _re

    XL_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zf:
        # 1) Nacti sharedStrings.xml (pokud existuje)
        shared_strings = []
        try:
            with zf.open("xl/sharedStrings.xml") as ss_f:
                for _, elem in ET.iterparse(ss_f, events=("end",)):
                    tag = elem.tag
                    if tag == XL_NS + "si":
                        parts = []
                        for t in elem.iter(XL_NS + "t"):
                            if t.text: parts.append(t.text)
                        shared_strings.append("".join(parts))
                        elem.clear()
        except KeyError:
            shared_strings = []

        # 2) Stream sheet1.xml a vrat radky jako pole stringu
        sheet_path = None
        for name in zf.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                sheet_path = name
                break
        if sheet_path is None:
            return

        # Helper: A1 styl letter -> 0-based column index
        def col_index(ref):
            m = _re.match(r"([A-Z]+)", ref)
            if not m: return 0
            letters = m.group(1)
            idx = 0
            for ch in letters:
                idx = idx * 26 + (ord(ch) - ord("A") + 1)
            return idx - 1

        with zf.open(sheet_path) as sh_f:
            current_row = []
            for event, elem in ET.iterparse(sh_f, events=("start", "end")):
                tag = elem.tag

                if event == "start" and tag == XL_NS + "row":
                    current_row = []

                elif event == "end" and tag == XL_NS + "c":
                    cell_ref = elem.get("r", "")
                    cell_type = elem.get("t", "")
                    v_el = elem.find(XL_NS + "v")
                    is_el = elem.find(XL_NS + "is")

                    if v_el is not None and v_el.text is not None:
                        if cell_type == "s":
                            try:
                                val = shared_strings[int(v_el.text)]
                            except (IndexError, ValueError):
                                val = ""
                        elif cell_type == "b":
                            val = "TRUE" if v_el.text == "1" else "FALSE"
                        else:
                            val = v_el.text
                    elif is_el is not None:
                        parts = []
                        for t in is_el.iter(XL_NS + "t"):
                            if t.text: parts.append(t.text)
                        val = "".join(parts)
                    else:
                        val = ""

                    col_idx = col_index(cell_ref)
                    while len(current_row) < col_idx:
                        current_row.append("")
                    current_row.append(val)
                    elem.clear()

                elif event == "end" and tag == XL_NS + "row":
                    yield current_row
                    elem.clear()
                    current_row = []


def parse_afrr_energy_xlsx(xlsx_bytes):
    """Parser RESULT_LIST_ANONYMOUS pro aFRR ENERGY market.
    Pouziva rychly streaming XML parser misto openpyxl.

    Realne sloupce v XLSX:
        DELIVERY_DATE, TYPE_OF_RESERVES, PRODUCT,
        ENERGY_PRICE_[EUR/MWh], ENERGY_PRICE_PAYMENT_DIRECTION,
        OFFERED_CAPACITY_[MW], ALLOCATED_CAPACITY_[MW], COUNTRY, NOTE

    Format produktu: 'POS_069' / 'NEG_069' (1-indexovane: POS_069 = 17:00-17:15).
    """
    import re as _re

    rows_iter = _xlsx_iter_rows(xlsx_bytes)
    try:
        headers = next(rows_iter)
    except StopIteration:
        return {"slots": {}, "directions_available": [], "raw_columns": []}

    headers = [h.strip() if h else "" for h in headers]
    headers_lower = [h.lower() for h in headers]

    def find_col(*candidates):
        for cand in candidates:
            cand_l = cand.lower()
            for i, h in enumerate(headers_lower):
                if h == cand_l: return i
        for cand in candidates:
            cand_l = cand.lower()
            for i, h in enumerate(headers_lower):
                if cand_l in h: return i
        return -1

    idx_product   = find_col("PRODUCT", "PRODUCT_NAME")
    idx_volume    = find_col("OFFERED_CAPACITY_[MW]", "OFFERED_CAPACITY",
                             "OFFERED_ENERGY_VOLUME_MW", "VOLUME_MW")
    idx_price     = find_col("ENERGY_PRICE_[EUR/MWh]", "ENERGY_PRICE",
                             "OFFERED_ENERGY_PRICE_EUR_MWH", "PRICE_EUR_MWH")
    idx_payment   = find_col("ENERGY_PRICE_PAYMENT_DIRECTION", "PAYMENT_DIRECTION")
    idx_country   = find_col("COUNTRY")
    idx_reserves  = find_col("TYPE_OF_RESERVES", "RESERVE_TYPE")
    idx_allocated = find_col("ALLOCATED_CAPACITY_[MW]", "ALLOCATED_CAPACITY")

    if idx_product < 0 or idx_volume < 0 or idx_price < 0:
        return {
            "slots": {}, "directions_available": [], "raw_columns": headers,
            "_error": (f"Pozadovane sloupce nenalezeny. "
                       f"product_idx={idx_product}, vol_idx={idx_volume}, "
                       f"price_idx={idx_price}")
        }

    raw = {}
    directions_seen = set()
    sample_products = []
    seen_products = set()
    skipped = {"non_afrr": 0, "non_de": 0, "bad_product": 0, "bad_qh": 0,
               "bad_value": 0, "bad_direction": 0}
    total_rows = 0
    max_idx = max(idx_product, idx_volume, idx_price)

    for row in rows_iter:
        total_rows += 1
        if not row or len(row) <= max_idx:
            continue
        product = (row[idx_product] or "").strip()
        volume_s = (row[idx_volume] or "").strip()
        price_s  = (row[idx_price] or "").strip()
        if not product or not volume_s or not price_s:
            continue

        # Filter aFRR
        if idx_reserves >= 0 and len(row) > idx_reserves:
            rtype = (row[idx_reserves] or "").strip().lower()
            if rtype and "afrr" not in rtype:
                skipped["non_afrr"] += 1
                continue

        # Filter DE only
        if idx_country >= 0 and len(row) > idx_country:
            ctry = (row[idx_country] or "").strip().upper()
            if ctry and ctry != "DE":
                skipped["non_de"] += 1
                continue

        if product not in seen_products and len(sample_products) < 8:
            sample_products.append(product)
            seen_products.add(product)

        # Smer
        direction = None
        pu = product.upper()
        if pu.startswith("POS"):
            direction = "POS"
        elif pu.startswith("NEG"):
            direction = "NEG"
        elif idx_payment >= 0 and len(row) > idx_payment:
            pay = (row[idx_payment] or "").strip().upper()
            if "GRID_TO_PROVIDER" in pay:
                direction = "POS"
            elif "PROVIDER_TO_GRID" in pay:
                direction = "NEG"

        if direction is None:
            skipped["bad_direction"] += 1
            continue

        # QH index
        qh_idx = None
        m = _re.match(r"(?:POS|NEG)_(\d{1,3})", product, _re.IGNORECASE)
        if not m:
            m = _re.search(r"QH[_\-]?(\d{1,3})", product, _re.IGNORECASE)
        if not m:
            m = _re.search(r"(\d{1,3})(?!.*\d)", product)
        if m:
            try:
                qh_idx = int(m.group(1))
            except ValueError:
                pass

        if qh_idx is None or qh_idx < 1 or qh_idx > 96:
            skipped["bad_qh"] += 1
            continue

        # 1-indexovani: POS_001 = 00:00-00:15, POS_069 = 17:00-17:15
        start_min = (qh_idx - 1) * 15
        sh, sm = divmod(start_min, 60)
        eh, em = divmod(start_min + 15, 60)
        if eh == 24: eh, em = 0, 0
        slot_key = f"{sh:02d}:{sm:02d}-{eh:02d}:{em:02d}"

        try:
            vol_f = float(volume_s.replace(",", "."))
            price_f = float(price_s.replace(",", "."))
        except (ValueError, TypeError):
            skipped["bad_value"] += 1
            continue

        slot_data = raw.setdefault(slot_key, {"POS": [], "NEG": []})
        slot_data[direction].append({"price": price_f, "volume_mw": vol_f})
        directions_seen.add(direction)

    # Merit-order sort + kumulativni MW pro graf
    slots_processed = {}
    for slot_key, dirs in raw.items():
        slot_out = {}
        for dir_key, bids in dirs.items():
            if not bids:
                continue
            bids.sort(key=lambda x: x["price"])
            cum = 0.0
            ladder = []
            for b in bids:
                cum += b["volume_mw"]
                ladder.append({
                    "price":     round(b["price"], 4),
                    "volume_mw": round(b["volume_mw"], 4),
                    "cum_mw":    round(cum, 4),
                })
            slot_out[dir_key.lower()] = {
                "bids":      ladder,
                "total_mw":  round(cum, 4),
                "min_price": ladder[0]["price"],
                "max_price": ladder[-1]["price"],
                "count":     len(ladder),
            }
        if slot_out:
            slots_processed[slot_key] = slot_out

    return {
        "slots":     slots_processed,
        "directions_available": sorted(directions_seen),
        "raw_columns": headers,
        "_debug": {
            "total_rows": total_rows,
            "sample_products": sample_products,
            "skipped": skipped,
            "parser": "fast-iterparse",
            "col_indices": {
                "product": idx_product, "volume": idx_volume, "price": idx_price,
                "payment": idx_payment, "country": idx_country,
                "reserves": idx_reserves, "allocated": idx_allocated,
            },
        },
    }


def _rl_refresh_in_background(delivery_date):
    """Stahne XLSX v background threadu. Idempotent."""
    key = ("aFRR", "ENERGY", delivery_date)
    with _RL_REFRESH_LOCK:
        if key in _RL_REFRESH_INFLIGHT:
            return
        _RL_REFRESH_INFLIGHT.add(key)

    def worker():
        try:
            print(f"  -> [bg] Regelleistung refresh (XLSX): {delivery_date}", flush=True)
            t0 = time.time()
            xlsx_bytes = fetch_regelleistung_xlsx("aFRR", "ENERGY", delivery_date)
            t_fetch = time.time() - t0
            parsed = parse_afrr_energy_xlsx(xlsx_bytes)
            parsed["date"] = delivery_date
            parsed["fetched_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            _rl_cache_set(key, parsed)
            print(f"  -> [bg] OK fetched={t_fetch:.1f}s "
                  f"total={time.time()-t0:.1f}s slots={len(parsed.get('slots', {}))}",
                  flush=True)
        except Exception as e:
            print(f"  -> [bg] FAIL: {e}", flush=True)
        finally:
            with _RL_REFRESH_LOCK:
                _RL_REFRESH_INFLIGHT.discard(key)

    threading.Thread(target=worker, daemon=True).start()


def get_afrr_energy_data(delivery_date):
    """Stale-while-revalidate cache."""
    key = ("aFRR", "ENERGY", delivery_date)

    fresh = _rl_cache_get(key, allow_stale=False)
    if fresh is not None:
        out = dict(fresh); out["_cache"] = "hit"; out["_age_sec"] = int(_rl_cache_age(key) or 0)
        return out

    stale = _rl_cache_get(key, allow_stale=True)
    if stale is not None:
        age = _rl_cache_age(key) or 0
        print(f"  -> Regelleistung: cache stale (age={int(age)}s), bg refresh + vracim stale", flush=True)
        _rl_refresh_in_background(delivery_date)
        out = dict(stale); out["_cache"] = "stale"; out["_age_sec"] = int(age)
        return out

    print(f"  -> Regelleistung XLSX fetch (sync): aFRR ENERGY {delivery_date}", flush=True)
    t0 = time.time()
    xlsx_bytes = fetch_regelleistung_xlsx("aFRR", "ENERGY", delivery_date)
    t_fetch = time.time() - t0
    print(f"     downloaded {len(xlsx_bytes)} bytes in {t_fetch:.1f}s", flush=True)

    parsed = parse_afrr_energy_xlsx(xlsx_bytes)
    parsed["date"] = delivery_date
    parsed["fetched_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    print(f"     parsed in {time.time()-t0:.1f}s total: "
          f"{len(parsed.get('slots', {}))} slots, "
          f"directions={parsed.get('directions_available')}", flush=True)

    _rl_cache_set(key, parsed)
    out = dict(parsed); out["_cache"] = "miss"; out["_age_sec"] = 0
    return out


# ============================================================
# HTTP Handler
# ============================================================

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} {args[1]}", flush=True)

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors(); self.end_headers()

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self._cors(); self.end_headers()

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        g = lambda k, d="": qs.get(k, [d])[0]

        if parsed.path == "/health":
            self._json({"status": "ok", "time": datetime.now().isoformat(),
                        "version": "regelleistung-xlsx-v20-forecast"}); return

        if parsed.path in ("/", "/index.html", "/live_odchylky.html"):
            self._html(); return

        if parsed.path == "/entsoe/solar":
            self._entsoe_solar(qs); return

        if parsed.path == "/ote/spot":
            self._ote_spot(qs); return

        if parsed.path == "/ote/debug":
            # Debug endpoint - ukaze stav historie a cache
            history = globals().get("_OTE_HISTORY", {})
            cache = globals().get("_OTE_SPOT_CACHE", {})
            history_summary = {}
            for date_str, hours in history.items():
                history_summary[date_str] = {
                    "count": len(hours),
                    "first": hours[0] if hours else None,
                    "last": hours[-1] if hours else None,
                }
            self._json({
                "history_dates": list(history.keys()),
                "history_summary": history_summary,
                "cache_age_sec": int(time.time() - cache.get("ts", 0)) if cache.get("ts") else None,
                "cache_has_data": cache.get("data") is not None,
            }); return

        if parsed.path == "/weather":
            self._weather(qs); return

        if parsed.path == "/wind-de":
            self._wind_de(qs); return

        if parsed.path == "/forecast/de":
            self._forecast_de(qs); return

        if parsed.path == "/regelleistung/afrr-energy":
            self._regelleistung_afrr_energy(qs); return

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
        # Pridame fetched_at = cas kdy server zavolal CEPS API
        data["fetched_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        print(f"  -> {method}: {len(data['rows'])} radku, cols={data['columns']}", flush=True)
        self._json(data)

    def _entsoe_solar(self, qs):
        try:
            day_str = qs.get("day", [None])[0]
            if day_str:
                day = datetime.strptime(day_str, "%Y-%m-%d")
            else:
                day = datetime.now(timezone.utc).replace(tzinfo=None)
            ps = day.replace(hour=0, minute=0, second=0, microsecond=0)
            pe = ps + timedelta(hours=23)

            base = {
                "documentType": "A69",
                "psrType":      "B16",
                "in_Domain":    CZ_DOMAIN,
                "periodStart":  fmt_entsoe_period(ps),
                "periodEnd":    fmt_entsoe_period(pe),
            }

            da_xml, da_st = call_entsoe({**base, "processType": "A01"})
            da_points = parse_entsoe_xml(da_xml, ps) if da_st == 200 else []
            print(f"  -> ENTSO-E Solar DA: status={da_st}, points={len(da_points)}", flush=True)
            if da_st != 200:
                tok_len = len(ENTSOE_TOKEN) if ENTSOE_TOKEN else 0
                tok_preview = (ENTSOE_TOKEN[:6] + "..." + ENTSOE_TOKEN[-4:]) if tok_len > 12 else "(empty)"
                print(f"     token_len={tok_len}, token={tok_preview}", flush=True)
                print(f"     response[:300]: {da_xml[:300]}", flush=True)

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

    def _ote_spot(self, qs):
        """Vraci aktualni spotovou cenu + statistiky pro cely den.
        Cache 5 minut. Vystup: {price_czk, price_eur, hour, day_stats: {...}}
        """
        try:
            nocache = qs.get("nocache", ["0"])[0] in ("1", "true", "yes")
            # Cache - drzi se 5 minut
            if "_OTE_SPOT_CACHE" not in globals():
                globals()["_OTE_SPOT_CACHE"] = {"ts": 0, "data": None}
            # Historie hodinovych cen napric dny - format: {"2026-05-06": [{"hour": 0, "priceEur": ...}, ...], ...}
            if "_OTE_HISTORY" not in globals():
                globals()["_OTE_HISTORY"] = {}
            cache = globals()["_OTE_SPOT_CACHE"]
            history = globals()["_OTE_HISTORY"]
            now = time.time()
            if not nocache and cache["data"] and (now - cache["ts"]) < 300:
                out = dict(cache["data"]); out["_cache"] = "hit"
                self._json(out); return

            # 1) Stahni aktualni cenu (pro hour info)
            r1 = _request_with_retry(
                requests.get,
                "https://spotovaelektrina.cz/api/v1/price/get-actual-price-json",
                timeout=15,
                headers={"User-Agent": "Mozilla/5.0 (compatible; ceps-dashboard)"}
            )
            if r1.status_code != 200:
                self._json({"error": f"OTE actual HTTP {r1.status_code}"}, 502); return
            actual = r1.json()

            current_hour = actual.get("hour")
            current_eur = actual.get("priceEUR")
            current_czk = actual.get("priceCZK")

            # API neposkytuje "hour" klic - vezmeme aktualni hodinu z Berlin timezone
            if current_hour is None:
                now_utc = datetime.now(timezone.utc)
                # Berlin = CET/CEST (zima UTC+1, leto UTC+2)
                month = now_utc.month
                berlin_offset = 2 if 4 <= month <= 10 else 1
                berlin_now = now_utc + timedelta(hours=berlin_offset)
                current_hour = berlin_now.hour

            # Initialize
            day_stats = None
            tomorrow_stats = None
            last_8h = None  # 8 poslednich hodin pro mini KPI bunky

            # 2) Stahni 24h ceny pro statistiky
            try:
                r2 = _request_with_retry(
                    requests.get,
                    "https://spotovaelektrina.cz/api/v1/price/get-prices-json",
                    timeout=15,
                    headers={"User-Agent": "Mozilla/5.0 (compatible; ceps-dashboard)"}
                )
                if r2.status_code == 200:
                    day_data = r2.json()
                    # Format: {"hoursToday": [{"hour": 0, "priceCZK":..., "priceEur":...}, ...], "hoursTomorrow": [...]}
                    # POZOR: API vraci "priceEur" (male r), ne "priceEUR"!
                    hours_today = day_data.get("hoursToday", [])
                    hours_yesterday = []  # spotovaelektrina.cz nevraci

                    # Spocitej dnesni datum (Berlin TZ)
                    now_utc = datetime.now(timezone.utc)
                    month_now = now_utc.month
                    berlin_offset = 2 if 4 <= month_now <= 10 else 1
                    berlin_now = now_utc + timedelta(hours=berlin_offset)
                    today_date_str = berlin_now.strftime("%Y-%m-%d")
                    yesterday_date_str = (berlin_now - timedelta(days=1)).strftime("%Y-%m-%d")

                    # ULOZ dnesni data do historie - uchovavame max 7 dni zpetne
                    if hours_today:
                        history[today_date_str] = hours_today
                        # Vyciste stara data starsi nez 7 dni
                        cutoff = (berlin_now - timedelta(days=7)).strftime("%Y-%m-%d")
                        for old_date in list(history.keys()):
                            if old_date < cutoff:
                                del history[old_date]

                    # Vcerejsek z historie (kdyz uz ho mame z drivejska)
                    if yesterday_date_str in history:
                        hours_yesterday = history[yesterday_date_str]
                        print(f"  -> OTE history: vcerejsek ({yesterday_date_str}) z pameti, {len(hours_yesterday)}h", flush=True)

                    # Fallback pro vcerejsek - stahnout z ENTSO-E (uz mame token, stejna API jako solar)
                    if not hours_yesterday and current_hour is not None and current_hour < 8:
                        try:
                            yesterday_dt = berlin_now - timedelta(days=1)
                            ps = yesterday_dt.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(hours=berlin_offset)
                            pe = ps + timedelta(hours=24)
                            entsoe_xml, entsoe_st = call_entsoe({
                                "documentType": "A44",
                                "in_Domain":    CZ_DOMAIN,
                                "out_Domain":   CZ_DOMAIN,
                                "periodStart":  fmt_entsoe_period(ps),
                                "periodEnd":    fmt_entsoe_period(pe),
                            })
                            if entsoe_st == 200:
                                # Parsuj XML - struktura je <Period><Point><position>1</position><price.amount>X</price.amount></Point>...
                                import re as _re
                                body_no_ns = _re.sub(r'\sxmlns="[^"]+"', '', entsoe_xml, count=1)
                                root = ET.fromstring(body_no_ns)
                                # Najdi prvni TimeSeries -> Period -> body
                                hours_yesterday = []
                                for ts_node in root.findall(".//TimeSeries"):
                                    period = ts_node.find("Period")
                                    if period is None: continue
                                    res = period.find("resolution")
                                    if res is None or res.text != "PT60M": continue
                                    for pt in period.findall("Point"):
                                        pos_el = pt.find("position")
                                        price_el = pt.find("price.amount")
                                        if pos_el is None or price_el is None: continue
                                        try:
                                            pos = int(pos_el.text)
                                            price = float(price_el.text)
                                            # position 1 = 00:00, position 24 = 23:00
                                            hr = pos - 1
                                            if 0 <= hr <= 23:
                                                hours_yesterday.append({
                                                    "hour": hr,
                                                    "priceEur": price,
                                                    "priceCZK": int(price * 24.35),  # priblizny EUR->CZK
                                                })
                                        except (ValueError, TypeError):
                                            continue
                                    if hours_yesterday:
                                        break  # mame data, ostatni TimeSeries jsou duplicates
                                if hours_yesterday:
                                    print(f"  -> ENTSO-E A44: {len(hours_yesterday)}h vcerejsich CZ cen", flush=True)
                                    history[yesterday_date_str] = hours_yesterday
                            else:
                                print(f"  -> ENTSO-E A44 status {entsoe_st}", flush=True)
                        except Exception as e:
                            print(f"  -> ENTSO-E A44 fallback failed: {e}", flush=True)

                    if hours_today:
                        prices_eur = [h.get("priceEur") for h in hours_today if h.get("priceEur") is not None]
                        if prices_eur:
                            min_eur = min(prices_eur)
                            max_eur = max(prices_eur)
                            avg_eur = sum(prices_eur) / len(prices_eur)
                            min_hour = next((h["hour"] for h in hours_today if h.get("priceEur") == min_eur), None)
                            max_hour = next((h["hour"] for h in hours_today if h.get("priceEur") == max_eur), None)
                            spread = max_eur - min_eur
                            # Vs prumer pro aktualni hodinu
                            vs_avg_pct = None
                            if current_eur is not None and avg_eur > 0:
                                vs_avg_pct = ((current_eur - avg_eur) / avg_eur) * 100
                            day_stats = {
                                "min_eur": round(min_eur, 2),
                                "max_eur": round(max_eur, 2),
                                "avg_eur": round(avg_eur, 2),
                                "min_hour": min_hour,
                                "max_hour": max_hour,
                                "spread_eur": round(spread, 2),
                                "current_vs_avg_pct": round(vs_avg_pct, 1) if vs_avg_pct is not None else None,
                            }
                    # Zitrejsi statistiky (pokud OTE uz publikovalo - obvykle po 14:00 CET)
                    hours_tomorrow = day_data.get("hoursTomorrow", [])
                    if hours_tomorrow:
                        prices_t = [h.get("priceEur") for h in hours_tomorrow if h.get("priceEur") is not None]
                        if prices_t:
                            tmin = min(prices_t)
                            tmax = max(prices_t)
                            tavg = sum(prices_t) / len(prices_t)
                            tmin_h = next((h["hour"] for h in hours_tomorrow if h.get("priceEur") == tmin), None)
                            tmax_h = next((h["hour"] for h in hours_tomorrow if h.get("priceEur") == tmax), None)
                            tomorrow_stats = {
                                "min_eur": round(tmin, 2),
                                "max_eur": round(tmax, 2),
                                "avg_eur": round(tavg, 2),
                                "min_hour": tmin_h,
                                "max_hour": tmax_h,
                                "spread_eur": round(tmax - tmin, 2),
                                "published": True,
                            }
                        else:
                            tomorrow_stats = {"published": False}
                    else:
                        tomorrow_stats = {"published": False}

                    # last_8h: VZDY posledni 8 hodin koncici aktualni hodinou
                    # ch=3 -> [yesterday-20, yesterday-21, yesterday-22, yesterday-23, today-0, today-1, today-2, today-3]
                    if current_hour is not None and hours_today:
                        ch = current_hour
                        last_8h_list = []
                        for offset in range(-7, 1):
                            target_hour = ch + offset
                            is_yesterday = target_hour < 0  # FIX: dle target_hour, NE offsetu
                            if not is_yesterday:
                                # Aktualni den
                                row = next((h for h in hours_today if h.get("hour") == target_hour), None)
                                display_hour = target_hour
                            else:
                                # Predchozi den
                                yesterday_hour = target_hour + 24
                                row = next((h for h in hours_yesterday if h.get("hour") == yesterday_hour), None)
                                display_hour = yesterday_hour
                            if row and row.get("priceEur") is not None:
                                last_8h_list.append({
                                    "hour": display_hour,
                                    "price_eur": round(row.get("priceEur"), 2),
                                    "price_czk": row.get("priceCZK"),
                                    "is_current": offset == 0,
                                    "is_yesterday": is_yesterday,
                                })
                            else:
                                last_8h_list.append({
                                    "hour": display_hour,
                                    "price_eur": None,
                                    "price_czk": None,
                                    "is_current": offset == 0,
                                    "is_yesterday": is_yesterday,
                                })
                        last_8h = last_8h_list
            except Exception as e:
                print(f"  -> OTE day stats fetch failed: {e}", flush=True)

            out = {
                "price_czk": current_czk,
                "price_eur": current_eur,
                "hour": current_hour,
                "day_stats": day_stats,
                "tomorrow_stats": tomorrow_stats,
                "last_8h": last_8h,
                "source": "spotovaelektrina.cz",
                "fetched_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "_cache": "miss",
            }
            cache["ts"] = now
            cache["data"] = {k: v for k, v in out.items() if k != "_cache"}
            self._json(out)
        except Exception as e:
            print(f"  -> OTE Spot ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    def _weather(self, qs):
        """Vraci aktualni pocasi pro Prahu + zitrejsi forecast.
        Open-Meteo API (free, bez API klice).
        Cache 30 minut.
        """
        try:
            if "_WEATHER_CACHE" not in globals():
                globals()["_WEATHER_CACHE"] = {"ts": 0, "data": None}
            cache = globals()["_WEATHER_CACHE"]
            now = time.time()
            if cache["data"] and (now - cache["ts"]) < 1800:
                out = dict(cache["data"]); out["_cache"] = "hit"
                self._json(out); return

            # Praha souradnice
            url = ("https://api.open-meteo.com/v1/forecast"
                   "?latitude=50.0755&longitude=14.4378"
                   "&current=temperature_2m,wind_speed_10m,weather_code,cloud_cover"
                   "&daily=temperature_2m_max,temperature_2m_min,wind_speed_10m_max,weather_code,sunshine_duration"
                   "&timezone=Europe%2FBerlin&forecast_days=2")

            r = _request_with_retry(
                requests.get, url, timeout=15,
                headers={"User-Agent": "Mozilla/5.0 (compatible; ceps-dashboard)"}
            )
            if r.status_code != 200:
                self._json({"error": f"Open-Meteo HTTP {r.status_code}"}, 502); return

            data = r.json()
            current = data.get("current", {})
            daily = data.get("daily", {})

            # WMO weather codes -> ikony emoji + text
            def wcode_info(code):
                if code is None: return ("", "—")
                c = int(code)
                if c == 0:                      return ("☀️", "jasno")
                if c in (1, 2):                 return ("🌤️", "polojasno")
                if c == 3:                      return ("☁️", "zataženo")
                if c in (45, 48):               return ("🌫️", "mlha")
                if c in (51, 53, 55, 56, 57):   return ("🌦️", "mrholení")
                if c in (61, 63, 65, 66, 67):   return ("🌧️", "déšť")
                if c in (71, 73, 75, 77):       return ("🌨️", "sněžení")
                if c in (80, 81, 82):           return ("🌦️", "přeháňky")
                if c in (85, 86):               return ("❄️", "sněhové přeháňky")
                if c in (95, 96, 99):           return ("⛈️", "bouřka")
                return ("", "—")

            # Today (index 0)
            today_max = daily.get("temperature_2m_max", [None, None])[0]
            today_min = daily.get("temperature_2m_min", [None, None])[0]
            today_wind = daily.get("wind_speed_10m_max", [None, None])[0]
            today_code = daily.get("weather_code", [None, None])[0]
            today_sun = daily.get("sunshine_duration", [None, None])[0]
            today_icon, today_desc = wcode_info(today_code)

            # Tomorrow (index 1)
            tom_max = daily.get("temperature_2m_max", [None, None])[1] if len(daily.get("temperature_2m_max", [])) > 1 else None
            tom_min = daily.get("temperature_2m_min", [None, None])[1] if len(daily.get("temperature_2m_min", [])) > 1 else None
            tom_wind = daily.get("wind_speed_10m_max", [None, None])[1] if len(daily.get("wind_speed_10m_max", [])) > 1 else None
            tom_code = daily.get("weather_code", [None, None])[1] if len(daily.get("weather_code", [])) > 1 else None
            tom_sun = daily.get("sunshine_duration", [None, None])[1] if len(daily.get("sunshine_duration", [])) > 1 else None
            tom_icon, tom_desc = wcode_info(tom_code)

            out = {
                "current": {
                    "temp_c": current.get("temperature_2m"),
                    "wind_ms": current.get("wind_speed_10m"),
                    "cloud_pct": current.get("cloud_cover"),
                    "weather_code": current.get("weather_code"),
                },
                "today": {
                    "temp_max": today_max,
                    "temp_min": today_min,
                    "wind_max_ms": today_wind,
                    "icon": today_icon,
                    "desc": today_desc,
                    "sunshine_h": round(today_sun / 3600, 1) if today_sun else None,
                },
                "tomorrow": {
                    "temp_max": tom_max,
                    "temp_min": tom_min,
                    "wind_max_ms": tom_wind,
                    "icon": tom_icon,
                    "desc": tom_desc,
                    "sunshine_h": round(tom_sun / 3600, 1) if tom_sun else None,
                },
                "location": "Praha",
                "source": "open-meteo.com",
                "fetched_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "_cache": "miss",
            }
            cache["ts"] = now
            cache["data"] = {k: v for k, v in out.items() if k != "_cache"}
            self._json(out)
        except Exception as e:
            print(f"  -> Weather ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    def _forecast_de(self, qs):
        """EpexPredictor.batzill.com - statisticky model pro DE day-ahead ceny.
        DE je proxy pro CZ (Market Coupling - silna korelace).
        Vraci dalsich N hodin predikce. Cache 30 min.
        """
        try:
            if "_FORECAST_DE_CACHE" not in globals():
                globals()["_FORECAST_DE_CACHE"] = {"ts": 0, "data": None}
            cache = globals()["_FORECAST_DE_CACHE"]
            now = time.time()
            if cache["data"] and (now - cache["ts"]) < 1800:
                out = dict(cache["data"]); out["_cache"] = "hit"
                self._json(out); return

            # EpexPredictor API - region=DE, 48h forecast
            r = _request_with_retry(
                requests.get,
                "https://epexpredictor.batzill.com/prices_short",
                params={
                    "region": "DE",
                    "hours": 48,
                    "unit": "EUR_PER_MWH",
                },
                timeout=15,
                headers={"User-Agent": "Mozilla/5.0 (compatible; ceps-dashboard)"}
            )
            if r.status_code != 200:
                self._json({"error": f"EpexPredictor HTTP {r.status_code}"}, 502); return

            data = r.json()
            # Format: {"s": [unix_timestamp, ...], "t": [price, ...]}
            timestamps = data.get("s", [])
            prices = data.get("t", [])
            if not timestamps or not prices or len(timestamps) != len(prices):
                self._json({"error": "EpexPredictor returned empty/invalid data",
                            "raw_keys": list(data.keys())}, 502); return

            # Spocitej Berlin TZ aktualni hodinu pro filtrovani "next N"
            now_utc = datetime.now(timezone.utc)
            month_now = now_utc.month
            berlin_offset = 2 if 4 <= month_now <= 10 else 1
            berlin_now = now_utc + timedelta(hours=berlin_offset)
            current_unix = int(now_utc.timestamp())

            # Vyrob seznam {hour, date, price_eur} pro vsechny budouci body
            # EpexPredictor vraci unix_seconds, prevedem na Berlin TZ
            forecast_list = []
            for i, ts in enumerate(timestamps):
                if ts <= current_unix - 1800:  # ignoruj < 30min stare
                    continue
                price = prices[i]
                if price is None: continue
                # Convert unix -> Berlin time
                point_utc = datetime.fromtimestamp(ts, tz=timezone.utc)
                point_berlin = point_utc + timedelta(hours=berlin_offset)
                forecast_list.append({
                    "unix": ts,
                    "hour": point_berlin.hour,
                    "date": point_berlin.strftime("%Y-%m-%d"),
                    "price_eur": round(float(price), 2),
                })

            out = {
                "next_8h": forecast_list[:8],  # nejblizsich 8 hodin
                "next_24h": forecast_list[:24],
                "all_count": len(forecast_list),
                "source": "epexpredictor.batzill.com",
                "region": "DE (proxy for CZ via Market Coupling)",
                "fetched_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "_cache": "miss",
            }
            cache["ts"] = now
            cache["data"] = {k: v for k, v in out.items() if k != "_cache"}
            self._json(out)
        except Exception as e:
            print(f"  -> /forecast/de ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    def _wind_de(self, qs):
        """Vraci aktualni vitr v severnim Nemecku - prumer ze 3 mest:
        Hamburg, Bremerhaven, Husum. Open-Meteo API. Cache 30 minut.
        """
        try:
            if "_WIND_DE_CACHE" not in globals():
                globals()["_WIND_DE_CACHE"] = {"ts": 0, "data": None}
            cache = globals()["_WIND_DE_CACHE"]
            now = time.time()
            if cache["data"] and (now - cache["ts"]) < 1800:
                out = dict(cache["data"]); out["_cache"] = "hit"
                self._json(out); return

            # 3 lokality v sev. Nemecku (blizko velkym vetrnym parkum)
            cities = [
                {"name": "Hamburg",     "lat": 53.5511, "lon": 9.9937},
                {"name": "Bremerhaven", "lat": 53.5396, "lon": 8.5810},
                {"name": "Husum",       "lat": 54.4858, "lon": 9.0524},
            ]

            def wcode_info(code):
                if code is None: return ("", "—")
                c = int(code)
                if c == 0:                      return ("☀️", "jasno")
                if c in (1, 2):                 return ("🌤️", "polojasno")
                if c == 3:                      return ("☁️", "zataženo")
                if c in (45, 48):               return ("🌫️", "mlha")
                if c in (51, 53, 55, 56, 57):   return ("🌦️", "mrholení")
                if c in (61, 63, 65, 66, 67):   return ("🌧️", "déšť")
                if c in (71, 73, 75, 77):       return ("🌨️", "sněžení")
                if c in (80, 81, 82):           return ("🌦️", "přeháňky")
                if c in (85, 86):               return ("❄️", "sněhové přeháňky")
                if c in (95, 96, 99):           return ("⛈️", "bouřka")
                return ("", "—")

            # Stahni data pro vsechna 3 mesta paralelne (sekvencne, ale rychle)
            results = []
            for city in cities:
                url = ("https://api.open-meteo.com/v1/forecast"
                       f"?latitude={city['lat']}&longitude={city['lon']}"
                       "&current=temperature_2m,wind_speed_10m,weather_code"
                       "&daily=temperature_2m_max,temperature_2m_min,wind_speed_10m_max,weather_code,sunshine_duration"
                       "&timezone=Europe%2FBerlin&forecast_days=2")
                try:
                    r = _request_with_retry(
                        requests.get, url, timeout=10,
                        headers={"User-Agent": "Mozilla/5.0 (compatible; ceps-dashboard)"}
                    )
                    if r.status_code == 200:
                        results.append(r.json())
                except Exception as e:
                    print(f"  -> wind-de {city['name']} fail: {e}", flush=True)

            if not results:
                self._json({"error": "all 3 cities failed"}, 502); return

            # Helpery pro prumerovani
            def avg_or_none(values):
                vals = [v for v in values if v is not None]
                if not vals: return None
                return sum(vals) / len(vals)

            def safe_at(d, key, idx=None):
                """daily nebo current value; idx=None pro current, idx=0/1 pro daily"""
                if idx is None:
                    return d.get("current", {}).get(key)
                arr = d.get("daily", {}).get(key, [])
                return arr[idx] if len(arr) > idx else None

            # Prumeruj hodnoty napric mesty
            cur_temp = avg_or_none([safe_at(d, "temperature_2m") for d in results])
            cur_wind = avg_or_none([safe_at(d, "wind_speed_10m") for d in results])
            # Pro weather_code vezmeme nejhorsi (nejvyssi cislo = nejvic stres)
            cur_codes = [safe_at(d, "weather_code") for d in results]
            cur_codes = [c for c in cur_codes if c is not None]
            cur_code = max(cur_codes) if cur_codes else None

            today_max = avg_or_none([safe_at(d, "temperature_2m_max", 0) for d in results])
            today_min = avg_or_none([safe_at(d, "temperature_2m_min", 0) for d in results])
            today_wind_max = avg_or_none([safe_at(d, "wind_speed_10m_max", 0) for d in results])
            today_codes = [safe_at(d, "weather_code", 0) for d in results]
            today_codes = [c for c in today_codes if c is not None]
            today_code = max(today_codes) if today_codes else None
            today_sun = avg_or_none([safe_at(d, "sunshine_duration", 0) for d in results])
            today_icon, today_desc = wcode_info(today_code)

            tom_max = avg_or_none([safe_at(d, "temperature_2m_max", 1) for d in results])
            tom_min = avg_or_none([safe_at(d, "temperature_2m_min", 1) for d in results])
            tom_wind_max = avg_or_none([safe_at(d, "wind_speed_10m_max", 1) for d in results])
            tom_codes = [safe_at(d, "weather_code", 1) for d in results]
            tom_codes = [c for c in tom_codes if c is not None]
            tom_code = max(tom_codes) if tom_codes else None
            tom_sun = avg_or_none([safe_at(d, "sunshine_duration", 1) for d in results])
            tom_icon, tom_desc = wcode_info(tom_code)

            out = {
                "current": {
                    "temp_c": round(cur_temp, 1) if cur_temp is not None else None,
                    "wind_ms": round(cur_wind, 1) if cur_wind is not None else None,
                    "weather_code": cur_code,
                },
                "today": {
                    "temp_max": round(today_max, 1) if today_max is not None else None,
                    "temp_min": round(today_min, 1) if today_min is not None else None,
                    "wind_max_ms": round(today_wind_max, 1) if today_wind_max is not None else None,
                    "icon": today_icon,
                    "desc": today_desc,
                    "sunshine_h": round(today_sun / 3600, 1) if today_sun else None,
                },
                "tomorrow": {
                    "temp_max": round(tom_max, 1) if tom_max is not None else None,
                    "temp_min": round(tom_min, 1) if tom_min is not None else None,
                    "wind_max_ms": round(tom_wind_max, 1) if tom_wind_max is not None else None,
                    "icon": tom_icon,
                    "desc": tom_desc,
                    "sunshine_h": round(tom_sun / 3600, 1) if tom_sun else None,
                },
                "location": "DE-sever (HH+HB+Husum)",
                "cities_ok": len(results),
                "source": "open-meteo.com",
                "fetched_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "_cache": "miss",
            }
            cache["ts"] = now
            cache["data"] = {k: v for k, v in out.items() if k != "_cache"}
            self._json(out)
        except Exception as e:
            print(f"  -> wind-de ERROR: {e}", flush=True)
            self._json({"error": str(e)}, 502)

    def _regelleistung_afrr_energy(self, qs):
        try:
            date_str = qs.get("date", [None])[0]
            nocache = qs.get("nocache", ["0"])[0] in ("1", "true", "yes")
            if not date_str:
                now_utc = datetime.now(timezone.utc)
                month = now_utc.month
                berlin_offset = 2 if 4 <= month <= 10 else 1
                berlin_now = now_utc + timedelta(hours=berlin_offset)
                date_str = berlin_now.strftime("%Y-%m-%d")

            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                self._json({"error": f"Invalid date format: {date_str}, use YYYY-MM-DD"}, 400)
                return

            if nocache:
                key = ("aFRR", "ENERGY", date_str)
                _RL_CACHE.pop(key, None)
                print(f"  -> Regelleistung: cache invalidated for {date_str}", flush=True)

            data = get_afrr_energy_data(date_str)
            self._json(data)
        except Exception as e:
            print(f"  -> Regelleistung aFRR ENERGY ERROR: {e}", flush=True)
            import traceback
            traceback.print_exc()
            self._json({"error": str(e)}, 502)

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
  setRefresh(15);
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
        try:
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self._cors(); self.end_headers()
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def handle_one_request(self):
        try:
            super().handle_one_request()
        except (BrokenPipeError, ConnectionResetError):
            pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    public_url = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")

    # Warmup: stahni dnesni Regelleistung data hned po startu
    def _warmup_regelleistung():
        time.sleep(3)
        try:
            now_utc = datetime.now(timezone.utc)
            month = now_utc.month
            berlin_offset = 2 if 4 <= month <= 10 else 1
            berlin_now = now_utc + timedelta(hours=berlin_offset)
            today = berlin_now.strftime("%Y-%m-%d")
            print(f"[warmup] Pre-fetching Regelleistung aFRR ENERGY for {today}", flush=True)
            t0 = time.time()
            get_afrr_energy_data(today)
            print(f"[warmup] Done in {time.time()-t0:.1f}s", flush=True)
        except Exception as e:
            print(f"[warmup] FAIL: {e}", flush=True)
    threading.Thread(target=_warmup_regelleistung, daemon=True).start()

    if public_url:
        def keepalive():
            time.sleep(60)
            while True:
                try:
                    r = requests.get(f"{public_url}/health", timeout=10)
                    print(f"[keepalive] ping {public_url}/health -> {r.status_code}", flush=True)
                except Exception as e:
                    print(f"[keepalive] FAIL: {e}", flush=True)
                time.sleep(10 * 60)
        threading.Thread(target=keepalive, daemon=True).start()
        print(f"[keepalive] thread started, pinging {public_url}/health every 10 min", flush=True)
    else:
        print("[keepalive] RENDER_EXTERNAL_URL not set - keepalive disabled", flush=True)
    print(f"CEPS API server -> port {port}", flush=True)
    print(f"VERSION: regelleistung-xlsx-v20-forecast", flush=True)
    HTTPServer(("0.0.0.0", port), Handler).serve_forever()
