"""
CEPS API server - Railway / cloud deployment
Lokalne: python server.py
Railway: automaticky pres Procfile
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import json, requests, xml.etree.ElementTree as ET, os, time, io
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook  # pip install openpyxl

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

# Regelleistung.net (DE+AT) aFRR ENERGY market
RL_BASE_URL = "https://www.regelleistung.net/apps/crds/api/v2/tenders/results/anonymous"
RL_CACHE_TTL_SEC = 10 * 60   # 10 minut
RL_REQUEST_TIMEOUT = 45      # XLSX muze byt vetsi soubor
_RL_CACHE = {}               # {(productType, market, date): (timestamp, data)}


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
    """
    body = xml_text
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
    """Format datetime -> 'YYYYMMDDhhmm' pro ENTSO-E."""
    return d.strftime("%Y%m%d%H%M")


# ============================================================
# Regelleistung.net (DE+AT) aFRR ENERGY bidy
# ============================================================

def _rl_cache_get(key):
    entry = _RL_CACHE.get(key)
    if not entry: return None
    ts, data = entry
    if time.time() - ts > RL_CACHE_TTL_SEC: return None
    return data


def _rl_cache_set(key, data):
    _RL_CACHE[key] = (time.time(), data)


def fetch_regelleistung_xlsx(product_type, market, delivery_date):
    """Stahne XLSX z regelleistung.net cpp-publisher API. Vraci bytes."""
    params = {
        "productType":  product_type,   # 'aFRR', 'mFRR', 'FCR'
        "market":       market,         # 'ENERGY' nebo 'CAPACITY'
        "exportFormat": "xlsx",
        "deliveryDate": delivery_date,  # 'YYYY-MM-DD'
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


def parse_afrr_energy_xlsx(xlsx_bytes):
    """Parser RESULT_LIST_ANONYMOUS pro aFRR ENERGY market.
    Format produktu: 'POS_QH_064_2026-05-06' (POS/NEG, QH index 0-95).
    """
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows_iter)
    except StopIteration:
        return {"slots": {}, "directions_available": [], "raw_columns": []}

    headers = [str(h).strip() if h is not None else "" for h in header_row]
    headers_lower = [h.lower() for h in headers]

    def find_col(*candidates):
        """Hledame fuzzy: nejdriv exact, pak substring (case-insensitive)."""
        for cand in candidates:
            cand_l = cand.lower()
            for i, h in enumerate(headers_lower):
                if h == cand_l: return i
        for cand in candidates:
            cand_l = cand.lower()
            for i, h in enumerate(headers_lower):
                if cand_l in h: return i
        return -1

    # Realne nazvy sloupcu z regelleistung.net (zjisteno z odpovedi):
    #   PRODUCT, ENERGY_PRICE_[EUR/MWh], ENERGY_PRICE_PAYMENT_DIRECTION,
    #   OFFERED_CAPACITY_[MW], ALLOCATED_CAPACITY_[MW], COUNTRY, TYPE_OF_RESERVES, NOTE
    idx_product   = find_col("PRODUCT", "PRODUCT_NAME")
    idx_volume    = find_col("OFFERED_CAPACITY_[MW]", "OFFERED_CAPACITY",
                             "OFFERED_ENERGY_VOLUME_MW", "VOLUME_MW", "OFFERED_VOLUME")
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

    slots = {}
    directions_seen = set()
    sample_products = []      # debug - prvnich 5 unikatnich produktovych nazvu
    seen_products = set()
    skipped_reasons = {"non_afrr": 0, "bad_product": 0, "bad_qh": 0, "bad_value": 0}
    total_rows = 0

    for row in rows_iter:
        total_rows += 1
        if row is None or len(row) <= max(idx_product, idx_volume, idx_price):
            continue
        product = row[idx_product]
        volume  = row[idx_volume]
        price   = row[idx_price]
        if product is None or volume is None or price is None: continue

        # Filter jen aFRR (TYPE_OF_RESERVES muze byt aFRR/mFRR/FCR)
        if idx_reserves >= 0 and len(row) > idx_reserves:
            rtype = str(row[idx_reserves] or "").strip().lower()
            if rtype and "afrr" not in rtype:
                skipped_reasons["non_afrr"] += 1
                continue

        # Filter jen Nemecko (regelleistung publikuje DE+AT joint, my chceme jen DE)
        if idx_country >= 0 and len(row) > idx_country:
            ctry = str(row[idx_country] or "").strip().upper()
            if ctry and ctry != "DE":
                skipped_reasons.setdefault("non_de", 0)
                skipped_reasons["non_de"] += 1
                continue

        product_str = str(product).strip()
        if product_str not in seen_products and len(sample_products) < 5:
            sample_products.append(product_str)
            seen_products.add(product_str)

        # Smer odvodime z PAYMENT_DIRECTION sloupce nebo z prefixu PRODUCT
        direction = None
        if idx_payment >= 0 and len(row) > idx_payment:
            pay = str(row[idx_payment] or "").strip().upper()
            # GRID_TO_PROVIDER = TSO plati providerovi za UP energy (POS)
            # PROVIDER_TO_GRID = provider plati TSO za moznost dodat DOWN (NEG)
            if "GRID_TO_PROVIDER" in pay or pay == "POS" or pay == "POSITIVE":
                direction = "POS"
            elif "PROVIDER_TO_GRID" in pay or pay == "NEG" or pay == "NEGATIVE":
                direction = "NEG"

        # Fallback: prefix v PRODUCT
        if direction is None:
            pu = product_str.upper()
            if pu.startswith("POS_") or "_POS_" in pu or pu.startswith("POS"):
                direction = "POS"
            elif pu.startswith("NEG_") or "_NEG_" in pu or pu.startswith("NEG"):
                direction = "NEG"

        if direction is None:
            skipped_reasons["bad_product"] += 1
            continue

        # QH index - hledame "QH_NNN" v product nazvu, nebo "_NNN_" jako 3-cifernou cast
        qh_idx = None
        import re as _re
        m = _re.search(r"QH[_\-]?(\d{1,3})", product_str, _re.IGNORECASE)
        if not m:
            m = _re.search(r"_(\d{3})_", product_str)
        if not m:
            m = _re.search(r"_(\d{1,3})$", product_str)
        if m:
            try:
                qh_idx = int(m.group(1))
            except ValueError:
                pass

        if qh_idx is None or qh_idx < 0 or qh_idx > 95:
            skipped_reasons["bad_qh"] += 1
            continue

        start_min = qh_idx * 15
        sh, sm = divmod(start_min, 60)
        eh, em = divmod(start_min + 15, 60)
        if eh == 24: eh, em = 0, 0
        slot_key = f"{sh:02d}:{sm:02d}-{eh:02d}:{em:02d}"

        try:
            vol_f = float(volume); price_f = float(price)
        except (ValueError, TypeError):
            skipped_reasons["bad_value"] += 1
            continue

        # Pro NEG bidy je cena "kolik provider plati TSO" - v balancing.services
        # vizualizaci jsou ukazany jako kladne hodnoty na merit-order curve.
        # Necham puvodni znamenko, frontend si poradi.

        slots.setdefault(slot_key, []).append({
            "volume_mw": vol_f, "price": price_f, "direction": direction,
        })
        directions_seen.add(direction)

    return {
        "slots": slots,
        "directions_available": sorted(directions_seen),
        "raw_columns": headers,
        "_debug": {
            "total_rows": total_rows,
            "sample_products": sample_products,
            "skipped": skipped_reasons,
            "col_indices": {
                "product": idx_product, "volume": idx_volume, "price": idx_price,
                "payment": idx_payment, "country": idx_country,
                "reserves": idx_reserves, "allocated": idx_allocated,
            },
        },
    }


def get_afrr_energy_data(delivery_date):
    """Hlavni funkce s cache."""
    key = ("aFRR", "ENERGY", delivery_date)
    cached = _rl_cache_get(key)
    if cached is not None:
        out = dict(cached); out["_cache"] = "hit"
        return out

    print(f"  -> Regelleistung XLSX fetch: aFRR ENERGY {delivery_date}", flush=True)
    xlsx_bytes = fetch_regelleistung_xlsx("aFRR", "ENERGY", delivery_date)
    print(f"     downloaded {len(xlsx_bytes)} bytes", flush=True)

    parsed = parse_afrr_energy_xlsx(xlsx_bytes)
    parsed["date"] = delivery_date
    parsed["fetched_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    print(f"     parsed: {len(parsed.get('slots', {}))} slots, "
          f"directions={parsed.get('directions_available')}", flush=True)

    _rl_cache_set(key, parsed)
    out = dict(parsed); out["_cache"] = "miss"
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
        # Render health check posila HEAD - vrat 200 OK
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
            self._json({"status": "ok", "time": datetime.now().isoformat()}); return

        if parsed.path in ("/", "/index.html", "/live_odchylky.html"):
            self._html(); return

        # ENTSO-E Solar Forecast (Day-ahead + Intraday) pro CR
        if parsed.path == "/entsoe/solar":
            self._entsoe_solar(qs); return

        # Regelleistung.net DE+AT aFRR energy bids (merit-order ladder)
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
        print(f"  -> {method}: {len(data['rows'])} radku, cols={data['columns']}", flush=True)
        self._json(data)

    def _entsoe_solar(self, qs):
        """Solar Forecast pro CR - vraci Day-ahead a Intraday serie pro dnesek.
        Optional ?day=YYYY-MM-DD pro jiny den.
        """
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

    def _regelleistung_afrr_energy(self, qs):
        """Vraci aFRR energy bidy pro DE+AT pro dany den.
        Optional ?date=YYYY-MM-DD (default: dnesek v Berlin tz).
        Optional ?nocache=1 pro vynucene znovustazeni (preskoci cache).
        """
        try:
            date_str = qs.get("date", [None])[0]
            nocache = qs.get("nocache", ["0"])[0] in ("1", "true", "yes")
            if not date_str:
                # Default: dnesni datum v Berlin tz (kde se aFRR aukce dela)
                now_utc = datetime.now(timezone.utc)
                month = now_utc.month
                # Zjednodusene CEST/CET (4-10 = leto, jinak zima)
                berlin_offset = 2 if 4 <= month <= 10 else 1
                berlin_now = now_utc + timedelta(hours=berlin_offset)
                date_str = berlin_now.strftime("%Y-%m-%d")

            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                self._json({"error": f"Invalid date format: {date_str}, use YYYY-MM-DD"}, 400)
                return

            if nocache:
                # Vyhod cache pro tento klic, takze get_afrr_energy_data
                # bude muset stahnout cerstva data
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
        """Override pro tichy handling BrokenPipe."""
        try:
            super().handle_one_request()
        except (BrokenPipeError, ConnectionResetError):
            pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    public_url = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")
    if public_url:
        import threading
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
    HTTPServer(("0.0.0.0", port), Handler).serve_forever()
