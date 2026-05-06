"""
CEPS API server - Railway / cloud deployment
Lokalne: python server.py
Railway: automaticky pres Procfile
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import json, requests, xml.etree.ElementTree as ET, os, time
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, timezone

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

        # ENTSO-E aFRR aktivace + ceny (z PICASSO/CZ)
        if parsed.path == "/entsoe/afrr":
            self._entsoe_afrr(qs); return

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
                day = datetime.now(timezone.utc).replace(tzinfo=None)
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
            if da_st != 200:
                # Vypsat prvnich 300 znaku odpovedi a delku tokenu pro debug
                tok_len = len(ENTSOE_TOKEN) if ENTSOE_TOKEN else 0
                tok_preview = (ENTSOE_TOKEN[:6] + "..." + ENTSOE_TOKEN[-4:]) if tok_len > 12 else "(empty)"
                print(f"     token_len={tok_len}, token={tok_preview}", flush=True)
                print(f"     response[:300]: {da_xml[:300]}", flush=True)

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

    def _entsoe_afrr(self, qs):
        """ENTSO-E aFRR Capacity Bids - zkousi vice kombinaci parametru pro CZ.
        Vraci JSON vcetne debug info aby slo videt jaka kombinace zafungovala.
        """
        try:
            day_str = qs.get("day", [None])[0]
            if day_str:
                day = datetime.strptime(day_str, "%Y-%m-%d")
            else:
                day = datetime.now(timezone.utc).replace(tzinfo=None)
            ps = day.replace(hour=0, minute=0, second=0, microsecond=0)
            pe = ps + timedelta(hours=23)

            base_period = {
                "periodStart": fmt_entsoe_period(ps),
                "periodEnd":   fmt_entsoe_period(pe),
            }

            import re as _re
            def fetch(extra):
                params = {**base_period, **extra}
                xml, st = call_entsoe(params)
                pts = parse_entsoe_xml(xml, ps) if st == 200 else []
                # Extract reason text z XML pokud je
                reason = ""
                rm = _re.search(r"<text>([^<]+)</text>", xml)
                if rm: reason = rm.group(1)[:120]
                return pts, st, xml, reason

            # Kombinace parametru pro CZ aFRR upward capacity bids
            # Spravne dle ENTSO-E dokumentace (potvrzene NL example):
            # documentType=A81 + businessType=A95(B95) + psrType=A04 + type_MarketAgreement.Type=A01
            #   + processType (A51=aFRR, A52=FCR, A47=mFRR, A46=RR)
            combos = [
                # aFRR (A51) - upward (defaultne) - sirsi
                {"label": "A81+A95+aFRR-A51+A04", "documentType": "A81", "businessType": "A95",
                 "psrType": "A04", "processType": "A51",
                 "controlArea_Domain": CZ_DOMAIN,
                 "type_MarketAgreement.Type": "A01"},
                # B95 (Procured capacity) misto A95
                {"label": "A81+B95+aFRR-A51+A04", "documentType": "A81", "businessType": "B95",
                 "psrType": "A04", "processType": "A51",
                 "controlArea_Domain": CZ_DOMAIN,
                 "type_MarketAgreement.Type": "A01"},
                # mFRR
                {"label": "A81+A95+mFRR-A47+A04", "documentType": "A81", "businessType": "A95",
                 "psrType": "A04", "processType": "A47",
                 "controlArea_Domain": CZ_DOMAIN,
                 "type_MarketAgreement.Type": "A01"},
                # FCR (A52) - sanity check ze CZ ma alespon neco
                {"label": "A81+A95+FCR-A52+A04", "documentType": "A81", "businessType": "A95",
                 "psrType": "A04", "processType": "A52",
                 "controlArea_Domain": CZ_DOMAIN,
                 "type_MarketAgreement.Type": "A01"},
                # Bez processType - vraci vsechny
                {"label": "A81+A95+A04+noProc", "documentType": "A81", "businessType": "A95",
                 "psrType": "A04",
                 "controlArea_Domain": CZ_DOMAIN,
                 "type_MarketAgreement.Type": "A01"},
                # A89 = Contracted reserve prices (jiny doctype, pro ceny)
                {"label": "A89+B95+aFRR+A04", "documentType": "A89", "businessType": "B95",
                 "psrType": "A04", "processType": "A51",
                 "controlArea_Domain": CZ_DOMAIN,
                 "type_MarketAgreement.Type": "A01"},
            ]

            debug = []
            best_pts = []
            best_label = None
            print(f"  -> ENTSO-E aFRR: {ps.strftime('%Y-%m-%d')} - zkousim {len(combos)} kombinaci...", flush=True)
            for c in combos:
                label = c.pop("label")
                pts, st, xml, reason = fetch(c)
                debug.append({
                    "label":  label,
                    "status": st,
                    "points": len(pts),
                    "reason": reason,
                    "xml_preview": xml[:200] if (not pts and st == 200) else None,
                })
                print(f"     [{label}] status={st}, points={len(pts)}, reason={reason!r}", flush=True)
                if pts and not best_pts:
                    best_pts = pts
                    best_label = label

            self._json({
                "day": ps.strftime("%Y-%m-%d"),
                "volumes_up":    best_pts,
                "volumes_down":  [],
                "prices_up":     [],
                "prices_down":   [],
                "best_combo":    best_label,
                "debug":         debug,
            })
        except Exception as e:
            print(f"  -> ENTSO-E aFRR ERROR: {e}", flush=True)
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
            # Client zavrel spojeni drive nez stihl server odpovedet - normalni
            pass

    def handle_one_request(self):
        """Override pro tichy handling BrokenPipe."""
        try:
            super().handle_one_request()
        except (BrokenPipeError, ConnectionResetError):
            pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    # Self-ping pro Render Free - aby sluzba neusla po 15 min necinnosti
    # Render automaticky nastavi RENDER_EXTERNAL_URL na verejnou URL sluzby
    public_url = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")
    if public_url:
        import threading
        def keepalive():
            time.sleep(60)  # pockat nez server nastartuje
            while True:
                try:
                    r = requests.get(f"{public_url}/health", timeout=10)
                    print(f"[keepalive] ping {public_url}/health -> {r.status_code}", flush=True)
                except Exception as e:
                    print(f"[keepalive] FAIL: {e}", flush=True)
                time.sleep(10 * 60)  # ping kazdych 10 minut
        threading.Thread(target=keepalive, daemon=True).start()
        print(f"[keepalive] thread started, pinging {public_url}/health every 10 min", flush=True)
    else:
        print("[keepalive] RENDER_EXTERNAL_URL not set - keepalive disabled", flush=True)
    print(f"CEPS API server -> port {port}", flush=True)
    HTTPServer(("0.0.0.0", port), Handler).serve_forever()
