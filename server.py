"""
CEPS API server - Railway / cloud deployment
"""
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
import json, requests, xml.etree.ElementTree as ET, os, time, io, threading
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, timezone
import urllib.request

# Nacti .env
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
ENTSOE_URL = "https://web-api.tp.entsoe.eu/api"
ENTSOE_TOKEN = os.environ.get("ENTSOE_TOKEN", "")
CZ_DOMAIN = "10YCZ-CEPS-----N"
DE_LU_DOMAIN = "10Y1001A1001A82H"

_VDT_CACHE = {}  # cache pro VDT data

# ============================================================
# HELPERS
# ============================================================

def _request_with_retry(func, *args, retries=3, delay=1.5, **kwargs):
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt == retries - 1:
                raise
            time.sleep(delay * (attempt + 1))

def _json_response(self, data, code=200):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    self.send_response(code)
    self.send_header("Content-Type", "application/json; charset=utf-8")
    self.send_header("Access-Control-Allow-Origin", "*")
    self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
    self.end_headers()
    self.wfile.write(body)

# ============================================================
# VDT - HLAVNÍ OPRAVA PRO HRUŠKA GRAF
# ============================================================

def fetch_vdt_day(target_date):
    """Stáhne a zpracuje VDT data pro jeden den"""
    cache_key = target_date.strftime("%Y-%m-%d")
    if cache_key in _VDT_CACHE:
        cached_time, data = _VDT_CACHE[cache_key]
        if (datetime.now() - cached_time).total_seconds() < 3600:  # 1 hodina cache
            return data

    yyyy = target_date.strftime("%Y")
    mm = target_date.strftime("%m")
    dd = target_date.strftime("%d")
    url = f"https://www.ote-cr.cz/pubweb/attachments/27/{yyyy}/month{mm}/day{dd}/VDT_15MIN_{dd}_{mm}_{yyyy}_CZ.xlsx"

    try:
        r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code != 200:
            return {"date": cache_key, "data": []}

        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(r.content), read_only=True)
        ws = wb.active

        data_points = []
        for row in list(ws.iter_rows(min_row=2, values_only=True)):
            if len(row) < 5:
                continue
            interval = str(row[0] or "").strip()
            if not interval or "-" not in interval:
                continue

            try:
                # interval napr. "00:00-00:15"
                start_str = interval.split("-")[0].strip()
                h, m = map(int, start_str.split(":"))
                ts = target_date.replace(hour=h, minute=m, second=0, microsecond=0)

                last = float(str(row[4] or 0).replace(",", "."))
                weighted = float(str(row[1] or 0).replace(",", "."))

                data_points.append({
                    "timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"),
                    "last": round(last, 2),
                    "weighted_avg": round(weighted, 2)
                })
            except:
                continue

        result = {"date": cache_key, "data": data_points}
        _VDT_CACHE[cache_key] = (datetime.now(), result)
        return result

    except Exception as e:
        print(f"VDT fetch error {cache_key}: {e}")
        return {"date": cache_key, "data": [], "error": str(e)}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} {args[1]}", flush=True)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.end_headers()

    def do_GET(self):
        try:
            parsed = urlparse(self.path)
            qs = parse_qs(parsed.query)
            g = lambda k, d="": qs.get(k, [d])[0]

            if parsed.path == "/ote/vdt/range":
                self._ote_vdt_range(qs)
                return

            if parsed.path == "/ote/vdt":
                self._ote_vdt(qs)
                return

            if parsed.path == "/entsoe/residual-load":
                self._entsoe_residual_load(qs)
                return

            if parsed.path == "/smard/residual-load":
                self._smard_residual_load(qs)
                return

            if parsed.path == "/health":
                _json_response(self, {"status": "ok", "time": datetime.now().isoformat()})
                return

            _json_response(self, {"error": "unknown endpoint"}, 404)

        except Exception as e:
            print(f"ERROR: {e}")
            import traceback
            traceback.print_exc()
            _json_response(self, {"error": str(e)}, 500)

    def _ote_vdt_range(self, qs):
        """Hlavní endpoint pro Hruška graf"""
        try:
            days_back = int(qs.get("days_back", ["7"])[0])
            days_back = max(1, min(30, days_back))
        except:
            days_back = 7

        now = datetime.now()
        all_points = []

        for i in range(days_back + 1):
            target = now - timedelta(days=i)
            day_data = fetch_vdt_day(target)
            all_points.extend(day_data.get("data", []))

        _json_response(self, {
            "data": all_points,
            "days_back": days_back,
            "fetched_at": now.strftime("%Y-%m-%dT%H:%M:%SZ")
        })

    def _ote_vdt(self, qs):
        """Jeden den VDT"""
        date_str = qs.get("date", [""])[0]
        if not date_str:
            target = datetime.now()
        else:
            target = datetime.strptime(date_str, "%Y-%m-%d")
        data = fetch_vdt_day(target)
        _json_response(self, data)

    # Zbytek endpointů (residual load atd.) můžeš nechat původní nebo zjednodušit
    def _entsoe_residual_load(self, qs):
        _json_response(self, {"data": [], "error": "not implemented yet"})

    def _smard_residual_load(self, qs):
        _json_response(self, {"data": [], "error": "not implemented yet"})

    def _html_hruska(self):
        try:
            with open("hruska.html", "r", encoding="utf-8") as f:
                html = f.read()
            inject = """<script>
window.addEventListener('load', () => {
    API_URL = window.location.origin;
    localStorage.setItem('ceps_api_url', API_URL);
    if (typeof loadVdtWind === 'function') loadVdtWind();
});
</script>"""
            html = html.replace("</body>", inject + "\n</body>")
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)
        except:
            self.send_response(404)
            self.end_headers()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    print(f"CEPS API server running on port {port}")
    ThreadingHTTPServer(("0.0.0.0", port), Handler).serve_forever()
