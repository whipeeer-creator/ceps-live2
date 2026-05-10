"""
CEPS API server pro Hruška graf
"""
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
import json, requests, os, time, io
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, timezone

# Cache pro VDT
_VDT_CACHE = {}

def _json_response(self, data, code=200):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    self.send_response(code)
    self.send_header("Content-Type", "application/json; charset=utf-8")
    self.send_header("Access-Control-Allow-Origin", "*")
    self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
    self.end_headers()
    self.wfile.write(body)

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {self.path} {args[1] if len(args)>1 else ''}")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.end_headers()

    def do_GET(self):
        try:
            parsed = urlparse(self.path)
            qs = parse_qs(parsed.query)

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
            if parsed.path in ("/hruska.html", "/hruska"):
                self._serve_hruska()
                return

            _json_response(self, {"error": "unknown endpoint", "path": parsed.path}, 404)

        except Exception as e:
            import traceback
            traceback.print_exc()
            _json_response(self, {"error": str(e)}, 500)

    # ==================== VDT RANGE - HLAVNÍ PRO GRAF ====================
    def _ote_vdt_range(self, qs):
        try:
            days_back = int(qs.get("days_back", ["7"])[0])
            days_back = max(1, min(30, days_back))
        except:
            days_back = 7

        now = datetime.now()
        all_points = []

        for i in range(days_back + 1):
            target = now - timedelta(days=i)
            day_data = self._fetch_vdt_day(target)
            all_points.extend(day_data.get("data", []))

        _json_response(self, {
            "data": sorted(all_points, key=lambda x: x.get("timestamp", "")),
            "days_back": days_back,
            "fetched_at": now.strftime("%Y-%m-%dT%H:%M:%SZ")
        })

    def _fetch_vdt_day(self, target_date):
        cache_key = target_date.strftime("%Y-%m-%d")
        if cache_key in _VDT_CACHE:
            cached_time, data = _VDT_CACHE[cache_key]
            if (datetime.now() - cached_time).total_seconds() < 1800:  # 30 min cache
                return data

        yyyy = target_date.strftime("%Y")
        mm = target_date.strftime("%m")
        dd = target_date.strftime("%d")
        url = f"https://www.ote-cr.cz/pubweb/attachments/27/{yyyy}/month{mm}/day{dd}/VDT_15MIN_{dd}_{mm}_{yyyy}_CZ.xlsx"

        try:
            r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code != 200:
                result = {"date": cache_key, "data": []}
                _VDT_CACHE[cache_key] = (datetime.now(), result)
                return result

            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(r.content), read_only=True)
            ws = wb.active

            data_points = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) < 5: continue
                interval = str(row[0] or "").strip()
                if not interval or "-" not in interval: continue

                try:
                    start_str = interval.split("-")[0].strip()
                    h, m = map(int, start_str.split(":"))
                    ts = target_date.replace(hour=h if h < 24 else 0, minute=m, second=0, microsecond=0)
                    if h >= 24:
                        ts += timedelta(days=1)

                    last_price = float(str(row[4] or 0).replace(",", "."))

                    data_points.append({
                        "timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"),
                        "last": round(last_price, 2)
                    })
                except:
                    continue

            result = {"date": cache_key, "data": data_points}
            _VDT_CACHE[cache_key] = (datetime.now(), result)
            return result

        except Exception as e:
            print(f"VDT fetch error {cache_key}: {e}")
            result = {"date": cache_key, "data": []}
            _VDT_CACHE[cache_key] = (datetime.now(), result)
            return result

    def _ote_vdt(self, qs):
        date_str = qs.get("date", [""])[0]
        if date_str:
            target = datetime.strptime(date_str, "%Y-%m-%d")
        else:
            target = datetime.now()
        data = self._fetch_vdt_day(target)
        _json_response(self, data)

    def _entsoe_residual_load(self, qs):
        _json_response(self, {"data": []})

    def _smard_residual_load(self, qs):
        _json_response(self, {"data": []})

    def _serve_hruska(self):
        try:
            with open("hruska.html", "r", encoding="utf-8") as f:
                html = f.read()
            # Auto-inject API URL
            inject = """<script>
window.addEventListener('load', function() {
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
        except Exception as e:
            self.send_response(404)
            self.end_headers()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    print(f"✅ Server běží na portu {port}")
    print(f"   Hruška graf: http://localhost:{port}/hruska.html")
    ThreadingHTTPServer(("0.0.0.0", port), Handler).serve_forever()
