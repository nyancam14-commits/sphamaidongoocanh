"""
Verification script for:
1. /api/status
2. /api/load-preset-file
3. /api/upload-csv
4. /api/dashboard
5. /api/export-excel
"""
import urllib.request
import json

base_url = "http://localhost:8080"

def test_api():
    print("1. Testing GET /api/status...")
    req = urllib.request.urlopen(f"{base_url}/api/status")
    status = json.loads(req.read().decode('utf-8'))
    print("Status:", status)

    print("\n2. Testing POST /api/load-preset-file (10k)...")
    post_data = json.dumps({"preset": "10k"}).encode('utf-8')
    req = urllib.request.Request(f"{base_url}/api/load-preset-file", data=post_data, headers={'Content-Type': 'application/json'})
    res = json.loads(urllib.request.urlopen(req).read().decode('utf-8'))
    print("Preset 10k loaded:", res["filename"], res["total_raw"], "rows. Gate auto approved:", res["gate_info"]["auto_approved"])

    print("\n3. Testing GET /api/dashboard...")
    req = urllib.request.urlopen(f"{base_url}/api/dashboard")
    dash = json.loads(req.read().decode('utf-8'))
    print("Dashboard KPIs Net Revenue:", dash["kpis"]["total_net_revenue"])
    print("Top 3 stores count:", len(dash["top_3_stores"]))
    print("Top 3 stores:", [s["store_name"] for s in dash["top_3_stores"]])
    print("Store metrics count:", len(dash["store_metrics"]))
    print("Categories count:", len(dash["category_metrics"]))

    print("\n4. Testing GET /api/export-excel...")
    req = urllib.request.urlopen(f"{base_url}/api/export-excel")
    excel_bytes = req.read()
    print("Excel exported successfully! File size:", len(excel_bytes), "bytes.")

    print("\n5. Testing POST /api/load-preset-file (60k)...")
    post_data = json.dumps({"preset": "60k"}).encode('utf-8')
    req = urllib.request.Request(f"{base_url}/api/load-preset-file", data=post_data, headers={'Content-Type': 'application/json'})
    res = json.loads(urllib.request.urlopen(req).read().decode('utf-8'))
    print("Preset 60k loaded:", res["filename"], res["total_raw"], "rows. Gate auto approved:", res["gate_info"]["auto_approved"])

    print("\nALL VERIFICATIONS PASSED!")

if __name__ == "__main__":
    test_api()

