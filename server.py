"""
Ngọc Ánh Data Studio - Multi-Enterprise AI Agent & Analytics Dashboard
Runs on http://localhost:8080
Zero external pip dependencies required (Python 3 standard library).
"""

import os
import io
import re
import csv
import json
import urllib.parse
import urllib.request
from http.server import HTTPServer, BaseHTTPRequestHandler
from agent_core import AuditLogger, DataCleansingAndAuditSkill, RetailDecisionIntelligenceSkill
from data_generator import generate_retail_dataset
from test_suite import run_tests
from excel_exporter import generate_2sheet_excel, generate_3sheet_excel, generate_executive_report_excel
from word_exporter import generate_word_executive_report

# Ngưỡng chuẩn trung bình hệ thống (System Benchmark Baseline)
SYSTEM_AVG_RETURN_RATE = 1.8   # Tỷ lệ đổi trả trung bình chuẩn: 1.8%
SYSTEM_AVG_CANCEL_RATE = 1.0   # Tỷ lệ đơn hủy/lỗi trung bình chuẩn: 1.0%

# Global Application State
STATE = {
    "audit_logger": AuditLogger(),
    "raw_dataset": None,
    "cleansed_result": None,
    "human_approved": False,
    "human_verified_clean": False,
    "human_clean_notes": "",
    "human_evaluated_analysis": False,
    "human_eval_notes": "",
    "business_model": "",
    "business_problem": "",
    "clarifications": {},
    "decision_result": None,
    "pipeline_stage": "INIT",
    "gate_info": None
}

def get_default_masters():
    """Khởi tạo bộ danh mục rỗng để tự động thích ứng 100% theo dữ liệu thực tế của từng doanh nghiệp"""
    return {
        "stores": {},
        "products": {},
        "targets": {},
        "period_config": None
    }

def _normalize_date_str(raw_d):
    s = str(raw_d or "").strip()
    if not s:
        return "2026-10-15"
    # Handle YYYY-MM-DD or YYYY/MM/DD
    m1 = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m1:
        return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
    # Handle DD/MM/YYYY or DD-MM-YYYY
    m2 = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})", s)
    if m2:
        return f"{m2.group(3)}-{int(m2.group(2)):02d}-{int(m2.group(1)):02d}"
    return s[:10]

def ensure_masters_for_transactions(txs, masters):
    """Tự động xây dựng danh mục đơn vị, sản phẩm/danh mục, mục tiêu và kỳ báo cáo từ dữ liệu thực tế"""
    stores = dict(masters.get("stores") or {})
    products = dict(masters.get("products") or {})
    targets = dict(masters.get("targets") or {})
    period_config = masters.get("period_config")

    store_rev_sum = {}
    dates = []
    for tx in txs:
        sid = tx.get("store_id")
        pid = tx.get("product_id")
        d = tx.get("date")
        if d and re.match(r"^\d{4}-\d{2}-\d{2}$", str(d)):
            dates.append(str(d))

        rev = abs(float(tx.get("unit_price", 0) or 0) * int(tx.get("qty", 1) or 1))
        if sid:
            store_rev_sum[sid] = store_rev_sum.get(sid, 0.0) + rev
            if sid not in stores:
                sname = tx.get("_store_name") or sid
                stores[sid] = {
                    "store_id": sid,
                    "name": sname,
                    "region": tx.get("_region") or "Toàn hệ thống",
                    "manager": tx.get("_manager") or f"Phụ trách {sid}"
                }

        if pid and pid not in products:
            price = float(tx.get("unit_price", 0) or 0)
            cogs = tx.get("_cogs")
            if cogs is None:
                cogs = round(price * 0.58, 2)
            products[pid] = {
                "name": tx.get("_product_name") or str(pid)[:40],
                "category": tx.get("_category") or str(pid)[:30] or "Danh mục chung",
                "price": price,
                "cogs": float(cogs)
            }

    for sid, actual_rev in store_rev_sum.items():
        if sid not in targets:
            base_rev = max(actual_rev, 1000000.0)
            # Tạo ngưỡng kế hoạch tham chiếu sát với quy mô thực tế của từng đơn vị
            factor = 1.05 if (sum(ord(c) for c in str(sid)) % 3 != 0) else 1.18
            rep_rev_tgt = round(base_rev * factor, 2)
            rep_prof_tgt = round(rep_rev_tgt * 0.42, 2)
            targets[sid] = {
                "report_rev_target": rep_rev_tgt,
                "report_profit_target": rep_prof_tgt,
                "compare_rev_target": round(rep_rev_tgt * 0.95, 2),
                "compare_profit_target": round(rep_prof_tgt * 0.95, 2)
            }

    if not period_config:
        min_d = min(dates) if dates else "2000-01-01"
        max_d = max(dates) if dates else "2099-12-31"
        period_config = {
            "report_start": min_d,
            "report_end": max_d,
            "compare_start": "",
            "compare_end": ""
        }

    return stores, products, targets, period_config

def reset_human_gates():
    """Đặt lại trạng thái các cổng kiểm duyệt của con người khi nạp bộ dữ liệu mới"""
    STATE["cleansed_result"] = None
    STATE["human_approved"] = False
    STATE["human_verified_clean"] = False
    STATE["human_clean_notes"] = ""
    STATE["human_evaluated_analysis"] = False
    STATE["human_eval_notes"] = ""
    STATE["decision_result"] = None
    STATE["gate_info"] = None

def init_benchmark(total_records=60000, seed=42, filename="Ngoc_Anh_Enterprise_Standard_60K.csv"):
    STATE["audit_logger"].clear()
    STATE["audit_logger"].log("Hệ thống", "SystemInit", "Khởi tạo dữ liệu mẫu", 
                              f"Nạp {total_records:,} bản ghi chuẩn hóa của Ngọc Ánh Data Studio từ {filename}", total_records)
    data = generate_retail_dataset(total_records=total_records, seed=seed)
    data["filename"] = filename
    STATE["raw_dataset"] = data
    reset_human_gates()
    STATE["pipeline_stage"] = "DATA_LOADED"
    return {"status": "success", "total_raw": len(data["transactions"]), "stores_count": len(data["stores"]), "filename": filename}

def _parse_num(val, default=0.0):
    if val is None or val == "":
        return default
    s = re.sub(r"[^\d.\-]", "", str(val).replace(",", ""))
    try:
        return float(s) if s else default
    except ValueError:
        return default

def parse_csv_to_transactions(csv_content):
    """Phân tích nội dung CSV thành danh sách dictionary giao dịch chuẩn (linh hoạt cho mọi mô hình doanh nghiệp)"""
    reader = csv.DictReader(io.StringIO(csv_content))
    txs = []
    for r in reader:
        cleaned_row = {str(k).strip().lower(): str(v).strip() for k, v in r.items() if k is not None}
        if not any(cleaned_row.values()):
            continue

        # Nhận diện mã giao dịch / hồ sơ / RFQ
        tid = (cleaned_row.get("trans_id") or cleaned_row.get("rfq_id") or 
               cleaned_row.get("mã gd") or cleaned_row.get("mã giao dịch") or
               cleaned_row.get("mã hồ sơ") or cleaned_row.get("hồ sơ") or
               cleaned_row.get("stt") or cleaned_row.get("id") or f"TX{1000000+len(txs)}")
               
        oid = (cleaned_row.get("order_id") or cleaned_row.get("mã đơn") or 
               cleaned_row.get("order") or cleaned_row.get("nguoi_lien_he") or 
               cleaned_row.get("phụ trách") or cleaned_row.get("khách hàng") or f"ORD{500000+len(txs)}")
               
        raw_date = (cleaned_row.get("date") or cleaned_row.get("ngay_yeu_cau") or 
                    cleaned_row.get("ngày gd") or cleaned_row.get("ngày") or
                    cleaned_row.get("kỳ") or cleaned_row.get("tháng") or "2026-10-15")
        date_str = _normalize_date_str(raw_date)
                    
        # Nhận diện Đơn vị / Chi nhánh / Kênh / Nguồn / Khu vực
        sid = (cleaned_row.get("store_id") or cleaned_row.get("mã ch") or 
               cleaned_row.get("kênh / nguồn") or cleaned_row.get("kênh") or 
               cleaned_row.get("nguồn") or cleaned_row.get("phân nhóm") or 
               cleaned_row.get("đơn vị") or cleaned_row.get("cửa hàng") or
               cleaned_row.get("chi nhánh") or cleaned_row.get("khu vực") or
               cleaned_row.get("ten_khach_hang") or cleaned_row.get("dia_diem_trien_khai") or "Đơn vị 01")
               
        pid = (cleaned_row.get("product_id") or cleaned_row.get("hang_muc_yeu_cau") or 
               cleaned_row.get("mã sp") or cleaned_row.get("sản phẩm") or
               cleaned_row.get("dịch vụ") or cleaned_row.get("danh mục") or
               cleaned_row.get("ngành hàng") or cleaned_row.get("sku") or "SP01")

        store_name = (cleaned_row.get("store_name") or cleaned_row.get("tên cửa hàng") or
                      cleaned_row.get("tên đơn vị") or cleaned_row.get("tên chi nhánh") or "")
        cat_name = (cleaned_row.get("category") or cleaned_row.get("danh mục") or
                    cleaned_row.get("ngành hàng") or cleaned_row.get("nhóm sản phẩm") or "")
        prod_name = (cleaned_row.get("product_name") or cleaned_row.get("tên sản phẩm") or
                     cleaned_row.get("tên dịch vụ") or "")
        region = (cleaned_row.get("region") or cleaned_row.get("khu vực") or cleaned_row.get("vùng") or "")

        qty = max(1, int(abs(_parse_num(cleaned_row.get("qty") or cleaned_row.get("số lượng"), 1))))
        unit_price = _parse_num(
            cleaned_row.get("unit_price") or cleaned_row.get("giá trị") or 
            cleaned_row.get("doanh thu") or cleaned_row.get("ngan_sach") or 
            cleaned_row.get("đơn giá") or cleaned_row.get("giá") or cleaned_row.get("amount"),
            350000.0
        )

        raw_cogs = cleaned_row.get("cogs") or cleaned_row.get("giá vốn") or cleaned_row.get("chi phí") or cleaned_row.get("cost")
        cogs_val = _parse_num(raw_cogs, None) if raw_cogs not in (None, "") else None
            
        raw_status = cleaned_row.get("status") or cleaned_row.get("trạng thái") or "COMPLETED"
        status = raw_status.upper()
        if any(w in raw_status.lower() for w in ["hủy", "cancel", "void", "failed"]):
            status = "CANCELLED"
        elif any(w in raw_status.lower() for w in ["hoàn thành", "thành công", "done", "success"]):
            status = "COMPLETED"
            
        raw_type = cleaned_row.get("type") or cleaned_row.get("loại") or ""
        if any(w in raw_type.lower() for w in ["trả", "return", "hoàn"]) or "hoàn" in raw_status.lower():
            ttype = "RETURN"
        else:
            ttype = "SALE"
            
        ref_order = cleaned_row.get("ref_order_id") or cleaned_row.get("đơn gốc") or ""

        tx = {
            "trans_id": tid,
            "order_id": oid,
            "date": date_str,
            "store_id": sid,
            "product_id": pid,
            "qty": qty,
            "unit_price": unit_price,
            "status": status,
            "type": ttype,
            "ref_order_id": ref_order
        }
        if store_name: tx["_store_name"] = store_name
        if cat_name: tx["_category"] = cat_name
        if prod_name: tx["_product_name"] = prod_name
        if region: tx["_region"] = region
        if cogs_val is not None: tx["_cogs"] = cogs_val
        txs.append(tx)
    return txs

def fetch_google_sheets_csv(url):
    """Chuyển đổi URL Google Sheets sang CSV Export và tải về nội dung"""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("URL Google Sheets không hợp lệ. Vui lòng dán link có dạng: https://docs.google.com/spreadsheets/d/<ID>/edit")
    sheet_id = match.group(1)
    gid_match = re.search(r"[#&]gid=([0-9]+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    
    req = urllib.request.Request(
        export_url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    )
    with urllib.request.urlopen(req, timeout=10) as response:
        return response.read().decode("utf-8")

def run_skill_1_internal():
    """Chạy làm sạch dữ liệu và tính toán điểm trung bình đổi trả / đơn hủy để xác định tự duyệt hay chờ con người"""
    if not STATE["raw_dataset"]:
        return None
        
    skill1 = DataCleansingAndAuditSkill(STATE["audit_logger"])
    data = STATE["raw_dataset"]
    cleansed = skill1.execute(
        data["transactions"],
        data["stores"],
        data["products"],
        data["period_config"]
    )
    STATE["cleansed_result"] = cleansed
    
    total_raw = len(data["transactions"])
    valid_report = len(cleansed["cleaned_report"])
    valid_compare = len(cleansed.get("cleaned_compare", []))
    total_cleaned_valid = valid_report + valid_compare
    
    return_txs = [t for t in cleansed["cleaned_report"] if t["type"] == "RETURN"]
    sale_txs = [t for t in cleansed["cleaned_report"] if t["type"] == "SALE"]
    total_returns_count = len(return_txs)
    total_sales_count = len(sale_txs)
    
    actual_return_rate = round((total_returns_count / total_sales_count * 100.0) if total_sales_count > 0 else 0.0, 2)
    cancelled_count = sum(1 for u in cleansed["unqualified"] if u["reason_code"] == "ERR_ORDER_CANCELLED")
    actual_cancel_rate = round((cancelled_count / total_raw * 100.0) if total_raw > 0 else 0.0, 2)
    unqualified_count = len(cleansed["unqualified"])
    unqualified_pct = round(unqualified_count / total_raw * 100.0, 2) if total_raw > 0 else 0.0
    
    is_return_safe = actual_return_rate <= SYSTEM_AVG_RETURN_RATE
    is_cancel_safe = actual_cancel_rate <= SYSTEM_AVG_CANCEL_RATE
    auto_approved = is_return_safe and is_cancel_safe
    clean_pct = round((total_cleaned_valid / total_raw * 100.0) if total_raw > 0 else 0.0, 2)
    
    gate_info = {
        "total_raw": total_raw,
        "valid_count": total_cleaned_valid,
        "total_cleaned_valid": total_cleaned_valid,
        "clean_pct": clean_pct,
        "valid_report_period": valid_report,
        "valid_compare_period": valid_compare,
        "unqualified_count": unqualified_count,
        "unqualified_pct": unqualified_pct,
        "total_returns": total_returns_count,
        "actual_return_rate": actual_return_rate,
        "system_avg_return_rate": SYSTEM_AVG_RETURN_RATE,
        "cancelled_count": cancelled_count,
        "actual_cancel_rate": actual_cancel_rate,
        "system_avg_cancel_rate": SYSTEM_AVG_CANCEL_RATE,
        "auto_approved": False,
        "risk_ok": auto_approved,
        "status": "WAITING_HUMAN_VERIFICATION",
        "message": (
            f"Đã làm sạch {total_cleaned_valid:,}/{total_raw:,} bản ghi ({clean_pct}%). "
            f"Tỷ lệ đổi trả ({actual_return_rate}%) & đơn hủy ({actual_cancel_rate}%) ở mức an toàn. "
            f"Vui lòng chuyên viên BI rà soát bảng dữ liệu và bấm Xác nhận tại Cổng Kiểm Duyệt 1."
            if auto_approved else
            f"⚠️ Cảnh báo: Tỷ lệ đổi trả ({actual_return_rate}%) hoặc đơn hủy ({actual_cancel_rate}%) "
            f"vượt ngưỡng chuẩn ({SYSTEM_AVG_RETURN_RATE}% & {SYSTEM_AVG_CANCEL_RATE}%). "
            f"Chuyên viên BI cần kiểm tra bảng ngoại lệ và bấm Xác nhận tại Cổng Kiểm Duyệt 1."
        )
    }
    STATE["gate_info"] = gate_info
    STATE["human_approved"] = False
    STATE["human_verified_clean"] = False
    STATE["human_evaluated_analysis"] = False
    STATE["pipeline_stage"] = "WAITING_HUMAN_VERIFICATION"
    STATE["audit_logger"].log("Giai đoạn 1", "HumanApprovalGate", "Chờ chuyên viên BI kiểm duyệt dữ liệu sạch", gate_info["message"])
    run_skill_2_internal()
    STATE["pipeline_stage"] = "WAITING_HUMAN_VERIFICATION"
        
    return {
        "cleansed": cleansed,
        "gate_info": gate_info
    }

def run_skill_2_internal():
    """Chạy tính toán chỉ số quản trị và xếp hạng Top các đơn vị / cửa hàng ưu tiên"""
    if not STATE["cleansed_result"]:
        return None
        
    skill2 = RetailDecisionIntelligenceSkill(STATE["audit_logger"])
    data = STATE["raw_dataset"]
    cleansed = STATE["cleansed_result"]

    decision = skill2.execute(
        cleansed["cleaned_report"],
        cleansed["cleaned_compare"],
        data["stores"],
        data["targets"],
        data["period_config"]
    )
    STATE["decision_result"] = decision
    STATE["pipeline_stage"] = "REPORTS_READY"
    return decision

def parse_json_body(body):
    if not body:
        return {}
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252", "cp1258"]:
        try:
            return json.loads(body.decode(enc))
        except Exception:
            continue
    return {}

def ingest_transactions_to_state(parsed_txs, filename, action_label, details):
    """Nạp danh sách giao dịch chuẩn hóa vào STATE, đặt lại cổng duyệt và chạy làm sạch sơ bộ"""
    masters = get_default_masters()
    stores, prods, targets, period_config = ensure_masters_for_transactions(parsed_txs, masters)
    STATE["audit_logger"].clear()
    STATE["audit_logger"].log("Hệ thống", "DataIngestion", action_label, details, len(parsed_txs))
    STATE["raw_dataset"] = {
        "transactions": parsed_txs,
        "stores": stores,
        "products": prods,
        "targets": targets,
        "period_config": period_config,
        "filename": filename
    }
    reset_human_gates()
    STATE["pipeline_stage"] = "DATA_LOADED"
    cleansed_res = run_skill_1_internal()
    return {
        "status": "success",
        "filename": filename,
        "total_raw": len(parsed_txs),
        "summary": cleansed_res["cleansed"]["summary"] if cleansed_res else {},
        "gate_info": cleansed_res["gate_info"] if cleansed_res else {}
    }

class RequestHandler(BaseHTTPRequestHandler):
    def _send_json(self, data, status_code=200):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def _send_html(self, html_content):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html_content.encode("utf-8"))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        params = urllib.parse.parse_qs(parsed.query)

        if path == "/" or path == "/index.html":
            with open("index.html", "r", encoding="utf-8") as f:
                self._send_html(f.read())
            return

        if path == "/api/status":
            raw_count = len(STATE["raw_dataset"]["transactions"]) if STATE["raw_dataset"] else 0
            cleansed_count = len(STATE["cleansed_result"]["cleaned_report"]) if STATE["cleansed_result"] else 0
            unqual_count = len(STATE["cleansed_result"]["unqualified"]) if STATE["cleansed_result"] else 0
            filename = STATE["raw_dataset"].get("filename", "Chưa nạp tệp") if STATE["raw_dataset"] else "Chưa nạp tệp"
            self._send_json({
                "pipeline_stage": STATE["pipeline_stage"],
                "human_approved": STATE["human_approved"],
                "human_verified_clean": STATE["human_verified_clean"],
                "human_clean_notes": STATE["human_clean_notes"],
                "human_evaluated_analysis": STATE["human_evaluated_analysis"],
                "human_eval_notes": STATE["human_eval_notes"],
                "business_model": STATE["business_model"],
                "business_problem": STATE["business_problem"],
                "clarifications": STATE["clarifications"],
                "raw_count": raw_count,
                "cleansed_count": cleansed_count,
                "unqualified_count": unqual_count,
                "filename": filename,
                "gate_info": STATE["gate_info"],
                "has_reports": STATE["decision_result"] is not None
            })
            return

        if path == "/api/logs":
            self._send_json({"logs": STATE["audit_logger"].get_logs()})
            return

        if path == "/api/unqualified":
            if not STATE["cleansed_result"]:
                self._send_json({"unqualified": [], "cleaned_sample": [], "gate_info": STATE["gate_info"]})
                return
            unqual = STATE["cleansed_result"]["unqualified"]
            cleaned_rep = STATE["cleansed_result"].get("cleaned_report", [])
            limit = int(params.get("limit", [150])[0])
            self._send_json({
                "total": len(unqual),
                "summary": STATE["cleansed_result"]["summary"],
                "gate_info": STATE["gate_info"],
                "records": unqual[:limit],
                "cleaned_sample": cleaned_rep[:80]
            })
            return

        if path == "/api/download-sample-file":
            ftype = params.get("file", ["new_60k"])[0]
            if ftype == "new_60k":
                fname = "Du_Lieu_Test_Ky_Moi_AI301_60K.csv"
            elif ftype == "25k":
                fname = "Du_Lieu_Test_Ranh_Gioi_25K.csv"
            elif ftype == "60k":
                fname = "Du_Lieu_Moi_Thang_11_PhucAn_60K.csv"
            else:
                fname = "Du_Lieu_Moi_Test_10K.csv"

            if not os.path.exists(fname):
                init_benchmark(total_records=10000, filename=fname)
            with open(fname, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        # 1. Xuất file Excel 2 Sheets: Du_Lieu_Da_Lam_Sach & Bao_Cao_Tong_Quan
        if path == "/api/export-excel":
            if not STATE["decision_result"]:
                if STATE["cleansed_result"]:
                    run_skill_2_internal()
                else:
                    self.send_error(400, "Vui lòng nạp tệp và làm sạch dữ liệu trước khi xuất Excel.")
                    return
            
            cleaned_report = STATE["cleansed_result"]["cleaned_report"] if STATE["cleansed_result"] else []
            store_metrics = STATE["decision_result"]["store_metrics"] if STATE["decision_result"] else []
            category_metrics = STATE["decision_result"]["category_metrics"] if STATE["decision_result"] else {}
            
            total_net_rev = sum(m["net_revenue"] for m in store_metrics)
            total_net_profit = sum(m["net_profit"] for m in store_metrics)
            total_gross_rev = sum(m["gross_revenue"] for m in store_metrics)
            total_ret_rev = sum(m["return_revenue"] for m in store_metrics)
            total_rev_target = sum(m["rev_target"] for m in store_metrics)
            total_profit_target = sum(m["profit_target"] for m in store_metrics)

            sys_kpis = {
                "total_net_revenue": total_net_rev,
                "total_gross_profit": total_net_profit,
                "avg_margin_pct": round((total_net_profit / total_net_rev * 100.0) if total_net_rev > 0 else 0.0, 2),
                "sys_return_rate_pct": round((total_ret_rev / total_gross_rev * 100.0) if total_gross_rev > 0 else 0.0, 2),
                "rev_target_completion_pct": round((total_net_rev / total_rev_target * 100.0) if total_rev_target > 0 else 0.0, 2),
                "profit_target_completion_pct": round((total_net_profit / total_profit_target * 100.0) if total_profit_target > 0 else 0.0, 2)
            }

            xlsx_bytes = generate_2sheet_excel(cleaned_report, store_metrics, category_metrics, sys_kpis)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="Ngoc_Anh_Bao_Cao_2Sheets.xlsx"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.end_headers()
            self.wfile.write(xlsx_bytes)
            return

        # 2. Xuất Báo Cáo Word (.docx) chuyên sâu chuẩn theo mẫu Ban Giám Đốc / Nhà Nước
        if path == "/api/export-word":
            if not STATE["decision_result"]:
                if STATE["cleansed_result"]:
                    run_skill_2_internal()
                else:
                    self.send_error(400, "Vui lòng nạp tệp và làm sạch dữ liệu trước khi xuất Báo cáo.")
                    return

            store_metrics = STATE["decision_result"]["store_metrics"] if STATE["decision_result"] else []
            category_metrics = STATE["decision_result"]["category_metrics"] if STATE["decision_result"] else {}
            gate_info = STATE["gate_info"] or {}
            summary = STATE["cleansed_result"]["summary"] if STATE["cleansed_result"] else {}
            dataset_name = STATE["raw_dataset"].get("filename", "Doanh nghiệp mẫu") if STATE["raw_dataset"] else "Doanh nghiệp"

            total_net_rev = sum(m["net_revenue"] for m in store_metrics)
            total_net_profit = sum(m["net_profit"] for m in store_metrics)
            total_gross_rev = sum(m["gross_revenue"] for m in store_metrics)
            total_ret_rev = sum(m["return_revenue"] for m in store_metrics)
            total_rev_target = sum(m["rev_target"] for m in store_metrics)

            sys_kpis = {
                "total_net_revenue": total_net_rev,
                "total_gross_profit": total_net_profit,
                "avg_margin_pct": round((total_net_profit / total_net_rev * 100.0) if total_net_rev > 0 else 0.0, 2),
                "sys_return_rate_pct": round((total_ret_rev / total_gross_rev * 100.0) if total_gross_rev > 0 else 0.0, 2),
                "rev_target_completion_pct": round((total_net_rev / total_rev_target * 100.0) if total_rev_target > 0 else 0.0, 2)
            }

            docx_bytes = generate_word_executive_report(
                dataset_name, sys_kpis, store_metrics, category_metrics, gate_info, summary,
                business_model=STATE.get("business_model", ""),
                business_problem=STATE.get("business_problem", ""),
                clarifications=STATE.get("clarifications", {})
            )
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", 'attachment; filename="Bao_Cao_Phan_Tich_Chuyen_Sau_Ngoc_Anh_Studio.docx"')
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)
            return

        # 2. Xuất file Báo cáo Kết quả Kinh doanh theo mẫu Dashboard MasterMOS
        if path == "/api/export-executive-report":
            if not STATE["decision_result"]:
                if STATE["cleansed_result"]:
                    run_skill_2_internal()
                else:
                    self.send_error(400, "Vui lòng nạp tệp và làm sạch dữ liệu trước khi xuất Báo cáo.")
                    return
            
            top_3 = STATE["decision_result"]["top_3_stores"] if STATE["decision_result"] else []
            store_metrics = STATE["decision_result"]["store_metrics"] if STATE["decision_result"] else []
            category_metrics = STATE["decision_result"]["category_metrics"] if STATE["decision_result"] else {}
            
            total_net_rev = sum(m["net_revenue"] for m in store_metrics)
            total_net_profit = sum(m["net_profit"] for m in store_metrics)
            total_gross_rev = sum(m["gross_revenue"] for m in store_metrics)
            total_ret_rev = sum(m["return_revenue"] for m in store_metrics)
            total_rev_target = sum(m["rev_target"] for m in store_metrics)
            
            sys_kpis = {
                "total_net_revenue": total_net_rev,
                "total_gross_profit": total_net_profit,
                "avg_margin_pct": round((total_net_profit / total_net_rev * 100.0) if total_net_rev > 0 else 0.0, 2),
                "sys_return_rate_pct": round((total_ret_rev / total_gross_rev * 100.0) if total_gross_rev > 0 else 0.0, 2),
                "rev_target_completion_pct": round((total_net_rev / total_rev_target * 100.0) if total_rev_target > 0 else 0.0, 2)
            }

            xlsx_bytes = generate_executive_report_excel(sys_kpis, store_metrics, category_metrics, top_3)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="Bao_Cao_Ket_Qua_Kinh_Doanh_Doanh_Nghiep.xlsx"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.end_headers()
            self.wfile.write(xlsx_bytes)
            return

        if path == "/api/dashboard":
            if not STATE["decision_result"]:
                if STATE["cleansed_result"]:
                    run_skill_2_internal()
                else:
                    self._send_json({"error": "Chưa hoàn tất giai đoạn làm sạch dữ liệu."}, 400)
                    return
                    
            res = STATE["decision_result"]
            metrics = res["store_metrics"]
            
            total_net_rev = sum(m["net_revenue"] for m in metrics)
            total_net_profit = sum(m["net_profit"] for m in metrics)
            total_gross_rev = sum(m["gross_revenue"] for m in metrics)
            total_ret_rev = sum(m["return_revenue"] for m in metrics)
            total_rev_target = sum(m["rev_target"] for m in metrics)
            total_profit_target = sum(m["profit_target"] for m in metrics)

            sys_kpis = {
                "total_net_revenue": round(total_net_rev, 2),
                "total_gross_profit": round(total_net_profit, 2),
                "total_cost": round(total_net_rev - total_net_profit, 2),
                "avg_margin_pct": round((total_net_profit / total_net_rev * 100.0) if total_net_rev > 0 else 0.0, 2),
                "sys_return_rate_pct": round((total_ret_rev / total_gross_rev * 100.0) if total_gross_rev > 0 else 0.0, 2),
                "rev_target_completion_pct": round((total_net_rev / total_rev_target * 100.0) if total_rev_target > 0 else 0.0, 2),
                "profit_target_completion_pct": round((total_net_profit / total_profit_target * 100.0) if total_profit_target > 0 else 0.0, 2)
            }

            summary_data = STATE["cleansed_result"]["summary"] if STATE["cleansed_result"] else {}
            self._send_json({
                "kpis": sys_kpis,
                "store_metrics": metrics,
                "category_metrics": res["category_metrics"],
                "top_3_stores": res["top_3_stores"],
                "period_config": res["period_config"],
                "gate_info": STATE["gate_info"],
                "summary": summary_data,
                "business_model": STATE["business_model"],
                "business_problem": STATE["business_problem"],
                "clarifications": STATE["clarifications"],
                "human_verified_clean": STATE["human_verified_clean"],
                "human_clean_notes": STATE["human_clean_notes"],
                "human_evaluated_analysis": STATE["human_evaluated_analysis"],
                "human_eval_notes": STATE["human_eval_notes"]
            })
            return

        if path == "/api/run-tests":
            test_results = run_tests()
            self._send_json({"results": test_results, "all_passed": all(t["passed"] for t in test_results)})
            return

        if path == "/api/store-drilldown":
            sid = params.get("store_id", [None])[0]
            if not sid:
                self._send_json({"error": "Store ID không hợp lệ."}, 400)
                return
            if not STATE["decision_result"] and STATE["cleansed_result"]:
                run_skill_2_internal()
            if not STATE["decision_result"]:
                self._send_json({"error": "Chưa hoàn tất tính toán."}, 400)
                return
            
            store_row = next((m for m in STATE["decision_result"]["store_metrics"] if m["store_id"] == sid), None)
            if not store_row:
                self._send_json({"error": "Không tìm thấy đơn vị / cửa hàng."}, 404)
                return

            cleaned_txs = [tx for tx in STATE["cleansed_result"]["cleaned_report"] if tx["store_id"] == sid][:25]
            self._send_json({
                "store": store_row,
                "sample_transactions": cleaned_txs,
                "total_tx_count": store_row["tx_count"]
            })
            return

        self.send_error(404, "Endpoint not found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length) if content_length > 0 else b""

        # 1. Nạp từ link Google Sheets
        if path == "/api/load-sheets-url":
            payload = parse_json_body(body)
            url = payload.get("url", "").strip()
            if not url:
                self._send_json({"error": "Vui lòng dán link Google Sheets."}, 400)
                return
            
            try:
                csv_content = fetch_google_sheets_csv(url)
                parsed_txs = parse_csv_to_transactions(csv_content)
                filename = "Google_Sheets_Live_Data.csv"
            except Exception as e:
                self._send_json({
                    "error": f"Không thể kết nối đến Google Sheets (vui lòng kiểm tra quyền chia sẻ công khai 'Bất kỳ ai có liên kết'): {str(e)}"
                }, 400)
                return

            res = ingest_transactions_to_state(
                parsed_txs, filename, "Tiếp nhận dữ liệu từ Google Sheets",
                f"Đã đọc {len(parsed_txs):,} giao dịch từ liên kết"
            )
            self._send_json(res)
            return

        # 2. Tải lên tệp CSV/Excel từ máy tính
        if path == "/api/upload-csv":
            payload = parse_json_body(body)
            filename = payload.get("filename", "Uploaded_Dataset.csv")
            csv_text = payload.get("csv_text", "")
            
            if not csv_text:
                try:
                    csv_text = body.decode("utf-8")
                except Exception:
                    csv_text = body.decode("latin-1")
                    
            if not csv_text.strip():
                self._send_json({"error": "Nội dung tệp trống."}, 400)
                return
                
            parsed_txs = parse_csv_to_transactions(csv_text)
            if not parsed_txs:
                self._send_json({"error": "Không thể phân tích dữ liệu giao dịch từ tệp CSV đã chọn."}, 400)
                return

            res = ingest_transactions_to_state(
                parsed_txs, filename, "Tiếp nhận tệp dữ liệu người dùng tải lên",
                f"Đã đọc {len(parsed_txs):,} giao dịch từ tệp {filename}"
            )
            self._send_json(res)
            return

        # 3. Nạp một trong các tệp có sẵn trên máy chủ
        if path == "/api/load-preset-file":
            payload = parse_json_body(body)
            preset_name = payload.get("preset", "new_60k")
            
            if preset_name == "new_60k":
                fname = "Du_Lieu_Test_Ky_Moi_AI301_60K.csv"
            elif preset_name == "25k":
                fname = "Du_Lieu_Test_Ranh_Gioi_25K.csv"
            elif preset_name == "60k":
                fname = "Du_Lieu_Moi_Thang_11_PhucAn_60K.csv"
            elif preset_name == "std_60k":
                res = init_benchmark(60000, seed=42, filename="Ngoc_Anh_Enterprise_Standard_60K.csv")
                cleansed_res = run_skill_1_internal()
                res["summary"] = cleansed_res["cleansed"]["summary"] if cleansed_res else {}
                res["gate_info"] = cleansed_res["gate_info"] if cleansed_res else {}
                self._send_json(res)
                return
            else:
                fname = "Du_Lieu_Moi_Test_10K.csv"
                
            if not os.path.exists(fname):
                init_benchmark(total_records=10000, filename=fname)
                    
            with open(fname, "r", encoding="utf-8") as f:
                csv_content = f.read()
                
            parsed_txs = parse_csv_to_transactions(csv_content)
            res = ingest_transactions_to_state(
                parsed_txs, fname, f"Nạp tệp dữ liệu kiểm thử: {fname}",
                f"Đã đọc {len(parsed_txs):,} giao dịch"
            )
            self._send_json(res)
            return

        if path == "/api/load-benchmark":
            res = init_benchmark(60000, seed=42, filename="Ngoc_Anh_Enterprise_Standard_60K.csv")
            cleansed_res = run_skill_1_internal()
            res["summary"] = cleansed_res["cleansed"]["summary"] if cleansed_res else {}
            res["gate_info"] = cleansed_res["gate_info"] if cleansed_res else {}
            self._send_json(res)
            return

        if path == "/api/run-skill-1":
            if not STATE["raw_dataset"]:
                self._send_json({"error": "Vui lòng nạp dữ liệu thô trước."}, 400)
                return
            cleansed_res = run_skill_1_internal()
            self._send_json({
                "status": "success",
                "total_raw": cleansed_res["gate_info"]["total_raw"] if cleansed_res else 0,
                "summary": cleansed_res["cleansed"]["summary"],
                "gate_info": cleansed_res["gate_info"],
                "unqualified_count": len(cleansed_res["cleansed"]["unqualified"])
            })
            return

        # Cổng Kiểm Duyệt 1: Kiểm duyệt dữ liệu sạch (Con người tự làm)
        if path == "/api/human-verify":
            payload = parse_json_body(body)
            decision = payload.get("decision", "APPROVE")
            notes = str(payload.get("notes") or "Chuyên viên BI đã kiểm tra bảng ngoại lệ và xác nhận dữ liệu hợp lệ.").strip()
            
            if decision == "APPROVE":
                STATE["human_approved"] = True
                STATE["human_verified_clean"] = True
                STATE["human_clean_notes"] = notes
                STATE["pipeline_stage"] = "HUMAN_VERIFIED"
                STATE["audit_logger"].log("Giai đoạn 1", "HumanVerificationGate", "Chuyên viên BI phê duyệt dữ liệu sạch (Cổng 1)", 
                                          f"Quyết định: CÓ HỢP LỆ | Ghi chú: {notes}")
                decision_res = run_skill_2_internal()
                self._send_json({
                    "status": "approved",
                    "human_verified_clean": True,
                    "human_clean_notes": notes,
                    "message": "Chuyên viên BI đã xác nhận dữ liệu sạch hợp lệ. Đang chuyển sang Phân tích dữ liệu!",
                    "has_reports": decision_res is not None
                })
            else:
                STATE["human_approved"] = False
                STATE["human_verified_clean"] = False
                STATE["human_evaluated_analysis"] = False
                STATE["human_clean_notes"] = notes
                STATE["pipeline_stage"] = "CLEANSED"
                STATE["audit_logger"].log("Giai đoạn 1", "HumanVerificationGate", "Chuyên viên BI yêu cầu làm sạch lại (Cổng 1)", 
                                          f"Quyết định: KHÔNG HỢP LỆ | Yêu cầu: {notes}")
                self._send_json({
                    "status": "rejected",
                    "human_verified_clean": False,
                    "human_clean_notes": notes,
                    "message": "Đã ghi nhận yêu cầu rà soát & quay lại bước làm sạch dữ liệu."
                })
            return

        # Cổng Kiểm Duyệt 2: Đánh giá kết quả phân tích & Insight (Con người tự làm)
        if path == "/api/human-evaluate":
            payload = parse_json_body(body)
            decision = payload.get("decision", "APPROVE")
            notes = str(payload.get("notes") or "Chuyên viên BI đánh giá kết quả phân tích & Insight đạt yêu cầu.").strip()

            if decision == "APPROVE":
                STATE["human_evaluated_analysis"] = True
                STATE["human_eval_notes"] = notes
                STATE["pipeline_stage"] = "EVALUATION_APPROVED"
                STATE["audit_logger"].log("Giai đoạn 2", "HumanEvaluationGate", "Chuyên viên BI đánh giá kết quả đạt yêu cầu (Cổng 2)",
                                          f"Quyết định: ĐẠT YÊU CẦU | Ghi chú: {notes}")
                self._send_json({
                    "status": "approved",
                    "human_evaluated_analysis": True,
                    "human_eval_notes": notes,
                    "message": "Chuyên viên BI đã đánh giá Kết quả phân tích ĐẠT YÊU CẦU. Mở khóa xuất báo cáo & Dashboard!"
                })
            else:
                STATE["human_evaluated_analysis"] = False
                STATE["human_eval_notes"] = notes
                STATE["pipeline_stage"] = "HUMAN_VERIFIED"
                STATE["audit_logger"].log("Giai đoạn 2", "HumanEvaluationGate", "Chuyên viên BI yêu cầu điều chỉnh & phân tích lại (Cổng 2)",
                                          f"Quyết định: KHÔNG ĐẠT | Yêu cầu điều chỉnh: {notes}")
                run_skill_2_internal()
                self._send_json({
                    "status": "rejected",
                    "human_evaluated_analysis": False,
                    "human_eval_notes": notes,
                    "message": "Đã ghi nhận yêu cầu điều chỉnh & phân tích lại từ chuyên viên BI."
                })
            return

        if path == "/api/run-skill-2":
            if not STATE["cleansed_result"]:
                self._send_json({"error": "Cần thực hiện làm sạch dữ liệu trước."}, 400)
                return
            decision = run_skill_2_internal()
            self._send_json({
                "status": "success",
                "top_3_count": len(decision["top_3_stores"]),
                "stores_count": len(decision["store_metrics"])
            })
            return

        # 4. Trợ lý AI Co-Pilot hỏi đáp thông minh dựa trên dữ liệu thật
        if path == "/api/copilot-chat":
            payload = parse_json_body(body)
            query = payload.get("query", "")
            response_data = handle_copilot_chat(query)
            self._send_json(response_data)
            return

        # 5. Bổ sung / Hiệu chỉnh trường dữ liệu bị thiếu (Data Enrichment & Remediation)
        if path == "/api/enrich-transaction":
            payload = parse_json_body(body)
            trans_id = payload.get("trans_id")
            if not trans_id:
                self._send_json({"error": "Thiếu mã giao dịch (trans_id)."}, 400)
                return
            
            raw_txs = STATE["raw_dataset"].get("transactions", []) if STATE["raw_dataset"] else []
            found = False
            for tx in raw_txs:
                if str(tx.get("trans_id")) == str(trans_id):
                    found = True
                    if "order_id" in payload and str(payload["order_id"]).strip():
                        tx["order_id"] = str(payload["order_id"]).strip()
                    if "store_id" in payload and str(payload["store_id"]).strip():
                        tx["store_id"] = str(payload["store_id"]).strip()
                    if "product_id" in payload and str(payload["product_id"]).strip():
                        tx["product_id"] = str(payload["product_id"]).strip()
                    if "date" in payload and str(payload["date"]).strip():
                        tx["date"] = str(payload["date"]).strip()
                    if "status" in payload and str(payload["status"]).strip():
                        tx["status"] = str(payload["status"]).strip().upper()
                    if "type" in payload and str(payload["type"]).strip():
                        tx["type"] = str(payload["type"]).strip().upper()
                    if "qty" in payload and payload["qty"] is not None:
                        try:
                            tx["qty"] = int(payload["qty"])
                        except Exception:
                            pass
                    if "unit_price" in payload and payload["unit_price"] is not None:
                        try:
                            tx["unit_price"] = float(payload["unit_price"])
                        except Exception:
                            pass
                    break
            
            if not found:
                new_tx = {
                    "trans_id": str(trans_id),
                    "order_id": str(payload.get("order_id") or f"ORD_{trans_id}"),
                    "store_id": str(payload.get("store_id") or "CH01"),
                    "product_id": str(payload.get("product_id") or "SP01"),
                    "date": str(payload.get("date") or "2026-10-15"),
                    "status": str(payload.get("status") or "COMPLETED").upper(),
                    "type": str(payload.get("type") or "SALE").upper(),
                    "qty": int(payload.get("qty", 1)),
                    "unit_price": float(payload.get("unit_price", 100000.0)),
                    "ref_order_id": str(payload.get("ref_order_id") or payload.get("order_id") or f"ORD_{trans_id}")
                }
                raw_txs.append(new_tx)
                if STATE["raw_dataset"]:
                    STATE["raw_dataset"]["transactions"] = raw_txs

            if STATE["raw_dataset"]:
                masters = {
                    "stores": STATE["raw_dataset"].get("stores", {}),
                    "products": STATE["raw_dataset"].get("products", {}),
                    "targets": STATE["raw_dataset"].get("targets", {}),
                    "period_config": STATE["raw_dataset"].get("period_config", {})
                }
                stores, prods, targets, pcfg = ensure_masters_for_transactions(raw_txs, masters)
                STATE["raw_dataset"]["stores"] = stores
                STATE["raw_dataset"]["products"] = prods
                STATE["raw_dataset"]["targets"] = targets
                STATE["raw_dataset"]["period_config"] = pcfg

            STATE["audit_logger"].log("Giai đoạn 1", "DataEnrichment", f"Bổ sung trường thông tin cho giao dịch {trans_id}",
                                      f"Bổ sung: Đơn vị/Chi nhánh: {payload.get('store_id')}, Trạng thái: {payload.get('status')}")

            cleansed_res = run_skill_1_internal()
            decision_res = run_skill_2_internal()

            self._send_json({
                "status": "success",
                "message": f"Đã bổ sung thành công trường thông tin cho bản ghi {trans_id}! Dữ liệu sạch đã được đồng bộ lại.",
                "summary": cleansed_res["cleansed"]["summary"] if cleansed_res else {},
                "gate_info": cleansed_res["gate_info"] if cleansed_res else {}
            })
            return

        # 6. Chỉnh sửa dữ liệu trực tiếp trên bảng (Inline Direct Edit)
        if path == "/api/inline-edit-record":
            payload = parse_json_body(body)
            trans_id = str(payload.get("trans_id", "")).strip()
            field = str(payload.get("field", "")).strip()
            val = payload.get("value")

            if not trans_id or not field:
                self._send_json({"error": "Thiếu thông tin mã giao dịch hoặc trường sửa."}, 400)
                return

            raw_txs = STATE["raw_dataset"].get("transactions", []) if STATE["raw_dataset"] else []
            updated = False
            for tx in raw_txs:
                if str(tx.get("trans_id")) == trans_id:
                    if field in ["qty"]:
                        try:
                            tx[field] = int(val)
                        except Exception:
                            pass
                    elif field in ["unit_price"]:
                        try:
                            tx[field] = float(val)
                        except Exception:
                            pass
                    elif field in ["status", "type"]:
                        tx[field] = str(val).strip().upper()
                    else:
                        tx[field] = str(val).strip()
                    updated = True
                    break

            if not updated:
                self._send_json({"error": f"Không tìm thấy bản ghi {trans_id}"}, 404)
                return

            if STATE["raw_dataset"]:
                masters = {
                    "stores": STATE["raw_dataset"].get("stores", {}),
                    "products": STATE["raw_dataset"].get("products", {}),
                    "targets": STATE["raw_dataset"].get("targets", {}),
                    "period_config": STATE["raw_dataset"].get("period_config", {})
                }
                stores, prods, targets, pcfg = ensure_masters_for_transactions(raw_txs, masters)
                STATE["raw_dataset"]["stores"] = stores
                STATE["raw_dataset"]["products"] = prods
                STATE["raw_dataset"]["targets"] = targets
                STATE["raw_dataset"]["period_config"] = pcfg

            STATE["audit_logger"].log("Giai đoạn 1", "InlineEdit", f"Chỉnh sửa trực tiếp ô '{field}' cho {trans_id} thành '{val}'",
                                      "Dữ liệu được cập nhật trực tiếp tại bảng và đồng bộ lại")

            cleansed_res = run_skill_1_internal()
            decision_res = run_skill_2_internal()

            self._send_json({
                "status": "success",
                "message": f"Đã cập nhật trực tiếp trường {field} của {trans_id}!",
                "summary": cleansed_res["cleansed"]["summary"] if cleansed_res else {},
                "gate_info": cleansed_res["gate_info"] if cleansed_res else {}
            })
            return

        # 7. Lưu mô hình, vấn đề doanh nghiệp & câu hỏi làm rõ
        if path in ["/api/set-business-model", "/api/set-business-context"]:
            payload = parse_json_body(body)
            if "business_model" in payload:
                STATE["business_model"] = str(payload.get("business_model", "")).strip()
            if "business_problem" in payload:
                STATE["business_problem"] = str(payload.get("business_problem", "")).strip()
            if "clarifications" in payload and isinstance(payload.get("clarifications"), dict):
                STATE["clarifications"] = {
                    k: str(v or "").strip() for k, v in payload["clarifications"].items()
                }
            self._send_json({
                "status": "success",
                "business_model": STATE["business_model"],
                "business_problem": STATE["business_problem"],
                "clarifications": STATE["clarifications"]
            })
            return

        self.send_error(404, "Endpoint not found")

def handle_copilot_chat(query):
    query_lower = query.lower().strip()
    
    # Đảm bảo có dữ liệu báo cáo để trả lời
    if not STATE["decision_result"] and STATE["cleansed_result"]:
        run_skill_2_internal()
        
    res = STATE["decision_result"]
    gate = STATE["gate_info"]
    raw_ds = STATE["raw_dataset"]
    cleansed = STATE["cleansed_result"]
    
    # 1. Câu hỏi về đơn vị / cửa hàng có nguy cơ hụt mục tiêu / rủi ro cao nhất
    if any(k in query_lower for k in ["hụt mục tiêu", "nguy cơ", "rủi ro", "ưu tiên", "hỗ trợ", "kém nhất", "thấp nhất", "đơn vị nào"]):
        if res and res.get("top_3_stores"):
            top_stores = res["top_3_stores"]
            top1 = top_stores[0]
            
            # Tìm đơn vị đạt % mục tiêu doanh thu thấp nhất trong metrics
            metrics = res.get("store_metrics", [])
            worst_rev_target = min(metrics, key=lambda m: m["rev_target_pct"]) if metrics else None
            
            reply = f"🎯 <strong>Phân Tích Đơn Vị Rủi Ro & Nguy Cơ Cao Nhất:</strong><br><br>"
            reply += f"Đơn vị đang đứng đầu danh sách rủi ro cần hỗ trợ khẩn cấp là <strong>{top1['store_name']} ({top1['store_id']})</strong> (Điểm rủi ro: <span style='color:#b05757; font-weight:700;'>{top1['risk_score']}/100</span>):<br>"
            reply += f"• <strong>Doanh thu thuần:</strong> {top1['net_revenue']:,.0f} đ (Đạt <strong>{top1['rev_target_pct']}%</strong> kế hoạch)<br>"
            reply += f"• <strong>Lợi nhuận gộp:</strong> {top1['net_profit']:,.0f} đ (Biên LN: <strong>{top1['margin_pct']}%</strong>)<br>"
            reply += f"• <strong>Tỷ lệ đổi trả:</strong> <strong>{top1['return_rate_pct']}%</strong><br>"
            reply += f"• <strong>Căn cứ định lượng:</strong> {'; '.join(top1['root_causes'])}<br>"
            reply += f"• <strong>Đề xuất khắc phục:</strong> {top1['recommendations'][0]}<br><br>"
            
            if worst_rev_target and worst_rev_target['store_id'] != top1['store_id']:
                reply += f"⚠️ Riêng về <strong>tỷ lệ hụt mục tiêu doanh thu sâu nhất</strong>: Là <strong>{worst_rev_target['store_name']} ({worst_rev_target['store_id']})</strong> chỉ đạt <strong>{worst_rev_target['rev_target_pct']}%</strong> so với kế hoạch."
            
            return {
                "reply": reply,
                "action": "GO_TO_STEP_3",
                "action_label": "👉 Xem Phân Tích & Insight"
            }
            
    # 2. Câu hỏi về tỷ lệ đổi trả, hoàn trả, đơn hủy, đơn lỗi
    if any(k in query_lower for k in ["đổi trả", "hoàn trả", "hủy", "đơn hủy", "lỗi", "ngoại lệ"]):
        if gate:
            reply = f"🛡️ <strong>Báo Cáo Đối Soát Đổi Trả & Đơn Hủy Của Tệp Này:</strong><br><br>"
            reply += f"• <strong>Tỷ lệ đơn đổi trả thực tế:</strong> <span style='color:#c68a4c; font-weight:700;'>{gate['actual_return_rate']}%</span> (So với ngưỡng trung bình chuẩn hệ thống: <strong>{gate['system_avg_return_rate']}%</strong>)<br>"
            reply += f"• <strong>Tổng số đơn hoàn trả hợp lệ:</strong> {gate['total_returns']:,} đơn<br>"
            reply += f"• <strong>Tỷ lệ đơn bị hủy / lỗi:</strong> <span style='color:#c94a4a; font-weight:700;'>{gate['actual_cancel_rate']}%</span> ({gate['cancelled_count']:,} đơn bị loại trừ)<br>"
            reply += f"• <strong>Trạng thái kiểm duyệt:</strong> {gate['message']}"
            return {
                "reply": reply,
                "action": "GO_TO_STEP_2",
                "action_label": "👉 Xem Làm Sạch & Kiểm Duyệt"
            }
            
    # 3. Câu hỏi về số dòng dữ liệu đã làm sạch
    if any(k in query_lower for k in ["làm sạch", "bao nhiêu dòng", "dòng sạch", "hợp lệ"]):
        if gate:
            total_clean = gate.get('total_cleaned_valid', gate['valid_count'])
            valid_rep = gate.get('valid_report_period', gate['valid_count'])
            valid_comp = gate.get('valid_compare_period', 0)
            raw_rows = gate['total_raw']
            unqual_rows = gate['unqualified_count']
            clean_pct = round(total_clean / raw_rows * 100.0, 1) if raw_rows > 0 else 0
            
            reply = f"🧹 <strong>Kết Quả Làm Sạch & Đối Soát Dữ Liệu:</strong><br><br>"
            reply += f"• <strong>Tổng số dòng thô nạp vào:</strong> <strong>{raw_rows:,} dòng</strong><br>"
            reply += f"• <strong>Tổng số dòng sạch hợp lệ giữ lại:</strong> <span style='color:#137333; font-weight:700; font-size:1.05rem;'>{total_clean:,} dòng</span> ({clean_pct}%)<br>"
            reply += f"  - Trong đó: <strong>{valid_rep:,} dòng</strong> thuộc kỳ báo cáo chính<br>"
            if valid_comp > 0:
                reply += f"  - Và <strong>{valid_comp:,} dòng</strong> thuộc kỳ đối chiếu tăng trưởng<br>"
            reply += f"• <strong>Số dòng ngoại lệ bị loại bỏ:</strong> <span style='color:#b05757; font-weight:700;'>{unqual_rows:,} dòng</span> ({round(unqual_rows/raw_rows*100.0, 1) if raw_rows>0 else 0}%) (Đơn hủy, trùng mã, thiếu thông tin)."
            return {
                "reply": reply,
                "action": "GO_TO_STEP_2",
                "action_label": "👉 Mở Bước Làm Sạch & Kiểm Duyệt"
            }

    # 4. Câu hỏi về doanh thu, lợi nhuận, chi phí, tài chính, KPI
    if any(k in query_lower for k in ["doanh thu", "lợi nhuận", "chi phí", "giá vốn", "kpi", "tổng"]):
        if res:
            metrics = res.get("store_metrics", [])
            total_rev = sum(m["net_revenue"] for m in metrics)
            total_profit = sum(m["net_profit"] for m in metrics)
            total_cost = total_rev - total_profit
            margin = round((total_profit / total_rev * 100.0), 1) if total_rev > 0 else 0
            
            reply = f"📊 <strong>Tổng Quan Chỉ Số Tài Chính Kỳ Này:</strong><br><br>"
            reply += f"• <strong>Tổng doanh thu thuần:</strong> <span style='color:#b8860b; font-weight:700;'>{total_rev:,.0f} đ</span><br>"
            reply += f"• <strong>Tổng chi phí giá vốn:</strong> {total_cost:,.0f} đ<br>"
            reply += f"• <strong>Tổng lợi nhuận gộp:</strong> <span style='color:#2a9d8f; font-weight:700;'>{total_profit:,.0f} đ</span><br>"
            reply += f"• <strong>Biên lợi nhuận gộp trung bình:</strong> <strong>{margin}%</strong>"
            return {
                "reply": reply,
                "action": "GO_TO_STEP_4",
                "action_label": "👉 Mở Bento Dashboard & Báo Cáo"
            }

    # 5. Câu hỏi về xuất file, tải file, excel, word, báo cáo
    if any(k in query_lower for k in ["excel", "word", "docx", "tải", "xuất", "file", "báo cáo"]):
        reply = "📥 <strong>Bạn có thể xuất 3 định dạng báo cáo tại bước Dashboard:</strong><br><br>"
        reply += "1. <strong>Excel 2 Sheets:</strong> Dữ liệu đã làm sạch & Báo cáo tổng quan.<br>"
        reply += "2. <strong>Báo Cáo Word (.docx):</strong> Văn bản phân tích chuyên sâu trình Lãnh đạo.<br>"
        reply += "3. <strong>Executive Dashboard (.xlsx):</strong> Bảng tổng hợp điều hành.<br><br>"
        reply += "<a href='/api/export-excel' style='color:#b8860b; font-weight:700; text-decoration:underline;'>[Tải Excel 2 Sheets]</a> · "
        reply += "<a href='/api/export-word' style='color:#1d70b8; font-weight:700; text-decoration:underline;'>[Tải Báo cáo Word]</a>"
        return {
            "reply": reply,
            "action": "GO_TO_STEP_4",
            "action_label": "👉 Chuyển Đến Dashboard & Xuất Tệp"
        }

    # Fallback thông minh dựa trên dữ liệu thực tế
    units_count = len(res["store_metrics"]) if (res and "store_metrics" in res) else 0
    reply = f"🤖 <strong>Ngọc Ánh AI Co-Pilot đã phân tích xong:</strong><br><br>"
    reply += f"Hệ thống hiện đang quản trị <strong>{units_count} đơn vị / chi nhánh</strong> với dữ liệu đã được làm sạch và đối soát. Bạn có thể hỏi tôi bất kỳ câu nào sau đây:<br>"
    reply += f"• <em>'Đơn vị nào có nguy cơ hụt mục tiêu cao nhất?'</em><br>"
    reply += f"• <em>'Báo cáo tỷ lệ đổi trả và đơn hủy của tệp này'</em><br>"
    reply += f"• <em>'Doanh thu và lợi nhuận toàn hệ thống là bao nhiêu?'</em>"
    return {
        "reply": reply,
        "action": None
    }

def run(port=8080):
    init_benchmark()
    run_skill_1_internal()
    server_address = ("", port)
    httpd = HTTPServer(server_address, RequestHandler)
    print(f"================================================================")
    print(f" Ngoc Anh Data Studio Server running at http://localhost:{port}")
    print(f"================================================================")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        httpd.server_close()

if __name__ == "__main__":
    run(8080)
