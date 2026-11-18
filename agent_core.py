"""
AI Agent Core Engine for Phuc An Retail Management System
Compliant with BPMN 3-phase Retail Analytics & Auditing Workflow
"""

import json
import math
from datetime import datetime
from collections import defaultdict

class AuditLogger:
    """Tracks step-by-step execution, rules applied, and data lineage."""
    def __init__(self):
        self.logs = []

    def log(self, phase, skill, action, details, records_affected=0):
        entry = {
            "timestamp": datetime.now().isoformat(),
            "phase": phase,
            "skill": skill,
            "action": action,
            "details": details,
            "records_affected": records_affected
        }
        self.logs.append(entry)
        return entry

    def get_logs(self):
        return self.logs

    def clear(self):
        self.logs = []


class DataCleansingAndAuditSkill:
    """
    Kỹ năng nghiệp vụ 1: Kiểm định chất lượng dữ liệu, lọc trùng,
    loại bỏ đơn hủy, đối chiếu đơn đổi trả theo quy tắc và chuẩn hóa kỳ báo cáo.
    """
    def __init__(self, audit_logger):
        self.logger = audit_logger

    def execute(self, raw_transactions, stores_master, products_master, period_config):
        """
        Input:
            raw_transactions: list of dicts (trans_id, order_id, store_id, product_id, date, type, qty, unit_price, status, ref_order_id)
            stores_master: dict store_id -> {name, region, manager}
            products_master: dict product_id -> {name, category, cogs}
            period_config: {report_start, report_end, compare_start, compare_end}
        Output:
            cleaned_report_data: list of valid transactions in reporting period
            cleaned_compare_data: list of valid transactions in comparison period
            unqualified_records: list of excluded records with specific rejection reason
            summary_stats: metadata on cleansing metrics
        """
        self.logger.log("Giai đoạn 1", "DataCleansingSkill", "Khởi động tiếp nhận dữ liệu", 
                        f"Tiếp nhận {len(raw_transactions)} bản ghi giao dịch thô", len(raw_transactions))

        cleaned_report = []
        cleaned_compare = []
        unqualified = []

        seen_trans_ids = set()
        completed_orders = {}  # order_id -> list of sale items for return matching

        # Bước 1: Phân loại cơ bản và lưu trữ đơn bán hợp lệ để đối chiếu trả hàng
        step1_valid_items = []
        for idx, item in enumerate(raw_transactions):
            trans_id = item.get("trans_id")
            order_id = item.get("order_id")
            status = str(item.get("status", "")).strip().upper()
            trans_type = str(item.get("type", "")).strip().upper()  # 'SALE' or 'RETURN'
            
            # Kiểm tra 1: Bản ghi trùng Transaction ID
            if trans_id in seen_trans_ids:
                unqualified.append({
                    "record": item,
                    "reason_code": "ERR_DUPLICATE_TRANS_ID",
                    "reason": f"Trùng mã giao dịch {trans_id}",
                    "severity": "HIGH",
                    "suggested_action": "Loại bỏ bản ghi trùng lặp"
                })
                continue
            seen_trans_ids.add(trans_id)

            # Kiểm tra 2: Đơn bị hủy (CANCELLED / VOID)
            if status in ["CANCELLED", "HUY", "VOID"]:
                unqualified.append({
                    "record": item,
                    "reason_code": "ERR_ORDER_CANCELLED",
                    "reason": f"Đơn hàng trạng thái hủy ({status})",
                    "severity": "INFO",
                    "suggested_action": "Không hạch toán doanh thu và giá vốn"
                })
                continue

            # Kiểm tra 3: Kiểm tra tính hợp lệ của cửa hàng (Store ID)
            store_id = item.get("store_id")
            if store_id not in stores_master:
                unqualified.append({
                    "record": item,
                    "reason_code": "ERR_UNKNOWN_STORE",
                    "reason": f"Mã cửa hàng {store_id} không tồn tại trong danh mục cửa hàng",
                    "severity": "HIGH",
                    "suggested_action": "Bổ sung thông tin cửa hàng hoặc kiểm tra lại mã POS"
                })
                continue

            # Kiểm tra 4: Kiểm tra tính hợp lệ của sản phẩm và giá vốn (Product ID & COGS)
            product_id = item.get("product_id")
            if product_id not in products_master:
                unqualified.append({
                    "record": item,
                    "reason_code": "ERR_UNKNOWN_PRODUCT",
                    "reason": f"Mã sản phẩm {product_id} không tồn tại trong danh mục sản phẩm",
                    "severity": "CRITICAL",
                    "suggested_action": "Cập nhật danh mục sản phẩm trước khi chốt số liệu"
                })
                continue

            prod_info = products_master[product_id]
            unit_cogs = prod_info.get("cogs")
            if unit_cogs is None or unit_cogs < 0:
                unqualified.append({
                    "record": item,
                    "reason_code": "ERR_MISSING_COGS",
                    "reason": f"Sản phẩm {product_id} thiếu giá vốn hợp lệ (COGS = {unit_cogs})",
                    "severity": "CRITICAL",
                    "suggested_action": "Kế toán giá thành cần cung cấp định mức giá vốn"
                })
                continue

            # Lưu vết đơn bán để đối chiếu hàng trả
            if trans_type == "SALE" and status in ["COMPLETED", "HOAN_THANH", "DELIVERED"]:
                if order_id not in completed_orders:
                    completed_orders[order_id] = []
                completed_orders[order_id].append(item)

            step1_valid_items.append(item)

        self.logger.log("Giai đoạn 1", "DataCleansingSkill", "Hoàn tất lọc cơ bản và đơn hủy", 
                        f"Phát hiện {len(unqualified)} bản ghi không đạt yêu cầu sơ bộ", len(unqualified))

        # Bước 2: Kiểm tra kỳ báo cáo & đối chiếu đổi trả theo quy tắc nghiệp vụ
        rep_start = period_config.get("report_start")
        rep_end = period_config.get("report_end")
        comp_start = period_config.get("compare_start")
        comp_end = period_config.get("compare_end")

        for item in step1_valid_items:

            trans_date = item.get("date")
            trans_type = str(item.get("type", "")).strip().upper()
            order_id = item.get("order_id")
            ref_order_id = item.get("ref_order_id") or order_id
            product_id = item.get("product_id")
            qty = abs(int(item.get("qty", 1)))
            unit_price = float(item.get("unit_price", 0.0))
            prod_info = products_master[product_id]
            unit_cogs = float(prod_info.get("cogs", 0.0))

            # Xác định kỳ
            is_report_period = (rep_start <= trans_date <= rep_end) if (rep_start and rep_end) else False
            is_compare_period = (comp_start <= trans_date <= comp_end) if (comp_start and comp_end) else False

            if not is_report_period and not is_compare_period:
                unqualified.append({
                    "record": item,
                    "reason_code": "ERR_OUT_OF_PERIOD",
                    "reason": f"Ngày giao dịch {trans_date} không thuộc kỳ báo cáo ({rep_start} -> {rep_end}) hoặc kỳ so sánh ({comp_start} -> {comp_end})",
                    "severity": "MEDIUM",
                    "suggested_action": "Tách giao dịch sang kỳ kế toán tương ứng"
                })
                continue

            # Quy tắc đối chiếu đổi trả (Return Matching Rules)
            # 1. Đơn trả hàng phải có ref_order_id tham chiếu đến đơn bán gốc
            # 2. Hàng trả trừ trực tiếp doanh thu và giảm giá vốn tương ứng (Net COGS)
            # 3. Cho phép trả hàng khác cửa hàng (Cross-store return) nhưng ghi nhận đúng cửa hàng tiếp nhận xử lý
            if trans_type == "RETURN":
                # Đối chiếu đơn bán gốc
                orig_sales = completed_orders.get(ref_order_id, [])
                matching_sale = None
                for sale in orig_sales:
                    if sale.get("product_id") == product_id:
                        matching_sale = sale
                        break

                if not matching_sale:
                    # Cảnh báo: Trả hàng không tìm thấy đơn bán gốc tương thích
                    unqualified.append({
                        "record": item,
                        "reason_code": "ERR_UNMATCHED_RETURN",
                        "reason": f"Đơn đổi trả {order_id} (Ref: {ref_order_id}) không tìm thấy đơn bán gốc chứa sản phẩm {product_id}",
                        "severity": "HIGH",
                        "suggested_action": "Kiểm tra chứng từ gốc của khách hàng hoặc phiếu tiếp nhận bảo hành/trả đổi"
                    })
                    continue

            # Bản ghi hợp lệ được phân bổ vào kỳ tương ứng
            enriched_record = dict(item)
            enriched_record["product_name"] = prod_info.get("name")
            enriched_record["category"] = prod_info.get("category")
            enriched_record["unit_cogs"] = unit_cogs
            enriched_record["gross_revenue"] = (unit_price * qty) if trans_type == "SALE" else 0.0
            enriched_record["return_revenue"] = (unit_price * qty) if trans_type == "RETURN" else 0.0
            enriched_record["net_revenue"] = (unit_price * qty) if trans_type == "SALE" else (-1.0 * unit_price * qty)
            enriched_record["net_cogs"] = (unit_cogs * qty) if trans_type == "SALE" else (-1.0 * unit_cogs * qty)
            enriched_record["net_profit"] = enriched_record["net_revenue"] - enriched_record["net_cogs"]

            if is_report_period:
                cleaned_report.append(enriched_record)
            elif is_compare_period:
                cleaned_compare.append(enriched_record)

        self.logger.log("Giai đoạn 1", "DataCleansingSkill", "Hoàn tất làm sạch & đồng bộ kỳ", 
                        f"Hợp lệ kỳ báo cáo: {len(cleaned_report)} dòng | Hợp lệ kỳ so sánh: {len(cleaned_compare)} dòng | Loại bỏ: {len(unqualified)} dòng", 
                        len(cleaned_report) + len(cleaned_compare))

        summary = {
            "total_raw": len(raw_transactions),
            "valid_report_period": len(cleaned_report),
            "valid_compare_period": len(cleaned_compare),
            "total_cleaned_valid": len(cleaned_report) + len(cleaned_compare),
            "total_unqualified": len(unqualified),
            "unqualified_by_code": defaultdict(int)
        }
        for u in unqualified:
            summary["unqualified_by_code"][u["reason_code"]] += 1
        summary["unqualified_by_code"] = dict(summary["unqualified_by_code"])

        return {
            "cleaned_report": cleaned_report,
            "cleaned_compare": cleaned_compare,
            "unqualified": unqualified,
            "summary": summary
        }


class RetailDecisionIntelligenceSkill:
    """
    Kỹ năng nghiệp vụ 2: Tính toán chỉ số quản trị (Doanh thu thuần, Lợi nhuận gộp, % hoàn thành),
    phân tích cơ cấu sản phẩm & đổi trả, phát hiện bất thường và đề xuất Top 3 cửa hàng ưu tiên.
    """
    def __init__(self, audit_logger):
        self.logger = audit_logger

    def execute(self, cleaned_report, cleaned_compare, stores_master, targets_master, period_config):
        """
        Input:
            cleaned_report: list of verified transactions (report period)
            cleaned_compare: list of verified transactions (comparison period)
            stores_master: dict store_id -> metadata
            targets_master: dict store_id -> {report_rev_target, report_profit_target, compare_rev_target, compare_profit_target}
            period_config: config info
        Output:
            store_metrics: table by store (revenue, profit, % target, return rate, growth)
            category_metrics: table by store & category
            top_3_stores: prioritized list with quantitative root causes and actionable recommendations
            reconciliation_audit: source-trace verification proofs
        """
        self.logger.log("Giai đoạn 2", "RetailDecisionSkill", "Tính toán chỉ số tài chính quản trị", 
                        f"Xử lý {len(cleaned_report)} giao dịch kỳ hiện tại và {len(cleaned_compare)} giao dịch kỳ so sánh", 
                        len(cleaned_report))

        # 1. Tính toán cho Kỳ Báo Cáo theo cửa hàng & ngành hàng
        report_store_agg = defaultdict(lambda: {
            "gross_revenue": 0.0,
            "return_revenue": 0.0,
            "net_revenue": 0.0,
            "net_cogs": 0.0,
            "net_profit": 0.0,
            "sale_tx_count": 0,
            "return_tx_count": 0,
            "transaction_ids": []
        })

        report_cat_agg = defaultdict(lambda: {
            "net_revenue": 0.0,
            "net_cogs": 0.0,
            "net_profit": 0.0,
            "gross_revenue": 0.0,
            "return_revenue": 0.0,
            "units_sold": 0,
            "units_returned": 0
        })

        report_store_cat_agg = defaultdict(lambda: defaultdict(lambda: {
            "net_revenue": 0.0,
            "net_profit": 0.0,
            "return_revenue": 0.0
        }))

        for tx in cleaned_report:
            sid = tx["store_id"]
            cat = tx["category"]
            tid = tx["trans_id"]
            ttype = tx["type"]
            qty = abs(int(tx["qty"]))

            st = report_store_agg[sid]
            st["gross_revenue"] += tx["gross_revenue"]
            st["return_revenue"] += tx["return_revenue"]
            st["net_revenue"] += tx["net_revenue"]
            st["net_cogs"] += tx["net_cogs"]
            st["net_profit"] += tx["net_profit"]
            st["transaction_ids"].append(tid)
            if ttype == "SALE":
                st["sale_tx_count"] += 1
            else:
                st["return_tx_count"] += 1

            ca = report_cat_agg[cat]
            ca["gross_revenue"] += tx["gross_revenue"]
            ca["return_revenue"] += tx["return_revenue"]
            ca["net_revenue"] += tx["net_revenue"]
            ca["net_cogs"] += tx["net_cogs"]
            ca["net_profit"] += tx["net_profit"]
            if ttype == "SALE":
                ca["units_sold"] += qty
            else:
                ca["units_returned"] += qty

            sca = report_store_cat_agg[sid][cat]
            sca["net_revenue"] += tx["net_revenue"]
            sca["net_profit"] += tx["net_profit"]
            sca["return_revenue"] += tx["return_revenue"]

        # 2. Tính toán cho Kỳ So Sánh theo cửa hàng
        compare_store_agg = defaultdict(lambda: {
            "net_revenue": 0.0,
            "net_cogs": 0.0,
            "net_profit": 0.0,
            "gross_revenue": 0.0,
            "return_revenue": 0.0
        })
        for tx in cleaned_compare:
            sid = tx["store_id"]
            st = compare_store_agg[sid]
            st["net_revenue"] += tx["net_revenue"]
            st["net_cogs"] += tx["net_cogs"]
            st["net_profit"] += tx["net_profit"]
            st["gross_revenue"] += tx["gross_revenue"]
            st["return_revenue"] += tx["return_revenue"]

        # 3. Tổng hợp Bảng Chỉ Số Cửa Hàng & Đánh giá Mục Tiêu
        store_metrics = []
        for sid, s_meta in stores_master.items():
            rep = report_store_agg[sid]
            cmp = compare_store_agg[sid]
            tgt = targets_master.get(sid, {})

            rev_target = tgt.get("report_rev_target", 0.0)
            profit_target = tgt.get("report_profit_target", 0.0)

            # Chỉ số kỳ báo cáo
            net_rev = rep["net_revenue"]
            net_profit = rep["net_profit"]
            gross_rev = rep["gross_revenue"]
            ret_rev = rep["return_revenue"]
            
            margin_pct = (net_profit / net_rev * 100.0) if net_rev > 0 else 0.0
            return_rate_pct = (ret_rev / gross_rev * 100.0) if gross_rev > 0 else 0.0

            rev_target_pct = (net_rev / rev_target * 100.0) if rev_target > 0 else 0.0
            profit_target_pct = (net_profit / profit_target * 100.0) if profit_target > 0 else 0.0

            # So sánh với kỳ trước (Growth rates)
            prev_net_rev = cmp["net_revenue"]
            prev_net_profit = cmp["net_profit"]
            rev_growth_pct = ((net_rev - prev_net_rev) / prev_net_rev * 100.0) if prev_net_rev > 0 else 0.0
            profit_growth_pct = ((net_profit - prev_net_profit) / abs(prev_net_profit) * 100.0) if prev_net_profit != 0 else 0.0

            # Phân tích cơ cấu sản phẩm của cửa hàng
            cat_breakdown = {}
            for cat, c_val in report_store_cat_agg[sid].items():
                cat_share = (c_val["net_revenue"] / net_rev * 100.0) if net_rev > 0 else 0.0
                cat_breakdown[cat] = {
                    "net_revenue": round(c_val["net_revenue"], 2),
                    "net_profit": round(c_val["net_profit"], 2),
                    "return_revenue": round(c_val["return_revenue"], 2),
                    "share_pct": round(cat_share, 1)
                }

            row = {
                "store_id": sid,
                "store_name": s_meta.get("name", sid),
                "region": s_meta.get("region", "N/A"),
                "manager": s_meta.get("manager", "N/A"),
                # Kỳ hiện tại
                "gross_revenue": round(gross_rev, 2),
                "return_revenue": round(ret_rev, 2),
                "net_revenue": round(net_rev, 2),
                "net_cogs": round(rep["net_cogs"], 2),
                "net_profit": round(net_profit, 2),
                "margin_pct": round(margin_pct, 2),
                "return_rate_pct": round(return_rate_pct, 2),
                # Mục tiêu
                "rev_target": rev_target,
                "profit_target": profit_target,
                "rev_target_pct": round(rev_target_pct, 2),
                "profit_target_pct": round(profit_target_pct, 2),
                # So sánh kỳ trước
                "prev_net_revenue": round(prev_net_rev, 2),
                "prev_net_profit": round(prev_net_profit, 2),
                "rev_growth_pct": round(rev_growth_pct, 2),
                "profit_growth_pct": round(profit_growth_pct, 2),
                # Vết truy xuất
                "tx_count": len(rep["transaction_ids"]),
                "raw_trans_sample": rep["transaction_ids"][:5],
                "all_trans_ids": rep["transaction_ids"],
                "categories": cat_breakdown
            }
            store_metrics.append(row)

        # 4. Kỹ năng Phân Tích & Xác Định Top 3 Cửa Hàng Cần Ưu Tiên Hỗ Trợ
        # Tiêu chí chấm điểm nguy cơ (Risk Scoring Model):
        # - Hoàn thành mục tiêu doanh thu thấp (< 85%) -> Điểm phạt cao
        # - Hoàn thành mục tiêu lợi nhuận thấp hoặc lỗ (< 80%) -> Trọng số lớn
        # - Tỷ lệ đổi trả cao bất thường (> 5%)
        # - Tăng trưởng âm so với kỳ trước
        self.logger.log("Giai đoạn 3", "RetailDecisionSkill", "Đánh giá xếp hạng rủi ro cửa hàng", 
                        "Áp dụng thuật toán Risk Scoring đa chiều để chọn Top 3 cửa hàng cần hỗ trợ", 18)

        scored_stores = []
        for s in store_metrics:
            # Điểm rủi ro (Risk Score - điểm càng cao càng đáng báo động)
            risk_score = 0.0

            # 1. Yếu tố hoàn thành doanh thu
            if s["rev_target_pct"] < 70:
                risk_score += 40
            elif s["rev_target_pct"] < 85:
                risk_score += 25
            elif s["rev_target_pct"] < 95:
                risk_score += 10

            # 2. Yếu tố hoàn thành lợi nhuận & biên lợi nhuận gộp
            if s["net_profit"] < 0:
                risk_score += 50  # Âm lợi nhuận là báo động đỏ
            elif s["profit_target_pct"] < 70:
                risk_score += 35
            elif s["profit_target_pct"] < 85:
                risk_score += 20

            # 3. Tỷ lệ đổi trả cao
            if s["return_rate_pct"] > 8.0:
                risk_score += 25
            elif s["return_rate_pct"] > 5.0:
                risk_score += 15

            # 4. Tăng trưởng âm
            if s["profit_growth_pct"] < -10.0:
                risk_score += 20
            elif s["rev_growth_pct"] < -5.0:
                risk_score += 10

            # Kiểm tra cơ cấu sản phẩm: nếu phụ thuộc vào ngành hàng lỗ/biên mỏng
            low_margin_cats = [c for c, d in s["categories"].items() if d["net_profit"] < 0 or (d["net_revenue"] > 0 and d["net_profit"]/d["net_revenue"] < 0.1)]
            if low_margin_cats:
                risk_score += 10

            scored_stores.append({
                "store": s,
                "risk_score": round(risk_score, 1),
                "low_margin_cats": low_margin_cats
            })

        # Sắp xếp theo rủi ro giảm dần
        scored_stores.sort(key=lambda x: x["risk_score"], reverse=True)
        top_3_candidates = scored_stores[:3]

        top_3_stores = []
        for rank, item in enumerate(top_3_candidates, 1):
            s = item["store"]
            score = item["risk_score"]

            # Phân tích nguyên nhân định lượng (Quantitative Root Causes)
            causes = []
            if s["rev_target_pct"] < 85:
                causes.append(f"Mức hoàn thành doanh thu thấp chỉ đạt {s['rev_target_pct']}% so với kế hoạch (kế hoạch {s['rev_target']:,.0f} đ, thực tế {s['net_revenue']:,.0f} đ).")
            if s["net_profit"] < 0:
                causes.append(f"Cửa hàng bị LỖ GỘP với lợi nhuận ròng -{abs(s['net_profit']):,.0f} đ (Biên lợi nhuận {s['margin_pct']}%).")
            elif s["profit_target_pct"] < 80:
                causes.append(f"Lợi nhuận gộp không đạt kỳ vọng ({s['profit_target_pct']}% mục tiêu lợi nhuận, đạt {s['net_profit']:,.0f} đ).")
            
            if s["return_rate_pct"] > 5.0:
                causes.append(f"Tỷ lệ hàng đổi trả báo động ở mức {s['return_rate_pct']}% (doanh thu trả lại {s['return_revenue']:,.0f} đ trên tổng bán {s['gross_revenue']:,.0f} đ).")

            if s["profit_growth_pct"] < 0:
                causes.append(f"Lợi nhuận suy giảm {abs(s['profit_growth_pct'])}% so với cùng kỳ trước ({s['prev_net_profit']:,.0f} đ -> {s['net_profit']:,.0f} đ).")

            # Đề xuất cải thiện hành động cụ thể (Prescriptive Actionable Recommendations)
            recommendations = []
            if s["return_rate_pct"] > 5.0:
                recommendations.append("Thanh tra chất lượng đóng gói và tư vấn size tại cửa hàng; rà soát lỗi sản phẩm ở ngành hàng có tỷ lệ trả cao nhất.")
            if s["margin_pct"] < 20 or s["net_profit"] < 0:
                recommendations.append("Tái cơ cấu danh mục trưng bày: giảm diện tích trưng bày hàng biên mỏng/hàng tồn kho chậm luân chuyển; đẩy mạnh các sản phẩm chủ lực có biên lợi nhuận > 35%.")
            if s["rev_target_pct"] < 85:
                recommendations.append("Thiết lập chương trình kích cầu tại chỗ (Local Activation), đào tạo kỹ năng up-sell/cross-sell cho đội ngũ bán hàng và tối ưu ca kíp theo khung giờ cao điểm.")
            recommendations.append("Giám đốc Kinh doanh khu vực trực tiếp làm việc với Cửa hàng trưởng trong 48h để rà soát cam kết P&L tuần tiếp theo.")

            top_3_stores.append({
                "rank": rank,
                "store_id": s["store_id"],
                "store_name": s["store_name"],
                "region": s["region"],
                "manager": s["manager"],
                "risk_score": score,
                "net_revenue": s["net_revenue"],
                "rev_target_pct": s["rev_target_pct"],
                "net_profit": s["net_profit"],
                "profit_target_pct": s["profit_target_pct"],
                "margin_pct": s["margin_pct"],
                "return_rate_pct": s["return_rate_pct"],
                "rev_growth_pct": s["rev_growth_pct"],
                "profit_growth_pct": s["profit_growth_pct"],
                "root_causes": causes,
                "recommendations": recommendations,
                "contributing_trans_ids_count": s["tx_count"],
                "sample_trans_ids": s["raw_trans_sample"]
            })

        self.logger.log("Giai đoạn 3", "RetailDecisionSkill", "Hoàn tất báo cáo & đề xuất Top 3 cửa hàng", 
                        f"Đã chọn: {[t['store_name'] for t in top_3_stores]}", 3)

        return {
            "store_metrics": store_metrics,
            "category_metrics": dict(report_cat_agg),
            "top_3_stores": top_3_stores,
            "period_config": period_config
        }
