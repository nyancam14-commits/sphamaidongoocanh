"""
Automated Test Suite for Phuc An Retail AI Agent
Verifies normal cases, missing data, and boundary rules.
"""

from agent_core import AuditLogger, DataCleansingAndAuditSkill, RetailDecisionIntelligenceSkill

def run_tests():
    logger = AuditLogger()
    cleansing_skill = DataCleansingAndAuditSkill(logger)
    decision_skill = RetailDecisionIntelligenceSkill(logger)

    results = []

    # Mock Master Data
    stores_master = {
        "CH01": {"store_id": "CH01", "name": "Phúc An Ba Đình", "region": "Miền Bắc"},
        "CH02": {"store_id": "CH02", "name": "Phúc An Cầu Giấy", "region": "Miền Bắc"}
    }
    products_master = {
        "SKU01": {"name": "Áo Sơ Mi", "category": "Thời trang", "cogs": 200000},
        "SKU02": {"name": "Quần Tây", "category": "Thời trang", "cogs": 300000}
    }
    period_config = {
        "report_start": "2026-10-01",
        "report_end": "2026-10-31",
        "compare_start": "2026-09-01",
        "compare_end": "2026-09-30"
    }

    # Test 1: Deduplication
    test1_raw = [
        {"trans_id": "TX01", "order_id": "ORD01", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-05", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": ""},
        {"trans_id": "TX01", "order_id": "ORD01", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-05", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": ""}
    ]
    res1 = cleansing_skill.execute(test1_raw, stores_master, products_master, period_config)
    pass1 = (len(res1["cleaned_report"]) == 1 and len(res1["unqualified"]) == 1 and res1["unqualified"][0]["reason_code"] == "ERR_DUPLICATE_TRANS_ID")
    results.append({"id": "TC01", "name": "Lọc bản ghi trùng mã giao dịch (Deduplication)", "passed": pass1, "details": "Đã phát hiện và loại bỏ bản ghi trùng TX01"})

    # Test 2: Filter Cancelled Orders
    test2_raw = [
        {"trans_id": "TX02", "order_id": "ORD02", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-06", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "CANCELLED", "ref_order_id": ""}
    ]
    res2 = cleansing_skill.execute(test2_raw, stores_master, products_master, period_config)
    pass2 = (len(res2["cleaned_report"]) == 0 and res2["unqualified"][0]["reason_code"] == "ERR_ORDER_CANCELLED")
    results.append({"id": "TC02", "name": "Loại bỏ đơn hàng bị hủy (Cancelled Orders Filter)", "passed": pass2, "details": "Đã loại bỏ đơn ORD02 có trạng thái CANCELLED"})

    # Test 3: Return Matching with Original Order
    test3_raw = [
        {"trans_id": "TX03", "order_id": "ORD03", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-07", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": ""},
        {"trans_id": "TX04", "order_id": "RET03", "store_id": "CH02", "product_id": "SKU01", "date": "2026-10-09", "type": "RETURN", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": "ORD03"}
    ]
    res3 = cleansing_skill.execute(test3_raw, stores_master, products_master, period_config)
    pass3 = (len(res3["cleaned_report"]) == 2 and len(res3["unqualified"]) == 0)
    results.append({"id": "TC03", "name": "Khớp đơn đổi trả với đơn bán gốc (Return Order Matching)", "passed": pass3, "details": "Khớp thành công đơn trả RET03 với đơn bán gốc ORD03 (hỗ trợ đổi trả chéo cửa hàng)"})

    # Test 4: Unmatched Return (Orphan Return)
    test4_raw = [
        {"trans_id": "TX05", "order_id": "RET04", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-10", "type": "RETURN", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": "NON_EXIST_ORDER"}
    ]
    res4 = cleansing_skill.execute(test4_raw, stores_master, products_master, period_config)
    pass4 = (len(res4["cleaned_report"]) == 0 and res4["unqualified"][0]["reason_code"] == "ERR_UNMATCHED_RETURN")
    results.append({"id": "TC04", "name": "Bắt lỗi đơn trả không có căn cứ (Unmatched Orphan Return)", "passed": pass4, "details": "Cô lập đơn trả RET04 không có hóa đơn bán gốc"})

    # Test 5: Out of Period Transactions
    test5_raw = [
        {"trans_id": "TX06", "order_id": "ORD05", "store_id": "CH01", "product_id": "SKU01", "date": "2026-08-15", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": ""}
    ]
    res5 = cleansing_skill.execute(test5_raw, stores_master, products_master, period_config)
    pass5 = (len(res5["cleaned_report"]) == 0 and res5["unqualified"][0]["reason_code"] == "ERR_OUT_OF_PERIOD")
    results.append({"id": "TC05", "name": "Kiểm soát ranh giới kỳ thời gian (Period Boundary Check)", "passed": pass5, "details": "Loại bỏ giao dịch ngoài phạm vi kỳ báo cáo và kỳ so sánh"})

    # Test 6: Missing COGS
    test6_raw = [
        {"trans_id": "TX07", "order_id": "ORD06", "store_id": "CH01", "product_id": "SKU_NO_COGS", "date": "2026-10-12", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": ""}
    ]
    res6 = cleansing_skill.execute(test6_raw, stores_master, products_master, period_config)
    pass6 = (len(res6["cleaned_report"]) == 0 and res6["unqualified"][0]["reason_code"] == "ERR_UNKNOWN_PRODUCT")
    results.append({"id": "TC06", "name": "Ngăn chặn tự tạo giá vốn khi thiếu danh mục (No Hallucinated COGS)", "passed": pass6, "details": "Từ chối đưa vào tính toán khi mã SKU không có định mức giá vốn xác thực"})

    # Test 7: Net Revenue & Gross Profit Calculation Integrity
    # CH01: 1 sale 500k (cogs 200k) + 1 return 500k (cogs 200k) -> Net Rev = 0, Net Profit = 0
    test7_raw = [
        {"trans_id": "TX08", "order_id": "ORD07", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-15", "type": "SALE", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": ""},
        {"trans_id": "TX09", "order_id": "RET07", "store_id": "CH01", "product_id": "SKU01", "date": "2026-10-16", "type": "RETURN", "qty": 1, "unit_price": 500000, "status": "COMPLETED", "ref_order_id": "ORD07"}
    ]
    res7 = cleansing_skill.execute(test7_raw, stores_master, products_master, period_config)
    targets_mock = {"CH01": {"report_rev_target": 1000000, "report_profit_target": 500000}, "CH02": {"report_rev_target": 1000000, "report_profit_target": 500000}}
    dec7 = decision_skill.execute(res7["cleaned_report"], [], stores_master, targets_mock, period_config)
    ch01_metric = next(m for m in dec7["store_metrics"] if m["store_id"] == "CH01")
    pass7 = (ch01_metric["net_revenue"] == 0.0 and ch01_metric["net_profit"] == 0.0 and ch01_metric["return_revenue"] == 500000)
    results.append({"id": "TC07", "name": "Tính đúng Doanh thu thuần & Lợi nhuận gộp sau đổi trả", "passed": pass7, "details": "Đúng công thức Doanh thu thuần = Bán - Trả; Lợi nhuận gộp = Doanh thu thuần - Giá vốn thuần"})

    # Test 8: Source Traceability (Every metric links to transaction IDs)
    pass8 = len(ch01_metric["all_trans_ids"]) == 2 and "TX08" in ch01_metric["all_trans_ids"] and "TX09" in ch01_metric["all_trans_ids"]
    results.append({"id": "TC08", "name": "Truy vết nguồn gốc giao dịch (Full Traceability Audit)", "passed": pass8, "details": "Mọi chỉ số tổng hợp đều gắn chặt với danh sách mã giao dịch nguồn gốc"})

    return results

if __name__ == "__main__":
    test_results = run_tests()
    all_passed = all(t["passed"] for t in test_results)
    print(f"Test Suite Summary: {sum(1 for t in test_results if t['passed'])}/{len(test_results)} passed.")
    for t in test_results:
        status = "PASSED" if t["passed"] else "FAILED"
        print(f"[{status}] {t['id']}: {t['name']} - {t['details']}")

