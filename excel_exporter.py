"""
Pure Python XLSX generator for:
1. 3-Sheet Workbook:
   Sheet 1: Du_Lieu_Da_Lam_Sach
   Sheet 2: Don_Vi_Can_Ho_Tro_Uu_Tien
   Sheet 3: Dashboard_Tong_Quan
2. Executive Business Performance Report Workbook:
   Sheet 1: Bao_Cao_Ket_Qua_Kinh_Doanh
No third-party packages required (uses zipfile + standard XML).
"""

import io
import zipfile
import html

def escape_xml(s):
    return html.escape(str(s))

def build_worksheet_xml(rows):
    """
    rows: list of lists (each element can be str, int, float)
    """
    out = []
    out.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    out.append('<sheetData>')
    
    for r_idx, row in enumerate(rows, 1):
        out.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, 1):
            col_letter = ""
            temp = c_idx
            while temp > 0:
                temp, remainder = divmod(temp - 1, 26)
                col_letter = chr(65 + remainder) + col_letter
            
            cell_ref = f"{col_letter}{r_idx}"
            
            if val is None:
                continue
            elif isinstance(val, (int, float)):
                out.append(f'<c r="{cell_ref}"><v>{val}</v></c>')
            else:
                esc = escape_xml(val)
                out.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{esc}</t></is></c>')
        out.append('</row>')
        
    out.append('</sheetData>')
    out.append('</worksheet>')
    return "".join(out).encode("utf-8")

def generate_2sheet_excel(cleaned_txs, store_metrics, category_metrics, sys_kpis):
    """
    Tạo file Excel gồm đúng 2 Sheets theo yêu cầu mới nhất:
    Sheet 1: Du_Lieu_Da_Lam_Sach
    Sheet 2: Bao_Cao_Tong_Quan
    """
    # Sheet 1: Dữ liệu đã làm sạch
    s1_rows = [
        ["MÃ GIAO DỊCH", "MÃ ĐƠN HÀNG / HỢP ĐỒNG", "ĐƠN VỊ / CHI NHÁNH / KÊNH", "MÃ SẢN PHẨM / DỊCH VỤ", "TÊN SẢN PHẨM / DỊCH VỤ", "PHÂN NHÓM / DANH MỤC", "NGÀY GIAO DỊCH", "LOẠI GIAO DỊCH", "SỐ LƯỢNG", "ĐƠN GIÁ (VNĐ)", "DOANH THU THUẦN (VNĐ)", "CHI PHÍ / GIÁ VỐN (VNĐ)", "LỢI NHUẬN GỘP (VNĐ)"]
    ]
    for tx in cleaned_txs[:3000]:
        s1_rows.append([
            tx.get("trans_id", ""),
            tx.get("order_id", ""),
            tx.get("store_id", ""),
            tx.get("product_id", ""),
            tx.get("product_name", ""),
            tx.get("category", ""),
            tx.get("date", ""),
            tx.get("type", ""),
            tx.get("qty", 0),
            tx.get("unit_price", 0),
            tx.get("net_revenue", 0),
            tx.get("net_cogs", 0),
            tx.get("net_profit", 0)
        ])

    # Sheet 2: Báo Cáo Tổng Quan (KPIs + Hiệu Suất Đơn Vị + Cơ Cấu Danh Mục)
    s2_rows = [
        ["NGỌC ÁNH DATA STUDIO - BÁO CÁO KẾT QUẢ KINH DOANH & HIỆU SUẤT VẬN HÀNH"],
        [""],
        ["I. TỔNG QUAN CHỈ SỐ TOÀN HỆ THỐNG", "GIÁ TRỊ (VNĐ / %)"],
        ["Tổng Doanh Thu Thuần", sys_kpis.get("total_net_revenue", 0)],
        ["Tổng Lợi Nhuận Gộp", sys_kpis.get("total_gross_profit", 0)],
        ["Tổng Chi Phí / Giá Vốn Bán Hàng", (sys_kpis.get("total_net_revenue", 0) - sys_kpis.get("total_gross_profit", 0))],
        ["Biên Lợi Nhuận Gộp Bình Quân", f"{sys_kpis.get('avg_margin_pct', 0)}%"],
        ["Tỷ Lệ Đổi Trả Thực Tế Hệ Thống", f"{sys_kpis.get('sys_return_rate_pct', 0)}%"],
        ["% Hoàn Thành Kế Hoạch Doanh Thu", f"{sys_kpis.get('rev_target_completion_pct', 0)}%"],
        ["% Hoàn Thành Kế Hoạch Lợi Nhuận", f"{sys_kpis.get('profit_target_completion_pct', 0)}%"],
        [""],
        ["II. BẢNG HIỆU SUẤT THEO ĐƠN VỊ / CHI NHÁNH / KÊNH PHÂN PHỐI"],
        ["MÃ ĐƠN VỊ", "TÊN ĐƠN VỊ / CHI NHÁNH", "KHU VỰC", "DOANH THU THUẦN", "MỤC TIÊU DOANH THU", "% ĐẠT KH DOANH THU", "LỢI NHUẬN GỘP", "MỤC TIÊU LỢI NHUẬN", "% ĐẠT KH LỢI NHUẬN", "BIÊN LN (%)", "TỶ LỆ ĐỔI TRẢ (%)", "TĂNG TRƯỞNG DT (%)", "ĐÁNH GIÁ TRẠNG THÁI"]
    ]
    for sm in store_metrics:
        rev_pct = sm.get('rev_target_pct', 100.0)
        status_eval = "Trụ cột hiệu quả cao" if rev_pct >= 95 else ("Ổn định" if rev_pct >= 85 else "Sụt giảm — Cần can thiệp")
        s2_rows.append([
            sm.get("store_id", ""),
            sm.get("store_name", ""),
            sm.get("region", ""),
            sm.get("net_revenue", 0),
            sm.get("rev_target", 0),
            f"{rev_pct}%",
            sm.get("net_profit", 0),
            sm.get("profit_target", 0),
            f"{sm.get('profit_target_pct', 0)}%",
            f"{sm.get('margin_pct', 0)}%",
            f"{sm.get('return_rate_pct', 0)}%",
            f"{sm.get('rev_growth_pct', 0)}%",
            status_eval
        ])

    s2_rows.append([""])
    s2_rows.append(["III. CƠ CẤU DOANH THU & LỢI NHUẬN THEO PHÂN NHÓM / NGÀNH HÀNG"])
    s2_rows.append(["PHÂN NHÓM / NGÀNH HÀNG", "DOANH THU THUẦN", "LỢI NHUẬN GỘP", "BIÊN LN (%)", "DOANH THU ĐỔI TRẢ", "SỐ LƯỢNG BÁN", "SỐ LƯỢNG TRẢ"])
    for cat_name, c in category_metrics.items():
        c_rev = c.get("net_revenue", 0)
        c_prof = c.get("net_profit", 0)
        c_margin = round(c_prof / c_rev * 100.0, 1) if c_rev > 0 else 0
        s2_rows.append([
            cat_name,
            c_rev,
            c_prof,
            f"{c_margin}%",
            c.get("return_revenue", 0),
            c.get("units_sold", 0),
            c.get("units_returned", 0)
        ])

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
        zf.writestr("[Content_Types].xml", content_types)

        rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
        zf.writestr("_rels/.rels", rels)

        wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"""
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)

        workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Du_Lieu_Da_Lam_Sach" sheetId="1" r:id="rId1"/>
    <sheet name="Bao_Cao_Tong_Quan" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""
        zf.writestr("xl/workbook.xml", workbook)

        zf.writestr("xl/worksheets/sheet1.xml", build_worksheet_xml(s1_rows))
        zf.writestr("xl/worksheets/sheet2.xml", build_worksheet_xml(s2_rows))

    buf.seek(0)
    return buf.getvalue()

def generate_3sheet_excel(cleaned_txs, top3_stores, store_metrics, sys_kpis):
    # Sheet 1: Dữ liệu đã làm sạch
    s1_rows = [
        ["MÃ GIAO DỊCH", "MÃ ĐƠN HÀNG", "ĐƠN VỊ / CHI NHÁNH", "MÃ SẢN PHẨM", "TÊN SẢN PHẨM", "NGÀNH HÀNG", "NGÀY GIAO DỊCH", "LOẠI", "SỐ LƯỢNG", "ĐƠN GIÁ (VNĐ)", "DOANH THU THUẦN (VNĐ)", "GIÁ VỐN THUẦN (VNĐ)", "LỢI NHUẬN GỘP (VNĐ)"]
    ]
    for tx in cleaned_txs[:1500]:
        s1_rows.append([
            tx.get("trans_id", ""),
            tx.get("order_id", ""),
            tx.get("store_id", ""),
            tx.get("product_id", ""),
            tx.get("product_name", ""),
            tx.get("category", ""),
            tx.get("date", ""),
            tx.get("type", ""),
            tx.get("qty", 0),
            tx.get("unit_price", 0),
            tx.get("net_revenue", 0),
            tx.get("net_cogs", 0),
            tx.get("net_profit", 0)
        ])

    # Sheet 2: Danh sách đơn vị / cửa hàng cần hỗ trợ ưu tiên
    s2_rows = [
        ["BÁO CÁO ĐƠN VỊ / CỬA HÀNG CẦN HỖ TRỢ ƯU TIÊN - NGỌC ÁNH DATA STUDIO"],
        [""],
        ["XẾP HẠNG", "MÃ ĐƠN VỊ", "TÊN ĐƠN VỊ / CHI NHÁNH", "KHU VỰC", "QUẢN LÝ", "ĐIỂM RỦI RO", "DOANH THU THUẦN (VNĐ)", "% ĐẠT KẾ HOẠCH DT", "LỢI NHUẬN GỘP (VNĐ)", "% ĐẠT KẾ HOẠCH LN", "BIÊN LN GỘP (%)", "TỶ LỆ ĐỔI TRẢ (%)", "TĂNG TRƯỞNG DT (%)", "NGUYÊN NHÂN GỐC RỄ (ROOT CAUSES)", "HÀNH ĐỘNG CẢI THIỆN ĐỀ XUẤT"]
    ]
    for st in top3_stores:
        causes = " | ".join(st.get("root_causes", []))
        recomms = " | ".join(st.get("recommendations", []))
        s2_rows.append([
            f"Ưu tiên #{st.get('rank', 1)}",
            st.get("store_id", ""),
            st.get("store_name", ""),
            st.get("region", ""),
            st.get("manager", ""),
            f"{st.get('risk_score', 0)}/100",
            st.get("net_revenue", 0),
            f"{st.get('rev_target_pct', 0)}%",
            st.get("net_profit", 0),
            f"{st.get('profit_target_pct', 0)}%",
            f"{st.get('margin_pct', 0)}%",
            f"{st.get('return_rate_pct', 0)}%",
            f"{st.get('rev_growth_pct', 0)}%",
            causes,
            recomms
        ])

    # Sheet 3: Dashboard Tổng Quan
    s3_rows = [
        ["NGỌC ÁNH DATA STUDIO - DASHBOARD CHỈ SỐ KINH DOANH TỔNG QUAN"],
        [""],
        ["CHỈ SỐ TOÀN HỆ THỐNG", "GIÁ TRỊ"],
        ["Tổng Doanh Thu Thuần", sys_kpis.get("total_net_revenue", 0)],
        ["Tổng Lợi Nhuận Gộp", sys_kpis.get("total_gross_profit", 0)],
        ["Biên Lợi Nhuận Gộp Trung Bình", f"{sys_kpis.get('avg_margin_pct', 0)}%"],
        ["Tỷ Lệ Đổi Trả Bình Quân Hệ Thống", f"{sys_kpis.get('sys_return_rate_pct', 0)}%"],
        [""],
        ["BẢNG CHỈ SỐ HIỆU QUẢ CÁC ĐƠN VỊ / CHI NHÁNH"],
        ["MÃ ĐƠN VỊ", "TÊN CHI NHÁNH / ĐƠN VỊ", "KHU VỰC", "DOANH THU THUẦN", "MỤC TIÊU DT", "% HOÀN THÀNH DT", "LỢI NHUẬN GỘP", "MỤC TIÊU LN", "% HOÀN THÀNH LN", "BIÊN LN (%)", "TỶ LỆ TRẢ (%)"]
    ]
    for sm in store_metrics:
        s3_rows.append([
            sm.get("store_id", ""),
            sm.get("store_name", ""),
            sm.get("region", ""),
            sm.get("net_revenue", 0),
            sm.get("rev_target", 0),
            f"{sm.get('rev_target_pct', 0)}%",
            sm.get("net_profit", 0),
            sm.get("profit_target", 0),
            f"{sm.get('profit_target_pct', 0)}%",
            f"{sm.get('margin_pct', 0)}%",
            f"{sm.get('return_rate_pct', 0)}%"
        ])

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
        zf.writestr("[Content_Types].xml", content_types)

        rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
        zf.writestr("_rels/.rels", rels)

        wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>"""
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)

        workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Du_Lieu_Da_Lam_Sach" sheetId="1" r:id="rId1"/>
    <sheet name="Don_Vi_Can_Ho_Tro_Uu_Tien" sheetId="2" r:id="rId2"/>
    <sheet name="Dashboard_Tong_Quan" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"""
        zf.writestr("xl/workbook.xml", workbook)

        zf.writestr("xl/worksheets/sheet1.xml", build_worksheet_xml(s1_rows))
        zf.writestr("xl/worksheets/sheet2.xml", build_worksheet_xml(s2_rows))
        zf.writestr("xl/worksheets/sheet3.xml", build_worksheet_xml(s3_rows))

    buf.seek(0)
    return buf.getvalue()

def generate_executive_report_excel(sys_kpis, store_metrics, category_metrics, top3_stores):
    """
    Tạo tệp Báo Cáo Kết Quả Kinh Doanh Doanh Nghiệp (Executive Report)
    mô phỏng đúng mẫu Dashboard MasterMOS trong ảnh người dùng cung cấp.
    """
    total_net_rev = sys_kpis.get("total_net_revenue", 0)
    total_profit = sys_kpis.get("total_gross_profit", 0)
    total_cost = total_net_rev - total_profit
    margin_pct = sys_kpis.get("avg_margin_pct", 0)
    return_rate = sys_kpis.get("sys_return_rate_pct", 0)
    rev_target_pct = sys_kpis.get("rev_target_completion_pct", 0)

    rows = [
        ["NGỌC ÁNH DATA STUDIO - BÁO CÁO KẾT QUẢ KINH DOANH DOANH NGHIỆP"],
        ["ĐƠN VỊ TÍNH: TRIỆU VNĐ / VNĐ", f"KỲ ĐỐI CHIẾU: THÁNG HIỆN HÀNH"],
        [""],
        ["=== KHỐI CHỈ SỐ KPI CHÍNH (KEY PERFORMANCE INDICATORS) ==="],
        ["DOANH THU BÁN HÀNG", "CHI PHÍ GIÁ VỐN", "LỢI NHUẬN BÁN HÀNG", "% HOÀN THÀNH MỤC TIÊU", "BIÊN LỢI NHUẬN GỘP", "TỶ LỆ ĐỔI TRẢ"],
        [
            total_net_rev,
            total_cost,
            total_profit,
            f"{rev_target_pct}%",
            f"{margin_pct}%",
            f"{return_rate}%"
        ],
        [""],
        ["=== 1. HIỆU QUẢ KINH DOANH THEO CHI NHÁNH / ĐƠN VỊ ==="],
        ["MÃ ĐƠN VỊ", "TÊN ĐƠN VỊ / CHI NHÁNH", "KHU VỰC", "DOANH THU THUẦN", "MỤC TIÊU DT", "% ĐẠT MỤC TIÊU", "LỢI NHUẬN GỘP", "BIÊN LN GỘP (%)", "TỶ LỆ ĐỔI TRẢ (%)", "TĂNG TRƯỞNG DT (%)"]
    ]

    for sm in store_metrics:
        rows.append([
            sm.get("store_id", ""),
            sm.get("store_name", ""),
            sm.get("region", ""),
            sm.get("net_revenue", 0),
            sm.get("rev_target", 0),
            f"{sm.get('rev_target_pct', 0)}%",
            sm.get("net_profit", 0),
            f"{sm.get('margin_pct', 0)}%",
            f"{sm.get('return_rate_pct', 0)}%",
            f"{sm.get('rev_growth_pct', 0)}%"
        ])

    rows.extend([
        [""],
        ["=== 2. CƠ CẤU DOANH THU & ĐỔI TRẢ THEO DANH MỤC / NGÀNH HÀNG ==="],
        ["DANH MỤC / NGÀNH HÀNG", "DOANH THU THUẦN", "LỢI NHUẬN GỘP", "BIÊN LN GỘP (%)", "DOANH THU ĐỔI TRẢ", "SỐ LƯỢNG BÁN", "SỐ LƯỢNG TRẢ", "TỶ LỆ TRẢ THEO SL (%)"]
    ])

    for cat, c in category_metrics.items():
        cat_margin = round((c.get("net_profit", 0) / c.get("net_revenue", 1) * 100.0), 1) if c.get("net_revenue", 0) > 0 else 0
        sold = c.get("units_sold", 0)
        ret = c.get("units_returned", 0)
        ret_rate = round((ret / sold * 100.0), 1) if sold > 0 else 0
        rows.append([
            cat,
            c.get("net_revenue", 0),
            c.get("net_profit", 0),
            f"{cat_margin}%",
            c.get("return_revenue", 0),
            sold,
            ret,
            f"{ret_rate}%"
        ])

    rows.extend([
        [""],
        ["=== 3. CÁC ĐƠN VỊ / CHI NHÁNH CẦN ƯU TIÊN HỖ TRỢ & HÀNH ĐỘNG CẢI THIỆN ==="],
        ["XẾP HẠNG", "MÃ ĐƠN VỊ", "TÊN ĐƠN VỊ", "ĐIỂM RỦI RO", "CĂN CỨ ĐỊNH LƯỢNG", "HÀNH ĐỘNG CẢI THIỆN ĐỀ XUẤT"]
    ])

    for st in top3_stores:
        causes = " | ".join(st.get("root_causes", []))
        recomms = " | ".join(st.get("recommendations", []))
        rows.append([
            f"Ưu tiên #{st.get('rank', 1)}",
            st.get("store_id", ""),
            st.get("store_name", ""),
            f"{st.get('risk_score', 0)}/100",
            causes,
            recomms
        ])

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
        zf.writestr("[Content_Types].xml", content_types)

        rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
        zf.writestr("_rels/.rels", rels)

        wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)

        workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Bao_Cao_Ket_Qua_Kinh_Doanh" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
        zf.writestr("xl/workbook.xml", workbook)

        zf.writestr("xl/worksheets/sheet1.xml", build_worksheet_xml(rows))

    buf.seek(0)
    return buf.getvalue()
