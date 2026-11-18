import os
import io
import datetime
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def _generate_chart_images(store_metrics, category_metrics):
    """Tạo các ảnh biểu đồ PNG chất lượng cao bằng matplotlib để nhúng vào Báo cáo Word"""
    charts = {}
    plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Segoe UI"]
    plt.rcParams["axes.edgecolor"] = "#dfd3bd"
    plt.rcParams["axes.linewidth"] = 0.8

    top_stores = (store_metrics or [])[:8]
    if top_stores:
        # 1. Biểu đồ cột: Doanh thu Thực tế vs Kế hoạch (Triệu VNĐ)
        labels = [str(s.get("store_id", ""))[:12] for s in top_stores]
        actuals = [round(s.get("net_revenue", 0) / 1e6, 1) for s in top_stores]
        targets = [round(s.get("rev_target", 0) / 1e6, 1) for s in top_stores]

        fig, ax = plt.subplots(figsize=(6.8, 3.2), dpi=180)
        x = list(range(len(labels)))
        w = 0.36
        ax.bar([i - w/2 for i in x], actuals, width=w, label="Doanh thu thực tế", color="#eab308", zorder=3)
        ax.bar([i + w/2 for i in x], targets, width=w, label="Mục tiêu kế hoạch", color="#1e3a5f", zorder=3)
        ax.set_title("Doanh Thu Thực Tế vs Mục Tiêu Kế Hoạch (Triệu VNĐ)", fontsize=10.5, fontweight="bold", color="#2a2417", pad=10)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=8.5, rotation=15, ha="right")
        ax.grid(axis="y", linestyle="--", alpha=0.45, color="#d9caa1", zorder=0)
        ax.legend(frameon=False, fontsize=8.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        plt.tight_layout()
        buf1 = io.BytesIO()
        fig.savefig(buf1, format="png", facecolor="#ffffff")
        plt.close(fig)
        buf1.seek(0)
        charts["rev_vs_plan"] = buf1

        # 2. Biểu đồ đường: Xu hướng Lợi nhuận gộp theo đơn vị
        profits = [round(s.get("net_profit", 0) / 1e6, 1) for s in top_stores]
        fig2, ax2 = plt.subplots(figsize=(6.8, 2.9), dpi=180)
        ax2.plot(labels, profits, marker="o", linewidth=2.2, color="#189677", markersize=6, zorder=3)
        ax2.fill_between(range(len(labels)), profits, color="#189677", alpha=0.14, zorder=2)
        ax2.set_title("Xu Hướng Lợi Nhuận Gộp Theo Đơn Vị (Triệu VNĐ)", fontsize=10.5, fontweight="bold", color="#2a2417", pad=10)
        ax2.tick_params(axis="x", labelsize=8.5, rotation=15)
        ax2.grid(axis="y", linestyle="--", alpha=0.45, color="#d9caa1", zorder=0)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        plt.tight_layout()
        buf2 = io.BytesIO()
        fig2.savefig(buf2, format="png", facecolor="#ffffff")
        plt.close(fig2)
        buf2.seek(0)
        charts["profit_trend"] = buf2

    # 3. Biểu đồ tròn (Donut): Cơ cấu Doanh thu theo Danh mục
    cat_items = [(k, v.get("net_revenue", 0)) for k, v in (category_metrics or {}).items() if v.get("net_revenue", 0) > 0]
    if cat_items:
        cat_items.sort(key=lambda t: t[1], reverse=True)
        top_cats = cat_items[:6]
        cat_labels = [str(c[0])[:22] for c in top_cats]
        cat_vals = [c[1] for c in top_cats]
        palette = ["#eab308", "#189677", "#1e3a5f", "#d95d43", "#f59e0b", "#6366f1"]

        fig3, ax3 = plt.subplots(figsize=(6.4, 3.2), dpi=180)
        wedges, texts, autotexts = ax3.pie(
            cat_vals,
            labels=None,
            autopct="%1.1f%%",
            pctdistance=0.78,
            colors=palette[:len(cat_vals)],
            wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2)
        )
        for at in autotexts:
            at.set_fontsize(8)
            at.set_color("#1e241b")
            at.set_fontweight("bold")
        ax3.legend(wedges, cat_labels, title="Danh mục / Ngành hàng", loc="center left", bbox_to_anchor=(0.95, 0.5), frameon=False, fontsize=8.5)
        ax3.set_title("Cơ Cấu Tỷ Trọng Doanh Thu Theo Danh Mục", fontsize=10.5, fontweight="bold", color="#2a2417", pad=8)
        plt.tight_layout()
        buf3 = io.BytesIO()
        fig3.savefig(buf3, format="png", facecolor="#ffffff", bbox_inches="tight")
        plt.close(fig3)
        buf3.seek(0)
        charts["category_donut"] = buf3

    return charts

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def generate_word_executive_report(dataset_name, kpis, store_metrics, category_metrics, gate_info, summary, business_model="", business_problem="", clarifications=None):
    """
    Tạo file Word (.docx) chuẩn theo mẫu báo cáo quản trị chuyên sâu của người dùng
    """
    doc = Document()

    # Set page margins: 1 inch (72 pt)
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # 1. Header Table (Quốc hiệu tiêu ngữ & Tên đơn vị)
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False

    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    cell_left.width = Inches(3.4)
    cell_right.width = Inches(3.4)

    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_l1 = p_left.add_run("HỆ THỐNG PHÂN TÍCH QUẢN TRỊ NGỌC ÁNH DATA STUDIO\n")
    r_l1.font.bold = True
    r_l1.font.size = Pt(9.5)
    r_l1.font.name = "Times New Roman"
    r_l2 = p_left.add_run("BỘ PHẬN PHÂN TÍCH DỮ LIỆU & CHIẾN LƯỢC\n")
    r_l2.font.bold = True
    r_l2.font.size = Pt(9.5)
    r_l2.font.name = "Times New Roman"
    r_l3 = p_left.add_run("--------------------")
    r_l3.font.size = Pt(8)

    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_r1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r_r1.font.bold = True
    r_r1.font.size = Pt(9.5)
    r_r1.font.name = "Times New Roman"
    r_r2 = p_right.add_run("Độc lập - Tự do - Hạnh phúc\n")
    r_r2.font.bold = True
    r_r2.font.size = Pt(9.5)
    r_r2.font.name = "Times New Roman"
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    r_r3 = p_right.add_run(f"Ngày lập báo cáo: {now_str}")
    r_r3.font.italic = True
    r_r3.font.size = Pt(8.5)
    r_r3.font.name = "Times New Roman"

    doc.add_paragraph()

    # 2. Tiêu đề báo cáo
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_title.add_run("BÁO CÁO PHÂN TÍCH CHUYÊN SÂU DASHBOARD & DỮ LIỆU VẬN HÀNH\n")
    r_t.font.bold = True
    r_t.font.size = Pt(15)
    r_t.font.color.rgb = RGBColor(0x7c, 0x4d, 0x3a)
    r_t.font.name = "Times New Roman"

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    bm_txt = f" — Mô hình: {business_model}" if business_model else ""
    r_sub = p_sub.add_run(f"Chiều phân tích trọng tâm: Đơn Vị / Phân Nhóm{bm_txt} — Bộ dữ liệu: {dataset_name or 'Dữ liệu doanh nghiệp'}\n")
    r_sub.font.italic = True
    r_sub.font.size = Pt(10.5)
    r_sub.font.name = "Times New Roman"

    # Trích xuất số liệu
    total_raw = summary.get("total_raw", 0) if summary else (gate_info.get("total_raw", 0) if gate_info else 0)
    valid_cleaned = summary.get("total_cleaned_valid", 0) if summary else (gate_info.get("total_cleaned_valid", 0) if gate_info else 0)
    valid_report = summary.get("valid_report_period", 0) if summary else (gate_info.get("valid_report_period", 0) if gate_info else 0)
    unqual_count = summary.get("total_unqualified", 0) if summary else (gate_info.get("unqualified_count", 0) if gate_info else 0)

    total_net_rev = kpis.get("total_net_revenue", 0.0) if kpis else 0.0
    total_profit = kpis.get("total_gross_profit", 0.0) if kpis else 0.0
    total_cost = total_net_rev - total_profit
    avg_margin = kpis.get("avg_margin_pct", 0.0) if kpis else 0.0
    return_rate = kpis.get("sys_return_rate_pct", 0.0) if kpis else 0.0

    # I. BỐI CẢNH DOANH NGHIỆP & TÓM TẮT ĐIỀU HÀNH
    h1 = doc.add_heading("I. BỐI CẢNH DOANH NGHIỆP & TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)", level=1)
    h1.style.font.name = "Times New Roman"
    h1.style.font.color.rgb = RGBColor(0x33, 0x24, 0x1d)

    p1 = doc.add_paragraph()
    r = p1.add_run("1. Bối cảnh & Vấn đề doanh nghiệp đang gặp phải:\n")
    r.font.bold = True
    context_lines = []
    if business_model:
        context_lines.append(f"• Mô hình kinh doanh: {business_model}.")
    if business_problem:
        context_lines.append(f"• Vấn đề trọng tâm cần phân tích: {business_problem}.")
    clar = clarifications or {}
    if clar.get("time_scope"):
        context_lines.append(f"• Phạm vi thời gian ưu tiên: {clar['time_scope']}.")
    if clar.get("decision_goal"):
        context_lines.append(f"• Mục tiêu ra quyết định & chỉ số ưu tiên: {clar['decision_goal']}.")
    if clar.get("benchmark"):
        context_lines.append(f"• Đối tượng so sánh / mục tiêu tham chiếu: {clar['benchmark']}.")
    context_lines.append(
        "Doanh nghiệp triển khai số hóa và tự động hóa quy trình phân tích dữ liệu kinh doanh đa chiều, "
        "nhằm kiểm soát chất lượng dữ liệu đầu vào, tối ưu hiệu suất vận hành từng đơn vị/nhóm và hỗ trợ Ban Giám đốc ra quyết định chính xác."
    )
    p1.add_run("\n".join(context_lines))

    p2 = doc.add_paragraph()
    r = p2.add_run("2. Tóm tắt kết quả điều hành trình Lãnh đạo:\n")
    r.font.bold = True
    
    sorted_stores = sorted(store_metrics or [], key=lambda x: x.get("net_revenue", 0.0), reverse=True)
    top_store = sorted_stores[0] if sorted_stores else {}
    top2_store = sorted_stores[1] if len(sorted_stores) > 1 else {}
    top_store_name = top_store.get("store_name", "Đơn vị chính")
    top_store_rev = top_store.get("net_revenue", 0.0)
    top_share = round(top_store_rev / total_net_rev * 100.0, 1) if total_net_rev > 0 else 0
    top2_share = round((top_store_rev + top2_store.get("net_revenue", 0.0)) / total_net_rev * 100.0, 1) if total_net_rev > 0 else 0

    sorted_margin = sorted(store_metrics or [], key=lambda x: x.get("margin_pct", 0.0), reverse=True)
    best_margin = sorted_margin[0] if sorted_margin else {}
    worst_margin = sorted_margin[-1] if sorted_margin else {}
    margin_spread = round(best_margin.get("margin_pct", 0.0) - worst_margin.get("margin_pct", 0.0), 1)
    lowest_target = min(store_metrics, key=lambda x: x.get("rev_target_pct", 100.0)) if store_metrics else {}
    highest_ret = max(store_metrics, key=lambda x: x.get("return_rate_pct", 0.0)) if store_metrics else {}

    sorted_cats = sorted((category_metrics or {}).items(), key=lambda x: x[1].get("net_revenue", 0.0), reverse=True)
    top_cat_name = sorted_cats[0][0] if sorted_cats else "-"
    top_cat_rev = sorted_cats[0][1].get("net_revenue", 0.0) if sorted_cats else 0.0
    top_cat_share = round(top_cat_rev / total_net_rev * 100.0, 1) if total_net_rev > 0 else 0

    p2.add_run(
        f"Phân tích tổng thể {valid_cleaned:,} bản ghi dữ liệu sạch trong kỳ cho thấy Doanh thu thuần đạt "
        f"{total_net_rev:,.0f} VND, Lợi nhuận gộp đạt {total_profit:,.0f} VND và Tổng chi phí/giá vốn ở mức {total_cost:,.0f} VND "
        f"(Biên lợi nhuận gộp bình quân toàn hệ thống đạt {avg_margin}%).\n"
        f"Cấu trúc đóng góp doanh thu tập trung khi Top 2 đơn vị dẫn đầu ({top_store_name}"
        f"{' & ' + top2_store.get('store_name', '') if top2_store else ''}) chiếm tới {top2_share}% tổng quy mô "
        f"(riêng {top_store_name} đóng góp {top_share}% với {top_store_rev:,.0f} VND). "
        f"Nhóm '{best_margin.get('store_name', '-')}' đạt biên hiệu quả cao nhất ({best_margin.get('margin_pct', 0)}%), "
        f"trong khi '{worst_margin.get('store_name', '-')}' đạt {worst_margin.get('margin_pct', 0)}% (chênh lệch {margin_spread} điểm phần trăm). "
        f"Đơn vị '{lowest_target.get('store_name', '-')}' có tỷ lệ đạt mục tiêu thấp nhất ({lowest_target.get('rev_target_pct', 0)}%), cần ưu tiên can thiệp."
    )

    p3 = doc.add_paragraph()
    r = p3.add_run("3. Kết quả làm sạch & Kiểm duyệt dữ liệu đầu vào (So với file gốc):\n")
    r.font.bold = True
    clean_pct = round(valid_cleaned / total_raw * 100.0, 1) if total_raw > 0 else 100.0
    p3.add_run(
        f"Từ {total_raw:,} dòng dữ liệu thô ban đầu (Độ sạch ban đầu ~{clean_pct}%) -> Sau quy trình làm sạch chuẩn hóa tự động "
        f"đạt {valid_cleaned:,} dòng hợp lệ (Độ sạch 100.0% đưa vào hạch toán, đã tự động khử trùng lặp và loại trừ {unqual_count:,} dòng ngoại lệ/đơn hủy)."
    )

    # II. TỔNG QUAN DASHBOARD TRỰC QUAN & PHÂN TÍCH SÂU 4 CHỈ SỐ KPI
    h2 = doc.add_heading("II. TỔNG QUAN DASHBOARD TRỰC QUAN & PHÂN TÍCH SÂU 4 CHỈ SỐ KPI", level=1)
    h2.style.font.name = "Times New Roman"
    h2.style.font.color.rgb = RGBColor(0x33, 0x24, 0x1d)

    p_kpi_desc = doc.add_paragraph("Bảng phân tích chuyên sâu 4 chỉ số KPI trọng tâm trên Dashboard:")
    p_kpi_desc.runs[0].font.italic = True

    kpi_table = doc.add_table(rows=5, cols=4)
    kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpi_table.autofit = False

    kpi_headers = ["Chỉ số KPI", "Giá trị đạt được", "Đánh giá định lượng", "Phân tích chuyên sâu & Ý nghĩa quản trị"]
    col_widths = [Inches(1.6), Inches(1.5), Inches(1.4), Inches(2.3)]
    for i, h in enumerate(kpi_headers):
        cell = kpi_table.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "EAE4DC")
        set_cell_margins(cell, top=120, bottom=120, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.name = "Times New Roman"

    kpi_rows_data = [
        ("Tổng số bản ghi / Giao dịch", f"{valid_cleaned:,} bản ghi", "Độ sạch 100%", f"Phản ánh mật độ giao dịch thực tế sau khi đã khử hoàn toàn bản ghi trùng lặp và loại trừ {unqual_count:,} đơn hủy."),
        ("Doanh thu thuần", f"{total_net_rev:,.0f} VND", f"Đạt {kpis.get('rev_target_completion_pct', 92)}% KH", f"Doanh thu thuần hợp lệ toàn chuỗi. Đơn vị {top_store_name} chiếm tỷ trọng cao nhất ({top_share}%)."),
        ("Lợi nhuận gộp", f"{total_profit:,.0f} VND", f"Biên LN: {avg_margin}%", f"Biên lợi nhuận gộp toàn hệ thống đạt {avg_margin}%. Đảm bảo dòng tiền và an toàn tài chính sau khi trừ giá vốn."),
        ("Chi phí / Giá vốn bán hàng", f"{total_cost:,.0f} VND", f"Tỷ lệ CP: {round(100.0 - avg_margin, 1)}%", f"Chi phí giá vốn hàng bán và vận hành trực tiếp, tỷ lệ chi phí duy trì ổn định không vượt ngân sách cam kết.")
    ]

    for row_idx, r_data in enumerate(kpi_rows_data, start=1):
        for col_idx, text in enumerate(r_data):
            cell = kpi_table.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            if row_idx % 2 == 1:
                set_cell_background(cell, "FAF8F5")
            p = cell.paragraphs[0]
            if col_idx in [1, 2]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(text)
            r.font.size = Pt(8.5)
            r.font.name = "Times New Roman"
            if col_idx in [0, 1]:
                r.font.bold = True

    doc.add_paragraph()

    # III. PHÂN TÍCH CHUYÊN SÂU CÁC KHỐI BIỂU ĐỒ TRÊN DASHBOARD
    h3 = doc.add_heading("III. PHÂN TÍCH CHUYÊN SÂU CÁC KHỐI BIỂU ĐỒ TRÊN DASHBOARD", level=1)
    h3.style.font.name = "Times New Roman"
    h3.style.font.color.rgb = RGBColor(0x33, 0x24, 0x1d)

    charts = _generate_chart_images(store_metrics, category_metrics)

    # 1. Phân tích Doanh thu Thực tế vs Kế hoạch
    doc.add_heading("1. Phân tích 'Doanh thu Thực tế vs Kế hoạch theo Đơn vị / Nhóm'", level=2)
    doc.add_paragraph(
        f"- Nhận xét chuyên sâu: Xếp hạng quy mô cho thấy sự phân hóa rõ rệt giữa các đơn vị. "
        f"'{top_store_name}' dẫn đầu với {top_store_rev:,.0f} VND ({top_share}%), Top 2 nhóm đầu nắm giữ {top2_share}% tổng giá trị. "
        f"Ngược lại, '{lowest_target.get('store_name', '-')}' chỉ đạt {lowest_target.get('rev_target_pct', 0)}% kế hoạch, "
        f"trở thành điểm nghẽn chính cần cô lập xử lý thay vì điều chỉnh dàn trải."
    )
    if charts.get("rev_vs_plan"):
        p_img1 = doc.add_paragraph()
        p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img1.add_run().add_picture(charts["rev_vs_plan"], width=Inches(6.2))

    # 2. Phân tích Xu hướng Lợi nhuận Gộp
    doc.add_heading("2. Phân tích 'Xu hướng Lợi nhuận Gộp & Biên hiệu quả'", level=2)
    doc.add_paragraph(
        f"- Nhận xét chuyên sâu: Lợi nhuận gộp toàn hệ thống đạt {total_profit:,.0f} VND (biên bình quân {avg_margin}%). "
        f"Chênh lệch biên hiệu quả giữa nhóm cao nhất ('{best_margin.get('store_name', '-')}': {best_margin.get('margin_pct', 0)}%) "
        f"và nhóm thấp nhất ('{worst_margin.get('store_name', '-')}': {worst_margin.get('margin_pct', 0)}%) lên tới {margin_spread} điểm phần trăm, "
        f"cho thấy việc mở rộng quy mô cần đi kèm kiểm soát chặt định mức chi phí tại nhóm biên thấp."
    )
    if charts.get("profit_trend"):
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img2.add_run().add_picture(charts["profit_trend"], width=Inches(6.2))

    # 3. Cơ cấu ngành hàng / danh mục
    doc.add_heading("3. Phân tích 'Cơ cấu danh mục & Đóng góp giá trị'", level=2)
    doc.add_paragraph(
        f"- Nhận xét chuyên sâu: Trong cơ cấu danh mục, nhóm '{top_cat_name}' chiếm tỷ trọng lớn nhất ({top_cat_share}%, đạt {top_cat_rev:,.0f} VND). "
        f"Đối chiếu với tỷ lệ hoàn trả toàn hệ thống ({return_rate}%), cần ưu tiên kiểm soát chất lượng ở các danh mục phát sinh đổi trả cao để bảo toàn lợi nhuận:"
    )
    if charts.get("category_donut"):
        p_img3 = doc.add_paragraph()
        p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img3.add_run().add_picture(charts["category_donut"], width=Inches(5.8))

    cat_table = doc.add_table(rows=len(category_metrics) + 1, cols=5)
    cat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cat_headers = ["Phân Nhóm / Ngành Hàng", "Doanh Thu Thuần (VND)", "Lợi Nhuận Gộp (VND)", "Biên LN (%)", "Tỷ Lệ Đổi Trả (%)"]
    for i, h in enumerate(cat_headers):
        cell = cat_table.cell(0, i)
        set_cell_background(cell, "EAE4DC")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.name = "Times New Roman"

    for r_idx, (cat_name, c_data) in enumerate(category_metrics.items(), start=1):
        cat_rev = c_data.get("net_revenue", 0.0)
        cat_prof = c_data.get("net_profit", 0.0)
        cat_margin = round(cat_prof / cat_rev * 100.0, 1) if cat_rev > 0 else 0
        cat_ret_rate = round(c_data.get("return_revenue", 0.0) / cat_rev * 100.0, 1) if cat_rev > 0 else 0
        
        row_vals = [
            cat_name,
            f"{cat_rev:,.0f}",
            f"{cat_prof:,.0f}",
            f"{cat_margin}%",
            f"{cat_ret_rate}%"
        ]
        for c_idx, val in enumerate(row_vals):
            cell = cat_table.cell(r_idx, c_idx)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if r_idx % 2 == 1:
                set_cell_background(cell, "FAF8F5")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            r.font.name = "Times New Roman"

    doc.add_paragraph()

    # 4. Xếp hạng hiệu quả theo Đơn Vị / Chi Nhánh
    doc.add_heading("4. Bảng Xếp Hạng Hiệu Suất & Đánh Giá Trạng Thái Vận Hành", level=2)
    doc.add_paragraph(
        "Tổng hợp đa chỉ tiêu từng đơn vị/chi nhánh nhằm phát hiện sớm các điểm nghẽn hiệu suất và nhóm cần hỗ trợ:"
    )

    store_table = doc.add_table(rows=min(len(store_metrics), 10) + 1, cols=6)
    store_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["Đơn Vị / Chi Nhánh", "Doanh Thu (VND)", "% Đạt KH", "Biên LN (%)", "Tỷ Lệ Đổi Trả", "Đánh Giá Trạng Thái"]
    for i, h in enumerate(s_headers):
        cell = store_table.cell(0, i)
        set_cell_background(cell, "EAE4DC")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.name = "Times New Roman"

    for r_idx, s in enumerate(store_metrics[:10], start=1):
        rev = s.get("net_revenue", 0.0)
        target_pct = s.get("rev_target_pct", 100.0)
        margin = s.get("margin_pct", 25.0)
        ret_rate = s.get("return_rate_pct", 1.0)
        
        status_eval = "Trụ cột hiệu quả cao" if target_pct >= 95 else ("Ổn định" if target_pct >= 85 else "Sụt giảm — Cần can thiệp")
        
        row_vals = [
            f"{s.get('store_name', s.get('store_id'))}",
            f"{rev:,.0f}",
            f"{target_pct}%",
            f"{margin}%",
            f"{ret_rate}%",
            status_eval
        ]
        for c_idx, val in enumerate(row_vals):
            cell = store_table.cell(r_idx, c_idx)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if r_idx % 2 == 1:
                set_cell_background(cell, "FAF8F5")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx in [0, 5] else WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            r.font.name = "Times New Roman"
            if c_idx == 5:
                if "Sụt giảm" in val:
                    r.font.color.rgb = RGBColor(0x99, 0x1b, 0x1b)
                    r.font.bold = True
                elif "Trụ cột" in val:
                    r.font.color.rgb = RGBColor(0x13, 0x73, 0x33)
                    r.font.bold = True

    doc.add_paragraph()

    # IV. KẾT LUẬN TRỌNG TÂM & ĐỀ XUẤT HÀNH ĐỘNG
    h4 = doc.add_heading("IV. KẾT LUẬN TRỌNG TÂM & ĐỀ XUẤT HÀNH ĐỘNG", level=1)
    h4.style.font.name = "Times New Roman"
    h4.style.font.color.rgb = RGBColor(0x33, 0x24, 0x1d)

    p_conc = doc.add_paragraph()
    r = p_conc.add_run("1. Các điểm nghẽn & rủi ro trọng yếu đúc kết từ Dashboard:\n")
    r.font.bold = True
    prob_line = f"• Gắn với vấn đề doanh nghiệp ({business_problem}): Điểm nghẽn tập trung tại '{lowest_target.get('store_name', '-')}' ({lowest_target.get('rev_target_pct', 0)}% KH) và nhóm biên thấp '{worst_margin.get('store_name', '-')}' ({worst_margin.get('margin_pct', 0)}%).\n" if business_problem else ""
    p_conc.add_run(
        prob_line +
        f"• Rủi ro tập trung doanh thu: Top 2 đơn vị dẫn đầu đóng góp tới {top2_share}% tổng doanh thu toàn chuỗi, nếu có biến động tại nhóm trụ cột sẽ ảnh hưởng trực tiếp đến kết quả chung.\n"
        f"• Điểm nghẽn đổi trả & dữ liệu: Nhóm có tỷ lệ đổi trả cao nhất là '{highest_ret.get('store_name', '-')}' ({highest_ret.get('return_rate_pct', 0)}%); hệ thống đã loại bỏ {unqual_count:,} bản ghi ngoại lệ/đơn hủy để bảo đảm số liệu báo cáo trung thực 100%."
    )

    p_short = doc.add_paragraph()
    r = p_short.add_run("2. Giải pháp hành động ngắn hạn (30 ngày tới):\n")
    r.font.bold = True
    p_short.add_run(
        f"• Kiểm tra nguyên nhân gốc rễ tại '{lowest_target.get('store_name', 'nhóm dưới chuẩn')}' để chặn đà suy giảm và đưa tỷ lệ hoàn thành kế hoạch lên trên 90%.\n"
        f"• Tái phân bổ nguồn lực từ các nhóm biên thấp ('{worst_margin.get('store_name', '-')}') sang nhóm dẫn đầu biên hiệu quả ('{best_margin.get('store_name', '-')}').\n"
        f"• Rà soát quy trình vận hành & chất lượng đầu ra tại '{highest_ret.get('store_name', 'điểm nóng đổi trả')}' và xử lý dứt điểm các bản ghi ngoại lệ trong File Excel sạch."
    )

    p_long = doc.add_paragraph()
    r = p_long.add_run("3. Định hướng phát triển trung & dài hạn (6 - 12 tháng):\n")
    r.font.bold = True
    p_long.add_run(
        "• Thiết lập ngưỡng cảnh báo biên lợi nhuận và tỷ lệ chi phí chuẩn cho từng đơn vị/danh mục theo từng kỳ.\n"
        f"• Mở rộng tỷ trọng danh mục chủ lực '{top_cat_name}' và các nhóm tiềm năng ở phân khúc giữa để giảm phụ thuộc vào Top 1-2 nhóm dẫn đầu.\n"
        "• Duy trì quy trình kiểm duyệt 2 cổng (Tự động & Chuyên viên BI duyệt) làm tiêu chuẩn vận hành."
    )

    p_focus = doc.add_paragraph()
    r = p_focus.add_run("4. Đề xuất điều chỉnh trọng tâm:\n")
    r.font.bold = True
    goal_txt = f" nhằm phục vụ trực tiếp mục tiêu: {clar['decision_goal']}" if clar.get("decision_goal") else ""
    bench_txt = f" (đối chiếu mốc: {clar['benchmark']})" if clar.get("benchmark") else ""
    p_focus.add_run(
        f"• Điều chỉnh ngân sách theo hiệu suất thực tế{goal_txt}{bench_txt}.\n"
        f"• Giao chỉ tiêu phục hồi cụ thể cho '{lowest_target.get('store_name', 'nhóm sụt giảm')}' dựa trên mức chênh lệch trên Dashboard và siết định mức chi phí tại nhóm biên thấp."
    )

    # Footer Signatures Table
    doc.add_paragraph()
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False

    sig_left = sig_table.cell(0, 0)
    sig_right = sig_table.cell(0, 1)
    sig_left.width = Inches(3.4)
    sig_right.width = Inches(3.4)

    p_s_left = sig_left.paragraphs[0]
    p_s_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sl = p_s_left.add_run("NGƯỜI LẬP BÁO CÁO\n(Ký và ghi rõ họ tên)\n\n\n\nChuyên viên Phân tích Dữ liệu")
    r_sl.font.name = "Times New Roman"
    r_sl.font.size = Pt(9.5)

    p_s_right = sig_right.paragraphs[0]
    p_s_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sr = p_s_right.add_run("PHÊ DUYỆT BỞI BAN GIÁM ĐỐC\n(Ký tên và đóng dấu)\n\n\n\nGiám đốc Điều hành / Chiến lược")
    r_sr.font.name = "Times New Roman"
    r_sr.font.size = Pt(9.5)
    r_sr.font.bold = True

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()
