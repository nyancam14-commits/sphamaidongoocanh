"""
Realistic Retail Data Generator for Phuc An Retail (18 Stores, 60,000 lines/month)
Includes normal cases, missing data, and boundary condition test cases.
"""

import random
from datetime import datetime, timedelta

def generate_retail_dataset(total_records=60000, seed=42):
    random.seed(seed)

    # 1. Store Master (18 units)
    stores = {
        f"CH{i:02d}": {
            "store_id": f"CH{i:02d}",
            "name": f"Đơn Vị / Chi Nhánh {i:02d}",
            "region": "Miền Bắc" if i <= 6 else ("Miền Trung" if i <= 11 else "Miền Nam"),
            "manager": f"Quản lý {i:02d}"
        } for i in range(1, 19)
    }

    # 2. Product Master (~20 SKUs across 5 categories)
    categories = ["Thời trang Nam", "Thời trang Nữ", "Phụ kiện", "Giày dép", "Đồ thể thao"]
    products = {
        "SKU001": {"name": "Áo Sơ Mi Nam Oxford", "category": "Thời trang Nam", "price": 450000, "cogs": 220000},
        "SKU002": {"name": "Quần Tây Nam Slimfit", "category": "Thời trang Nam", "price": 550000, "cogs": 280000},
        "SKU003": {"name": "Áo Thun Polo Cotton", "category": "Thời trang Nam", "price": 320000, "cogs": 150000},
        "SKU004": {"name": "Áo Khoác Bomber Nam", "category": "Thời trang Nam", "price": 850000, "cogs": 480000},
        
        "SKU005": {"name": "Đầm Xòe Nữ Voan Hoa", "category": "Thời trang Nữ", "price": 680000, "cogs": 310000},
        "SKU006": {"name": "Chân Váy Chữ A Công Sở", "category": "Thời trang Nữ", "price": 420000, "cogs": 190000},
        "SKU007": {"name": "Áo Kiểu Nữ Tay Bồng", "category": "Thời trang Nữ", "price": 390000, "cogs": 180000},
        "SKU008": {"name": "Set Vest Nữ Thanh Lịch", "category": "Thời trang Nữ", "price": 1250000, "cogs": 650000},
        
        "SKU009": {"name": "Thắt Lưng Da Bò Nam", "category": "Phụ kiện", "price": 350000, "cogs": 120000},
        "SKU010": {"name": "Túi Xách Nữ Da Mềm", "category": "Phụ kiện", "price": 750000, "cogs": 320000},
        "SKU011": {"name": "Ví Cầm Tay Mini", "category": "Phụ kiện", "price": 280000, "cogs": 90000},
        "SKU012": {"name": "Kính Mát Chống UV", "category": "Phụ kiện", "price": 250000, "cogs": 80000},
        
        "SKU013": {"name": "Giày Sneaker Phố Classic", "category": "Giày dép", "price": 890000, "cogs": 480000},
        "SKU014": {"name": "Giày Cao Gót 7cm", "category": "Giày dép", "price": 620000, "cogs": 300000},
        "SKU015": {"name": "Giày Lười Loafer Da", "category": "Giày dép", "price": 950000, "cogs": 520000},
        "SKU016": {"name": "Dép Quai Ngang Đúc", "category": "Giày dép", "price": 190000, "cogs": 70000},
        
        "SKU017": {"name": "Áo Tập Gym Dry-Fit", "category": "Đồ thể thao", "price": 280000, "cogs": 110000},
        "SKU018": {"name": "Quần Legging Co Giãn 4 Chiều", "category": "Đồ thể thao", "price": 360000, "cogs": 140000},
        "SKU019": {"name": "Bộ Nỉ Thể Thao Mùa Thu", "category": "Đồ thể thao", "price": 720000, "cogs": 380000},
        "SKU020": {"name": "Áo Gió Thể Thao Chống Nước", "category": "Đồ thể thao", "price": 580000, "cogs": 270000}
    }

    # 3. Targets Master (Tháng 10/2026 - Kỳ Báo cáo; Tháng 09/2026 - Kỳ So sánh)
    # Cài đặt một số cửa hàng có chỉ tiêu thách thức hoặc hiệu suất chênh lệch để kiểm thử logic
    targets = {}
    for sid in stores:
        # Mục tiêu doanh thu: ~ 1 tỷ đến 1.8 tỷ / tháng
        rep_rev_tgt = random.randint(1100, 1600) * 1000000
        # Mục tiêu lợi nhuận gộp: ~ 45% - 52%
        rep_prof_tgt = int(rep_rev_tgt * random.uniform(0.46, 0.52))
        
        cmp_rev_tgt = int(rep_rev_tgt * random.uniform(0.92, 1.05))
        cmp_prof_tgt = int(cmp_rev_tgt * random.uniform(0.46, 0.52))
        
        targets[sid] = {
            "report_rev_target": rep_rev_tgt,
            "report_profit_target": rep_prof_tgt,
            "compare_rev_target": cmp_rev_tgt,
            "compare_profit_target": cmp_prof_tgt
        }

    # Cố định đặc điểm cho một số cửa hàng để kiểm thử phát hiện rủi ro:
    # CH04 (Hà Nội Cầu Giấy): Tỷ lệ trả hàng rất cao do vấn đề kiểm soát size, doanh thu hụt mạnh
    # CH09 (Đà Nẵng): Lỗ gộp do xả hàng giảm giá sâu dưới giá vốn, biên lợi nhuận âm
    # CH15 (TP.HCM Tân Bình): Doanh thu sụt giảm 40% so với kỳ trước, hụt mục tiêu nghiêm trọng
    targets["CH04"]["report_rev_target"] = 1500000000
    targets["CH04"]["report_profit_target"] = 750000000
    targets["CH09"]["report_rev_target"] = 1400000000
    targets["CH09"]["report_profit_target"] = 700000000
    targets["CH15"]["report_rev_target"] = 1600000000
    targets["CH15"]["report_profit_target"] = 800000000

    # 4. Period Configuration
    period_config = {
        "report_period_name": "Tháng 10/2026",
        "report_start": "2026-10-01",
        "report_end": "2026-10-31",
        "compare_period_name": "Tháng 09/2026",
        "compare_start": "2026-09-01",
        "compare_end": "2026-09-30"
    }

    # 5. Generate Raw Transactions (~60,000 lines)
    transactions = []
    
    # 60,000 records partitioned:
    # - ~54% Report Period (Oct 2026) -> ~32,400 records
    # - ~42% Compare Period (Sep 2026) -> ~25,200 records
    # - ~4% Out of period / Edge cases / Injected anomalies -> ~2,400 records
    
    sku_list = list(products.keys())
    store_list = list(stores.keys())

    start_rep = datetime(2026, 10, 1)
    end_rep = datetime(2026, 10, 31)
    start_cmp = datetime(2026, 9, 1)
    end_cmp = datetime(2026, 9, 30)

    order_history_rep = []
    order_history_cmp = []

    trans_id_counter = 1000001
    order_id_counter = 500001

    num_rep = int(total_records * 0.54)
    num_cmp = int(total_records * 0.42)
    num_edge = total_records - num_rep - num_cmp

    # Tạo giao dịch Kỳ So Sánh (Tháng 9)
    for _ in range(num_cmp):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        order_id = f"ORD{order_id_counter}"
        order_id_counter += 1

        sid = random.choice(store_list)
        pid = random.choice(sku_list)
        prod = products[pid]
        
        # Ngày tháng ngẫu nhiên trong tháng 9
        days_offset = random.randint(0, 29)
        tx_date = (start_cmp + timedelta(days=days_offset)).strftime("%Y-%m-%d")
        
        qty = random.choices([1, 2, 3], weights=[80, 15, 5])[0]
        price = prod["price"]
        status = "COMPLETED"
        ttype = "SALE"

        tx = {
            "trans_id": trans_id,
            "order_id": order_id,
            "store_id": sid,
            "product_id": pid,
            "date": tx_date,
            "type": ttype,
            "qty": qty,
            "unit_price": price,
            "status": status,
            "ref_order_id": ""
        }
        transactions.append(tx)
        order_history_cmp.append(tx)

    # Tạo giao dịch Kỳ Báo Cáo (Tháng 10)
    for _ in range(num_rep):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        order_id = f"ORD{order_id_counter}"
        order_id_counter += 1

        # Cân chỉnh xác suất cho các cửa hàng đặc biệt:
        # CH15 có ít giao dịch hơn (sụt giảm doanh số)
        sid_weights = [0.4 if s == "CH15" else 1.0 for s in store_list]
        sid = random.choices(store_list, weights=sid_weights)[0]
        
        pid = random.choice(sku_list)
        prod = products[pid]

        days_offset = random.randint(0, 30)
        tx_date = (start_rep + timedelta(days=days_offset)).strftime("%Y-%m-%d")
        qty = random.choices([1, 2, 3], weights=[80, 15, 5])[0]
        
        # CH09 bán xả hàng giá sốc dưới giá vốn (gây âm lợi nhuận gộp)
        if sid == "CH09" and random.random() < 0.6:
            price = int(prod["cogs"] * 0.75) # Bán lỗ 25% dưới giá vốn
        else:
            price = prod["price"]

        status = "COMPLETED"
        ttype = "SALE"

        tx = {
            "trans_id": trans_id,
            "order_id": order_id,
            "store_id": sid,
            "product_id": pid,
            "date": tx_date,
            "type": ttype,
            "qty": qty,
            "unit_price": price,
            "status": status,
            "ref_order_id": ""
        }
        transactions.append(tx)
        order_history_rep.append(tx)

    # Thêm các giao dịch Đổi Trả Hàng (RETURN) cho kỳ tháng 10
    # CH04 có tỷ lệ đổi trả cực cao (~16%), các cửa hàng khác ~ 2.5%
    for sale_tx in order_history_rep[:3500]:
        sid = sale_tx["store_id"]
        return_prob = 0.25 if sid == "CH04" else 0.025
        if random.random() < return_prob:
            ret_trans_id = f"TX{trans_id_counter}"
            trans_id_counter += 1
            ret_order_id = f"RET{order_id_counter}"
            order_id_counter += 1

            # Có thể trả hàng tại cửa hàng khác (Cross-store return)
            ret_store_id = sid if random.random() < 0.85 else random.choice(store_list)
            
            # Ngày trả sau ngày mua 1-5 ngày
            sale_dt = datetime.strptime(sale_tx["date"], "%Y-%m-%d")
            ret_dt = min(sale_dt + timedelta(days=random.randint(1, 4)), end_rep)

            ret_tx = {
                "trans_id": ret_trans_id,
                "order_id": ret_order_id,
                "store_id": ret_store_id,
                "product_id": sale_tx["product_id"],
                "date": ret_dt.strftime("%Y-%m-%d"),
                "type": "RETURN",
                "qty": 1,
                "unit_price": sale_tx["unit_price"],
                "status": "COMPLETED",
                "ref_order_id": sale_tx["order_id"]
            }
            transactions.append(ret_tx)

    # 6. Inject Anomaly & Boundary Test Cases (Để kiểm thử năng lực làm sạch và đối soát của Agent)
    # a. Giao dịch trùng mã (Duplicate transaction ID)
    for i in range(120):
        target_tx = random.choice(order_history_rep)
        dup_tx = dict(target_tx)
        transactions.append(dup_tx)

    # b. Đơn bị hủy (Cancelled / Void)
    for i in range(250):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        cancel_tx = {
            "trans_id": trans_id,
            "order_id": f"ORD{order_id_counter}",
            "store_id": random.choice(store_list),
            "product_id": random.choice(sku_list),
            "date": "2026-10-15",
            "type": "SALE",
            "qty": 1,
            "unit_price": 500000,
            "status": random.choice(["CANCELLED", "HUY", "VOID"]),
            "ref_order_id": ""
        }
        order_id_counter += 1
        transactions.append(cancel_tx)

    # c. Đơn đổi trả không tìm thấy đơn gốc (Unmatched Return / Fraud / Orphan Return)
    for i in range(85):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        orphan_return = {
            "trans_id": trans_id,
            "order_id": f"RET{order_id_counter}",
            "store_id": random.choice(store_list),
            "product_id": random.choice(sku_list),
            "date": "2026-10-20",
            "type": "RETURN",
            "qty": 1,
            "unit_price": 450000,
            "status": "COMPLETED",
            "ref_order_id": f"NON_EXISTENT_ORDER_{i}"
        }
        order_id_counter += 1
        transactions.append(orphan_return)

    # d. Giao dịch ngoài kỳ báo cáo (Out of period: tháng 8 hoặc tháng 11)
    for i in range(180):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        oop_tx = {
            "trans_id": trans_id,
            "order_id": f"ORD{order_id_counter}",
            "store_id": random.choice(store_list),
            "product_id": random.choice(sku_list),
            "date": random.choice(["2026-08-15", "2026-08-28", "2026-11-02", "2026-11-10"]),
            "type": "SALE",
            "qty": 1,
            "unit_price": 400000,
            "status": "COMPLETED",
            "ref_order_id": ""
        }
        order_id_counter += 1
        transactions.append(oop_tx)

    # e. Mã sản phẩm không tồn tại trong danh mục hoặc thiếu giá vốn
    for i in range(45):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        bad_sku_tx = {
            "trans_id": trans_id,
            "order_id": f"ORD{order_id_counter}",
            "store_id": random.choice(store_list),
            "product_id": f"UNKNOWN_SKU_{i:03d}",
            "date": "2026-10-18",
            "type": "SALE",
            "qty": 1,
            "unit_price": 600000,
            "status": "COMPLETED",
            "ref_order_id": ""
        }
        order_id_counter += 1
        transactions.append(bad_sku_tx)

    # f. Mã cửa hàng không hợp lệ
    for i in range(25):
        trans_id = f"TX{trans_id_counter}"
        trans_id_counter += 1
        bad_store_tx = {
            "trans_id": trans_id,
            "order_id": f"ORD{order_id_counter}",
            "store_id": "CH99_INVALID",
            "product_id": random.choice(sku_list),
            "date": "2026-10-19",
            "type": "SALE",
            "qty": 1,
            "unit_price": 300000,
            "status": "COMPLETED",
            "ref_order_id": ""
        }
        order_id_counter += 1
        transactions.append(bad_store_tx)

    # Trộn ngẫu nhiên thứ tự giao dịch để mô phỏng dữ liệu thô từ nhiều nguồn POS đẩy về
    random.shuffle(transactions)

    return {
        "stores": stores,
        "products": products,
        "targets": targets,
        "period_config": period_config,
        "transactions": transactions
    }

if __name__ == "__main__":
    data = generate_retail_dataset(total_records=60000)
    print(f"Generated {len(data['transactions'])} transactions across {len(data['stores'])} stores.")
    print("Sample transaction:", data["transactions"][0])

