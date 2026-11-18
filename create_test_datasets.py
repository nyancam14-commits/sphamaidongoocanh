"""
Script to generate fresh test datasets for Phuc An Retail AI Agent:
1. Du_Lieu_Moi_Thang_11_PhucAn_60K.csv (~60,000 records)
2. Du_Lieu_Moi_Test_10K.csv (~10,000 records)
3. Danh_Muc_San_Pham_Gia_Von.csv (SKU, Name, Category, Price, COGS)
4. Danh_Sach_18_Cua_Hang_Muc_Tieu.csv (Store Targets)
"""

import csv
import json
from data_generator import generate_retail_dataset

def export_to_csv():
    # 1. Generate fresh 60K dataset with seed 2026
    print("Generating 60,000 transaction dataset (seed=2026)...")
    dataset_60k = generate_retail_dataset(total_records=60000, seed=2026)
    
    csv_file_60k = "Du_Lieu_Moi_Thang_11_PhucAn_60K.csv"
    with open(csv_file_60k, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["trans_id", "order_id", "date", "store_id", "product_id", "qty", "unit_price", "status", "type", "ref_order_id"])
        for tx in dataset_60k["transactions"]:
            writer.writerow([
                tx["trans_id"],
                tx["order_id"],
                tx["date"],
                tx["store_id"],
                tx["product_id"],
                tx["qty"],
                tx["unit_price"],
                tx["status"],
                tx["type"],
                tx.get("ref_order_id") or ""
            ])
    print(f"Created {csv_file_60k} with {len(dataset_60k['transactions'])} transactions.")

    # 2. Generate 10K test dataset with seed 777
    print("Generating 10,000 transaction test dataset (seed=777)...")
    dataset_10k = generate_retail_dataset(total_records=10000, seed=777)
    
    csv_file_10k = "Du_Lieu_Moi_Test_10K.csv"
    with open(csv_file_10k, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["trans_id", "order_id", "date", "store_id", "product_id", "qty", "unit_price", "status", "type", "ref_order_id"])
        for tx in dataset_10k["transactions"]:
            writer.writerow([
                tx["trans_id"],
                tx["order_id"],
                tx["date"],
                tx["store_id"],
                tx["product_id"],
                tx["qty"],
                tx["unit_price"],
                tx["status"],
                tx["type"],
                tx.get("ref_order_id") or ""
            ])
    print(f"Created {csv_file_10k} with {len(dataset_10k['transactions'])} transactions.")

    # 3. Product Master CSV
    prod_file = "Danh_Muc_San_Pham_Gia_Von.csv"
    with open(prod_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["product_id", "name", "category", "price", "cogs"])
        for pid, p in dataset_60k["products"].items():
            writer.writerow([pid, p["name"], p["category"], p["price"], p["cogs"]])
    print(f"Created {prod_file}.")

    # 4. Store Targets CSV
    store_file = "Danh_Sach_18_Cua_Hang_Muc_Tieu.csv"
    with open(store_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["store_id", "name", "region", "manager", "report_rev_target", "report_profit_target", "compare_rev_target", "compare_profit_target"])
        for sid, s in dataset_60k["stores"].items():
            tgt = dataset_60k["targets"][sid]
            writer.writerow([
                sid, s["name"], s["region"], s["manager"],
                tgt["report_rev_target"], tgt["report_profit_target"],
                tgt["compare_rev_target"], tgt["compare_profit_target"]
            ])
    print(f"Created {store_file}.")

if __name__ == "__main__":
    export_to_csv()

