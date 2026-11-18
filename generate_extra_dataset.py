"""
Script to generate additional test datasets:
1. Du_Lieu_Test_Ky_Moi_AI301_60K.csv (60,000 transactions, fresh seed 8888)
2. Du_Lieu_Test_Ranh_Gioi_25K.csv (25,000 transactions, boundary cases, seed 5555)
"""

import csv
from data_generator import generate_retail_dataset

def generate_extra():
    # 1. Dataset 1: 60,000 records, seed 8888
    print("Generating Du_Lieu_Test_Ky_Moi_AI301_60K.csv...")
    ds_60k = generate_retail_dataset(total_records=60000, seed=8888)
    fname_60k = "Du_Lieu_Test_Ky_Moi_AI301_60K.csv"
    with open(fname_60k, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["trans_id", "order_id", "date", "store_id", "product_id", "qty", "unit_price", "status", "type", "ref_order_id"])
        for tx in ds_60k["transactions"]:
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
    print(f"-> Created {fname_60k} with {len(ds_60k['transactions'])} transactions.")

    # 2. Dataset 2: 25,000 records, seed 5555
    print("Generating Du_Lieu_Test_Ranh_Gioi_25K.csv...")
    ds_25k = generate_retail_dataset(total_records=25000, seed=5555)
    fname_25k = "Du_Lieu_Test_Ranh_Gioi_25K.csv"
    with open(fname_25k, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["trans_id", "order_id", "date", "store_id", "product_id", "qty", "unit_price", "status", "type", "ref_order_id"])
        for tx in ds_25k["transactions"]:
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
    print(f"-> Created {fname_25k} with {len(ds_25k['transactions'])} transactions.")

if __name__ == "__main__":
    generate_extra()

