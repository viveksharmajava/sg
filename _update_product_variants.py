from pathlib import Path
import openpyxl

ROOT = Path(r"C:\vivek\project\catalog\templates\pp\pp\SG")

VARIANT_MAP = {
    "batting-gloves": ("CRICKET_ACS_SIZE", "Junior,Boys,Youth,Adult"),
    "batting-pads": ("CRICKET_ACS_SIZE", "Junior,Boys,Youth,Adult"),
    "keeping-gloves": ("CRICKET_ACS_SIZE", "Junior,Boys,Youth,Adult"),
    "thigh-pads": ("CRICKET_ACS_SIZE", "Junior,Boys,Youth,Adult"),
    "inner-gloves": ("CRICKET_ACS_SIZE", "Junior,Boys,Youth,Adult"),
    "elbow-guards": ("CRICKET_ACS_SIZE", "Junior,Boys,Youth,Adult"),
    "english-willow": ("C_BAT_SIZE", "2,3,4,5,6,H,SH"),
    "kashmir-willow": ("C_BAT_SIZE", "2,3,4,5,6,H,SH"),
    "shoes": ("SHOE_SIZE", "1,2,3,4,5,6,7,8,9,10,11"),
    "helmets": ("TSHIRT_SIZE", "S,M,L,XL,XXL"),
}

HEADERS = [
    "Product_id",
    "Variant_type1",
    "Variant_type1_value",
    "Variant_type2",
    "Variant_type2_value",
    "Variant_type3",
    "Variant_type3_value",
]


def main():
    summaries = []

    for folder, (vtype, vvalue) in VARIANT_MAP.items():
        folder_path = ROOT / folder
        if not folder_path.is_dir():
            print(f"SKIP missing folder: {folder}")
            continue
        for xlsx in sorted(folder_path.glob("*.xlsx")):
            if xlsx.name.startswith("~$"):
                continue
            wb = openpyxl.load_workbook(xlsx)
            if "Products" not in wb.sheetnames:
                print(f"SKIP no Products sheet: {xlsx}")
                wb.close()
                continue

            products = wb["Products"]
            header_row = [cell.value for cell in products[1]]
            header_lower = [
                (str(h).strip().lower() if h is not None else "") for h in header_row
            ]

            try:
                pid_idx = header_lower.index("product_id")
            except ValueError:
                print(f"SKIP no product_id column: {xlsx}")
                wb.close()
                continue

            flag_idx = None
            flag_col = None
            if "is_virtual" in header_lower:
                flag_idx = header_lower.index("is_virtual")
                flag_col = "is_virtual"
            elif "is_variant" in header_lower:
                flag_idx = header_lower.index("is_variant")
                flag_col = "is_variant"

            product_ids = []
            rows_to_flag = []
            for r in range(2, products.max_row + 1):
                pid = products.cell(r, pid_idx + 1).value
                if pid is None or str(pid).strip() == "":
                    continue
                product_ids.append(str(pid).strip())
                rows_to_flag.append(r)

            if "Product_variants" in wb.sheetnames:
                del wb["Product_variants"]
            ws = wb.create_sheet("Product_variants")
            for c, h in enumerate(HEADERS, 1):
                ws.cell(1, c, h)

            for i, pid in enumerate(product_ids, start=2):
                ws.cell(i, 1, pid)
                ws.cell(i, 2, vtype)
                ws.cell(i, 3, vvalue)

            virtual_updates = 0
            if flag_idx is not None:
                for r in rows_to_flag:
                    products.cell(r, flag_idx + 1).value = "Y"
                    virtual_updates += 1

            wb.save(xlsx)
            wb.close()
            summaries.append(
                {
                    "path": str(xlsx),
                    "folder": folder,
                    "products": len(product_ids),
                    "variants": len(product_ids),
                    "virtual_updates": virtual_updates,
                    "flag_col": flag_col,
                    "vtype": vtype,
                    "vvalue": vvalue,
                }
            )

    print("\n=== SUMMARY ===")
    for s in summaries:
        print(s["path"])
        print(
            f"  products={s['products']} variant_rows={s['variants']} "
            f"virtual_flag_updates={s['virtual_updates']} (col={s['flag_col']})"
        )

    print("\n=== VERIFICATION SAMPLES ===")
    groups = {
        "CRICKET_ACS_SIZE": "elbow-guards",
        "C_BAT_SIZE": "english-willow",
        "SHOE_SIZE": "shoes",
        "TSHIRT_SIZE": "helmets",
    }
    for group_name, folder in groups.items():
        match = next((s for s in summaries if s["folder"] == folder), None)
        if not match:
            print(f"No file for group {group_name}")
            continue
        wb = openpyxl.load_workbook(match["path"], read_only=True)
        ws = wb["Product_variants"]
        rows = list(ws.iter_rows(min_row=1, max_row=3, max_col=7, values_only=True))
        print(f"\n[{group_name}] {match['path']}")
        print(f"  headers: {rows[0]}")
        for r in rows[1:]:
            print(f"  sample: {r}")
        wb.close()


if __name__ == "__main__":
    main()
