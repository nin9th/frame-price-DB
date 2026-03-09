"""
update_data.py
──────────────
Run this whenever you update Frame_Price_DB.xlsx.
Requires: pip install pywin32
Usage:
    python update_data.py
"""
import os, sys, json
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

EXCEL_PATH  = "Frame_Price_DB.xlsx"
OUTPUT_PATH = "prices.json"
WOOD_TYPES  = [
    "TT-1", "TT-2", "TT-3", "TT-4", "TT-5",
    "TT-6/8", "TT-7/9", "TT-10", "TT-11", "TT-12", "TT-97", "TT-98", "TT-99"
]

try:
    import win32com.client
except ImportError:
    print("❌  pywin32 not found. Run:")
    print("    pip install pywin32")
    sys.exit(1)

def to_int(val):
    return int(Decimal(str(val)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))

def extract_all(excel_path):
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        print(f"❌  ไม่พบไฟล์: {abs_path}")
        sys.exit(1)
    print(f"📂  เปิด Excel: {abs_path}\n")
    xl  = win32com.client.Dispatch("Excel.Application")
    xl.Visible        = False
    xl.DisplayAlerts  = False
    wb  = xl.Workbooks.Open(abs_path)
    ws1 = wb.Sheets("Sheet1")
    all_data = {}
    for wood in WOOD_TYPES:
        print(f"   ⚙️  {wood}...", end=" ", flush=True)
        try:
            ws1.Range("B11").Value = wood
            xl.Calculate()
            ws_price = wb.Sheets("ใบราคา")
            def cell(r, c):
                return ws_price.Cells(r, c).Value
            data = {
                "description":   str(cell(1, 2) or ""),
                "perInchPlain":  round(float(cell(3, 2) or 0), 2),
                "perInchGroove": round(float(cell(3, 5) or 0), 2),
                "grooveLabel":   str(cell(6, 5) or "ใส่ฝ้าย 1 นิ้ว"),
                "sizes": []
            }
            for r in range(9, 27):
                size   = cell(r, 1)
                plain  = cell(r, 2)
                groove = cell(r, 5)
                if size and plain:
                    data["sizes"].append({
                        "size":   str(size),
                        "plain":  to_int(plain),
                        "groove": to_int(groove) if groove else None
                    })
            all_data[wood] = data
            print(f"✓  ({len(data['sizes'])} ขนาด, {data['perInchPlain']}฿/นิ้ว)")
        except Exception as e:
            print(f"❌  {e}")
    wb.Close(SaveChanges=False)
    xl.Quit()
    return all_data

def main():
    all_data = extract_all(EXCEL_PATH)
    output = {
        "updated": date.today().isoformat(),
        "woods":   all_data
    }
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    size_kb = os.path.getsize(OUTPUT_PATH) / 1024
    print(f"\n✅  สำเร็จ! {OUTPUT_PATH} ({size_kb:.1f} KB)")
    print(f"    ส่งไฟล์นี้ให้ทีมงานผ่าน LINE ได้เลย 📤")

if __name__ == "__main__":
    main()
