"""Create contacts_sample.xlsx with expected column names. Run after: pip install -r requirements.txt"""

from pathlib import Path

from openpyxl import Workbook


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["Comapany Name", "EmailID"])
    ws.append(["Acme Corporation", "contact@acme.example"])
    ws.append(["Beta Industries", "hello@beta.example"])
    out = Path(__file__).resolve().parent / "contacts_sample.xlsx"
    wb.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
