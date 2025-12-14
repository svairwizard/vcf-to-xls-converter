import re
from pathlib import Path
import openpyxl
from openpyxl import Workbook

VCF_PATH = "contacts.vcf"
XLSX_PATH = "contacts.xlsx"

def parse_vcards(text: str):
    contacts = []
    current = {}

    # временное хранилище itemN.* пока не поймём, что это именно Telegram
    items = {}

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        upper = line.upper()

        if upper.startswith("BEGIN:VCARD"):
            current = {}
            items = {}
        elif upper.startswith("END:VCARD"):
            # сматчить все itemN.URL + itemN.X-ABLabel:Telegram
            for key, data in items.items():
                label = data.get("label", "").lower()
                url = data.get("url", "")
                if "telegram" in label and "t.me" in url:
                    current["TelegramUrl"] = url
                    # вытащим «хэндл»/ID после https://t.me/
                    m = re.search(r"https?://t\\.me/([^/?\\s]+)", url)
                    if m:
                        current["TelegramHandle"] = m.group(1)
            if current:
                contacts.append(current)
        else:
            # Имя
            if upper.startswith("FN:"):
                current["ФИО"] = line[3:].strip()
            elif upper.startswith("N:"):
                parts = line[2:].split(";")
                last_name = parts[0].strip() if len(parts) > 0 else ""
                first_name = parts[1].strip() if len(parts) > 1 else ""
                if not current.get("FullName"):
                    current["ФИО"] = (first_name + " " + last_name).strip()

            # Телефоны
            elif upper.startswith("TEL"):
                m = re.search(r":(.+)$", line)
                if m:
                    phone = m.group(1).strip()
                    if phone:
                        current.setdefault("Телефон", []).append(phone)

            # itemN.URL;type=pref:...
            elif ".URL" in upper:
                # пример: item1.URL;type=pref:https://t.me/@id1234567890
                # отделим группу item1 и значение
                left, value = line.split(":", 1)
                group = left.split(".")[0]  # item1
                items.setdefault(group, {})["url"] = value.strip()

            # itemN.X-ABLabel:Telegram
            elif ".X-ABLABEL" in upper:
                left, value = line.split(":", 1)
                group = left.split(".")[0]  # item1
                items.setdefault(group, {})["label"] = value.strip()

    return contacts


def export_to_xlsx(contacts, path: str):
    wb: Workbook = Workbook()
    ws = wb.active
    ws.title = "Contacts"

    ws.append(
        ["ФИО", "Телефон1", "Телефон2", "Телеграм"]
    )

    for c in contacts:
        phones = c.get("Телефон", [])
        row = [
            c.get("ФИО", ""),
            phones[0] if len(phones) > 0 else "",
            phones[1] if len(phones) > 1 else "",
            c.get("TelegramUrl", ""),
            c.get("TelegramHandle", ""),
        ]
        ws.append(row)

    wb.save(path)


def main():
    vcf_text = Path(VCF_PATH).read_text(encoding="utf-8", errors="ignore")
    contacts = parse_vcards(vcf_text)
    export_to_xlsx(contacts, XLSX_PATH)
    print(f"Готово. Найдено контактов: {len(contacts)}")


if __name__ == "__main__":
    main()
