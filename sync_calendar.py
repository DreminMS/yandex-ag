#!/usr/bin/env python3
"""
Скрипт автономной синхронизации графика с Яндекс.Диска
и генерации iCalendar (.ics) файлов для подписки в Google / Apple Календаре.
Работает как локально, так и автономно в GitHub Actions 24/7.
"""

import os
import sys
import json
import re
import ssl
import time
import hashlib
import urllib.request
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
import datetime

PUBLIC_URL = "https://disk.yandex.ru/i/2NYAsJXEPM_ejg"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, "latest_schedule.xlsx")

SSL_CTX = ssl._create_unverified_context()

MONTHS_RU = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
]

def download_file_with_retry(url, target_path, retries=4, timeout=30):
    for attempt in range(1, retries + 1):
        try:
            print(f"Попытка {attempt}: получение ссылки для {url}...")
            enc = urllib.parse.quote(url.strip())
            api = f"https://cloud-api.yandex.net/v1/disk/public/resources?public_key={enc}"
            req = urllib.request.Request(api, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            })
            with urllib.request.urlopen(req, context=SSL_CTX, timeout=timeout) as resp:
                meta = json.loads(resp.read().decode("utf-8"))
            file_url = meta.get("file")
            if not file_url:
                raise RuntimeError(f"В ответе API нет прямой ссылки 'file': {meta}")
            
            print(f"Скачивание файла таблицы ({meta.get('size', 0)} байт)...")
            down_req = urllib.request.Request(file_url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            })
            with urllib.request.urlopen(down_req, context=SSL_CTX, timeout=timeout * 2) as resp:
                data = resp.read()
            with open(target_path, "wb") as f:
                f.write(data)
            print(f"Успешно сохранено: {target_path} ({len(data)} байт)")
            return True
        except Exception as e:
            print(f"Ошибка при попытке {attempt}: {e}")
            if attempt < retries:
                time.sleep(3)
            elif os.path.exists(target_path) and os.path.getsize(target_path) > 10000:
                print(f"Используем ранее сохранённый кэш {target_path}")
                return True
            else:
                raise
    return False

def extract_time_and_clean_title(text):
    """
    Извлекает время начала/окончания и очищенный заголовок.
    Поддерживает:
    - 10:30, 10.30, 17:00
    - 8.20-11.00, 15.30-19.20, 08:20 - 11:00
    - до 12.00, к 14:00
    - Исключает ложные срабатывания на даты (отгул за 06.06, 15.08, 29.08).
    """
    clean_lines = [l.strip() for l in text.split("\n") if l.strip()]
    first_line = clean_lines[0] if clean_lines else text

    # If day off / vacation / duty without hours, it's all-day
    lower_line = first_line.lower()
    if lower_line.startswith("отгул") or lower_line.startswith("отпуск") or re.match(r"^деж\b", lower_line):
        return None, None, first_line

    # Range: 8.20-11.00 or 15.30-19.20 or 9:00 - 18:00
    m_range = re.search(r"(?:^|\s|[сc]\s*)(\b\d{1,2})[.:](\d{2})\s*[-—–до]\s*(\d{1,2})[.:](\d{2})", first_line)
    if m_range:
        sh, sm, eh, em = int(m_range.group(1)), int(m_range.group(2)), int(m_range.group(3)), int(m_range.group(4))
        if 0 <= sh <= 23 and 0 <= sm <= 59 and 0 <= eh <= 23 and 0 <= em <= 59:
            title = first_line[:m_range.start()] + " " + first_line[m_range.end():]
            title = re.sub(r"^[,\s–—-]+", "", title).strip()
            # If title is empty or short, use original first_line
            return (sh, sm), (eh, em), title if len(title) > 3 else first_line

    # "до 12.00"
    m_until = re.search(r"\bдо\s*(\d{1,2})[.:](\d{2})", first_line, re.IGNORECASE)
    if m_until:
        eh, em = int(m_until.group(1)), int(m_until.group(2))
        if 0 <= eh <= 23 and 0 <= em <= 59:
            sh = max(0, eh - 2)
            sm = em
            return (sh, sm), (eh, em), first_line

    # Single time: 10:30, 17:00, 10.30
    # Must NOT be preceded by "за " or "от " (which means a date)
    m_single = re.search(r"(?<!за\s)(?<!от\s)(?:^|\s|[вb]\s*)(\b\d{1,2})[.:](\d{2})(?!\s*г|\s*года)", first_line)
    if m_single:
        sh, sm = int(m_single.group(1)), int(m_single.group(2))
        if 0 <= sh <= 23 and 0 <= sm <= 59:
            # Default duration 2.5 hours
            eh = sh + 2
            em = sm + 30
            if em >= 60:
                em -= 60
                eh += 1
            if eh >= 24:
                eh = 23
                em = 59
            title = first_line[:m_single.start()] + " " + first_line[m_single.end():]
            title = re.sub(r"^[,\s–—-]+", "", title).strip()
            return (sh, sm), (eh, em), title if len(title) > 3 else first_line

    return None, None, first_line

def extract_location(text):
    """
    Ищет адрес/локацию в тексте задачи и убирает лишнее.
    """
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    for line in lines:
        m = re.search(r"локаци[яи]\s*:\s*(.+)", line, re.IGNORECASE)
        if m:
            loc = m.group(1).strip()
            loc = re.sub(r"\s*ТЗ\s*:.*$", "", loc, flags=re.IGNORECASE).strip()
            return loc
        if any(w in line.lower() for w in ["ул.", "улица", "пер.", "переулок", "просп.", "проспект", "б-р", "шоссе", "д."]):
            if not line.startswith("http") and not line.lower().startswith("сдача") and not line.lower().startswith("обработка"):
                clean = re.sub(r"\s*ТЗ\s*:.*$", "", line, flags=re.IGNORECASE).strip()
                return clean
    return None

def extract_urls(text):
    return re.findall(r"https?://[^\s]+", text)

def parse_workbook(xlsx_path):
    print("Парсинг Excel файла таблицы...")
    base_date = datetime.date(1899, 12, 30)

    with zipfile.ZipFile(xlsx_path, "r") as z:
        sst = []
        if "xl/sharedStrings.xml" in z.namelist():
            tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            for si in tree.findall(".//ns:si", ns):
                sst.append("".join([t.text or "" for t in si.findall(".//ns:t", ns)]))

        tree = ET.fromstring(z.read("xl/workbook.xml"))
        ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rels_tree = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        r_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_tree.findall(".//r:Relationship", r_ns)}

        today = datetime.date.today()
        # Scan sheets for current month, next month, and previous month
        target_dates = [
            (today.year, today.month),
            (today.year if today.month > 1 else today.year - 1, today.month - 1 if today.month > 1 else 12),
            (today.year if today.month < 12 else today.year + 1, today.month + 1 if today.month < 12 else 1),
            (today.year if today.month < 11 else today.year + 1, today.month + 2 if today.month < 11 else (today.month + 2) % 12)
        ]

        target_sheets = []
        for s in tree.findall(".//ns:sheet", ns):
            s_name = s.attrib.get("name", "")
            s_lower = s_name.lower()
            for y, m in target_dates:
                m_ru = MONTHS_RU[m - 1]
                if m_ru in s_lower and str(y) in s_name:
                    r_id = s.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    t_path = "xl/" + rel_map[r_id] if not rel_map[r_id].startswith("xl/") else rel_map[r_id]
                    if t_path in z.namelist() and (s_name, t_path) not in target_sheets:
                        target_sheets.append((s_name, t_path))

        print(f"Найдены целевые листы для обработки: {[s[0] for s in target_sheets]}")

        all_events_by_person = {
            "misha": [],
            "tanya": [],
            "ilya": [],
            "dima": []
        }

        for sheet_name, target_path in target_sheets:
            stree = ET.fromstring(z.read(target_path))
            rows = stree.findall(".//ns:row", ns)
            if not rows:
                continue

            # Determine column mapping from header row
            col_map = {}
            for c in rows[0].findall("ns:c", ns):
                coord = c.attrib.get("r")
                col = "".join([ch for ch in coord if ch.isalpha()])
                t = c.attrib.get("t")
                v = c.find("ns:v", ns)
                val = ""
                if v is not None and v.text:
                    val = sst[int(v.text)] if t == "s" else v.text
                val_lower = val.lower()
                if "миша" in val_lower:
                    col_map[col] = ("misha", "Михаил Дремин")
                elif "таня" in val_lower:
                    col_map[col] = ("tanya", "Таня")
                elif "илья" in val_lower:
                    col_map[col] = ("ilya", "Илья Фесунов")
                elif "дима" in val_lower or "дмитрий" in val_lower:
                    col_map[col] = ("dima", "Дмитрий Беляев")

            cur_date = None
            for r in rows:
                row_idx = r.attrib.get("r")
                cells = {}
                for c in r.findall("ns:c", ns):
                    coord = c.attrib.get("r")
                    col = "".join([ch for ch in coord if ch.isalpha()])
                    t = c.attrib.get("t")
                    v = c.find("ns:v", ns)
                    val = ""
                    if v is not None and v.text:
                        val = sst[int(v.text)] if t == "s" else v.text
                    if val:
                        cells[col] = val

                # Date in Col A
                if "A" in cells:
                    try:
                        s = float(cells["A"])
                        if 40000 < s < 50000:
                            cur_date = base_date + datetime.timedelta(days=int(s))
                    except:
                        pass

                if not cur_date:
                    continue

                for col, (person_key, person_name) in col_map.items():
                    if col in cells:
                        text_val = cells[col].strip()
                        try:
                            s = float(text_val)
                            if 40000 < s < 50000:
                                continue
                        except:
                            pass

                        lower_val = text_val.lower()
                        if not text_val or "обозначения" in lower_val or "фотограф" in lower_val or "видеограф" in lower_val:
                            continue

                        start_t, end_t, clean_title = extract_time_and_clean_title(text_val)
                        loc = extract_location(text_val)
                        urls = extract_urls(text_val)

                        # Clean Title: short & expressive
                        title = clean_title.split("\n")[0].strip()
                        # Remove trailing URLs from title
                        title = re.sub(r"https?://[^\s]+", "", title).strip()
                        if len(title) > 90:
                            title = title[:87] + "..."

                        # Format detailed description
                        desc_parts = [text_val]
                        if loc:
                            desc_parts.append(f"📍 Адрес: {loc}")
                        if urls:
                            desc_parts.append(f"🔗 Ссылка на ТЗ: {urls[0]}")
                        desc = "\n\n".join(desc_parts)

                        # Stable deterministic UID based on person, date and content hash
                        uid_raw = f"{person_key}-{cur_date}-{text_val.splitlines()[0]}"
                        uid = hashlib.sha256(uid_raw.encode("utf-8")).hexdigest()[:16] + "@yandex-ag"

                        event_data = {
                            "uid": uid,
                            "date": str(cur_date),
                            "person_key": person_key,
                            "person_name": person_name,
                            "full_text": text_val,
                            "title": title or "Событие",
                            "description": desc,
                            "location": loc or "",
                            "urls": urls,
                            "start_time": f"{start_t[0]:02d}:{start_t[1]:02d}" if start_t else None,
                            "end_time": f"{end_t[0]:02d}:{end_t[1]:02d}" if end_t else None,
                            "is_all_day": start_t is None,
                            "row": row_idx,
                            "sheet": sheet_name
                        }
                        all_events_by_person[person_key].append(event_data)

        # Sort events by date ascending
        for p in all_events_by_person:
            all_events_by_person[p].sort(key=lambda x: (x["date"], x["start_time"] or "00:00", int(x["row"])))

        return all_events_by_person

def escape_ics_text(text):
    if not text:
        return ""
    text = text.replace("\\", "\\\\")
    text = text.replace(";", "\\;")
    text = text.replace(",", "\\,")
    text = text.replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n")
    return text

def generate_ics(events, calendar_name, calendar_id):
    """
    Генерирует RFC 5545 iCalendar (.ics) файл для подписки.
    """
    now_utc = datetime.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//DreminMS//Yandex Schedule {calendar_id}//RU",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{escape_ics_text(calendar_name)}",
        "X-WR-TIMEZONE:Europe/Moscow",
        "REFRESH-INTERVAL;VALUE=DURATION:PT1H",
        "X-PUBLISHED-TTL:PT1H",
        "BEGIN:VTIMEZONE",
        "TZID:Europe/Moscow",
        "X-LIC-LOCATION:Europe/Moscow",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0300",
        "TZOFFSETTO:+0300",
        "TZNAME:MSK",
        "DTSTART:19700101T000000",
        "END:STANDARD",
        "END:VTIMEZONE"
    ]

    for ev in events:
        d = datetime.datetime.strptime(ev["date"], "%Y-%m-%d").date()
        summary = escape_ics_text(ev["title"])
        desc = escape_ics_text(ev.get("description", ev["full_text"]))
        loc = escape_ics_text(ev["location"])
        
        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{ev['uid']}")
        lines.append(f"DTSTAMP:{now_utc}")
        
        if ev["is_all_day"]:
            # All-day event: DTSTART date, DTEND next date
            next_d = d + datetime.timedelta(days=1)
            lines.append(f"DTSTART;VALUE=DATE:{d.strftime('%Y%m%d')}")
            lines.append(f"DTEND;VALUE=DATE:{next_d.strftime('%Y%m%d')}")
        else:
            sh, sm = [int(x) for x in ev["start_time"].split(":")]
            eh, em = [int(x) for x in ev["end_time"].split(":")]
            dt_start = datetime.datetime(d.year, d.month, d.day, sh, sm)
            dt_end = datetime.datetime(d.year, d.month, d.day, eh, em)
            if dt_end <= dt_start:
                dt_end = dt_start + datetime.timedelta(hours=2)
            lines.append(f"DTSTART;TZID=Europe/Moscow:{dt_start.strftime('%Y%m%dT%H%M%S')}")
            lines.append(f"DTEND;TZID=Europe/Moscow:{dt_end.strftime('%Y%m%dT%H%M%S')}")

        lines.append(f"SUMMARY:{summary}")
        lines.append(f"DESCRIPTION:{desc}")
        if loc:
            lines.append(f"LOCATION:{loc}")
        if ev["urls"]:
            lines.append(f"URL:{ev['urls'][0]}")
        lines.append("STATUS:CONFIRMED")
        lines.append("TRANSP:OPAQUE")
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"

def main():
    print(f"[{datetime.datetime.now()}] Старт синхронизации расписания...")

    # If run in GitHub Actions or table file is missing, download fresh copy
    is_ci = os.getenv("CI") == "true" or os.getenv("GITHUB_ACTIONS") == "true"
    needs_download = is_ci or not os.path.exists(XLSX_PATH) or (datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getmtime(XLSX_PATH))).total_seconds() > 3600

    if needs_download:
        download_file_with_retry(PUBLIC_URL, XLSX_PATH)
    else:
        print(f"Используем кэшированный файл таблицы: {XLSX_PATH}")

    events_by_person = parse_workbook(XLSX_PATH)

    person_titles = {
        "misha": "Расписание съемок (Миша)",
        "tanya": "Расписание съемок (Таня)",
        "ilya": "Расписание съемок (Илья)",
        "dima": "Расписание съемок (Дима)"
    }

    # Generate individual .ics files
    for p_key, events in events_by_person.items():
        ics_content = generate_ics(events, person_titles.get(p_key, p_key), p_key)
        ics_file = os.path.join(BASE_DIR, f"{p_key}.ics")
        with open(ics_file, "w", encoding="utf-8") as f:
            f.write(ics_content)
        print(f"Сгенерирован {ics_file}: {len(events)} событий")

    # Generate all.ics
    all_events = []
    for evs in events_by_person.values():
        all_events.extend(evs)
    all_events.sort(key=lambda x: (x["date"], x["start_time"] or "00:00"))
    
    all_ics = generate_ics(all_events, "Расписание съемок (Вся команда)", "all")
    all_file = os.path.join(BASE_DIR, "all.ics")
    with open(all_file, "w", encoding="utf-8") as f:
        f.write(all_ics)
    print(f"Сгенерирован {all_file}: {len(all_events)} событий")

    # Generate events.json for web PWA
    json_file = os.path.join(BASE_DIR, "events.json")
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump({
            "updated_at": datetime.datetime.now().isoformat(),
            "events_by_person": events_by_person,
            "all_events": all_events
        }, f, ensure_ascii=False, indent=2)
    print(f"Сохранен JSON: {json_file}")

    print("Синхронизация успешно завершена!")

if __name__ == "__main__":
    main()
