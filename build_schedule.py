import datetime
import hashlib
import json
import os
import re
import subprocess
from docx import Document
import requests

COURSES_CONFIG = {
    "1": {
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/ACkK0WsKu-WA81W7FYHt3YA/1kurs.doc?rlkey=yp9pp95tget8jo9oct0yppp7n&st=971tfg0i&dl=0",
        "raw_file": "1kurs.doc",
        "docx_file": "1kurs.docx",
        "needs_conversion": True,
    },
    "2": {
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/ANeggPvlbzFonxQ6QxDE0qQ/2kurs.doc?rlkey=yp9pp95tget8jo9oct0yppp7n&st=funb9d22&dl=0",
        "raw_file": "2kurs.doc",
        "docx_file": "2kurs.docx",
        "needs_conversion": True,
    },
    "3": {
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/AOtlH0FGL7cybx5XbUDT2IM/3kurs.doc?rlkey=yp9pp95tget8jo9oct0yppp7n&st=s92322wo&dl=0",
        "raw_file": "3kurs.doc",
        "docx_file": "3kurs.docx",
        "needs_conversion": True,
    },
    "4": {
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/AA480GB96NvLPj1FCveHNO4/4kurs.docx?rlkey=yp9pp95tget8jo9oct0yppp7n&st=g9bhzpdc&dl=0",
        "raw_file": "4kurs.docx",
        "docx_file": "4kurs.docx",
        "needs_conversion": False,
    },
}

def format_datetime_ru(dt):
    months = ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
    return f"{dt.day} {months[dt.month - 1]} в {dt.strftime('%H:%M')}"

def download_file(course_id, cfg):
    url = cfg["url"].replace("&dl=0", "&dl=1").replace("?dl=0", "?dl=1")
    res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    res.raise_for_status()
    with open(cfg["raw_file"], "wb") as f:
        f.write(res.content)
    
    lm = res.headers.get("last-modified")
    if lm:
        try:
            dt = datetime.datetime.strptime(lm, "%a, %d %b %Y %H:%M:%S GMT") + datetime.timedelta(hours=3)
            return format_datetime_ru(dt)
        except Exception:
            pass
    return format_datetime_ru(datetime.datetime.now())

def convert_docs(files):
    if not files:
        return
    paths = ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "soffice"]
    cmd = next((p for p in paths if os.path.exists(p) or subprocess.run(["which", p], stdout=subprocess.PIPE).returncode == 0), None)
    if not cmd:
        print("LibreOffice не найден")
        return
    subprocess.run([cmd, "--headless", "--convert-to", "docx", *files, "--outdir", "."], check=True)

def format_time(t):
    digits = re.sub(r"\s+", "", t)
    m = re.match(r"(\d{2})[:.]?(\d{2})[-–](\d{2})[:.]?(\d{2})", digits)
    return f"{m.group(1)}:{m.group(2)}–{m.group(3)}:{m.group(4)}" if m else digits

def clean_cell_text(text):
    if not text:
        return ""
    text = text.replace('\xad', '').replace('\u200b', '')
    text = re.sub(r"([а-яА-ЯёЁa-zA-Z]+)[-–—]\s*\n\s*([а-яА-ЯёЁa-zA-Z]+)", r"\1\2", text)
    text = re.sub(r"(ОБЩЕЕ\s+ЗЕМЛЕВЕДЕ)\s*\n\s*(НИЕ)", r"\1\2", text, flags=re.I)
    text = re.sub(r"(ОБЩЕЕ)\s*\n\s*(ЗЕМЛЕВЕДЕНИЕ)", r"\1 \2", text, flags=re.I)
    text = re.sub(r"(ИСТОРИЧЕ)\s*\n\s*(СКАЯ)", r"\1\2", text, flags=re.I)
    text = re.sub(r"(ЯЗЫК\s+ПРОФЕССИО)\s*\n\s*(НАЛЬНОГО)", r"\1\2", text, flags=re.I)
    text = re.sub(r"(ГРУНТОВЕ)\s*\n\s*(ДЕНИЯ)", r"\1\2", text, flags=re.I)

    legit = [
        'социально-эконом',
        'физико-географ',
        'экономико-географ',
        'научно-технич',
        'аппаратно-программ',
        'гис-картограф',
        'гис-технолог',
        'соц-эконом',
        'препод-стаж',
        'препод.-стаж',
        'творонович-севрук',
    ]

    def protect(m):
        full = m.group(0)
        lower = full.lower()
        for l in legit:
            if l in lower:
                return full.replace('-', '__HYPH__').replace('–', '__HYPH__').replace('—', '__HYPH__')
        return full

    text = re.sub(r'[а-яА-ЯёЁa-zA-Z]+[-–—][а-яА-ЯёЁa-zA-Z]+', protect, text)
    while re.search(r'([а-яА-ЯёЁa-zA-Z])[-–—]+([а-яА-ЯёЁa-zA-Z])', text):
        text = re.sub(r'([а-яА-ЯёЁa-zA-Z])[-–—]+([а-яА-ЯёЁa-zA-Z])', r'\1\2', text)
    text = text.replace('__HYPH__', '-')

    typos = [
        (re.compile(r'гибробиолог', re.I), 'гидробиолог'),
        (re.compile(r'автоматизи-?рованые', re.I), 'автоматизированные'),
        (re.compile(r'геоинформа-?цинное', re.I), 'геоинформационное'),
        (re.compile(r'иностраный\s+язык', re.I), 'иностранный язык'),
        (re.compile(r'генезис\s+месторожде-?ние\s+полезных', re.I), 'генезис месторождений полезных'),
    ]
    for p, r in typos:
        text = p.sub(r, text)

    lines = [re.sub(r"\s+", " ", l).strip() for l in text.split("\n") if l.strip()]
    return "\n".join([l for l in lines if not (len(re.sub(r'[:. –—-]', '', l)) in (7, 8) and re.sub(r'[:. –—-]', '', l).isdigit())])

def parse_docx(path):
    if not os.path.exists(path):
        return {}
    doc = Document(path)
    res = {}
    for table in doc.tables:
        if len(table.rows) < 3:
            continue
        group_cols = {}
        h_idx = 0
        for r_idx in range(min(3, len(table.rows))):
            for c_idx, cell in enumerate(table.rows[r_idx].cells):
                txt = re.sub(r"\s+", " ", cell.text).strip().upper()
                if "ГРУППА" in txt:
                    group_cols[c_idx] = txt.split("\n")[0].strip()
            if group_cols:
                h_idx = r_idx
                break
        if not group_cols:
            continue
        cur_day, cur_time = None, None
        for row in table.rows[h_idx + 1:]:
            cells = row.cells
            if len(cells) < 3:
                continue
            day_txt = re.sub(r"\s+", " ", cells[0].text).strip().upper()
            if day_txt and "ВРЕМЯ" not in day_txt and "ДНИ" not in day_txt and len(day_txt) > 2:
                cur_day = day_txt
            if not cur_day:
                continue
            time_txt = cells[1].text.strip()
            if time_txt and any(c.isdigit() for c in time_txt):
                cur_time = format_time(time_txt)
            if not cur_time:
                continue
            for col_idx, g_name in group_cols.items():
                if col_idx >= len(cells):
                    continue
                lesson = clean_cell_text(cells[col_idx].text)
                if "СПЕЦИАЛЬНОСТЬ" in lesson.upper() or "НАПРАВЛЕНИЕ" in lesson.upper():
                    continue
                res.setdefault(g_name, {}).setdefault(cur_day, {}).setdefault(cur_time, [])
                if lesson and lesson not in res[g_name][cur_day][cur_time]:
                    res[g_name][cur_day][cur_time].append(lesson)
    return res

def build_all():
    print("Загрузка файлов деканата...")
    updated_times = {}
    to_convert = []

    for cid, cfg in COURSES_CONFIG.items():
        updated_times[cid] = download_file(cid, cfg)
        if cfg["needs_conversion"]:
            to_convert.append(cfg["raw_file"])

    convert_docs(to_convert)

    full_data = {"courses": {}, "updated_times": updated_times}
    for cid, cfg in COURSES_CONFIG.items():
        full_data["courses"][cid] = parse_docx(cfg["docx_file"])

    with open("schedule_data.json", "w", encoding="utf-8") as f:
        json.dump(full_data, f, ensure_ascii=False, indent=2)

    # Clean up temporary doc and docx files
    for cid, cfg in COURSES_CONFIG.items():
        for fpath in [cfg.get("raw_file"), cfg.get("docx_file")]:
            if fpath and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass

    print("schedule_data.json успешно собран и очищен от временных файлов!")

if __name__ == "__main__":
    build_all()
