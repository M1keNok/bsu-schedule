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
    cache_buster = int(datetime.datetime.now().timestamp())
    url = f"{url}&_ts={cache_buster}"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache"
    }
    res = requests.get(url, headers=headers, timeout=30)
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
        (re.compile(r'высшая\s+математика\s+с\s+основами\s+информатики', re.I), 'ВЫШМАТ'),
        (re.compile(r'высшая\s+математика', re.I), 'ВЫШМАТ'),
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
                if lesson:
                    # Separate multiple subjects crammed into a single cell by Decanat
                    sub_lessons = re.split(r"\n(?=(?:Немецкий язык|Русский язык как иностранный|ИНОСТРАННЫЙ ЯЗЫК|Французский язык))", lesson, flags=re.I)
                    for sub in sub_lessons:
                        sub = sub.strip()
                        if sub and sub not in res[g_name][cur_day][cur_time]:
                            res[g_name][cur_day][cur_time].append(sub)
    return res

def detect_lesson_language(text):
    if not text:
        return None
    lower = text.lower()
    # 1. French: teacher Гулезова or word франц
    if "гулезов" in lower or "франц" in lower:
        return "FR"
    # 2. German: teachers Пригодич, Хадасевич or word нем (avoiding неметаллическ)
    if ("нем" in lower and "неметал" not in lower) or "пригодич" in lower or "хадасевич" in lower:
        return "DE"
    # 3. Russian as foreign language (RKI)
    if "русск" in lower or "рки" in lower:
        if any(w in lower for w in ["иностр", "язык", "рки"]):
            return "RKI"
    # 4. English: teachers or explicit tokens or generic foreign language
    if any(t in lower for t in ["шарейко", "селиванова", "дворкина", "боровикова", "давыдова", "пышняк", "англ", "eduengscience"]):
        return "EN"
    if "иностранн" in lower or "язык профессионального общения" in lower or "практика иностранного языка" in lower:
        return "EN"
    return None

def propagate_stream_languages(courses_data):
    """
    In university schedules, foreign language stream classes (French, German)
    are often entered by Decanat in only 1 or 2 group columns, but attendable
    by students from any group of the same course who study that language.
    This function shares stream language lessons across groups of the same course
    if the group's slot is empty or only contains other foreign language / PE classes.
    """
    for cid, groups in courses_data.items():
        if not isinstance(groups, dict):
            continue
        stream_map = {}
        for g_name, days in groups.items():
            if not isinstance(days, dict):
                continue
            for day_key, slots in days.items():
                if not isinstance(slots, dict):
                    continue
                for time_str, lessons in slots.items():
                    for l in lessons:
                        lang = detect_lesson_language(l)
                        if lang in ("FR", "DE"):
                            key = (day_key, time_str, lang)
                            if key not in stream_map:
                                stream_map[key] = l

        for (day_key, time_str, lang), stream_lesson in stream_map.items():
            for g_name, days in groups.items():
                if not isinstance(days, dict) or day_key not in days:
                    continue
                slots = days[day_key]
                current_lessons = slots.get(time_str, [])
                has_lang = any(detect_lesson_language(l) == lang for l in current_lessons)
                if has_lang:
                    continue
                is_available = not current_lessons or all(
                    detect_lesson_language(l) is not None or bool(re.search(r"физическ.*культ", l, re.I))
                    for l in current_lessons
                )
                if is_available:
                    slots.setdefault(time_str, []).append(stream_lesson)

def parse_day_date(day_key, ref_date=None):
    if not day_key:
        return None
    if not ref_date:
        ref_date = datetime.date.today()
    m = re.search(r"(\d{1,2})[.](\d{1,2})(?:[.](\d{2,4}))?", day_key)
    if not m:
        return None
    d = int(m.group(1))
    mo = int(m.group(2))
    if m.group(3):
        y = int(m.group(3))
        if y < 100:
            y += 2000
    else:
        y = ref_date.year
        if ref_date.month == 12 and mo == 1:
            y += 1
        elif ref_date.month == 1 and mo == 12:
            y -= 1
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None

def merge_schedule_history(new_data, existing_path="schedule_data.json", max_history_days=14):
    today = datetime.date.today()
    old_data = {}
    if os.path.exists(existing_path):
        try:
            with open(existing_path, "r", encoding="utf-8") as f:
                old_data = json.load(f)
        except Exception as e:
            print(f"Предупреждение: не удалось прочитать {existing_path}: {e}")

    merged_courses = {}
    all_courses = set(old_data.get("courses", {}).keys()) | set(new_data.get("courses", {}).keys())

    for cid in sorted(all_courses):
        merged_courses[cid] = {}
        old_groups = old_data.get("courses", {}).get(cid, {})
        new_groups = new_data.get("courses", {}).get(cid, {})
        all_groups = set(old_groups.keys()) | set(new_groups.keys())

        for g_name in sorted(all_groups):
            merged_days = {}

            # 1. Take historical days that are not older than max_history_days
            for day_key, slots in old_groups.get(g_name, {}).items():
                dt = parse_day_date(day_key, today)
                if dt and (today - dt).days > max_history_days:
                    continue  # Prune days older than 14 days
                merged_days[day_key] = slots

            # 2. Overwrite / append newly parsed days
            for day_key, slots in new_groups.get(g_name, {}).items():
                dt = parse_day_date(day_key, today)
                if dt and (today - dt).days > max_history_days:
                    continue
                merged_days[day_key] = slots

            # Chronological sort of days
            def sort_key(k):
                d = parse_day_date(k, today)
                return d if d else datetime.date.max

            sorted_days = {k: merged_days[k] for k in sorted(merged_days.keys(), key=sort_key)}
            merged_courses[cid][g_name] = sorted_days

    # Merge updated_times (new takes precedence)
    merged_updated_times = dict(old_data.get("updated_times", {}))
    merged_updated_times.update(new_data.get("updated_times", {}))

    return {
        "courses": merged_courses,
        "updated_times": merged_updated_times
    }

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

    propagate_stream_languages(full_data["courses"])

    final_data = merge_schedule_history(full_data, "schedule_data.json", max_history_days=14)
    propagate_stream_languages(final_data["courses"])

    with open("schedule_data.json", "w", encoding="utf-8") as f:
        json.dump(final_data, f, ensure_ascii=False, indent=2)

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
