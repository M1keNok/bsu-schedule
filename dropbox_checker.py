#!/usr/bin/env python3
"""
dropbox_checker.py
Ultra-fast, lightweight change detector for BSU Geography Faculty schedule files on Dropbox.
Performs concurrent HTTP Range/HEAD checks to inspect ETags and Content-Range (size)
in ~1-2 seconds without downloading full documents.
"""

import sys
import os
import json
import time
import argparse
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor
import requests

COURSES_CONFIG = {
    "1": {
        "name": "1 курс",
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/ACkK0WsKu-WA81W7FYHt3YA/1kurs.doc?rlkey=yp9pp95tget8jo9oct0yppp7n&st=971tfg0i&dl=1",
        "raw_file": "1kurs.doc",
        "docx_file": "1kurs.docx",
    },
    "2": {
        "name": "2 курс",
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/ANeggPvlbzFonxQ6QxDE0qQ/2kurs.doc?rlkey=yp9pp95tget8jo9oct0yppp7n&st=funb9d22&dl=1",
        "raw_file": "2kurs.doc",
        "docx_file": "2kurs.docx",
    },
    "3": {
        "name": "3 курс",
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/AOtlH0FGL7cybx5XbUDT2IM/3kurs.doc?rlkey=yp9pp95tget8jo9oct0yppp7n&st=s92322wo&dl=1",
        "raw_file": "3kurs.doc",
        "docx_file": "3kurs.docx",
    },
    "4": {
        "name": "4 курс",
        "url": "https://www.dropbox.com/scl/fo/ltk4mvd0monarv8c9k70a/AA480GB96NvLPj1FCveHNO4/4kurs.docx?rlkey=yp9pp95tget8jo9oct0yppp7n&st=g9bhzpdc&dl=1",
        "raw_file": "4kurs.docx",
        "docx_file": "4kurs.docx",
    },
}

STATE_FILE = "dropbox_state.json"

def fetch_single_metadata(cid: str, cfg: dict, retries: int = 2) -> tuple:
    """Fetches HTTP headers (ETag, size) for a single course link on Dropbox."""
    url = cfg["url"]
    headers = {
        "Range": "bytes=0-0",
        "User-Agent": "Mozilla/5.0 (BSU-Schedule-Checker/2.0)",
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
    }
    for attempt in range(retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=12, allow_redirects=True)
            if resp.status_code in (200, 206):
                cr = resp.headers.get("content-range", "")
                size = 0
                if "/" in cr:
                    try:
                        size = int(cr.split("/")[-1])
                    except ValueError:
                        pass
                if not size:
                    cl = resp.headers.get("content-length", "")
                    if cl and cl.isdigit():
                        size = int(cl)
                etag = resp.headers.get("etag", "").strip('"').strip()
                return cid, {
                    "etag": etag,
                    "size": size,
                    "status": resp.status_code,
                    "checked_at": datetime.now(timezone.utc).isoformat()
                }
        except Exception as e:
            if attempt == retries:
                return cid, {"error": str(e), "etag": "", "size": 0, "status": 0}
            time.sleep(1)
    return cid, {"error": "unknown", "etag": "", "size": 0, "status": 0}

def get_current_dropbox_state() -> dict:
    """Concurrently fetches metadata for all courses."""
    results = {}
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(fetch_single_metadata, cid, cfg) for cid, cfg in COURSES_CONFIG.items()]
        for f in futures:
            cid, meta = f.result()
            results[cid] = meta
    return results

def load_saved_state(path: str = STATE_FILE) -> dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_state(state: dict, path: str = STATE_FILE):
    minsk_tz = timezone(timedelta(hours=3))
    now_minsk = datetime.now(minsk_tz)
    full_state = {
        "last_checked_utc": datetime.now(timezone.utc).isoformat(),
        "last_checked_minsk": now_minsk.strftime("%Y-%m-%d %H:%M:%S MSK"),
        "courses": state
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(full_state, f, ensure_ascii=False, indent=2)

def compare_states(current: dict, saved: dict) -> tuple:
    """
    Compares current Dropbox metadata with saved state.
    Returns (has_changes, changed_courses, details)
    """
    saved_courses = saved.get("courses", {})
    if not saved_courses:
        # No saved state -> treat as changed to bootstrap state
        return True, list(COURSES_CONFIG.keys()), "Initial run (no previous state found)"

    changed = []
    details = []

    for cid, cur_meta in current.items():
        if cur_meta.get("error"):
            # Skip error in single check or log
            continue
        old_meta = saved_courses.get(cid, {})
        old_etag = old_meta.get("etag", "")
        old_size = old_meta.get("size", 0)

        cur_etag = cur_meta.get("etag", "")
        cur_size = cur_meta.get("size", 0)

        if not old_etag:
            changed.append(cid)
            details.append(f"Курс {cid}: новый файл в состоянии")
        elif cur_etag and old_etag and cur_etag != old_etag:
            changed.append(cid)
            details.append(f"Курс {cid}: ETag изменился ({old_etag} -> {cur_etag})")
        elif cur_size and old_size and cur_size != old_size:
            changed.append(cid)
            details.append(f"Курс {cid}: размер изменился ({old_size} -> {cur_size} байт)")

    return len(changed) > 0, changed, "; ".join(details) if details else "No changes"

def run_build():
    """Runs build_schedule.py to download, convert and parse the full schedule."""
    import subprocess
    print("🚀 Запуск сборщика расписания build_schedule.py...")
    res = subprocess.run([sys.executable, "build_schedule.py"])
    return res.returncode == 0

def main():
    parser = argparse.ArgumentParser(description="BSU Dropbox schedule change detector")
    parser.add_argument("--check-only", action="store_true", help="Exit code 0 if up to date, 10 if changes detected")
    parser.add_argument("--update-state", action="store_true", help="Fetch and save current Dropbox state without building")
    parser.add_argument("--sync-if-changed", action="store_true", help="Check Dropbox, run build_schedule.py if changed, and update state")
    parser.add_argument("--force-sync", action="store_true", help="Unconditionally run build_schedule.py and update state")
    parser.add_argument("--status", action="store_true", help="Print human-readable comparison status")
    args = parser.parse_args()

    if args.force_sync:
        print("⚡ Принудительная синхронизация...")
        if run_build():
            cur_state = get_current_dropbox_state()
            save_state(cur_state)
            print("✅ Расписание собрано, состояние dropbox_state.json сохранено.")
            sys.exit(0)
        else:
            print("❌ Ошибка при сборке расписания.")
            sys.exit(1)

    print("🔍 Проверка состояния файлов в Dropbox...")
    t0 = time.time()
    current_state = get_current_dropbox_state()
    elapsed = time.time() - t0
    saved_state = load_saved_state()

    has_changes, changed_courses, reason = compare_states(current_state, saved_state)
    print(f"⏱️ Проверка Dropbox заняла: {elapsed:.2f} сек.")

    if args.status or (not args.check_only and not args.update_state and not args.sync_if_changed):
        print(f"📊 Результат проверки: {'Есть изменения' if has_changes else 'Изменений нет'}")
        print(f"Причина: {reason}")
        for cid, cfg in COURSES_CONFIG.items():
            cur = current_state.get(cid, {})
            old = saved_state.get("courses", {}).get(cid, {})
            ch = "🔄 ИЗМЕНЕН" if cid in changed_courses else "✔️ Без изменений"
            print(f"  [{cfg['name']}] {ch} | ETag: {cur.get('etag')} (был: {old.get('etag')}) | Размер: {cur.get('size')} байт")

    if args.update_state:
        save_state(current_state)
        print("💾 Текущее состояние успешно сохранено в dropbox_state.json")
        sys.exit(0)

    if args.check_only:
        if has_changes:
            print(f"🔔 ОБНАРУЖЕНЫ ИЗМЕНЕНИЯ: {reason}")
            sys.exit(10)
        else:
            print("✅ Все файлы в Dropbox актуальны, изменений нет.")
            sys.exit(0)

    if args.sync_if_changed:
        if has_changes:
            print(f"🔔 Изменения обнаружены ({reason}). Запускаем сборку...")
            if run_build():
                save_state(current_state)
                print("✅ Расписание обновлено, состояние сохранено.")
                sys.exit(0)
            else:
                print("❌ Ошибка при сборке расписания.")
                sys.exit(1)
        else:
            print("✅ Файлы в Dropbox не менялись. Сборка не требуется.")
            sys.exit(0)

if __name__ == "__main__":
    main()
