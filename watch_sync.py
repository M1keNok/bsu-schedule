#!/usr/bin/env python3
"""
watch_sync.py
Ultra-fast continuous watcher for Decanat Dropbox schedule changes.
Checks Dropbox every 45-60 seconds.
When a change is detected:
  - Builds locally with build_schedule.py (if LibreOffice is present) and pushes to GitHub
  - OR sends repository_dispatch to GitHub Actions to build in the cloud.
"""

import sys
import os
import time
import argparse
import subprocess
import requests
from datetime import datetime, timezone, timedelta
from dropbox_checker import get_current_dropbox_state, load_saved_state, save_state, compare_states

def get_github_token():
    token = os.getenv("GITHUB_TOKEN")
    if token:
        return token
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("GITHUB_TOKEN="):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""

DEFAULT_REPO = "M1keNok/bsu-schedule"

def check_local_libreoffice():
    paths = ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "soffice"]
    return any(os.path.exists(p) or subprocess.run(["which", p], stdout=subprocess.PIPE, stderr=subprocess.PIPE).returncode == 0 for p in paths)

def trigger_github_dispatch(token=None, repo=DEFAULT_REPO, event_type="sync-schedule"):
    if not token:
        token = get_github_token()
    if not token:
        print("⚠️ GITHUB_TOKEN не задан. Невозможно отправить dispatches в GitHub Actions.")
        return False
    url = f"https://api.github.com/repos/{repo}/dispatches"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "BSU-Schedule-Watcher"
    }
    data = {"event_type": event_type}
    try:
        r = requests.post(url, headers=headers, json=data, timeout=10)
        if r.status_code == 204:
            print(f"📡 [GitHub Actions] Событие {event_type} успешно отправлено!")
            return True
        else:
            print(f"⚠️ Ошибка вызова GitHub API: {r.status_code} {r.text}")
    except Exception as e:
        print(f"⚠️ Ошибка сети при отправке dispatches: {e}")
    return False

def local_build_and_push():
    print("🛠️ Локальная сборка расписания...")
    t0 = time.time()
    res = subprocess.run([sys.executable, "build_schedule.py"])
    if res.returncode != 0:
        print("❌ Ошибка при выполнении build_schedule.py")
        return False
    print(f"✅ Расписание успешно собрано за {time.time() - t0:.2f} сек.")

    # Git commit and push
    subprocess.run(["git", "add", "schedule_data.json", "dropbox_state.json"])
    diff = subprocess.run(["git", "diff", "--staged", "--quiet"])
    if diff.returncode != 0:
        print("📤 Отправка обновленного расписания в GitHub...")
        subprocess.run(["git", "commit", "-m", "Auto-update schedule from Decanat [instant sync]"])
        push = subprocess.run(["git", "push"])
        if push.returncode == 0:
            print("🚀 Изменения успешно запушены в GitHub!")
            return True
        else:
            print("⚠️ Ошибка push в репозиторий.")
            return False
    else:
        print("ℹ️ Нет изменений для коммита в git.")
        return True

def watch_loop(interval=45, force_cloud=False):
    has_local_lo = check_local_libreoffice() and not force_cloud
    minsk_tz = timezone(timedelta(hours=3))
    print(f"👀 Запуск мониторинга Dropbox деканата...")
    print(f"   Интервал проверки: {interval} сек.")
    print(f"   Режим сборки: {'Локальный LibreOffice (сверхбыстро ~15 сек)' if has_local_lo else 'Облачный (GitHub Actions repository_dispatch)'}")

    while True:
        try:
            now_str = datetime.now(minsk_tz).strftime("%H:%M:%S")
            cur_state = get_current_dropbox_state()
            saved_state = load_saved_state()
            has_changes, changed, details = compare_states(cur_state, saved_state)

            if has_changes:
                print(f"[{now_str}] 🔔 ДЕКАНАТ ОБНОВИЛ ФАЙЛЫ В DROPBOX! Курсы: {changed}")
                print(f"   Причина: {details}")

                if has_local_lo:
                    success = local_build_and_push()
                else:
                    success = trigger_github_dispatch()

                if success:
                    save_state(cur_state)
                    print(f"[{now_str}] 🎉 Синхронизация завершена успешно.")
            else:
                print(f"[{now_str}] ✔️ Dropbox проверен: изменений нет.")

        except KeyboardInterrupt:
            print("\n🛑 Мониторинг остановлен пользователем.")
            break
        except Exception as e:
            print(f"⚠️ Ошибка в цикле проверки: {e}")

        time.sleep(interval)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="BSU Schedule Dropbox Watcher")
    parser.add_argument("--interval", type=int, default=45, help="Check interval in seconds (default: 45)")
    parser.add_argument("--cloud", action="store_true", help="Force cloud build via GitHub Actions instead of local")
    parser.add_argument("--once", action="store_true", help="Check once and exit")
    args = parser.parse_args()

    if args.once:
        cur_state = get_current_dropbox_state()
        saved_state = load_saved_state()
        has_changes, changed, details = compare_states(cur_state, saved_state)
        if has_changes:
            print(f"Обнаружены изменения: {changed}")
            if check_local_libreoffice() and not args.cloud:
                local_build_and_push()
            else:
                trigger_github_dispatch()
            save_state(cur_state)
        else:
            print("Изменений нет.")
    else:
        watch_loop(interval=args.interval, force_cloud=args.cloud)
