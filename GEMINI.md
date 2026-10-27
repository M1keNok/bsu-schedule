# Workspace Rules: BSU Geography Faculty Schedule (bsu_bot)

## 1. Проект и Архитектура
- **Назначение**: Telegram Mini App расписания учебных занятий географического факультета БГУ.
- **Архитектура**: Zero-Build Static Serverless (HTML5 + Vanilla JavaScript + CSS).
- **Хостинг**: GitHub Pages. Никаких Node.js сборщиков (Vite, Webpack, npm-скриптов сборки). Приложение запускается напрямую из статических файлов.
- **Данные**:
  - `schedule_data.json` — расписание занятий, генерируется исключительно скриптом `build_schedule.py`.
  - `teachers_info.json` — справочник преподавателей (кафедра, аудитория, описание).

---

## 2. Стек UI: Tailwind CSS + daisyUI + Liquid Glass
- **Фреймворк компонентов**: **daisyUI** (подключен через CDN) поверх **Tailwind CSS** (Play CDN).
- **Стилистическая парадигма**: **Apple Liquid Glass** (стеклянные карточки с `backdrop-filter: blur(...)`, тонкие полупрозрачные границы `rgba(255, 255, 255, 0.1..0.9)`, внутренние фаски `inner-bevel`, мягкие тени и высокая читаемость).
- **Использование компонентов daisyUI**:
  - Для новых элементов интерфейса и рефакторинга использовать семантические классы daisyUI:
    - Кнопки: `.btn`, `.btn-primary`, `.btn-secondary`, `.btn-ghost`, `.btn-sm`, `.btn-circle`
    - Бейджи и плашки: `.badge`, `.badge-primary`, `.badge-ghost`, `.badge-outline`, `.badge-sm`
    - Модальные окна и шторки: сочетание стилей Liquid Glass и классов daisyUI (`.modal`, `.modal-box`, `.modal-action`)
    - Вкладки и навигация: `.tabs`, `.tabs-boxed`, `.tab`
    - Индикаторы и счетчики: `.indicator`, `.indicator-item`, `.stats`, `.stat`
    - Переключатели и чекбоксы: `.toggle`, `.toggle-primary`, `.checkbox`
    - Спойлеры и аккордеоны: `.collapse`, `.collapse-arrow`
  - Не загромождать DOM лишними обертками; сохранять легковесность и высокую скорость рендеринга.

---

## 3. Синхронизация темы (Dark / Light Mode) и OLED Pure Black
- В приложении реализована светлая и темная темы (автоматически или вручную через переключатель daisyUI Theme Controller).
- **Спецификация темной темы**:
  - Строгий **OLED Pure Black** (`#000000`, `base-100: #000000`, стеклянные карточки `rgba(14, 14, 16, 0.88)`).
  - Никаких синих/индиго паразитных оттенков (`#090c13`, `#6366f1` и т.д. запрещены в темной теме). Использовать нейтральную палитру (Zinc/Slate), монохромный белый текст `#f4f4f5` и органические акценты (изумруд `#10b981` для активных занятий, янтарь `#f59e0b` для перемен).
- **Обязательное правило**: при переключении темы (функции `initTheme()` и `setTheme(mode)`) необходимо синхронизировать:
  1. Класс на `document.documentElement` (`.dark`) для кастомных CSS-переменных.
  2. Атрибут `data-theme` на `document.documentElement`:
     ```javascript
     document.documentElement.classList.toggle('dark', mode === 'dark');
     document.documentElement.setAttribute('data-theme', mode === 'dark' ? 'dark' : 'light');
     ```
  3. Состояние чекбокса `#headerThemeToggle` (daisyUI Theme Controller в шапке).
  Это гарантирует, что все компоненты daisyUI мгновенно меняют цветовую палитру в унисон с интерфейсом приложения.

---

## 4. Офлайн-режим и Service Worker (`sw.js`)
- Приложение должно полноценно запускаться и работать в Telegram WebApp даже в авиарежиме (офлайн).
- Любые подключаемые внешние CDN-библиотеки (`cdn.tailwindcss.com`, `cdn.jsdelivr.net`, `telegram.org`) **обязаны** перехватываться и кэшироваться в `sw.js` по стратегии **Cache-First**:
  - Если ресурс есть в кэше — отдавать немедленно.
  - Если ресурса нет в кэше — загрузить по сети, сохранить клон в `caches` и вернуть клиенту.
- Расписание `schedule_data.json` и `teachers_info.json` используют стратегию **Network-First** с автоматическим откатом на закэшированную версию при отсутствии соединения.
- Оболочка `index.html` обслуживается по стратегии **Stale-While-Revalidate**.

---

## 5. Мобильный UX и Telegram Mini App
- **Безопасные зоны (Safe Areas)**: Учитывать вырезы Dynamic Island и нижнюю панель жестов (`env(safe-area-inset-top)`, `env(safe-area-inset-bottom)`).
- **Тактильный отклик (Haptic Feedback)**:
  - Вызывать `Telegram.WebApp?.HapticFeedback?.impactOccurred('light' | 'medium')` при кликах на важные переключатели, выбор дней, фильтры и модалки.
- **Производительность**:
  - Целевой фреймрейт: стабильные 60–120 FPS на мобильных устройствах.
  - Избегать принудительных reflow/layout thrashing.
  - Анимации строить на `transform` и `opacity` с аппаратным ускорением.
