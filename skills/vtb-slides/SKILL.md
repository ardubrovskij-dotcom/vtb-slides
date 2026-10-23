---
name: vtb-slides
description: Создаёт презентации PowerPoint (.pptx) в фирменном стиле ВТБ — титулы с фото и фирменными плашками, контентные слайды, инфографика (воронка, пирамида, SWOT, матрица 2×2, таймлайн, дорожная карта/Гант, оргструктура, цикл, хаб, Венн, Исикава, айсберг, шкалы прогресса), KPI-карточки, дашборды, таблицы и нативные графики. Используй всякий раз, когда пользователь просит слайды, презентацию, деку, питч, отчёт в слайдах, «в стиле ВТБ», «в корпоративном шаблоне ВТБ», переделать/адаптировать инфографику под ВТБ — даже если слово «ВТБ» не прозвучало, но речь о корпоративной банковской презентации этого бренда.
---

# VTB Slides

Скилл собирает `.pptx` на базе официального шаблона ВТБ (мастер, логотип, плашка номера страницы, 20 титульных макетов) и библиотеки из 37 типов слайдов, перерисованных в фирменной стилистике. Схемы из библиотеки инфографики (72 категории) адаптированы под ВТБ — см. `references/infographics-map.md`.

## Быстрый старт

1. Разложи содержание по слайдам (одна мысль — один слайд), выбери тип каждого слайда по таблице ниже.
2. Напиши JSON-спецификацию (или Python-скрипт) и собери:
   ```bash
   python3 <skill>/scripts/build_deck.py spec.json out.pptx
   ```
   Python API: `sys.path.insert(0, "<skill>/scripts"); from vtb import Deck`.
3. Прочитай предупреждения `WARNING [slide N] text overflow…` — сократи текст или раздели слайд, пересобери.
4. Визуальная проверка (обязательна): `soffice` → PDF → `pdftoppm` → посмотреть каждый слайд (детали — в скилле pptx, раздел QA). Валидатор: `python <pptx-skill>/scripts/office/validate.py out.pptx --original <skill>/assets/vtb_template.pptx`.

Пример со всеми типами: `examples/demo.json` (превью — `examples/demo_preview.jpg`).

## Типы слайдов

| Задача | `type` | Ключевые параметры |
|---|---|---|
| Титул | `cover` | `title` (≤3 строк), `subtitle`, `variant` 1–20 (`references/covers.jpg`) |
| Финальный слайд | `closing` | `title`, `subtitle`, `contacts`, `variant` |
| Разделитель раздела | `section` | `number`, `title`, `subtitle` |
| Повестка | `agenda` | `items` (str или {title,text}), `active` |
| Тезисы с иконками (+фото) | `bullets` | `items` [{icon,title,text,highlight}], `photo`: people/meeting/имя фото/путь |
| Сетка карточек | `cards` | `items` [{icon,title,text}] 2–8, `cols`, `highlight` |
| Колонки с буллетами | `columns` | `cols` [{title,icon,items,highlight}], `color_scheme` navy/mixed |
| Нумерованный список | `numbered` | `items`, `panel` {icon,text}, `header` |
| Два списка (риски/меры) | `two_lists` | `left`/`right` {title, items} |
| Крупные цифры | `stats` | `items` [{value,unit,label,icon}], `text` (вывод) |
| KPI-карточки | `kpi_cards` | `items` [{name,icon,label,value,unit,text \| bullets,highlight}] 1–4 |
| Шкалы прогресса / ADKAR | `progress` | `items` [{label,value 0–100,text,icon}] |
| Этапы с карточками | `process` | `steps` [{title,items \| text}] 2–5 |
| Шевроны / путь клиента | `chevrons` | `steps` [{title,text,icon}] 3–6 |
| Проблема → решение | `flow_rows` | `rows` [{icon,label,text,highlight}] |
| Таймлайн | `timeline` | `events` [{date,title,text,highlight}] 3–8 |
| Дорожная карта / Гант | `roadmap` | `periods`, `rows` [{name,start,end,highlight,label}] |
| Воронка | `funnel` | `stages` [{label,value,text}] 3–6 |
| Пирамида | `pyramid` | `levels` [{label,text}] 3–6 |
| Матрица 2×2 | `matrix` | `quadrants` ×4, `x_axis`, `y_axis`, `highlight` |
| SWOT | `swot` | `strengths`, `weaknesses`, `opportunities`, `threats` |
| Сравнение / было-стало | `comparison` | `left`/`right` {title,items,icon} |
| Тарифы / пакеты | `pricing` | `plans` [{name,price,period,items,highlight}] |
| Центр + спутники | `hub` | `center`, `items` [{title,text,icon}] 4–8 |
| Цикл (PDCA) | `cycle` | `steps` [{title,text}] 3–6, `center` |
| Диаграмма Венна | `venn` | `sets` 2–3, `center` |
| Оргструктура | `org_chart` | `root`, `children` [{name,role,children}] |
| Исикава (причины) | `fishbone` | `effect`, `causes` [{title,items}] 4–6 |
| Айсберг | `iceberg` | `above`, `below` |
| Команда | `team` | `people` [{name,role,text,photo}] 2–5 |
| Таблица | `table` | `header`, `rows`, `highlight_col`, `highlight_row`, `group_rows` |
| График | `chart` | `kind` column/bar/line/pie/doughnut/stacked/stacked_bar/area, `categories`, `series`, `side`, `header`, `conclusion` |
| Каскад / факторный анализ | `waterfall` | `categories`, `values` (приращения), `totals` (индексы итогов), `side`, `conclusion` |
| Дашборд | `dashboard` | `kpis` [{value,label,delta}], `chart` {...}, `bullets` |
| Текст + фото | `text_image` | `text`, `image`, `kicker` |
| Ключевой тезис | `statement` | `text`, `kicker`, `source` |
| Цитата | `quote` | `text`, `author`, `role`, `photo` |

Общие параметры любого контентного слайда: `title`, `note` (строка «Примечание/Источник» внизу), `notes` (заметки докладчика). Полная сигнатура каждого метода — в `references/components.md` и докстрингах `scripts/vtb.py`.

**Выделение в тексте:** `**текст**` → полужирный в фирменном синем (на синих плашках — светло-циановый). Так в шаблоне ВТБ выделяют главное в строке.

**Иконки:** только из фирменного набора ВТБ, 34 шт. (`references/icons.png`): award, bank, briefcase-chart, calculator, calendar, card, case, certificate, check, coins, consultant, cross, document-edit, document-stamp, globe, layers, leader, mail, medal, money, monitor-search, report, search, speedometer, star, strategy, structure, target, team, transform, truck, user-calculator, user-search, wallet. Нет подходящей — не ставь иконку (номер или пусто), не подменяй сторонней.

**Фото:** `assets/photos/` — moscow_city, towers, team, office, mountains, mountains_lake, finance_chart, analytics, globe; можно передать имя или путь к файлу пользователя.

## Правила стиля (кратко; подробно — `references/style-guide.md`)

- Холст 16:9, 13,333×7,5″. Контент: x 0,56–12,77″, y 1,35–6,8″. Заголовок слайда — вверху слева, ПРОПИСНЫМИ, синий #002882, 28 pt (≤2 строк). Логотип и номер страницы ставит макет — не рисуй их.
- Палитра: синий #002882 (доминирует) → циан #00AAFF (акцент, 1 выделение на слайд) → серые для вторичного. Градиенты только внутри фирменной пары синий→тёмно-синий или циан→синий.
- Фирменный мотив — **косой срез 20°** (плашки, панели KPI, параллелограммы на титулах). Не добавляй другие декоративные полосы, тени-«бары», подчёркивания заголовков.
- Карточки: белые, тонкий серый контур, мягкая тень. Разделители — тонкие серые линии.
- Шрифты темы: VTB Group Demi Bold (заголовки) / VTB Group Book (текст). Если у получателя нет шрифтов ВТБ — `"font_mode": "safe"` (всё в Arial).
- Текст: основной 14–16 pt, подписи ≥10 pt. Не уменьшай ниже — режь текст или дели слайд.
- Числа: «12,4 млн», «+18%», неразрывный пробел в тысячах («1 904»). Данные не выдумывай: если цифр нет — спроси или оставь явный плейсхолдер «XX».

## Когда нужно нестандартное

Собери слайд из примитивов `Deck`: `new_slide(title)`, `plate()` (фирменная плашка), `card()`, `text()`, `icon()`, `circle()`, `poly()`, `line()/hline()/vline()`, `picture()` (обрезка, круглая маска), `number_badge()`. Цвета — константы `NAVY, CYAN, …` из `vtb.py`. Держи геометрию в контентной зоне и проверяй рендер.

Карты России/мира из исходного шаблона ВТБ в библиотеку не вошли (векторные карты регионов); для них возьми исходный файл шаблона и правь слайд напрямую через скилл pptx.
