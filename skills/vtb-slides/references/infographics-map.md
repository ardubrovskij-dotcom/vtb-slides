# Адаптация библиотеки инфографики под стиль ВТБ

Источник: папка Google Drive «ВТБ слайды / 01-Powerpoint-Infographics» — 72 категории, 98 файлов .pptx, 1 844 слайда. Исходные слайды сделаны в сторонней палитре (зелёный/бирюзовый/оранжевый), с заголовком по центру и шрифтами Poppins/Montserrat. Поэтому они **не копируются**, а сведены к схемам-архетипам, которые перерисованы компонентами `vtb.Deck` в палитре, сетке и мотиве ВТБ (синий/циан, косой срез 20°, карточки с тенью, заголовок слева ПРОПИСНЫМИ).

Принцип выбора: определи, **какую логику** показывает исходная схема (последовательность, иерархия, сравнение, доли, связь с центром…), и возьми компонент с той же логикой. Декоративные формы исходника (3D, силуэты людей, лампочки, головы, пазлы) в стиле ВТБ не используются: смысл передаётся структурой, а не иллюстрацией.

## Категории → компоненты

| Категория библиотеки (слайдов) | Логика схемы | Компонент ВТБ |
|---|---|---|
| ADKAR (20) | 5 этапов изменений / уровень готовности | `progress` (уровни, %), `chevrons` (этапы), `cards` |
| Agenda (25) | список разделов | `agenda` |
| Agile (20), SCRUM (20) | итеративный цикл | `cycle`, `process` |
| Arrow (20), Flow (20) | последовательность, направление | `chevrons`, `flow_rows`, `process` |
| Brain (20) | группы идей вокруг темы | `hub`, `cards` |
| Budget (20), Profit and Loss (20), Finance (20) | структура и динамика денег | `waterfall`, `chart` (column/stacked), `table`, `stats` |
| Business Plan (50), Business (50), Strategy (20) | смешанные бизнес-схемы | `cards`, `columns`, `numbered`, `matrix`, `process`, `stats` |
| Buyer Persona (20) | портрет клиента | `team` (1–2 карточки) + `columns` (потребности/боли) или `text_image` |
| Calendar 2020/2021 (108) | календарная сетка | `roadmap` (периоды по месяцам/кварталам) или `table` |
| Chart (36), Pie Chart (20), Mekko Chart (20) | количественные данные | `chart` (column/bar/line/pie/doughnut/stacked/stacked_bar/area) |
| Waterfall Chart (30) | вклад факторов в итог | `waterfall` |
| Circle (20), Mindmap (20) | центр и ветви | `hub` |
| Cycle (16) | замкнутый процесс | `cycle` |
| Company Profile (51) | о компании: цифры, история, команда | `stats`, `timeline`, `team`, `text_image`, `cards` |
| Comparison (20) | 2 варианта | `comparison`, `table` (`highlight_col`) |
| Customer Journey Map (20) | этапы пути клиента | `chevrons` (этапы) + `table` (этап × действие/боль/решение) |
| Dashboard (20), KPI (20) | набор показателей | `dashboard`, `kpi_cards`, `stats` |
| Decision Tree (20), Org Chart (20) | дерево | `org_chart` (до 3 уровней) |
| Digital Marketing (20), Marketing (60), Social Media (21) | каналы, метрики, воронки | `funnel`, `cards`, `chart`, `stats` |
| E-Learning (20), Education (20), Training (19) | программа обучения | `process`, `timeline`, `cards` |
| Fishbone (20) | причины проблемы | `fishbone` |
| Funnel (35) | конверсия по этапам | `funnel` |
| Gantt (27), Product Roadmap (30), Roadmap (30), Project Management (20) | работы во времени | `roadmap`, `timeline` |
| Goals (20) | цели и шаги к ним | `pyramid` (иерархия), `numbered`, `chevrons` |
| Harvey Ball (20) | степень выполнения | `progress`, `table` с текстовыми оценками |
| Human Resources (50) | структура, роли, процессы HR | `org_chart`, `team`, `process`, `cards` |
| Iceberg (20) | видимое и скрытое | `iceberg` |
| Matrix (20) | 2×2 / сетка оценок | `matrix`, `table` |
| Milestones (20), Timeline (30) | вехи | `timeline` |
| Pricing Table (25) | пакеты и опции | `pricing`, `table` |
| Process (30) | шаги | `process`, `chevrons` |
| Puzzle (20) | составные части целого | `cards`, `hub` |
| Pyramid (16) | иерархия уровней | `pyramid` |
| Risk Management (20) | риски, вероятность × влияние | `matrix` (оси вероятность/влияние), `two_lists` (риски/меры), `table` |
| Supply Chain (20), Value Chain (20), Manufacturing (20) | цепочка звеньев | `chevrons`, `process`, `flow_rows` |
| SWOT (20) | 4 квадранта анализа | `swot` |
| Table (25) | табличные данные | `table` |
| Technology (30) | архитектура, компоненты | `hub`, `columns`, `cards` |
| Venn (20) | пересечение множеств | `venn` |
| World Map (20) | география | нет готового компонента: карты есть в исходном шаблоне ВТБ (Россия, ФО, регионы мира, слайды 69–88) — правь их через скилл pptx |
| Coronavirus (5), Ecology (20), Fitness (20), Food (66), Health (20), Healthcare (40), Medical (20), Real Estate (45), Travel (20) | отраслевые иллюстрации | тематические иллюстрации в стиль ВТБ не переносятся; содержание — через `cards`, `stats`, `chart`, `process` |
| Resume Templates (14) | резюме A4 | вне формата 16:9; для слайда о человеке — `team` или `text_image` |

## Как адаптировать конкретный слайд из библиотеки

1. Открой исходный слайд (`markitdown file.pptx` или рендер) и сформулируй его логику одной фразой: «5 этапов слева направо», «центр и 6 элементов», «4 квадранта».
2. Найди компонент в таблице выше. Не хватает элементов (например, 7 шагов в `chevrons`) — раздели на 2 слайда или выбери `process`/`timeline`.
3. Перенеси **текст и данные**, а не форму: иконки — из фирменного набора ВТБ, цвета — из палитры ВТБ, одно выделение — циан.
4. Нет подходящего компонента — собери слайд из примитивов (`plate`, `card`, `poly`, `circle`, `text`, `icon`) по правилам `style-guide.md`.
