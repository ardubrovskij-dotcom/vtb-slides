# Справочник компонентов `vtb.Deck`

Сгенерировано из докстрингов `scripts/vtb.py`. Размеры — в дюймах (холст 13,333×7,5″). Любой метод слайда принимает `note` (строка источника/примечания внизу) и `notes` (заметки докладчика).

## Слайды

### `cover(title, subtitle='', variant=4, notes=None)`

Титульный слайд. variant 1–20 (см. references/covers.md):
плашка на фото: 1,2,5,6,9,10,11,12,13,17,18; тёмная панель с диагональю: 3,4,7,8,14,15,16,19,20.

### `closing(title='Спасибо за внимание!', subtitle='', variant=3, contacts=None, notes=None)`

Финальный слайд на основе обложки. contacts: list[str] — строки контактов.

### `section(title, number=None, subtitle='', notes=None)`

Разделитель раздела: тёмный фирменный фон, номер в циановой плашке.

### `agenda(items, title='Содержание', active=None, note=None, notes=None)`

Повестка: список разделов с номерами. active — индекс (с 0) текущего раздела (выделяется плашкой).

### `bullets(title, items, photo=None, note=None, notes=None)`

Строки «иконка + заголовок + текст» с разделителями (слайд 1–2 шаблона).
items: [{icon, title, text}], 2–5 шт. photo: 'people' | 'meeting' | путь к картинке | None.
Первый элемент с highlight=True выделяется синей плашкой.

### `text_image(title, text, image, note=None, notes=None, kicker=None)`

Текст слева, фото справа с диагональю. text: str | list (абзацы); **выделение**.

### `statement(title, text, kicker=None, source=None, note=None, notes=None)`

Ключевой тезис / цитата крупно: циановая плашка-маркер слева, текст 28pt.

### `quote(text, author, role='', title=None, photo=None, note=None, notes=None)`

Цитата: крупные кавычки в плашке, текст, автор (с фото по желанию).

### `cards(title, items, cols=None, highlight=None, note=None, notes=None)`

Сетка карточек «иконка + заголовок + текст» (слайды 3–5). 2–8 элементов.
highlight: индекс карточки, выделяемой циановым параллелограммом.

### `columns(title, cols, color_scheme='navy', note=None, notes=None)`

Колонки с плашками-заголовками и буллетами в карточках (слайды 21–23).
cols: [{title, icon?, items: [str], highlight?}] 2–4 шт. color_scheme: navy | mixed (последняя циан).

### `numbered(title, items, panel=None, header=None, accent='navy', note=None, notes=None)`

Нумерованный список 01–07 с разделителями (слайды 13–15).
panel: {icon?, text} — циановая панель слева. header — плашка над списком.

### `two_lists(title, left, right, note=None, notes=None)`

Два нумерованных списка в карточках с плашками (слайд 13). left/right: {title, items}.

### `kpi_cards(title, items, note=None, notes=None)`

Карточки показателей с диагональной синей панелью (слайд 30). 2–4 шт.
items: [{name, icon?, value?, unit?, text?, bullets?: [str], label?, highlight?}]

### `stats(title, items, text=None, note=None, notes=None)`

Крупные цифры в ряд (2–4): значение 54pt, подпись, иконка. text — вывод под рядом.

### `progress(title, items, note=None, notes=None)`

Горизонтальные шкалы прогресса (ADKAR/навыки/выполнение плана).
items: [{label, value: 0–100, text?, icon?}]

### `process(title, steps, label='Шаг', note=None, notes=None)`

Этапы с шевронными плашками «ШАГ 1 → 2 → 3» и карточками (слайд 63). 2–5 шагов.
steps: [{title, text?, items?: [str]}]

### `chevrons(title, steps, note=None, notes=None)`

Поток шевронов с иконками и описанием под ними (категории Arrow/Flow/Process). 3–6 шагов.
steps: [{title, text, icon?}]

### `flow_rows(title, rows, note=None, notes=None)`

Строки «иконка в квадрате → стрелка → пояснение» (слайды 33–34). 2–5 строк.
rows: [{icon?, label, text, highlight?}]

### `timeline(title, events, note=None, notes=None)`

Горизонтальная шкала событий, подписи чередуются сверху/снизу (слайды 64–66). 3–8 событий.
events: [{date, title, text?, highlight?}]

### `roadmap(title, periods, rows, header='Направление', note=None, notes=None)`

Диаграмма Ганта / дорожная карта (слайды 67–68).
periods: ['Q1', 'Q2', ...]; rows: [{name, start, end, highlight?, label?}] (start/end — индексы периодов, end включительно).

### `funnel(title, stages, note=None, notes=None)`

Воронка (слайд 62): сужающиеся сегменты от синего к циану, подписи чередуются слева/справа.
stages: [{label, text?, value?}] 3–6 шт.

### `pyramid(title, levels, note=None, notes=None)`

Пирамида (вершина — первый элемент): уровни от тёмно-синего к циану, подписи справа.
levels: [{label, text?}] 3–6 шт.

### `matrix(title, quadrants, x_axis=None, y_axis=None, highlight=None, note=None, notes=None)`

Матрица 2×2 (приоритизация, BCG и т.п.). quadrants: 4×{title, text|items} в порядке
верх-лево, верх-право, низ-лево, низ-право. x_axis/y_axis: подписи осей («Низкая → Высокая»).

### `swot(title, strengths, weaknesses, opportunities, threats, labels=None, note=None, notes=None)`

SWOT-анализ: 4 карточки с буквенными плашками S/W/O/T.

### `comparison(title, left, right, arrow=True, note=None, notes=None)`

Сравнение двух вариантов / «Было → Стало». left/right: {title, items: [str], icon?}.
Левая колонка серая, правая — фирменная синяя/циан.

### `pricing(title, plans, note=None, notes=None)`

Тарифы / пакеты (Pricing Table): колонки с ценой и списком опций. highlight — рекомендованный.
plans: [{name, price, period?, items: [str], highlight?}] 2–4 шт.

### `hub(title, center, items, note=None, notes=None)`

Центр + спутники (слайды 56–58): круг-центр с градиентом, кольца, 4–8 блоков вокруг.
items: [{title, text?, icon?}]

### `cycle(title, steps, center=None, note=None, notes=None)`

Цикл (PDCA, жизненный цикл клиента): 3–6 узлов на кольце со стрелками, описания снаружи.
steps: [{title, text?}]

### `venn(title, sets, center=None, text=None, note=None, notes=None)`

Диаграмма Венна для 2–3 множеств. sets: [{title, text?}], center — подпись пересечения.

### `org_chart(title, root, children, note=None, notes=None)`

Оргструктура в 3 уровня. root: {name, role}; children: [{name, role, children?: [{name, role}]}] (2–5).

### `fishbone(title, effect, causes, note=None, notes=None)`

Диаграмма Исикавы: проблема справа, 4–6 групп причин сверху/снизу.
causes: [{title, items: [str]}]

### `iceberg(title, above, below, labels=('Видимая часть', 'Скрытая часть'), note=None, notes=None)`

Айсберг: видимые факторы над «водой», скрытые — под ней. above/below: [str].

### `team(title, people, note=None, notes=None)`

Команда / спикеры (слайд 39): карточки с круглым фото (или инициалами), имя, роль, описание.
people: [{name, role, text?, photo?}] 2–5 шт.

### `table(title, header, rows, highlight_col=None, highlight_row=None, col_widths=None, group_rows=None, note=None, notes=None, font_size=None)`

Таблица (слайды 40–46): синяя шапка, тонкие разделители, выделенный столбец/строка циановой заливкой.
group_rows: индексы строк-подзаголовков (заливка циан, белый текст).

### `chart(title, kind, categories, series, side=None, header=None, number_format='0', legend=None, labels=True, note=None, notes=None, colors=None, conclusion=None)`

Нативная диаграмма PowerPoint в фирменных цветах (слайды 48–55).
kind: column | bar | line | pie | doughnut | stacked | stacked_bar | area.
series: {name: [values]}; side: [{value, label, icon?}] (2–3 KPI справа) или list[str] (буллеты);
header — плашка над графиком; conclusion — строка вывода снизу.

### `waterfall(title, categories, values, totals=(), header=None, number_format='0', side=None, conclusion=None, note=None, notes=None)`

Каскадная диаграмма (Waterfall / факторный анализ). values — приращения; индексы из totals —
итоговые столбцы (значение = уровень, а не приращение). Рост — циан, снижение — серый, итоги — синий.

### `dashboard(title, kpis, chart, bullets=None, note=None, notes=None)`

Дашборд: ряд KPI сверху + график + блок выводов. kpis: [{value, label, delta?}] 3–4;
chart: {kind, categories, series, header?}; bullets: list[str].


## Примитивы

### `new_slide(title=None, layout='VTB Content', note=None, notes=None)`

Blank content slide with VTB frame (logo, page number). Returns slide.

### `text(s, x, y, w, h, text, size=16, color='1C1C1C', bold=False, font=None, align='l', anchor='t', min_size=None, bullets=False, bullet_color='002882', spacing=1.0, para_space=4, hl_color='002882', name=None, fit=True)`

Text box. `text` may be str (\n = new paragraph) or list (one paragraph each).
**двойные звёздочки** -> полужирный в цвете hl_color (фирменное выделение ВТБ).

### `plate(s, x, y, w, h, text='', color='navy', slant='right', size=18, text_color='FFFFFF', pad=0.22, bold=False, align='l', name=None)`

Фирменная плашка: прямоугольник с косым срезом 20° и градиентом.
color: 'navy' | 'cyan' | 'gray' | 'light' | hex; slant: right | left | both | none.

### `card(s, x, y, w, h, fill='FFFFFF', name=None)`

Белая карточка с тонким контуром и мягкой тенью (как в шаблоне).

### `shape(s, kind, x, y, w, h, fill=None, line=None, line_w=0.75, grad=None, shadow=False, name=None, alpha=None, adj=None)`



### `poly(s, pts, fill=None, line=None, grad=None, shadow=False, name=None, line_w=0.75, alpha=None)`

Closed polygon from (x, y) inch points.

### `circle(s, x, y, d, fill=None, line=None, line_w=0.75, grad=None, shadow=False, alpha=None)`



### `number_badge(s, n, x, y, d=0.5, color='navy', size=None)`



### `icon(s, name, x, y, size=0.5, color='navy')`

Иконка из фирменного набора ВТБ (assets/icons). color: navy|cyan|white|gray.

### `picture(s, path, x, y, w, h, circle=False)`

Picture cropped to fill w x h (center crop); circle=True -> round mask.

### `image_panel(s, path, x0=6.4)`

Фото справа с фирменной диагональю и тонкой циановой линией (как на обложках).

### `logo(s, white=False)`

Повторить логотип ВТБ поверх содержимого слайда (если фото/фон перекрыли логотип макета).

### `hline(s, x, y, w, color='D9D9D9', width=0.75)`



### `vline(s, x, y, h, color='D9D9D9', width=0.75)`



### `line(s, x1, y1, x2, y2, color='D9D9D9', width=0.75, dash=False, arrow=False)`



### `section_header(s, x, y, w, text, color='navy', h=0.48, size=16)`

Плашка-заголовок блока («Заголовок блока»).

### `save(path)`



## Константы цвета

| Имя | HEX | Назначение |
|---|---|---|
| `NAVY` | #002882 | главный синий |
| `NAVY_D` | #001E62 | тёмный конец градиента |
| `CYAN` | #00AAFF | акцент |
| `CYAN_L` | #5FC8FF | светлый циан, выделение на синем |
| `CYAN_BG` | #E5F6FF | фон выделенных строк/столбцов |
| `BLUE2` | #33539B | доп. синий |
| `BLUE3` | #667EB4 | доп. сине-серый |
| `TEXT` | #1C1C1C | основной текст |
| `GRAY` | #878787 | подписи |
| `GRAY_D` | #58595B | вторичный текст |
| `LINE` | #D9D9D9 | разделители |
| `CARD_LINE` | #BFBFBF | контур карточек |
| `GREEN` | #78B497 | позитив в графиках |
| `RED` | #E62632 | негатив в графиках |
