#!/usr/bin/env python3
"""Fill the OTUS defense template for the end-to-end planning project.

The source deck is the uploaded template (garbled filename). Slides 1–10 are
the defense skeleton. Instruction slides are dropped. Three library layouts
are kept and filled: problem list, process, before/after, metrics table.
"""

from __future__ import annotations

import os
from copy import deepcopy

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "zashchita-skvoznoe-planirovanie.pptx")


def template_path() -> str:
    matches = [
        os.path.join(ROOT, name)
        for name in os.listdir(ROOT)
        if name.endswith(".pptx")
    ]
    if not matches:
        raise SystemExit("template pptx not found in repo root")
    return matches[0]


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type is not None and shape.shape_type == 6:
            yield from iter_shapes(shape.shapes)


def find_shape(slide, name: str):
    for shape in iter_shapes(slide.shapes):
        if shape.name == name:
            return shape
    raise KeyError(name)


def set_paragraph_text(paragraph, text: str) -> None:
    p = paragraph._p
    runs = p.findall(qn("a:r"))
    if not runs:
        end = p.find(qn("a:endParaRPr"))
        run = etree.Element(qn("a:r"))
        rpr = etree.Element(qn("a:rPr"))
        rpr.set("lang", "ru-RU")
        if end is not None and end.get("sz"):
            rpr.set("sz", end.get("sz"))
        run.append(rpr)
        node = etree.SubElement(run, qn("a:t"))
        node.text = text
        if end is not None:
            end.addprevious(run)
        else:
            p.append(run)
        return
    first = runs[0]
    node = first.find(qn("a:t"))
    if node is None:
        node = etree.SubElement(first, qn("a:t"))
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in runs[1:]:
        t = extra.find(qn("a:t"))
        if t is not None:
            t.text = ""


def set_lines(shape, lines: list[str]) -> None:
    tf = shape.text_frame
    while len(tf.paragraphs) < len(lines):
        src = tf.paragraphs[-1]._p
        src.addnext(deepcopy(src))
    for index, paragraph in enumerate(tf.paragraphs):
        set_paragraph_text(paragraph, lines[index] if index < len(lines) else "")


def set_font_size(shape, hundredths: int) -> None:
    for rpr in shape._element.findall(".//" + qn("a:rPr")):
        rpr.set("sz", str(hundredths))


def first_table(slide):
    for shape in slide.shapes:
        if shape.has_table:
            return shape.table
    raise KeyError("table")


def fill_table(slide, rows: list[list[str]]) -> None:
    table = first_table(slide)
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            cell = table.cell(r, c)
            paragraphs = cell.text_frame.paragraphs
            set_paragraph_text(paragraphs[0], value)
            for extra in paragraphs[1:]:
                set_paragraph_text(extra, "")


def paint_like(src_cell, dst_cell, text: str) -> None:
    src_run = src_cell.text_frame.paragraphs[0]._p.find(qn("a:r"))
    dst_p = dst_cell.text_frame.paragraphs[0]._p
    for run in dst_p.findall(qn("a:r")):
        dst_p.remove(run)
    new_run = deepcopy(src_run)
    new_run.find(qn("a:t")).text = text
    end = dst_p.find(qn("a:endParaRPr"))
    if end is not None:
        end.addprevious(new_run)
    else:
        dst_p.append(new_run)


def silence_paragraph(paragraph) -> None:
    set_paragraph_text(paragraph, "")
    ppr = paragraph._p.find(qn("a:pPr"))
    if ppr is None:
        return
    for tag in ("buAutoNum", "buChar", "buFont", "buClr", "buSzPts"):
        for node in ppr.findall(qn(f"a:{tag}")):
            ppr.remove(node)
    ppr.append(etree.Element(qn("a:buNone")))
    ppr.set("lvl", "0")
    ppr.set("marL", "0")
    ppr.set("indent", "0")


def set_notes(slide, text: str) -> None:
    slide.notes_slide.notes_text_frame.text = text


def reorder_and_drop(prs: Presentation, keep: list[int]) -> None:
    lst = prs.slides._sldIdLst
    current = list(lst)
    for element in current:
        lst.remove(element)
    keep_set = set(keep)
    for index, element in enumerate(current):
        if index in keep_set:
            continue
        rel_id = element.get(qn("r:id"))
        prs.part.drop_rel(rel_id)
    for index in keep:
        lst.append(current[index])


def build() -> None:
    src = template_path()
    prs = Presentation(src)

    # --- 1. Title ---
    title = find_shape(prs.slides[0], "Google Shape;147;p35")
    set_lines(title, ["Повышение предсказуемости", "от квартала до спринта"])
    set_font_size(title, 3200)
    title.height = Inches(1.45)
    course = find_shape(prs.slides[0], "Google Shape;150;p35")
    set_lines(course, ["Team Lead"])
    set_notes(
        prs.slides[0],
        "Полная тема: «Повышение предсказуемости разработки и бизнес-ценности "
        "через сквозное планирование: от квартальных целей до спринтов». "
        "Перед защитой замените фамилию на слайде 3 и стоковое фото на своё.",
    )

    # --- 3. Intro ---
    cover = find_shape(prs.slides[2], "Google Shape;163;p37")
    set_lines(cover, ["Защита проекта", "Тема: от квартала до спринта"])
    bio = find_shape(prs.slides[2], "Google Shape;166;p37")
    bio.height = Inches(1.55)
    set_lines(
        bio,
        [
            "Team Lead, команда из 8 человек.",
            "Веду план от квартала до спринта.",
            "Scrum, спринт — две недели.",
        ],
    )
    set_notes(
        prs.slides[2],
        "Устно: веду продуктовую команду из восьми человек и отвечаю за "
        "обязательства перед бизнесом. Проект про то, как связать цели квартала "
        "с бэклогом так, чтобы обещание совпадало с фактом.",
    )

    # --- 4. Plan: keep the five beats, name the problem in the first ---
    plan_first = find_shape(prs.slides[3], "Google Shape;172;p38")
    set_lines(plan_first, ["Проблема, цель и задачи"])
    set_notes(
        prs.slides[3],
        "Пять блоков шаблона сохранены. Первый чуть расширен: сначала проблема, "
        "затем цель. Дальше идут подходы, артефакты, выводы и вопросы.",
    )

    # --- 25. Problem list (library) ---
    problem = prs.slides[24]
    set_lines(find_shape(problem, "Google Shape;410;p59"), ["Где ломается план"])
    fill_table(
        problem,
        [
            ["1.", "Скоуп квартала без оценки"],
            ["2.", "Бизнес ждёт 100% объёма"],
            ["3.", "Требования дописывают в работе"],
            ["4.", "В спринт берут сырой бэклог"],
            ["5.", "В бэклоге задачи без ценности"],
        ],
    )
    set_notes(
        problem,
        "Это as-is одной продуктовой команды. Скоуп квартала собирали списком "
        "желаний, без ёмкости и без связи с целью. Отсюда срывы, возвраты "
        "требований и задачи, которые команда додумала сама.",
    )

    # --- 5. Goal and tasks ---
    goal_title = find_shape(prs.slides[4], "Google Shape;185;p39")
    set_lines(goal_title, ["Цель и задачи"])
    goal = find_shape(prs.slides[4], "Google Shape;187;p39")
    set_lines(
        goal,
        [
            "Поднять предсказуемость поставки",
            "и долю задач, закрывающих KR.",
        ],
    )
    fill_table(
        prs.slides[4],
        [
            ["1.", "Зафиксировать 4 измеримых KR"],
            ["2.", "Обещать 70% ёмкости, не список"],
            ["3.", "Ввести DoR из семи пунктов"],
            ["4.", "Отбирать квартал через WSJF"],
            ["5.", "Считать пять метрик каждый спринт"],
        ],
    )
    bubble = find_shape(prs.slides[4], "Google Shape;190;p39")
    set_lines(
        bubble,
        [
            "Пять задач замыкают контур: от цифры квартала до готового спринта. Порог успеха — сдвиг каждой метрики сильнее 10%.",
        ],
    )
    set_notes(
        prs.slides[4],
        "Цель: внедрить сквозное планирование от квартальных OKR до спринт-бэклога. "
        "Commitment = средняя velocity × число спринтов × 0,7. Остаток ёмкости — "
        "stretch, его не обещают. Шестая операционная задача, «в спринт только Ready», "
        "входит в DoR: неготовая задача не оценивается в коммит.",
    )

    # --- 6. Tools ---
    fill_table(
        prs.slides[5],
        [
            ["1.", "OKR: цель квартала, не фичи"],
            ["2.", "DoR: в спринт только готовое"],
            ["3.", "WSJF: выше ценность, ниже размер"],
            ["4.", "Jira: дашборд пяти метрик"],
        ],
    )
    tech_tbl = first_table(prs.slides[5])._tbl
    empty_row = tech_tbl.findall(qn("a:tr"))[-1]
    tech_tbl.remove(empty_row)
    why = find_shape(prs.slides[5], "Google Shape;198;p40")
    set_lines(
        why,
        [
            "Четыре правила на четыре сбоя: обещание, требования, приоритет и факт.",
        ],
    )
    set_notes(
        prs.slides[5],
        "DoR, семь пунктов: 1) проблема и результат для пользователя; "
        "2) ссылка на KR квартала; 3) проверяемые критерии приёмки; "
        "4) макет, если есть интерфейс; 5) зависимости сняты или вынесены; "
        "6) есть оценка, задача входит в один спринт; 7) QA понимает, как проверить. "
        "Владелец чеклиста — аналитик, подпись PO. "
        "WSJF = (ценность + срочность + снятие риска) / размер, шкала 1–2–3–5–8–13. "
        "Оценку ставят PO, team lead и аналитик вместе, не один человек.",
    )

    # --- 7. Artifacts ---
    artifacts = find_shape(prs.slides[6], "Google Shape;207;p41")
    set_lines(
        artifacts,
        [
            "Дальше три слайда с артефактами.",
            "Схема: KR → ёмкость → WSJF → DoR → спринт.",
            "DoR: семь пунктов, владелец — аналитик.",
            "Кейс: заявка с 5 дней до 4.",
            "Цель спринта: статус без звонка.",
            "Метрики: 3 спринта до и 3 после.",
            "",
            "",
        ],
    )
    set_notes(
        prs.slides[6],
        "Пример маппинга. KR: сократить срок заявки с 5 дней до 4 к концу квартала. "
        "Эпик: убрать ручные передачи статуса. Sprint Goal: клиент видит статус заявки "
        "без звонка оператору. Все истории эпика ссылаются на этот KR и проходят DoR.",
    )

    # --- 34. Process ---
    process = prs.slides[33]
    set_lines(find_shape(process, "Google Shape;526;p68"), ["От KR до спринта"])
    stages = [
        ("Google Shape;527;p68", "KR"),
        ("Google Shape;528;p68", "Ёмкость"),
        ("Google Shape;529;p68", "WSJF"),
        ("Google Shape;530;p68", "DoR"),
        ("Google Shape;531;p68", "Спринт"),
    ]
    for name, label in stages:
        set_lines(find_shape(process, name), [label])
    fill_table(
        process,
        [
            [
                "4 измеримых KR",
                "Обещаем 70% ёмкости",
                "В план только с KR",
                "Без DoR нет входа",
                "Один KR на спринт",
            ]
        ],
    )
    set_notes(
        process,
        "Порядок жёсткий. Сначала цифра квартала, потом ёмкость, потом отбор, "
        "потом проработка, и только затем планирование спринта. "
        "Sprint Goal всегда закрывает один KR, а не набор несвязанных задач.",
    )

    # --- 28. Before / after ---
    compare = prs.slides[27]
    set_lines(find_shape(compare, "Google Shape;438;p62"), ["Было и стало"])
    set_lines(find_shape(compare, "Google Shape;443;p62"), ["До"])
    set_lines(find_shape(compare, "Google Shape;444;p62"), ["После"])
    set_lines(
        find_shape(compare, "Google Shape;441;p62"),
        [
            "Скоуп без оценки",
            "Ждут исполнения на 100%",
            "Требования дописывают в работе",
            "В спринт берут сырой бэклог",
        ],
    )
    set_lines(
        find_shape(compare, "Google Shape;442;p62"),
        [
            "Обещание — 70% ёмкости",
            "В квартал входят только KR",
            "DoR закрывают до плана",
            "В спринт берём только Ready",
        ],
    )
    set_notes(
        compare,
        "Справа не «мы стали стараться», а четыре запрета. "
        "Нельзя обещать неизмеренный список. Нельзя брать в квартал задачу без KR. "
        "Нельзя планировать спринт из неготового. Нельзя начинать разработку, пока DoR открыт.",
    )

    # --- 42. Metrics table ---
    metrics = prs.slides[41]
    set_lines(find_shape(metrics, "Google Shape;608;p76"), ["Метрики: до и после"])
    fill_table(
        metrics,
        [
            ["", "Метрика", "До → после", "Лучше на"],
            ["1.", "Say-Do", "62% → 84%", "35%"],
            ["2.", "Готовность бэклога", "41% → 79%", "93%"],
            ["3.", "Доля возвратов", "35% → 18%", "49%"],
            ["4.", "Задачи с KR", "48% → 80%", "67%"],
            ["5.", "CV velocity", "0,40 → 0,20", "50%"],
        ],
    )
    set_notes(
        metrics,
        "Окно: 3 спринта до внедрения и 3 после. Округление до целых. "
        "«Лучше на» — относительное изменение. Для возвратов и CV снижение и есть улучшение. "
        "Все пять сдвигов сильнее порога 10%.\n"
        "Say-Do = закрытые SP из исходного коммита / SP коммита. "
        "Добавленное после старта спринта в дробь не входит. "
        "Сырые суммы: 84 из 136 SP → 94 из 112 SP. "
        "Знаменатель упал, потому что перестали обещать неизмеренный список. "
        "Числитель не упал: 84 → 94 SP. Предсказуемость выросла не за счёт меньшей поставки.\n"
        "Готовность = задачи, прошедшие DoR, среди кандидатов на планировании. 22/54 → 46/58.\n"
        "Возвраты = задачи, вернувшиеся в анализ из-за требований, / завершённые. 17/48 → 9/51.\n"
        "Задачи с KR = завершённые со ссылкой на KR квартала / все завершённые. 23/48 → 41/51. "
        "Буфер 30% на инциденты сознательно снижает эту долю, если у инцидента нет KR.\n"
        "CV = выборочное стандартное отклонение velocity / среднее. "
        "До: 18, 40, 26 SP, среднее 28, CV 0,40. "
        "После: 26, 38, 30 SP, среднее 31, CV 0,20.\n"
        "Если пилот шёл на живом трекере, замените таблицу выгрузкой Jira до защиты. "
        "Формулы оставлены так, чтобы новый расчёт встал в те же ячейки.",
    )

    # --- 8. Conclusions + reflection in the bubble ---
    conclusions = prs.slides[7]
    fill_table(
        conclusions,
        [
            ["1.", "Say-Do: 62% → 84%"],
            ["2.", "Готовность: 41% → 79%"],
            ["3.", "Возвраты: 35% → 18%"],
            ["4.", "Ценность: 48% → 80%"],
        ],
    )
    # The template's fourth number cell is empty and has no run style.
    number_table = first_table(conclusions)
    paint_like(number_table.cell(0, 0), number_table.cell(3, 0), "4.")
    reflection = find_shape(conclusions, "Google Shape;216;p42")
    set_lines(
        reflection,
        [
            "Цель взята: все метрики лучше 10%.",
            "Легко легли KR и дашборд.",
            "Трудно было отпустить 100% скоупа.",
            "Срок: 6 спринтов, 12 недель.",
            "Польза 9/10. Дальше — сценарии квартала.",
            "",
            "",
            "",
            "",
        ],
    )
    for paragraph in reflection.text_frame.paragraphs[5:]:
        silence_paragraph(paragraph)
    set_notes(
        prs.slides[7],
        "Пятая метрика, CV velocity 0,40 → 0,20, лежит на предыдущем слайде: "
        "в шаблоне выводов только четыре строки. "
        "Почему не 10 из 10: одно квартальное окно, привычка бизнеса ещё догоняет процесс. "
        "Дальше — прогноз квартала в трёх сценариях (commitment, stretch, риск) "
        "и пересмотр KR раз в месяц. Ограничение: одна команда, шесть спринтов, "
        "без контрольной группы. Для курса этого достаточно, причинный вывод шире не делаю.",
    )

    # --- 9. Questions: keep +/-, replace the instruction bubble ---
    q_bubble = find_shape(prs.slides[8], "Google Shape;229;p43")
    set_lines(
        q_bubble,
        ["Готов разобрать формулы, DoR и правило 70% ёмкости."],
    )
    set_notes(
        prs.slides[8],
        "Частые вопросы. Почему не velocity как главная метрика: боль была в сорванном "
        "обещании, а не в скорости. Не уменьшили ли поставку ради Say-Do: нет, "
        "закрытые SP выросли с 84 до 94, а коммит снизился со 136 до 112. "
        "Куда девается техдолг и инциденты: в буфер 30%, не в обещание. "
        "Кто владелец: PO — KR и приоритет, аналитик — DoR, team lead — ёмкость и Say-Do.",
    )

    reorder_and_drop(
        prs,
        [0, 1, 2, 3, 24, 4, 5, 6, 33, 27, 41, 7, 8, 9],
    )

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    print(OUT, os.path.getsize(OUT))


if __name__ == "__main__":
    build()
