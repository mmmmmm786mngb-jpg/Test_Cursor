#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Run the optimized query but with inline date literals to avoid date parameters.
Connection: Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;
"""

import sys
import pythoncom
import win32com.client


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("ascii", "replace").decode("ascii"))


def connect():
    pythoncom.CoInitialize()
    try:
        com = win32com.client.Dispatch("V83.COMConnector")
        return com.Connect("Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;")
    finally:
        pythoncom.CoUninitialize()


def find_portfolio(connection, name: str):
    catalogs = getattr(connection, "Справочники")
    for cat_name in ("Портфели", "Портфель"):
        manager = getattr(catalogs, cat_name, None)
        if manager is None:
            continue
        finder = getattr(manager, "НайтиПоНаименованию", None)
        if finder is None:
            continue
        ref = finder(name)
        if not ref.Пустая():
            return ref
    # Create if missing to satisfy parameter typing
    for cat_name in ("Портфели", "Портфель"):
        manager = getattr(catalogs, cat_name, None)
        if manager is None:
            continue
        obj = manager.СоздатьЭлемент()
        obj.Наименование = name
        obj.Записать()
        return obj.Ссылка
    raise RuntimeError("Portfolio catalog not found")


def main() -> int:
    safe_print("Connecting...")
    conn = connect()

    safe_print("Resolving portfolio...")
    портфель = find_portfolio(conn, "TEST_PORTFOLIO")

    ПС = conn.ПланыСчетов.Хозрасчетный
    СчетаЦБ = conn.NewObject("Массив")
    for acc in (
        ПС.АкцииТело,
        ПС.ОблигацииТело,
        ПС.ПаиТело,
        ПС.ИпотечныеСертификатыТело,
        ПС.ДепозитарныеРаспискиТело,
        ПС.СтоимостьСтруктурнойНоты,
    ):
        СчетаЦБ.Добавить(acc)

    СчетаДополнительные = conn.NewObject("Массив")
    СчетаДополнительные.Добавить(ПС.ОбеспеченияОбязательствПолученные)
    СчетаДополнительные.Добавить(ПС.ПолученныеЗаймыПоОперациямРЕПОТело)

    ВидыСубконто = conn.NewObject("Массив")
    try:
        ВидыСубконто.Добавить(
            conn.ПланыВидовХарактеристик.ВидыСубконтоХозрасчетные.ЦенныеБумаги
        )
    except Exception:
        pass

    ВидыАктивов = conn.NewObject("Массив")

    # Inline date literal to bypass date parameters
    D1 = "ДАТАВРЕМЯ(2000,1,1,0,0,0)"
    D2 = "ДАТАВРЕМЯ(2000,12,31,23,59,59)"

    query_text = (
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОбороты.Субконто1 КАК Субконто1\n"
        "ПОМЕСТИТЬ ВТ_ОбъектыОсновныхСчетов\n"
        "ИЗ\n"
        f"\tРегистрБухгалтерии.Хозрасчетный.Обороты({D1}, {D2}, Период, Счет В (&СчетаЦБ), , Портфель = &Портфель, , ) КАК ХозрасчетныйОбороты\n"
        "ГДЕ\n"
        "\tХозрасчетныйОбороты.КоличествоОборот <> 0\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОстатки.Субконто1\n"
        "ИЗ\n"
        f"\tРегистрБухгалтерии.Хозрасчетный.Остатки({D2}, Счет В (&СчетаЦБ), , Портфель = &Портфель) КАК ХозрасчетныйОстатки\n"
        "ГДЕ\n"
        "\tХозрасчетныйОстатки.КоличествоОстаток <> 0\n"
        "\n"
        "ИНДЕКСИРОВАТЬ ПО\n"
        "\tСубконто1;\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tАктивы.Ссылка КАК Актив\n"
        "ИЗ\n"
        "\tВТ_ОбъектыОсновныхСчетов КАК ВТ_ОбъектыОсновныхСчетов\n"
        "\t\tВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.Активы КАК Активы\n"
        "\t\tПО ВТ_ОбъектыОсновныхСчетов.Субконто1 = Активы.Объект\n"
        "\t\t\tИ (Активы.ВидАктива В ИЕРАРХИИ (&ВидыАктивов))\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОбороты.Субконто1\n"
        "ИЗ\n"
        f"\tРегистрБухгалтерии.Хозрасчетный.Обороты({D1}, {D2}, Период, Счет В (&СчетаДополнительные), &ВидыСубконто, Портфель = &Портфель, , ) КАК ХозрасчетныйОбороты\n"
        "ГДЕ\n"
        "\tХозрасчетныйОбороты.КоличествоОборот <> 0\n"
        "\tИ ВЫРАЗИТЬ(ХозрасчетныйОбороты.Субконто1 КАК Справочник.Активы) В ИЕРАРХИИ (&ВидыАктивов)\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОстатки.Субконто1\n"
        "ИЗ\n"
        f"\tРегистрБухгалтерии.Хозрасчетный.Остатки({D2}, Счет В (&СчетаДополнительные), &ВидыСубконто, Портфель = &Портфель) КАК ХозрасчетныйОстатки\n"
        "ГДЕ\n"
        "\tХозрасчетныйОстатки.КоличествоОстаток <> 0\n"
        "\tИ ВЫРАЗИТЬ(ХозрасчетныйОстатки.Субконто1 КАК Справочник.Активы) В ИЕРАРХИИ (&ВидыАктивов)\n"
    )

    q = conn.NewObject("Запрос", query_text)
    q.УстановитьПараметр("Портфель", портфель)
    q.УстановитьПараметр("СчетаЦБ", СчетаЦБ)
    q.УстановитьПараметр("СчетаДополнительные", СчетаДополнительные)
    q.УстановитьПараметр("ВидыСубконто", ВидыСубконто)
    q.УстановитьПараметр("ВидыАктивов", ВидыАктивов)

    safe_print("Executing query...")
    t = q.Выполнить().Выгрузить()
    safe_print(f"OK: rows = {t.Количество()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())


