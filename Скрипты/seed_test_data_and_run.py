#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Seed minimal test data in 1C (empty DB) and run the optimized query.

Actions:
  1) Ensure a portfolio exists: TEST_PORTFOLIO
  2) Ensure at least one asset exists: TEST_ASSET_1 (with any available ВидАктива)
  3) Execute the query from СписокЦБ_БУ_Оптимизированный.bsl with today's date

Note: Creating accounting register movements is configuration-specific.
This script does not create postings; result set may be empty if no movements/balances exist.
"""

import sys
import datetime as dt

import pythoncom
import win32com.client


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("ascii", "replace").decode("ascii"))


def connect() -> object:
    pythoncom.CoInitialize()
    try:
        com = win32com.client.Dispatch("V83.COMConnector")
        return com.Connect("Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;")
    finally:
        pythoncom.CoUninitialize()


def ensure_portfolio(connection, name: str):
    catalogs = getattr(connection, "Справочники")
    for cat_name in ("Портфели", "Портфель"):
        manager = getattr(catalogs, cat_name, None)
        if manager is None:
            continue
        find = getattr(manager, "НайтиПоНаименованию", None)
        if find is None:
            continue
        ref = find(name)
        if not ref.Пустая():
            return ref
        obj = manager.СоздатьЭлемент()
        obj.Наименование = name
        obj.Записать()
        return obj.Ссылка
    raise RuntimeError("No suitable portfolio catalog found")


def pick_any_asset_type(connection):
    q = connection.NewObject("Запрос",
        "ВЫБРАТЬ ПЕРВЫЕ 1\n"
        "    Активы.ВидАктива КАК Вид\n"
        "ИЗ\n"
        "    Справочник.Активы КАК Активы\n"
        "ГДЕ НЕ ПУСТО(Активы.ВидАктива)")
    t = q.Выполнить().Выгрузить()
    if t.Количество() > 0:
        return t[0].Вид
    # Try enumeration or catalog fallback via metadata? Fallback to None
    return None


def ensure_asset(connection, name: str):
    catalogs = getattr(connection, "Справочники")
    assets = getattr(catalogs, "Активы", None)
    if assets is None:
        raise RuntimeError("Catalog 'Активы' not found")
    find = getattr(assets, "НайтиПоНаименованию", None)
    if find is None:
        raise RuntimeError("Method НайтиПоНаименованию not found on 'Активы'")
    ref = find(name)
    if not ref.Пустая():
        return ref
    manager = connection.NewObject("СправочникМенеджер.Активы")
    obj = manager.СоздатьЭлемент()
    obj.Наименование = name
    # Try to set ВидАктива if such requisite exists
    try:
        asset_type = pick_any_asset_type(connection)
        if asset_type is not None:
            obj.ВидАктива = asset_type
    except Exception:
        pass
    obj.Записать()
    return obj.Ссылка


def run_query(connection, portfolio_ref, date_obj):
    ПС = connection.ПланыСчетов.Хозрасчетный

    СчетаЦБ = connection.NewObject("Массив")
    for acc in (
        ПС.АкцииТело,
        ПС.ОблигацииТело,
        ПС.ПаиТело,
        ПС.ИпотечныеСертификатыТело,
        ПС.ДепозитарныеРаспискиТело,
        ПС.СтоимостьСтруктурнойНоты,
    ):
        СчетаЦБ.Добавить(acc)

    Счет66_04_1 = ПС.ПолученныеЗаймыПоОперациямРЕПОТело
    СчетОбеспечения = ПС.ОбеспеченияОбязательствПолученные

    СчетаДополнительные = connection.NewObject("Массив")
    СчетаДополнительные.Добавить(СчетОбеспечения)
    СчетаДополнительные.Добавить(Счет66_04_1)

    ВидыСубконто = connection.NewObject("Массив")
    ВидыСубконто.Добавить(connection.ПланыВидовХарактеристик.ВидыСубконто.ЦенныеБумаги)

    # Build all asset types from catalog as default
    ВидыАктивов = connection.NewObject("Массив")
    q_types = connection.NewObject("Запрос",
        "ВЫБРАТЬ РАЗЛИЧНЫЕ\n"
        "    Активы.ВидАктива КАК Вид\n"
        "ИЗ\n"
        "    Справочник.Активы КАК Активы")
    t = q_types.Выполнить().Выгрузить()
    for row in t:
        if not (row.Вид is None):
            ВидыАктивов.Добавить(row.Вид)

    # Build dates
    Дата = win32com.client.Dispatch("V83.Date")(date_obj.year, date_obj.month, date_obj.day)

    query_text = (
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОбороты.Субконто1 КАК Субконто1\n"
        "ПОМЕСТИТЬ ВТ_ОбъектыОсновныхСчетов\n"
        "ИЗ\n"
        "\tРегистрБухгалтерии.Хозрасчетный.Обороты(&ДатаН, &ДатаК, Период, Счет В (&СчетаЦБ), , Портфель = &Портфель, , ) КАК ХозрасчетныйОбороты\n"
        "ГДЕ\n"
        "\tХозрасчетныйОбороты.КоличествоОборот <> 0\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОстатки.Субконто1\n"
        "ИЗ\n"
        "\tРегистрБухгалтерии.Хозрасчетный.Остатки(&ДатаК, Счет В (&СчетаЦБ), , Портфель = &Портфель) КАК ХозрасчетныйОстатки\n"
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
        "\tРегистрБухгалтерии.Хозрасчетный.Обороты(&ДатаН, &ДатаК, Период, Счет В (&СчетаДополнительные), &ВидыСубконто, Портфель = &Портфель, , ) КАК ХозрасчетныйОбороты\n"
        "ГДЕ\n"
        "\tХозрасчетныйОбороты.КоличествоОборот <> 0\n"
        "\tИ ВЫРАЗИТЬ(ХозрасчетныйОбороты.Субконто1 КАК Справочник.Активы) В ИЕРАРХИИ (&ВидыАктивов)\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОстатки.Субконто1\n"
        "ИЗ\n"
        "\tРегистрБухгалтерии.Хозрасчетный.Остатки(&ДатаК, Счет В (&СчетаДополнительные), &ВидыСубконто, Портфель = &Портфель) КАК ХозрасчетныйОстатки\n"
        "ГДЕ\n"
        "\tХозрасчетныйОстатки.КоличествоОстаток <> 0\n"
        "\tИ ВЫРАЗИТЬ(ХозрасчетныйОстатки.Субконто1 КАК Справочник.Активы) В ИЕРАРХИИ (&ВидыАктивов)\n"
    )

    q = connection.NewObject("Запрос", query_text)
    q.УстановитьПараметр("Портфель", portfolio_ref)
    q.УстановитьПараметр("ДатаН", connection.НачалоДня(Дата))
    q.УстановитьПараметр("ДатаК", connection.КонецДня(Дата))
    q.УстановитьПараметр("СчетаЦБ", СчетаЦБ)
    q.УстановитьПараметр("СчетаДополнительные", СчетаДополнительные)
    q.УстановитьПараметр("ВидыСубконто", ВидыСубконто)
    q.УстановитьПараметр("ВидыАктивов", ВидыАктивов)

    table = q.Выполнить().Выгрузить()
    safe_print(f"OK: rows = {table.Количество()}")
    return table


def main() -> int:
    safe_print("Connecting...")
    try:
        connection = connect()
    except Exception as e:  # noqa: BLE001
        safe_print("ERROR: connect failed")
        safe_print(type(e).__name__)
        return 2

    safe_print("Seeding portfolio...")
    portfolio = ensure_portfolio(connection, "TEST_PORTFOLIO")

    safe_print("Seeding asset...")
    ensure_asset(connection, "TEST_ASSET_1")

    safe_print("Running query...")
    today = dt.date.today()
    table = run_query(connection, portfolio, today)

    limit = min(10, table.Количество())
    for i in range(limit):
        row = table[i]
        try:
            safe_print(str(row.Актив))
        except Exception:
            safe_print("<ref>")
    return 0


if __name__ == "__main__":
    sys.exit(main())


