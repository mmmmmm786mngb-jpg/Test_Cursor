#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Execute the query from СписокЦБ_БУ_Оптимизированный.bsl inside 1C via COM.

Params:
  --portfolio "ИмяПортфеля" (required)
  --date "YYYY-MM-DD" (optional, default: today)
  --asset-types "Имя1,Имя2" (optional; names from Справочник.Активы.ВидАктива)

Connection: Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;

Console output uses ASCII-safe text.
"""

import argparse
import datetime as dt
import sys

import pythoncom
import win32com.client


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("ascii", "replace").decode("ascii"))


def connect_1c(conn_string: str):
    pythoncom.CoInitialize()
    try:
        com = win32com.client.Dispatch("V83.COMConnector")
        return com.Connect(conn_string)
    finally:
        pythoncom.CoUninitialize()


def find_portfolio(connection, name: str):
    # Try common catalog names: "Портфели", "Портфель"
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
    raise RuntimeError("Portfolio not found by name")


def build_asset_types(connection, names_csv: str):
    # Returns 1C Array of Активы.ВидАктива references based on names
    if not names_csv:
        return None
    catalogs = getattr(connection, "Справочники")
    assets_mgr = getattr(catalogs, "Активы", None)
    if assets_mgr is None:
        raise RuntimeError("Catalog 'Активы' not found")
    arr = connection.NewObject("Массив")
    for name in [n.strip() for n in names_csv.split(",") if n.strip()]:
        # Find first asset with matching ВидАктива by name
        q = connection.NewObject("Запрос")
        q.Текст = (
            "ВЫБРАТЬ ПЕРВЫЕ 1\n"
            "    Активы.ВидАктива КАК Вид\n"
            "ИЗ\n"
            "    Справочник.Активы КАК Активы\n"
            "ГДЕ\n"
            "    Активы.ВидАктива.Наименование = &Имя"
        )
        q.УстановитьПараметр("Имя", name)
        table = q.Выполнить().Выгрузить()
        if table.Количество() > 0:
            arr.Добавить(table[0].Вид)
    return arr if arr.Количество() > 0 else None


def main() -> int:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--portfolio", required=True)
    parser.add_argument("--date", default=dt.date.today().isoformat())
    parser.add_argument("--asset-types", default="")
    args = parser.parse_args()

    conn_string = "Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;"

    safe_print("Connecting to 1C...")
    try:
        connection = connect_1c(conn_string)
    except Exception as e:  # noqa: BLE001
        safe_print("ERROR: connect failed")
        safe_print(type(e).__name__)
        return 2

    safe_print("Resolving portfolio...")
    try:
        портфель = find_portfolio(connection, args.portfolio)
    except Exception as e:  # noqa: BLE001
        safe_print("ERROR: portfolio not found")
        safe_print(type(e).__name__)
        return 3

    # Date will be resolved in 1C via DevOps helper
    date_obj = None

    safe_print("Building parameters...")
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
    try:
        ВидыСубконто.Добавить(
            connection.ПланыВидовХарактеристик.ВидыСубконтоХозрасчетные.ЦенныеБумаги
        )
    except Exception:
        pass

    ВидыАктивов = build_asset_types(connection, args["asset_types"] if isinstance(args, dict) else args.asset_types)
    if ВидыАктивов is None:
        # If not provided, fall back to taking all types present in Активы
        ВидыАктивов = connection.NewObject("Массив")
    try:
        q_types = connection.NewObject("Запрос",
                "ВЫБРАТЬ РАЗЛИЧНЫЕ\n"
                "    Активы.ВидАктива КАК Вид\n"
                "ИЗ\n"
                "    Справочник.Активы КАК Активы")
        t = q_types.Выполнить().Выгрузить()
        for row in t:
            if getattr(row, "Вид", None) is not None:
                ВидыАктивов.Добавить(row.Вид)
    except Exception:
        pass

    # Период: берем через DevOps функцией, выражение БЕЗ ";"
    Период = None
    try:
        devops = getattr(connection, "ВТБ_DevOps")
        Период = devops.ВычислитьКод("ТекущаяДатаСеанса()")
    except Exception:
        Период = None
    if Период is None:
        # Резерв: собрать дату из ОС внутри 1С
        from datetime import date as _d
        _today = _d.today()
        try:
            qd = connection.NewObject("Запрос",
                "ВЫБРАТЬ\n"
                "    Дата(&y, &m, &d) КАК D")
            qd.УстановитьПараметр("y", _today.year)
            qd.УстановитьПараметр("m", _today.month)
            qd.УстановитьПараметр("d", _today.day)
            Период = qd.Выполнить().Выгрузить()[0].D
        except Exception:
            Период = None

    # Текст запроса: используем один параметр &Период
    query_text = (
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОбороты.Субконто1 КАК Субконто1\n"
        "ПОМЕСТИТЬ ВТ_ОбъектыОсновныхСчетов\n"
        "ИЗ\n"
        "\tРегистрБухгалтерии.Хозрасчетный.Обороты(&Период, &Период, Период, Счет В (&СчетаЦБ), , Портфель = &Портфель, , ) КАК ХозрасчетныйОбороты\n"
        "ГДЕ\n"
        "\tХозрасчетныйОбороты.КоличествоОборот <> 0\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОстатки.Субконто1\n"
        "ИЗ\n"
        "\tРегистрБухгалтерии.Хозрасчетный.Остатки(&Период, Счет В (&СчетаЦБ), , Портфель = &Портфель) КАК ХозрасчетныйОстатки\n"
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
        "\tРегистрБухгалтерии.Хозрасчетный.Обороты(&Период, &Период, Период, Счет В (&СчетаДополнительные), &ВидыСубконто, Портфель = &Портфель, , ) КАК ХозрасчетныйОбороты\n"
        "ГДЕ\n"
        "\tХозрасчетныйОбороты.КоличествоОборот <> 0\n"
        "\tИ ВЫРАЗИТЬ(ХозрасчетныйОбороты.Субконто1 КАК Справочник.Активы) В ИЕРАРХИИ (&ВидыАктивов)\n"
        "\n"
        "ОБЪЕДИНИТЬ\n"
        "\n"
        "ВЫБРАТЬ\n"
        "\tХозрасчетныйОстатки.Субконто1\n"
        "ИЗ\n"
        "\tРегистрБухгалтерии.Хозрасчетный.Остатки(&Период, Счет В (&СчетаДополнительные), &ВидыСубконто, Портфель = &Портфель) КАК ХозрасчетныйОстатки\n"
        "ГДЕ\n"
        "\tХозрасчетныйОстатки.КоличествоОстаток <> 0\n"
        "\tИ ВЫРАЗИТЬ(ХозрасчетныйОстатки.Субконто1 КАК Справочник.Активы) В ИЕРАРХИИ (&ВидыАктивов)\n"
    )

    q = connection.NewObject("Запрос", query_text)
    q.УстановитьПараметр("Портфель", портфель)
    if Период is None:
        raise RuntimeError("Не удалось получить параметр Период")
    q.УстановитьПараметр("Период", Период)
    q.УстановитьПараметр("СчетаЦБ", СчетаЦБ)
    q.УстановитьПараметр("СчетаДополнительные", СчетаДополнительные)
    q.УстановитьПараметр("ВидыСубконто", ВидыСубконто)
    q.УстановитьПараметр("ВидыАктивов", ВидыАктивов)

    safe_print("Executing query...")
    table = q.Выполнить().Выгрузить()
    safe_print(f"OK: rows = {table.Количество()}")

    # Print first 10 result refs as strings
    limit = min(10, table.Количество())
    for i in range(limit):
        row = table[i]
        try:
            safe_print(str(row.Актив))
        except Exception:  # noqa: BLE001
            safe_print("<ref>")

    return 0


if __name__ == "__main__":
    sys.exit(main())


