#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal test: get current session date via ВТБ_DevOps and print it.
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


def main() -> int:
    pythoncom.CoInitialize()
    try:
        com = win32com.client.Dispatch("V83.COMConnector")
        conn = com.Connect("Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;")

        devops = getattr(conn, "ВТБ_DevOps", None)
        if devops is None:
            safe_print("ERROR: ВТБ_DevOps not found")
            return 2

        # Получаем дату без "Возврат"
        raw_date = None
        try:
            raw_date = devops.ВыполнитьКод("ТекущаяДатаСеанса();")
        except Exception:
            safe_print("ERROR: cannot obtain session date")
            return 1

        safe_print("SESSION DATE (raw):")
        safe_print(str(raw_date))
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())


