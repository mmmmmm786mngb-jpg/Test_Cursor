#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test: get session date using ВТБ_DevOps.ВычислитьКод and print it.
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

        # 1) Raw current session date via function
        try:
            raw = devops.ВычислитьКод("ТекущаяДатаСеанса()")
            safe_print("EVAL SESSION DATE (raw):")
            safe_print(str(raw))
        except Exception as e:
            safe_print(f"ERROR raw: {type(e).__name__}")

        # 2) Start of day from current date
        try:
            sod = devops.ВычислитьКод("НачалоДня(ТекущаяДата())")
            safe_print("EVAL START OF DAY (raw):")
            safe_print(str(sod))
        except Exception as e:
            safe_print(f"ERROR sod: {type(e).__name__}")

        # 3) Formatted string (for console stability)
        try:
            fmt = devops.ВычислитьКод("Формат(ТекущаяДатаСеанса(), \"ДФ=yyyy-MM-dd HH:mm:ss\")")
            safe_print("EVAL SESSION DATE (formatted):")
            safe_print(str(fmt))
        except Exception as e:
            safe_print(f"ERROR fmt: {type(e).__name__}")

        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())


