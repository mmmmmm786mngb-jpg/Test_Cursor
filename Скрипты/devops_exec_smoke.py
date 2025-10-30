#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal smoke test: connect via COM and execute small code through ВТБ_DevOps.ВыполнитьКод.
Connection: Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;
Console output is ASCII-only.
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
        safe_print("Connecting...")
        com = win32com.client.Dispatch("V83.COMConnector")
        conn = com.Connect("Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;")

        safe_print("Locating DevOps module...")
        devops = getattr(conn, "ВТБ_DevOps", None)
        if devops is None:
            safe_print("ERROR: module 'ВТБ_DevOps' not found in this base")
            return 2

        # 1) Simple arithmetic expression
        try:
            res1 = devops.ВыполнитьКод("1 + 1;")
            safe_print(f"OK: simple expr -> {res1}")
        except Exception as e:  # noqa: BLE001
            safe_print("ERROR: simple expr failed")
            safe_print(type(e).__name__)

        # 2) Current session date
        try:
            res2 = devops.ВыполнитьКод("ТекущаяДатаСеанса();")
            safe_print("OK: session date obtained")
        except Exception as e:  # noqa: BLE001
            safe_print("ERROR: session date expr failed")
            safe_print(type(e).__name__)

        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())



