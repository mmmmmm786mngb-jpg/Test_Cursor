#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test: execute exact code block (run_spisok_cb_bu_query.py L157-L168) to get Период via ВТБ_DevOps.
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

        # ---- BEGIN exact snippet (157-168) ----
        # Период: берем через DevOps (текущее время сеанса)
        Период = None
        try:
            devops = getattr(conn, "ВТБ_DevOps")
            Период = devops.ВыполнитьКод("ТекущаяДатаСеанса();")
            if Период is None:
                try:
                    Период = devops.ВыполнитьКод("Возврат ТекущаяДатаСеанса();")
                except Exception:
                    pass
        except Exception:
            Период = None
        # ---- END snippet ----

        if Период is None:
            safe_print("RESULT: None (failed to obtain Период)")
            return 1
        safe_print("RESULT: OK (Период obtained)")
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())



