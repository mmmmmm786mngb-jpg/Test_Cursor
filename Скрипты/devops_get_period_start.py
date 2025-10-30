#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Get period start via DevOps: Возврат НачалоДня(ТекущаяДата()); and print it.
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
        try:
            res = devops.ВыполнитьКод("Возврат НачалоДня(ТекущаяДата());")
        except Exception as e:  # noqa: BLE001
            safe_print(f"ERROR: {type(e).__name__}")
            return 1
        safe_print("PERIOD START (raw):")
        safe_print(str(res))
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())



