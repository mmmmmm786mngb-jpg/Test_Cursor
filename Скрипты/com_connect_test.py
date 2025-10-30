#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simple COM connection test to 1C:Enterprise.
Connects to Srvr="localhost";Ref="WIM_FIN" and executes a trivial query.

Console output uses ASCII-only text to avoid encoding issues.
"""

import sys
import pythoncom
import win32com.client


def safe_print(text: str) -> None:
    """Print using ASCII-safe fallback."""
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("ascii", "replace").decode("ascii"))


def try_connect(connection_string: str):
    pythoncom.CoInitialize()
    try:
        com = win32com.client.Dispatch("V83.COMConnector")
        connection = com.Connect(connection_string)
        return connection
    finally:
        pythoncom.CoUninitialize()


def run_query(connection) -> int:
    # 1C query: SELECT 1 AS One
    query_text = """
    ВЫБРАТЬ
        1 КАК Один
    """.strip()

    q = connection.NewObject("Запрос", query_text)
    result = q.Execute().Unload()
    # result is a collection; take first row and read column "Один"
    return int(result[0].Один)


def main() -> int:
    base_conn = "Srvr='localhost';Ref='WIM_FIN';"
    alt_conn = "Srvr='localhost';Ref='WIM_FIN';App='PyCOM';Locale=ru_RU;"

    safe_print("Connecting to 1C base...")

    connection = None

    # Try with App/Locale first (more robust for some contexts)
    try:
        connection = try_connect(alt_conn)
        safe_print("OK: Connected with App/Locale")
    except Exception as e1:  # noqa: BLE001
        safe_print("WARN: App/Locale connection failed, trying basic string")
        safe_print(f"Reason: {type(e1).__name__}")
        try:
            connection = try_connect(base_conn)
            safe_print("OK: Connected with basic connstring")
        except Exception as e2:  # noqa: BLE001
            safe_print("ERROR: Failed to connect to 1C")
            safe_print(f"Reason: {type(e2).__name__}")
            return 2

    # Validate session by executing a trivial query
    try:
        val = run_query(connection)
        if val == 1:
            safe_print("OK: Query result = 1")
            return 0
        safe_print("ERROR: Unexpected query result")
        return 3
    except Exception as e:  # noqa: BLE001
        safe_print("ERROR: Query execution failed")
        safe_print(f"Reason: {type(e).__name__}")
        return 4


if __name__ == "__main__":
    sys.exit(main())



