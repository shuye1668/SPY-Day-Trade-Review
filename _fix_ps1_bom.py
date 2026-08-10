#!/usr/bin/env python3
"""確保所有 .ps1 都是 UTF-8 with BOM + CRLF。

Windows PowerShell 5.1 讀沒有 BOM 的 .ps1 時會當成 ANSI(cp950)，
中文註解的位元組會被拆成含 } 或引號的字元 → 整支腳本語法錯誤。
每次編輯 .ps1 之後跑一次這支，避免再踩。

    python _fix_ps1_bom.py
"""
import glob
import os
import sys

BOM = b"\xef\xbb\xbf"


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    changed = 0
    for path in sorted(glob.glob(os.path.join(base, "*.ps1"))):
        raw = open(path, "rb").read()
        body = raw[len(BOM):] if raw.startswith(BOM) else raw
        try:
            text = body.decode("utf-8")
        except UnicodeDecodeError:
            print(f"  SKIP (非 UTF-8): {os.path.basename(path)}")
            continue
        want = BOM + text.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8")
        if want != raw:
            open(path, "wb").write(want)
            print(f"  fixed: {os.path.basename(path)}")
            changed += 1
        else:
            print(f"  ok   : {os.path.basename(path)}")
    print(f"\n{changed} 個檔案已修正")
    return 0


if __name__ == "__main__":
    sys.exit(main())
