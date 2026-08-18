#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ตรวจโครงสร้างรายงาน TULIBS DOCX โดยไม่ render."""
import argparse
import os
import sys

from make_report import verify_report_structure


def main():
    parser = argparse.ArgumentParser(description="ตรวจว่ารายงานมี 3 ตารางแยกและ schema ครบ")
    parser.add_argument("report")
    args = parser.parse_args()
    if not os.path.exists(args.report):
        sys.exit(f"File not found: {args.report}")
    issues = verify_report_structure(args.report)
    if issues:
        for issue in issues:
            print(f"ERROR: {issue}", file=sys.stderr)
        sys.exit(1)
    print(f"Report structure OK: {args.report}")


if __name__ == "__main__":
    main()
