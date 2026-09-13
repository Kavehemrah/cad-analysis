import argparse
import os
from datetime import datetime

import pythoncom
import win32com.client

from cad_analysis import Detector, DEFAULT_CLEARANCE_M
from report_excel import write_xlsx


def parse_args():
    parser = argparse.ArgumentParser(
        description="Analyze Drainage Pipe E against AG sleepers"
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Create colored moved copies in AutoCAD; originals are preserved",
    )
    parser.add_argument(
        "--save",
        action="store_true",
        help="Save the DWG after --apply",
    )
    parser.add_argument(
        "--clearance-mm",
        type=float,
        default=DEFAULT_CLEARANCE_M * 1000.0,
        help="Required final clearance in millimeters",
    )
    parser.add_argument(
        "--report",
        default=None,
        help="Output XLSX path",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        clearance_m = args.clearance_mm / 1000.0
        detector = Detector(doc, clearance_m=clearance_m)

        print("=" * 100)
        print("CAD DRAINAGE ANALYSIS")
        print("=" * 100)
        print(f"Drawing : {doc.Name}")
        print(f"Mode    : {'APPLY' if args.apply else 'ANALYZE ONLY'}")
        print(f"Clearance: {args.clearance_mm:.1f} mm")
        print("Scope   : FULL DRAWING")
        print()

        results = detector.analyze_all()

        print(f"Drainages analyzed: {len(results)}")
        for result in results:
            print(
                f"{result.drainage:>6}  "
                f"{result.status:<10}  "
                f"{result.confidence:<6}  "
                f"sleeper={result.sleeper or '-':>6}  "
                f"move={result.required_move_mm:8.1f} mm  "
                f"dir=({result.move_dir_x:+.4f},{result.move_dir_y:+.4f})"
            )

        moved = 0
        if args.apply:
            print()
            print("Creating moved copies; original drainage objects will be preserved...")
            moved = detector.apply_moves(results, save=args.save)
            print(f"Moved copies created: {moved}")

        report_path = args.report
        if not report_path:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = os.path.join(
                os.getcwd(),
                f"drainage_analysis_{stamp}.xlsx",
            )

        write_xlsx(report_path, results, args.clearance_mm)

        print()
        print(f"Excel report: {report_path}")
        if args.save and args.apply:
            print("DWG saved.")
        else:
            print("DWG not saved automatically.")
        print("=" * 100)

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
