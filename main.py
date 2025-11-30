# main.py

from loader import browse_for_file, load_togal_export
from parse_logic import parse_all_rows
from pier_logic import build_pier_metrics, print_pier_metrics_summary
from pier_template_writer import write_piers_to_template

import tkinter as tk
from tkinter import filedialog
from pathlib import Path


def main():
    # 1) Select the Togal export
    print("Select your Togal export (.xlsx) file...")
    export_path = browse_for_file("Select Togal export (.xlsx) file")

    if not export_path:
        print("No export file selected. Exiting.")
        return

    # 2) Select the BLANK estimating template
    print("\nSelect your BLANK estimating template (.xlsx/.xlsm) file...")
    template_path = browse_for_file("Select BLANK estimating template (.xlsx/.xlsm) file")

    if not template_path:
        print("No template file selected. Exiting.")
        return

    # -------------------------------------------------------
    # Layer 0: Load Togal export into DataFrame
    # -------------------------------------------------------
    df = load_togal_export(export_path)

    # -------------------------------------------------------
    # Layer 1: Parse DataFrame rows into ParsedRow objects
    # -------------------------------------------------------
    parsed_rows = parse_all_rows(df)

    print("\n=== First 5 parsed rows (Python view of each Togal row) ===")
    for pr in parsed_rows[:5]:
        print(pr)

    # -------------------------------------------------------
    # Layer 2: Build semantic drilled pier metrics
    # -------------------------------------------------------
    pier_metrics = build_pier_metrics(parsed_rows)
    print_pier_metrics_summary(pier_metrics)

    # -------------------------------------------------------
    # Layer 3: User chooses where to save completed file
    # -------------------------------------------------------

    # Detect template extension (.xlsm → keep macros, .xlsx → normal)
    template_suffix = Path(template_path).suffix.lower()
    if template_suffix == ".xlsm":
        default_ext = ".xlsm"
    else:
        default_ext = ".xlsx"

    print(f"\nSelect where to save the COMPLETED estimate ({default_ext})...")

    root = tk.Tk()
    root.withdraw()

    output_path = filedialog.asksaveasfilename(
        title="Save completed estimate as...",
        defaultextension=default_ext,
        filetypes=[
            ("Excel Macro-Enabled Workbook", "*.xlsm"),
            ("Excel Workbook", "*.xlsx"),
            ("All Files", "*.*"),
        ],
    )

    root.destroy()

    if not output_path:
        print("No output file selected; skipping template write.")
        return

    # -------------------------------------------------------
    # Layer 3: Write pier data into template
    # -------------------------------------------------------
    write_piers_to_template(
        template_path=template_path,
        output_path=output_path,
        pier_metrics=pier_metrics,
    )

    print(f"\nWrote pier data into template: {output_path}")
    print("\n=== Done in main.py ===")


if __name__ == "__main__":
    main()
