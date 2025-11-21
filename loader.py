import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog


def browse_for_file(title: str = "Select File") -> str:
    """Open a Windows file-browse dialog and return the selected path."""
    root = tk.Tk()
    root.withdraw()  # hide main tk window

    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel Files", "*.xlsx *.xlsm"),
            ("All Files", "*.*"),
        ]
    )

    root.destroy()
    return file_path


def load_togal_export(path_str: str) -> pd.DataFrame:
    """Read a Togal export Excel file into a pandas DataFrame and print basic info."""
    path = Path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    print(f"\n=== Reading Togal export ===")
    print(f"File: {path}")

    df = pd.read_excel(path)

    print("\n=== Columns ===")
    print(list(df.columns))

    print("\n=== First 5 rows ===")
    print(df.head(5))

    # Quick sanity check on important columns
    key_cols = ["Classification Folder", "Classification"]
    print("\n=== Column Check ===")
    for col in key_cols:
        if col in df.columns:
            print(f"[OK] {col!r} found")
        else:
            print(f"[WARN] {col!r} not found!")

    return df
