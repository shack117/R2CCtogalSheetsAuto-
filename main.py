# main.py

from loader import browse_for_file, load_togal_export
from parse_logic import parse_all_rows  # <-- NEW


def main():
    print("Select your Togal export (.xlsx) file...")
    file_path = browse_for_file()

    if not file_path:
        print("No file selected. Exiting.")
        return

    df = load_togal_export(file_path)

    parsed_rows = parse_all_rows(df)

    print("\n=== First 5 parsed rows (Python view of each Togal row) ===")
    for pr in parsed_rows[:5]:
        print(pr)

    print("\n=== Done in main.py ===")


if __name__ == "__main__":
    main()
