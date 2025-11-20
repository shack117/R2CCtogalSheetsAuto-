# main.py

from loader import browse_for_file, load_togal_export
from parse_logic import parse_all_rows
from pier_logic import build_pier_metrics, print_pier_metrics_summary


def main():
    print("Select your Togal export (.xlsx) file...")
    file_path = browse_for_file()

    if not file_path:
        print("No file selected. Exiting.")
        return

    # Load Togal export into a DataFrame
    df = load_togal_export(file_path)

    # Layer 1: parse raw DataFrame rows into ParsedRow objects
    parsed_rows = parse_all_rows(df)

    print("\n=== First 5 parsed rows (Python view of each Togal row) ===")
    for pr in parsed_rows[:5]:
        print(pr)

    # Layer 2: build semantic drilled pier metrics from ParsedRow list
    pier_metrics = build_pier_metrics(parsed_rows)
    print_pier_metrics_summary(pier_metrics)

    print("\n=== Done in main.py ===")


if __name__ == "__main__":
    main()
