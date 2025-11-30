# parse_logic.py

from dataclasses import dataclass
from typing import Optional, List

import pandas as pd


@dataclass
class ParsedRow:
    """One Togal row, cleaned up into a Python-friendly structure."""
    folder: str
    classification: str

    quantity1: Optional[float]
    quantity1_uom: Optional[str]

    quantity2: Optional[float]
    quantity2_uom: Optional[str]

    height: Optional[float]
    height_uom: Optional[str]

    width: Optional[float]
    width_uom: Optional[str]

    thickness: Optional[float]
    thickness_uom: Optional[str]

    length: Optional[float]
    length_uom: Optional[str]

    breakdown_tier: Optional[str]
    breakdown_item: Optional[str]

def _get_number(row: pd.Series, col_name: str) -> Optional[float]:
    """Return a float if the cell is numeric, otherwise None."""
    if col_name not in row or pd.isna(row[col_name]):
        return None
    try:
        return float(row[col_name])
    except (TypeError, ValueError):
        return None


def _get_text(row: pd.Series, col_name: str) -> Optional[str]:
    """Return a stripped string if present, otherwise None."""
    if col_name not in row or pd.isna(row[col_name]):
        return None
    text = str(row[col_name]).strip()
    return text if text else None


def parse_row(row: pd.Series) -> ParsedRow:
    """
    Take a single pandas row from the Togal export and turn it into a ParsedRow
    that matches how you look at the row as a human.
    """

    folder = _get_text(row, "Classification Folder") or ""
    classification = _get_text(row, "Classification") or ""

    quantity1 = _get_number(row, "Quantity 1")
    quantity1_uom = _get_text(row, "Quantity1 UOM")

    quantity2 = _get_number(row, "Quantity 2")
    quantity2_uom = _get_text(row, "Quantity2 UOM")

    height = _get_number(row, "Height")
    height_uom = _get_text(row, "Height UOM")

    width = _get_number(row, "Width")
    width_uom = _get_text(row, "Width UOM")

    thickness = _get_number(row, "Thickness")
    thickness_uom = _get_text(row, "Thickness UOM")

    length = _get_number(row, "Length")
    length_uom = _get_text(row, "Length UOM")

    breakdown_tier = _get_text(row, "Breakdown Tier")
    breakdown_item = _get_text(row, "Breakdown Item")

    return ParsedRow(
        folder=folder,
        classification=classification,
        quantity1=quantity1,
        quantity1_uom=quantity1_uom,
        quantity2=quantity2,
        quantity2_uom=quantity2_uom,
        height=height,
        height_uom=height_uom,
        width=width,
        width_uom=width_uom,
        thickness=thickness,
        thickness_uom=thickness_uom,
        length=length,
        length_uom=length_uom,
        breakdown_tier=breakdown_tier,
        breakdown_item=breakdown_item,
    )


def parse_all_rows(df: pd.DataFrame) -> List[ParsedRow]:
    """
    Walk the entire Togal DataFrame row-by-row and return a list of ParsedRow
    objects. This mirrors how you normally work: one row at a time.
    """
    parsed: List[ParsedRow] = []
    for _, row in df.iterrows():
        parsed_row = parse_row(row)
        parsed.append(parsed_row)
    return parsed
