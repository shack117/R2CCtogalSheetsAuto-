# pier_logic.py

from dataclasses import dataclass
from typing import Dict, Tuple, List, Optional

from parse_logic import ParsedRow


@dataclass
class PierMetrics:
    """
    Semantic representation of a drilled pier condition,
    derived from one or more ParsedRow entries.
    """
    tier: str                   # e.g. "TIER 1"
    classification: str         # e.g. "PIER - 1"

    shaft_dia_in: Optional[float]   # inches
    bell_dia_in: Optional[float]    # inches
    depth_ft: Optional[float]       # feet
    count: float                    # number of piers (EA)
    total_length_lf: Optional[float]  # depth_ft * count (LF)


# ---------- helpers ----------

def _normalize(text: Optional[str]) -> str:
    return (text or "").strip()


def _normalize_upper(text: Optional[str]) -> str:
    return _normalize(text).upper()


def _normalize_tier(tier: Optional[str]) -> str:
    """
    Normalize breakdown tier text so we can safely use it as a key.
    """
    t = _normalize_upper(tier)
    if t.startswith("TIER"):
        t = t.replace("TIER", "TIER ").replace("  ", " ").strip()
    if not t:
        t = "UNASSIGNED"
    return t


def _ft_to_in(value_ft: float) -> float:
    """Convert feet to inches with full precision."""
    return value_ft * 12.0


def is_pier_row(row: ParsedRow) -> bool:
    """
    Decide if a ParsedRow represents a drilled pier condition.
    """
    name = _normalize_upper(row.classification)
    folder = _normalize_upper(row.folder)

    if "PIER" in name:
        return True
    if "PIER" in folder:
        return True

    return False


# ---------- core layer-2 logic ----------

def build_pier_metrics(rows: List[ParsedRow]) -> Dict[Tuple[str, str], PierMetrics]:
    """
    Aggregate ParsedRow objects into PierMetrics, grouped by:
       (tier, classification)
    """
    acc: Dict[Tuple[str, str], Dict[str, Optional[float]]] = {}

    for r in rows:
        if not is_pier_row(r):
            continue

        tier = _normalize_tier(r.breakdown_tier)
        cls = _normalize(r.classification)
        key = (tier, cls)

        if key not in acc:
            acc[key] = {
                "shaft_dia_in": None,
                "bell_dia_in": None,
                "depth_ft": None,
                "count": 0.0,
            }

        m = acc[key]

        # --- Shaft diameter from Width (+ UOM) ---
        if r.width is not None and r.width_uom:
            uom = _normalize_upper(r.width_uom)

            if uom in ("IN", "INCH", "INCHES"):
                shaft = r.width

            elif uom in ("FT", "FEET", "FOOT"):
                # Convert exactly — no rounding here
                shaft = _ft_to_in(r.width)

            else:
                shaft = None

            if shaft is not None:
                m["shaft_dia_in"] = shaft

                # Bell = shaft if nothing else provided
                if m["bell_dia_in"] is None:
                    m["bell_dia_in"] = shaft

        # --- Depth from Height ---
        if r.height is not None and r.height_uom:
            uom = _normalize_upper(r.height_uom)

            if uom in ("FT", "FEET", "FOOT"):
                depth_ft = r.height
            elif uom in ("IN", "INCH", "INCHES"):
                depth_ft = r.height / 12.0
            else:
                depth_ft = None

            if depth_ft is not None:
                m["depth_ft"] = depth_ft

        # --- Count (EA) ---
        if r.quantity1 is not None:
            if not r.quantity1_uom:
                m["count"] += r.quantity1
            else:
                q_uom = _normalize_upper(r.quantity1_uom)
                if q_uom in ("EA", "EACH", "COUNT", "#"):
                    m["count"] += r.quantity1

    # Final conversion into PierMetrics objects
    result: Dict[Tuple[str, str], PierMetrics] = {}

    for (tier, cls), m in acc.items():
        depth_ft = m["depth_ft"]
        count = m["count"]

        total_length = depth_ft * count if depth_ft is not None else None

        result[(tier, cls)] = PierMetrics(
            tier=tier,
            classification=cls,
            shaft_dia_in=m["shaft_dia_in"],
            bell_dia_in=m["bell_dia_in"],
            depth_ft=depth_ft,
            count=count,
            total_length_lf=total_length,
        )

    return result


def print_pier_metrics_summary(piers: Dict[Tuple[str, str], PierMetrics]) -> None:
    """
    Debug helper to verify what we extracted for each pier condition.
    """
    print("\n=== Pier Metrics Summary ===")
    for (tier, cls), m in piers.items():
        print(
            f"{tier} | {cls}: "
            f"shaft={m.shaft_dia_in} in, "
            f"bell={m.bell_dia_in} in, "
            f"depth={m.depth_ft} ft, "
            f"count={m.count}, "
            f"total LF={m.total_length_lf}"
        )
