"""Random data generation for BOM components.
   Only fields in fields_to_populate are set; all others are left blank in the output.
   Part numbers follow the pattern A001, A002, ... A999, B001, ... with description "{PartNo} Desc".
   When use_long_partno=True, part numbers are 20-50 characters (alphanumeric, unique).
"""

import random
import string
from typing import Dict, Any, List, Set

# Global counter for sequential part numbers (A001, A002, ...); reset at start of each BOM build.
_part_number_counter = 0


def reset_part_number_counter() -> None:
    """Reset the part number sequence. Call at the start of BOM generation."""
    global _part_number_counter
    _part_number_counter = 0


def _prefix_from_index(i: int) -> str:
    """Convert index to letter prefix: 0->A, 1->B, ... 25->Z, 26->AA, 27->AB, ..."""
    if i < 26:
        return chr(65 + i)
    first = (i - 26) // 26
    second = (i - 26) % 26
    return chr(65 + first) + chr(65 + second)


def get_next_partno() -> str:
    """Return next part number in sequence: A001, A002, ... A999, B001, ... CB092, ..."""
    global _part_number_counter
    prefix_idx = _part_number_counter // 1000
    num = (_part_number_counter % 1000) + 1
    _part_number_counter += 1
    prefix = _prefix_from_index(prefix_idx)
    return f"{prefix}{num:03d}"


def get_next_partno_long() -> str:
    """Return a unique part number between 20 and 50 characters (alphanumeric)."""
    global _part_number_counter
    length = random.randint(20, 50)
    # Unique prefix (e.g. P000001) + random chars to reach desired length
    prefix = "P" + str(_part_number_counter).zfill(6)
    _part_number_counter += 1
    need = length - len(prefix)
    if need <= 0:
        return prefix[:length]
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=need))
    return prefix + suffix

# Sample data for random generation
LOCATIONS = ["GS", "WH", "FL", "RM", "WS", "DC"]
PRODUCTLINES = ["JM", "FG", "RM", "CM", "CP"]
SOURCES = ["M", "P", "B", "C"]
SORTCODES = ["COMPBX", "HARDWARE", "LEVEL-1", "LEVEL-2", "ELECTRIC", "ELWR", "SHTCRS", "BARSS", "SHTALUM"]


def random_revision() -> str:
    """e.g. R02."""
    return f"R{random.randint(1, 5):02d}"


def random_row_for_child(
    parent_partno: str,
    sequence: int,
    level: int = 1,
    fields_to_populate: Set[str] | None = None,
    use_long_partno: bool = False,
) -> Dict[str, Any]:
    """Generate one random child row. Only keys in fields_to_populate are set; others are omitted (blank)."""
    all_fields = {
        "PartNo", "Revision", "Description", "AltDescription1", "AltDescription2", "DescExtra",
        "Quantity", "IssueUM", "ConsumptionConv", "UM", "Cost", "Source", "Drawing", "Leadtime",
        "Level", "Location", "Memo1", "Memo2", "Parent", "Productline", "Sequence", "SortCode",
        "Tag", "Category", "BomComplete", "BomComments", "Router"
    }
    fields = fields_to_populate if fields_to_populate else all_fields

    partno = get_next_partno_long() if use_long_partno else get_next_partno()
    # Ensure child PartNo is never equal to parent (would cause blank Parent in BOM and "does not have a parent part number" error)
    while partno == parent_partno:
        partno = get_next_partno_long() if use_long_partno else get_next_partno()
    row = {}

    if "PartNo" in fields:
        row["PartNo"] = partno
    if "Revision" in fields:
        row["Revision"] = random_revision()
    if "Description" in fields:
        row["Description"] = f"{partno} Desc"
    if "AltDescription1" in fields:
        row["AltDescription1"] = f"ALT-DESC-{random.randint(1, 3)}"
    if "AltDescription2" in fields:
        row["AltDescription2"] = f"ALT-DESC-{random.randint(1, 3)}"
    if "DescExtra" in fields:
        row["DescExtra"] = random.choice(["EXTRA", "OPTION", "VARIANT"])
    if "Quantity" in fields:
        row["Quantity"] = random.randint(1, 10)
    if "IssueUM" in fields:
        row["IssueUM"] = "EA"
    if "ConsumptionConv" in fields:
        row["ConsumptionConv"] = round(random.uniform(0.25, 2.0), 2)
    if "UM" in fields:
        row["UM"] = random.choice(["EA", "FT", "M", "KG", "L", "P", "J", "F", "SF", "FT", "SI"])
    if "Cost" in fields:
        row["Cost"] = round(random.uniform(0.5, 250.0), 2)
    if "Source" in fields:
        row["Source"] = random.choice(SOURCES)
    if "Drawing" in fields:
        row["Drawing"] = f"DRAW{random.randint(1, 99)}"
    if "Leadtime" in fields:
        row["Leadtime"] = random.randint(1, 21)
    if "Level" in fields:
        row["Level"] = level
    if "Location" in fields:
        row["Location"] = random.choice(LOCATIONS)
    if "Memo1" in fields:
        row["Memo1"] = f"MEM{random.randint(1, 3)}"
    if "Memo2" in fields:
        row["Memo2"] = f"MEM{random.randint(1, 3)}"
    if "Parent" in fields:
        row["Parent"] = parent_partno
    if "Productline" in fields:
        row["Productline"] = random.choice(PRODUCTLINES)
    if "Sequence" in fields:
        row["Sequence"] = sequence * 100 if sequence > 0 else 100
    if "SortCode" in fields:
        row["SortCode"] = random.choice(SORTCODES)
    if "Tag" in fields:
        row["Tag"] = random.choice(["TG", "TAG1", "TAG2"])
    if "Category" in fields:
        row["Category"] = "Y" if random.random() > 0.2 else ""
    if "BomComplete" in fields:
        row["BomComplete"] = ""
    if "BomComments" in fields:
        row["BomComments"] = f"BOMCOMMENTS-{random.randint(1, 5)}"
    if "Router" in fields:
        row["Router"] = ""

    return row


def random_child_rows(
    parent_partno: str,
    count: int,
    fields_to_populate: Set[str] | None = None,
    level: int = 1,
    use_long_partno: bool = False,
) -> List[Dict[str, Any]]:
    """Generate multiple random child rows. Only selected fields get values; others stay blank."""
    return [
        random_row_for_child(parent_partno, i + 1, level=level, fields_to_populate=fields_to_populate, use_long_partno=use_long_partno)
        for i in range(count)
    ]
