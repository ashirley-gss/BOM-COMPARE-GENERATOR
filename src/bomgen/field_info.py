"""Field metadata from BOM_COMPARE_TEMPLATE Field_Info tab.
   Defines field type and maximum length for validation and random data generation.
"""

from typing import Any

# Field type: String, Double, or Int32
FIELD_TYPE = {
    "PartNo": "String",
    "Revision": "String",
    "Description": "String",
    "AltDescription1": "String",
    "AltDescription2": "String",
    "DescExtra": "String",
    "Quantity": "Double",
    "IssueUM": "String",
    "ConsumptionConv": "Double",
    "UM": "String",
    "Cost": "Double",
    "Source": "String",
    "Drawing": "String",
    "Leadtime": "Double",
    "Level": "Int32",
    "Location": "String",
    "Memo1": "String",
    "Memo2": "String",
    "Parent": "String",
    "Productline": "String",
    "Sequence": "Int32",
    "SortCode": "String",
    "Tag": "String",
    "Category": "String",
    "BomComplete": "String",
    "BomComments": "String",
    "Router": "String",
}

# Maximum length for String fields; numeric fields use this for display/decimal places where applicable
MAX_LENGTH = {
    "PartNo": 17,
    "Revision": 3,
    "Description": 30,
    "AltDescription1": 30,
    "AltDescription2": 30,
    "DescExtra": 30,
    "Quantity": 4,  # decimal places for Double
    "IssueUM": 2,
    "ConsumptionConv": 4,
    "UM": 2,
    "Cost": 4,
    "Source": 1,
    "Drawing": 20,
    "Leadtime": 1,  # single digit for Leadtime
    "Level": 2,
    "Location": 30,
    "Memo1": 30,
    "Memo2": 30,
    "Parent": 17,
    "Productline": 2,
    "Sequence": None,  # Int32, no string length
    "SortCode": 12,
    "Tag": 6,
    "Category": 1,
    "BomComplete": 1,
    "BomComments": 1,
    "Router": 20,
}


def apply_field_constraints(value: Any, field_name: str, max_length_override: int | None = None) -> Any:
    """Apply Field_Info type and max length to a value. Returns constrained value.
    Use max_length_override for PartNo/Parent when Use Long Part is enabled (e.g. 50)."""
    ftype = FIELD_TYPE.get(field_name, "String")

    if value is None or value == "":
        return "" if ftype == "String" else None

    max_len = max_length_override if max_length_override is not None else MAX_LENGTH.get(field_name)

    if ftype == "Int32":
        return int(value)
    if ftype == "Double":
        v = float(value)
        if max_len == 1:  # Leadtime: single digit 1-9
            v = max(1, min(9, int(v)))
        return v

    # String: truncate to max length
    s = str(value).strip()
    if max_len is not None and len(s) > max_len:
        return s[:max_len]
    return s
