"""Data models for BOM generator."""

from dataclasses import dataclass
from typing import List, Optional, Dict, Any
from datetime import datetime


@dataclass
class BOMItem:
    """Represents a single item in a Bill of Materials."""
    
    part_number: str
    description: str
    quantity: float
    unit: str = "EA"
    reference_designator: Optional[str] = None
    notes: Optional[str] = None
    
    def __str__(self) -> str:
        return f"{self.part_number} - {self.description} (Qty: {self.quantity})"


@dataclass
class BOM:
    """Represents a complete Bill of Materials."""
    
    name: str
    version: str
    date: datetime
    items: List[BOMItem]
    metadata: Optional[Dict[str, Any]] = None
    
    def __post_init__(self):
        if self.metadata is None:
            self.metadata = {}
    
    def add_item(self, item: BOMItem) -> None:
        """Add an item to the BOM."""
        self.items.append(item)
    
    def get_item_by_part_number(self, part_number: str) -> Optional[BOMItem]:
        """Get an item by part number."""
        for item in self.items:
            if item.part_number == part_number:
                return item
        return None
    
    def __len__(self) -> int:
        return len(self.items)


@dataclass
class BOMComparison:
    """Represents a comparison between two BOMs."""
    
    bom1: BOM
    bom2: BOM
    added_items: List[BOMItem]
    removed_items: List[BOMItem]
    modified_items: List[tuple[BOMItem, BOMItem]]  # (old_item, new_item)
    unchanged_items: List[BOMItem]
    
    def __post_init__(self):
        """Calculate differences between BOMs."""
        if not self.added_items and not self.removed_items and not self.modified_items:
            self._calculate_differences()
    
    def _calculate_differences(self) -> None:
        """Calculate the differences between the two BOMs."""
        bom1_dict = {item.part_number: item for item in self.bom1.items}
        bom2_dict = {item.part_number: item for item in self.bom2.items}
        
        bom1_part_numbers = set(bom1_dict.keys())
        bom2_part_numbers = set(bom2_dict.keys())
        
        # Items only in BOM1 (removed)
        removed_parts = bom1_part_numbers - bom2_part_numbers
        self.removed_items = [bom1_dict[part] for part in removed_parts]
        
        # Items only in BOM2 (added)
        added_parts = bom2_part_numbers - bom1_part_numbers
        self.added_items = [bom2_dict[part] for part in added_parts]
        
        # Items in both (check for modifications)
        common_parts = bom1_part_numbers & bom2_part_numbers
        self.modified_items = []
        self.unchanged_items = []
        
        for part in common_parts:
            item1 = bom1_dict[part]
            item2 = bom2_dict[part]
            
            if (item1.quantity != item2.quantity or 
                item1.description != item2.description or
                item1.unit != item2.unit):
                self.modified_items.append((item1, item2))
            else:
                self.unchanged_items.append(item1)
