"""Simple data models for label processing."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Optional


@dataclass
class RawRow:
    DM: str
    GTIN: str
    NAME: str


@dataclass
class EnrichedRow:
    DM: str
    GTIN: str
    NAME: str
    ShortGTIN: str
    ShortName: str
    PROD_DATE: str
    NUM: str
    EXP_DATE: str
    PART_NUM: str
    _FORMAT: str
    _SHELF_LOG: str = ""
    _SHORT_SRC: str = ""

    def as_dict(self) -> Dict[str, str]:
        return self.__dict__.copy()


@dataclass
class ProductInfo:
    FORMAT: str = ""
    SHELF: Dict[str, Optional[int]] = field(default_factory=dict)
    PART_TEMPLATE: str = ""
    SHORTNAME: str = ""


@dataclass
class AppConfig:
    batch_size: int = 1830
    show_print_dialog: bool = False
    formats: Dict[str, str] = field(default_factory=dict)
    product_map_path: str = ""
