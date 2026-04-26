"""Data models for the PDF Tag Extractor application.

Defines data classes used across the application for type safety,
configuration management, and structured data transfer between layers.
"""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class TagRecord:
    """Represents a single extracted tag from an engineering diagram.

    Attributes:
        page: Page number where the tag was found (1-indexed).
        area: Primary area (e.g., 'TOPSIDE', 'ACCOMODATION') based on font size ~14pt.
        subarea: Sub-area (e.g., 'M-06 - GAS DEHYDRATATION...', 'F-DECK') ~12pt.
        tag: The engineering tag identifier (e.g., 'COR-M06-5518406A').
        location: Physical location/room where the tag is installed
            (e.g., '4-MEN CABIN (A621)'). Empty string if not found.
    """

    page: int
    area: str
    subarea: str
    tag: str
    location: str = ""


@dataclass
class ExtractionConfig:
    """Configuration parameters for the PDF tag extraction engine.

    These values are calibrated for standard engineering one-line diagrams.
    Adjust zone boundaries and font size ranges to match different document
    layouts.

    Attributes:
        tag_pattern: Regex pattern to identify engineering tags.
        zone_min_x: Left boundary (x0) of the central content zone in points.
        zone_max_x: Right boundary (x0) of the central content zone in points.
        header1_size_range: Font size range (min, max) for primary headers.
        header2_size_range: Font size range (min, max) for secondary headers.
        exclusion_keywords: Words that, when found on the same line as a tag,
            cause the tag to be excluded (e.g., cross-reference labels).
        location_x_tolerance: Horizontal padding (pts) added to each side of
            the tag's x-range when searching for location text above.
        location_y_max_distance: Max vertical distance (pts) to look above
            a tag for bold location text. Larger values capture multi-line
            locations (e.g., 'TELECOM CONTROL' + 'ROOM (A707)').
    """

    tag_pattern: str = r"[A-Z]{2,3}-(?:[A-Z0-9]+-)?\d{7}[A-Z]?"
    zone_min_x: float = 600.0
    zone_max_x: float = 1400.0
    header1_size_range: tuple[float, float] = (13.6, 14.1)
    header2_size_range: tuple[float, float] = (11.6, 12.1)
    exclusion_keywords: list[str] = field(
        default_factory=lambda: ["FROM", "TO"]
    )
    location_x_tolerance: float = 15.0
    location_y_max_distance: float = 35.0

