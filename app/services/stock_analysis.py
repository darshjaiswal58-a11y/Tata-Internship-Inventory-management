"""Stock analysis and dashboard summary services."""

from app.legacy import (
    analyze_group_critical,
    analyze_stock_zone_materials,
    build_summary,
    process_upload,
)

__all__ = [
    "analyze_group_critical",
    "analyze_stock_zone_materials",
    "build_summary",
    "process_upload",
]

