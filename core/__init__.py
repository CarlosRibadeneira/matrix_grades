"""Core module for grading matrix generation."""

from .config_schema import DEFAULT_CONFIG, get_default_config, merge_config
from .validators import validate_config, validate_grades, validate_students
from .excel_generator import generate_workbook

__all__ = [
    "DEFAULT_CONFIG",
    "get_default_config",
    "merge_config",
    "validate_config",
    "validate_grades",
    "validate_students",
    "generate_workbook",
]

