"""Configuration schema and defaults for the grading matrix."""

from typing import Any
import copy

DEFAULT_CONFIG: dict[str, Any] = {
    "scale": {
        "min": 0,
        "max": 100,
        "decimal_places": 1
    },
    "set_a": {
        "name": "Set A",
        "weight": 0.70,
        "projects": ["Project 1", "Project 2", "Project 3", "Project 4"]
    },
    "set_b": {
        "name": "Set B",
        "weight": 0.30,
        "projects": ["Project 1", "Project 2"]
    },
    "qualitative_grades": [
        {"label": "Consistent Impact", "min": 90, "max": 100},
        {"label": "Developing Impact", "min": 80, "max": 89},
        {"label": "Emerging", "min": 70, "max": 79},
        {"label": "Needs Support", "min": 60, "max": 69},
        {"label": "Not Yet Meeting", "min": 0, "max": 59}
    ],
    "trimesters": ["Trimester 1", "Trimester 2", "Trimester 3"],
    "output_file": "grades.xlsx"
}


def get_default_config() -> dict[str, Any]:
    """Return a deep copy of the default configuration."""
    return copy.deepcopy(DEFAULT_CONFIG)


def merge_config(user_config: dict[str, Any]) -> dict[str, Any]:
    """
    Merge user configuration with defaults.
    
    User config values override defaults. Missing keys use default values.
    """
    result = get_default_config()
    
    if "scale" in user_config:
        result["scale"].update(user_config["scale"])
    
    if "set_a" in user_config:
        result["set_a"].update(user_config["set_a"])
    
    if "set_b" in user_config:
        result["set_b"].update(user_config["set_b"])
    
    if "qualitative_grades" in user_config:
        result["qualitative_grades"] = user_config["qualitative_grades"]
    
    if "trimesters" in user_config:
        result["trimesters"] = user_config["trimesters"]
    
    if "output_file" in user_config:
        result["output_file"] = user_config["output_file"]
    
    return result

