"""Validation utilities for grades and configuration."""

from typing import Any
import pandas as pd


def validate_config(config: dict[str, Any]) -> list[dict[str, str]]:
    """
    Validate configuration and return list of issues.
    
    Returns:
        List of dicts with 'type' (error/warning) and 'message'.
    """
    issues = []
    
    scale = config.get("scale", {})
    scale_min = scale.get("min", 0)
    scale_max = scale.get("max", 100)
    
    # Check weights sum to 100%
    weight_a = config.get("set_a", {}).get("weight", 0)
    weight_b = config.get("set_b", {}).get("weight", 0)
    total_weight = weight_a + weight_b
    
    if abs(total_weight - 1.0) > 0.001:
        issues.append({
            "type": "warning",
            "message": f"Weights sum to {total_weight * 100:.0f}% (should be 100%)"
        })
    
    # Check qualitative grade ranges are within scale
    for grade in config.get("qualitative_grades", []):
        grade_min = grade.get("min", 0)
        grade_max = grade.get("max", 100)
        label = grade.get("label", "Unknown")
        
        if grade_min < scale_min or grade_max > scale_max:
            issues.append({
                "type": "error",
                "message": f"Qualitative grade '{label}' range ({grade_min}-{grade_max}) is outside scale ({scale_min}-{scale_max})"
            })
    
    # Check for empty project lists
    if not config.get("set_a", {}).get("projects"):
        issues.append({
            "type": "error",
            "message": "Set A has no projects defined"
        })
    
    if not config.get("set_b", {}).get("projects"):
        issues.append({
            "type": "error",
            "message": "Set B has no projects defined"
        })
    
    # Check for empty trimesters
    if not config.get("trimesters"):
        issues.append({
            "type": "error",
            "message": "No trimesters defined"
        })
    
    return issues


def validate_grades(
    grades_df: pd.DataFrame,
    config: dict[str, Any]
) -> list[dict[str, Any]]:
    """
    Validate grades in a DataFrame against the configured scale.
    
    Returns:
        List of dicts with 'row', 'column', 'value', and 'message'.
    """
    issues = []
    
    scale = config.get("scale", {})
    scale_min = scale.get("min", 0)
    scale_max = scale.get("max", 100)
    
    # Get project columns (exclude student name and calculated columns)
    set_a_projects = config.get("set_a", {}).get("projects", [])
    set_b_projects = config.get("set_b", {}).get("projects", [])
    
    # Create unique column names for Set B if there's overlap
    set_a_cols = [f"A_{p}" for p in set_a_projects]
    set_b_cols = [f"B_{p}" for p in set_b_projects]
    grade_columns = set_a_cols + set_b_cols
    
    for col in grade_columns:
        if col not in grades_df.columns:
            continue
            
        for idx, value in grades_df[col].items():
            if pd.isna(value) or value == "":
                continue
            
            try:
                num_value = float(value)
                if num_value < scale_min or num_value > scale_max:
                    issues.append({
                        "row": idx,
                        "column": col,
                        "value": value,
                        "message": f"Grade {value} is outside scale ({scale_min}-{scale_max})"
                    })
            except (ValueError, TypeError):
                issues.append({
                    "row": idx,
                    "column": col,
                    "value": value,
                    "message": f"Invalid grade value: {value}"
                })
    
    return issues


def validate_students(students: list[str]) -> list[dict[str, str]]:
    """
    Validate student list and return issues.
    
    Returns:
        List of dicts with 'type' and 'message'.
    """
    issues = []
    
    if not students:
        issues.append({
            "type": "error",
            "message": "No students provided"
        })
        return issues
    
    # Check for duplicates
    seen = set()
    duplicates = []
    for student in students:
        if student in seen:
            duplicates.append(student)
        seen.add(student)
    
    if duplicates:
        issues.append({
            "type": "warning",
            "message": f"Duplicate student names: {', '.join(duplicates)}"
        })
    
    # Check for empty names
    empty_count = sum(1 for s in students if not s.strip())
    if empty_count:
        issues.append({
            "type": "warning",
            "message": f"{empty_count} empty student name(s) found"
        })
    
    return issues

