"""
Streamlit Grading Matrix Generator

A simple, reliable web application for creating student grading matrices.
All inputs in sidebar, results in main content.
"""

import streamlit as st
import pandas as pd
import json
import io
from typing import Any

from core import (
    get_default_config,
    merge_config,
    validate_config,
    validate_grades,
    validate_students,
    generate_workbook,
)


# Page configuration
st.set_page_config(
    page_title="Grading Matrix Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


def init_session_state():
    """Initialize session state with default values."""
    if "config" not in st.session_state:
        st.session_state.config = get_default_config()
    if "students" not in st.session_state:
        st.session_state.students = []
    if "grades" not in st.session_state:
        st.session_state.grades = {}
    if "chart_columns" not in st.session_state:
        st.session_state.chart_columns = []
    if "chart_type" not in st.session_state:
        st.session_state.chart_type = "Bar Chart"


def reset_app():
    """Reset all session state to defaults."""
    st.session_state.config = get_default_config()
    st.session_state.students = []
    st.session_state.grades = {}
    st.session_state.chart_columns = []
    st.session_state.chart_type = "Bar Chart"


def parse_students(content: str) -> list[str]:
    """Parse student names from text content."""
    students = []
    for line in content.strip().split("\n"):
        line = line.strip()
        if line and not line.startswith("#"):
            students.append(line)
    return students


def get_grade_columns(config: dict) -> list[str]:
    """Get list of grade column names based on config."""
    columns = ["Student Name"]
    for proj in config["set_a"]["projects"]:
        columns.append(f"A_{proj}")
    for proj in config["set_b"]["projects"]:
        columns.append(f"B_{proj}")
    return columns


def create_template_df(students: list[str], config: dict) -> pd.DataFrame:
    """Create an empty grades DataFrame template."""
    columns = get_grade_columns(config)
    data = {"Student Name": students}
    for col in columns[1:]:
        data[col] = [""] * len(students)
    return pd.DataFrame(data)


def calculate_row_stats(row: pd.Series, config: dict) -> dict:
    """Calculate Avg A, Avg B, Final Grade, and Qualitative for a row."""
    set_a_cols = [f"A_{p}" for p in config["set_a"]["projects"]]
    set_b_cols = [f"B_{p}" for p in config["set_b"]["projects"]]
    
    set_a_vals = []
    for col in set_a_cols:
        if col in row and pd.notna(row[col]) and row[col] != "":
            try:
                set_a_vals.append(float(row[col]))
            except (ValueError, TypeError):
                pass
    
    set_b_vals = []
    for col in set_b_cols:
        if col in row and pd.notna(row[col]) and row[col] != "":
            try:
                set_b_vals.append(float(row[col]))
            except (ValueError, TypeError):
                pass
    
    avg_a = sum(set_a_vals) / len(set_a_vals) if set_a_vals else None
    avg_b = sum(set_b_vals) / len(set_b_vals) if set_b_vals else None
    
    final_grade = None
    weight_a = config["set_a"]["weight"]
    weight_b = config["set_b"]["weight"]
    
    if avg_a is not None and avg_b is not None:
        final_grade = avg_a * weight_a + avg_b * weight_b
    elif avg_a is not None:
        final_grade = avg_a * weight_a
    elif avg_b is not None:
        final_grade = avg_b * weight_b
    
    qualitative = ""
    if final_grade is not None:
        for grade in sorted(config.get("qualitative_grades", []), key=lambda x: x["min"], reverse=True):
            if final_grade >= grade["min"]:
                qualitative = grade["label"]
                break
    
    dp = config["scale"]["decimal_places"]
    return {
        "Avg A": round(avg_a, dp) if avg_a is not None else "",
        "Avg B": round(avg_b, dp) if avg_b is not None else "",
        "Final Grade": round(final_grade, dp) if final_grade is not None else "",
        "Qualitative": qualitative
    }


def get_chart_column_options(config: dict) -> list[str]:
    """Get available columns for charting."""
    columns = []
    for proj in config["set_a"]["projects"]:
        columns.append(f"A_{proj}")
    for proj in config["set_b"]["projects"]:
        columns.append(f"B_{proj}")
    columns.extend(["Avg A", "Avg B", "Final Grade"])
    return columns


def build_chart_data(grades: dict, config: dict, selected_columns: list[str]) -> pd.DataFrame:
    """Build DataFrame for charting."""
    students = st.session_state.students
    if not students or not grades or not selected_columns:
        return pd.DataFrame()
    
    chart_data = {"Student": students}
    
    for trimester in config["trimesters"]:
        if trimester not in grades:
            continue
        
        df = grades[trimester]
        
        for col in selected_columns:
            col_key = f"{trimester} - {col}"
            values = []
            
            for student in students:
                student_row = df[df["Student Name"] == student]
                if student_row.empty:
                    values.append(None)
                    continue
                
                row = student_row.iloc[0]
                
                if col in ["Avg A", "Avg B", "Final Grade"]:
                    stats = calculate_row_stats(row, config)
                    val = stats.get(col, "")
                    values.append(float(val) if val != "" else None)
                elif col in row.index:
                    val = row[col]
                    try:
                        values.append(float(val) if pd.notna(val) and val != "" else None)
                    except (ValueError, TypeError):
                        values.append(None)
                else:
                    values.append(None)
            
            chart_data[col_key] = values
    
    return pd.DataFrame(chart_data)


def render_sidebar():
    """Render sidebar with all inputs."""
    config = st.session_state.config
    
    st.sidebar.title("📊 Grading Matrix")
    
    # ===== STEP 1: Configuration =====
    st.sidebar.header("Step 1: Configuration")
    
    config_file = st.sidebar.file_uploader(
        "Upload config JSON (optional)",
        type=["json"],
        key="config_file",
        help="Upload a JSON file to customize settings, or use defaults"
    )
    
    if config_file is not None:
        try:
            user_config = json.load(config_file)
            st.session_state.config = merge_config(user_config)
            st.sidebar.success("✓ Config loaded!")
        except json.JSONDecodeError:
            st.sidebar.error("Invalid JSON file")
    
    # Download current config
    config_json = json.dumps(st.session_state.config, indent=2)
    st.sidebar.download_button(
        "📥 Download current config",
        data=config_json,
        file_name="config.json",
        mime="application/json"
    )
    
    st.sidebar.divider()
    
    # ===== STEP 2: Students =====
    st.sidebar.header("Step 2: Students")
    
    students_file = st.sidebar.file_uploader(
        "Upload student list (TXT or CSV)",
        type=["txt", "csv"],
        key="students_file",
        help="One student name per line"
    )
    
    if students_file is not None:
        content = students_file.read().decode("utf-8")
        if students_file.name.endswith(".csv"):
            try:
                df = pd.read_csv(io.StringIO(content))
                if len(df.columns) >= 1:
                    st.session_state.students = df.iloc[:, 0].dropna().astype(str).tolist()
            except Exception:
                st.session_state.students = parse_students(content)
        else:
            st.session_state.students = parse_students(content)
        
        if st.session_state.students:
            st.sidebar.success(f"✓ {len(st.session_state.students)} students loaded")
    
    st.sidebar.divider()
    
    # ===== STEP 3: Grades =====
    st.sidebar.header("Step 3: Grades")
    
    # Refresh config reference
    config = st.session_state.config
    students = st.session_state.students
    
    if not students:
        st.sidebar.info("Upload students first (Step 2)")
    else:
        for trimester in config["trimesters"]:
            st.sidebar.subheader(f"📁 {trimester}")
            
            # Download template
            template_df = create_template_df(students, config)
            csv_template = template_df.to_csv(index=False)
            st.sidebar.download_button(
                f"📥 Download {trimester} template",
                data=csv_template,
                file_name=f"{trimester.replace(' ', '_')}_template.csv",
                mime="text/csv",
                key=f"template_{trimester}"
            )
            
            # Upload grades
            grades_file = st.sidebar.file_uploader(
                f"Upload {trimester} grades",
                type=["csv"],
                key=f"grades_{trimester}",
                label_visibility="collapsed"
            )
            
            if grades_file is not None:
                try:
                    grades_df = pd.read_csv(grades_file)
                    st.session_state.grades[trimester] = grades_df
                    st.sidebar.success(f"✓ {trimester} grades loaded")
                except Exception as e:
                    st.sidebar.error(f"Error: {e}")
    
    st.sidebar.divider()
    
    # ===== STEP 4: Charts (Optional) =====
    st.sidebar.header("Step 4: Charts (Optional)")
    
    chart_options = get_chart_column_options(config)
    
    st.session_state.chart_columns = st.sidebar.multiselect(
        "Select columns to chart",
        options=chart_options,
        default=st.session_state.chart_columns,
        key="chart_select"
    )
    
    st.session_state.chart_type = st.sidebar.selectbox(
        "Chart type",
        options=["Bar Chart", "Line Chart", "Area Chart"],
        key="chart_type_select"
    )
    
    st.sidebar.divider()
    
    # ===== STEP 5: Generate =====
    st.sidebar.header("Step 5: Generate Excel")
    
    can_generate = len(students) > 0
    
    if st.sidebar.button("🚀 Generate Excel", disabled=not can_generate, type="primary", use_container_width=True):
        st.session_state.generate_excel = True
    
    if not can_generate:
        st.sidebar.info("Upload students to enable")
    
    st.sidebar.divider()
    
    # ===== RESET =====
    if st.sidebar.button("🔄 Reset App", use_container_width=True):
        reset_app()
        st.rerun()


def render_main_content():
    """Render main content area with results."""
    config = st.session_state.config
    students = st.session_state.students
    grades = st.session_state.grades
    
    st.title("Grading Matrix Generator")
    
    # ===== Configuration Summary =====
    st.header("Configuration")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader(f"{config['set_a']['name']} ({int(config['set_a']['weight']*100)}%)")
        for proj in config["set_a"]["projects"]:
            st.write(f"• {proj}")
    
    with col2:
        st.subheader(f"{config['set_b']['name']} ({int(config['set_b']['weight']*100)}%)")
        for proj in config["set_b"]["projects"]:
            st.write(f"• {proj}")
    
    with col3:
        st.subheader("Settings")
        st.write(f"**Scale:** {config['scale']['min']} - {config['scale']['max']}")
        st.write(f"**Trimesters:** {', '.join(config['trimesters'])}")
        st.write(f"**Qualitative grades:** {len(config.get('qualitative_grades', []))}")
    
    st.divider()
    
    # ===== Students =====
    st.header(f"Students ({len(students)})")
    
    if students:
        # Show in columns for compactness
        cols = st.columns(4)
        for i, student in enumerate(students):
            cols[i % 4].write(f"{i+1}. {student}")
    else:
        st.info("No students loaded. Upload a student list in the sidebar (Step 2).")
    
    st.divider()
    
    # ===== Grades Preview =====
    st.header("Grades Preview")
    
    if not grades:
        st.info("No grades loaded. Upload grade CSVs in the sidebar (Step 3).")
    else:
        # Create tabs for each trimester with grades
        loaded_trimesters = [t for t in config["trimesters"] if t in grades]
        
        if loaded_trimesters:
            tabs = st.tabs(loaded_trimesters)
            
            for i, trimester in enumerate(loaded_trimesters):
                with tabs[i]:
                    df = grades[trimester]
                    
                    # Add calculated columns
                    preview_data = []
                    for _, row in df.iterrows():
                        stats = calculate_row_stats(row, config)
                        row_dict = row.to_dict()
                        row_dict.update(stats)
                        preview_data.append(row_dict)
                    
                    preview_df = pd.DataFrame(preview_data)
                    st.dataframe(preview_df, hide_index=True, use_container_width=True)
                    
                    # Validation
                    issues = validate_grades(df, config)
                    if issues:
                        st.warning(f"⚠️ {len(issues)} grade(s) outside valid range")
    
    st.divider()
    
    # ===== Chart Preview =====
    st.header("Chart Preview")
    
    chart_columns = st.session_state.chart_columns
    chart_type = st.session_state.chart_type
    
    if not chart_columns:
        st.info("Select columns to chart in the sidebar (Step 4).")
    elif not grades:
        st.info("Upload grades first to see chart preview.")
    else:
        chart_df = build_chart_data(grades, config, chart_columns)
        
        if not chart_df.empty and len(chart_df.columns) > 1:
            chart_display = chart_df.set_index("Student")
            chart_display = chart_display.dropna(axis=1, how='all')
            
            if not chart_display.empty:
                if chart_type == "Bar Chart":
                    st.bar_chart(chart_display)
                elif chart_type == "Line Chart":
                    st.line_chart(chart_display)
                else:
                    st.area_chart(chart_display)
            else:
                st.warning("No data available for selected columns")
        else:
            st.warning("No chart data available")
    
    st.divider()
    
    # ===== Generate Excel =====
    st.header("Generate Excel")
    
    if getattr(st.session_state, 'generate_excel', False):
        st.session_state.generate_excel = False
        
        with st.spinner("Generating Excel file..."):
            # Prepare data
            grades_by_trimester = {t: df for t, df in grades.items() if df is not None}
            
            # Build chart data
            chart_data = None
            if chart_columns:
                chart_data = build_chart_data(grades, config, chart_columns)
            
            chart_config = {
                "columns": chart_columns,
                "chart_type": chart_type.lower().replace(" ", "_")
            }
            
            # Generate
            wb = generate_workbook(
                config,
                students,
                grades_by_trimester,
                chart_data=chart_data,
                chart_config=chart_config
            )
            
            # Save to buffer
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            
            st.success("✓ Excel file generated!")
            
            st.download_button(
                "📥 Download Excel File",
                data=buffer,
                file_name=config.get("output_file", "grades.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
    else:
        # Show validation status
        config_issues = validate_config(config)
        student_issues = validate_students(students) if students else []
        
        errors = [i for i in config_issues + student_issues if i["type"] == "error"]
        warnings = [i for i in config_issues + student_issues if i["type"] == "warning"]
        
        if errors:
            for e in errors:
                st.error(f"❌ {e['message']}")
        if warnings:
            for w in warnings:
                st.warning(f"⚠️ {w['message']}")
        
        if not errors and students:
            st.success("✓ Ready to generate! Click 'Generate Excel' in the sidebar.")
        elif not students:
            st.info("Upload students in the sidebar to get started.")


def main():
    """Main application entry point."""
    init_session_state()
    render_sidebar()
    render_main_content()


if __name__ == "__main__":
    main()
