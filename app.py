"""
Streamlit Grading Matrix Generator

A flexible web application for creating and managing student grading matrices.
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
    initial_sidebar_state="collapsed"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        border-radius: 4px;
    }
    .step-header {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .step-complete {
        color: #28a745;
    }
    .step-pending {
        color: #6c757d;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialize session state with default values."""
    if "config" not in st.session_state:
        st.session_state.config = get_default_config()
    
    if "config_loaded" not in st.session_state:
        st.session_state.config_loaded = False
    
    if "students" not in st.session_state:
        st.session_state.students = []
    
    if "grades" not in st.session_state:
        st.session_state.grades = {}
    
    if "chart_config" not in st.session_state:
        st.session_state.chart_config = {
            "columns": [],
            "chart_type": "bar"
        }


def config_to_json(config: dict) -> str:
    """Convert config dict to JSON string."""
    return json.dumps(config, indent=2)


def parse_students_text(text: str) -> list[str]:
    """Parse student names from text (one per line)."""
    students = []
    for line in text.strip().split("\n"):
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


def create_empty_grades_df(students: list[str], config: dict) -> pd.DataFrame:
    """Create an empty grades DataFrame with proper columns."""
    columns = get_grade_columns(config)
    data = {"Student Name": students}
    for col in columns[1:]:
        data[col] = [None] * len(students)
    return pd.DataFrame(data)


def calculate_row_stats(row: pd.Series, config: dict) -> dict:
    """Calculate Avg A, Avg B, Final Grade, and Qualitative for a row."""
    set_a_cols = [f"A_{p}" for p in config["set_a"]["projects"]]
    set_b_cols = [f"B_{p}" for p in config["set_b"]["projects"]]
    
    # Get numeric values
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
    
    # Calculate averages
    avg_a = sum(set_a_vals) / len(set_a_vals) if set_a_vals else None
    avg_b = sum(set_b_vals) / len(set_b_vals) if set_b_vals else None
    
    # Calculate final grade
    final_grade = None
    weight_a = config["set_a"]["weight"]
    weight_b = config["set_b"]["weight"]
    
    if avg_a is not None and avg_b is not None:
        final_grade = avg_a * weight_a + avg_b * weight_b
    elif avg_a is not None:
        final_grade = avg_a * weight_a
    elif avg_b is not None:
        final_grade = avg_b * weight_b
    
    # Get qualitative grade
    qualitative = ""
    if final_grade is not None:
        for grade in sorted(config.get("qualitative_grades", []), key=lambda x: x["min"], reverse=True):
            if final_grade >= grade["min"]:
                qualitative = grade["label"]
                break
    
    return {
        "Avg A": round(avg_a, config["scale"]["decimal_places"]) if avg_a is not None else "",
        "Avg B": round(avg_b, config["scale"]["decimal_places"]) if avg_b is not None else "",
        "Final Grade": round(final_grade, config["scale"]["decimal_places"]) if final_grade is not None else "",
        "Qualitative": qualitative
    }


def get_available_chart_columns(config: dict) -> list[str]:
    """Get list of columns available for charting."""
    columns = []
    # Individual project columns
    for proj in config["set_a"]["projects"]:
        columns.append(f"A_{proj}")
    for proj in config["set_b"]["projects"]:
        columns.append(f"B_{proj}")
    # Calculated columns
    columns.extend(["Avg A", "Avg B", "Final Grade"])
    return columns


def build_chart_data(grades: dict, config: dict, selected_columns: list[str]) -> pd.DataFrame:
    """Build a DataFrame for charting from grades data."""
    students = st.session_state.students
    if not students or not grades:
        return pd.DataFrame()
    
    chart_data = {"Student": students}
    
    # For each trimester, calculate stats and extract selected columns
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
                
                # Check if it's a calculated column
                if col in ["Avg A", "Avg B", "Final Grade"]:
                    stats = calculate_row_stats(row, config)
                    val = stats.get(col, "")
                    values.append(float(val) if val != "" else None)
                else:
                    # Direct column from data
                    if col in row.index:
                        val = row[col]
                        try:
                            values.append(float(val) if pd.notna(val) and val != "" else None)
                        except (ValueError, TypeError):
                            values.append(None)
                    else:
                        values.append(None)
            
            chart_data[col_key] = values
    
    return pd.DataFrame(chart_data)


def get_workflow_status() -> dict:
    """Get the completion status of each workflow step."""
    config = st.session_state.config
    students = st.session_state.students
    grades = st.session_state.grades
    
    # Check grades status per trimester
    grades_status = {}
    for trimester in config.get("trimesters", []):
        grades_status[trimester] = trimester in grades and grades[trimester] is not None and not grades[trimester].empty
    
    return {
        "config": st.session_state.config_loaded,
        "students": len(students) > 0,
        "grades": grades_status,
        "grades_any": any(grades_status.values()),
        "charts": len(st.session_state.chart_config.get("columns", [])) > 0
    }


def render_sidebar():
    """Render a minimal sidebar for quick config access."""
    st.sidebar.header("Quick Access")
    
    # Config upload
    uploaded_config = st.sidebar.file_uploader(
        "Upload config JSON",
        type=["json"],
        key="sidebar_config_uploader",
        help="Quick upload for config file"
    )
    
    if uploaded_config is not None:
        try:
            user_config = json.load(uploaded_config)
            st.session_state.config = merge_config(user_config)
            st.session_state.config_loaded = True
            st.sidebar.success("✓ Config loaded!")
            st.rerun()
        except json.JSONDecodeError:
            st.sidebar.error("Invalid JSON file")
    
    # Config download
    config_json = config_to_json(st.session_state.config)
    st.sidebar.download_button(
        "📥 Download Config",
        data=config_json,
        file_name="grading_config.json",
        mime="application/json"
    )


def render_step1_config():
    """Render Step 1: Configuration."""
    status = get_workflow_status()
    
    if status["config"]:
        st.header("Step 1: Configuration ✓")
    else:
        st.header("Step 1: Configuration")
    
    st.markdown("Upload a JSON configuration file or use the default settings.")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        uploaded_config = st.file_uploader(
            "Upload config JSON",
            type=["json"],
            key="config_uploader",
            help="Upload a JSON file with your grading configuration"
        )
        
        if uploaded_config is not None:
            try:
                user_config = json.load(uploaded_config)
                st.session_state.config = merge_config(user_config)
                st.session_state.config_loaded = True
                st.success("✓ Config loaded successfully!")
                st.rerun()
            except json.JSONDecodeError:
                st.error("Invalid JSON file")
        
        # Download current config
        config_json = config_to_json(st.session_state.config)
        st.download_button(
            "📥 Download Current Config",
            data=config_json,
            file_name="grading_config.json",
            mime="application/json"
        )
    
    with col2:
        st.subheader("Current Configuration")
        config = st.session_state.config
        
        # Display config summary in columns
        c1, c2, c3 = st.columns(3)
        
        with c1:
            st.markdown(f"""
**Scale:** {config['scale']['min']} - {config['scale']['max']}

**{config['set_a']['name']}** ({int(config['set_a']['weight'] * 100)}%)
""")
            for proj in config['set_a']['projects']:
                st.markdown(f"- {proj}")
        
        with c2:
            st.markdown(f"""
**Decimal Places:** {config['scale']['decimal_places']}

**{config['set_b']['name']}** ({int(config['set_b']['weight'] * 100)}%)
""")
            for proj in config['set_b']['projects']:
                st.markdown(f"- {proj}")
        
        with c3:
            st.markdown("**Trimesters:**")
            for t in config['trimesters']:
                st.markdown(f"- {t}")
            
            st.markdown(f"\n**Qualitative Grades:** {len(config.get('qualitative_grades', []))}")
    
    if status["config"]:
        st.success("✓ Configuration loaded - proceed to Step 2")
    else:
        st.info("Using default configuration. Upload a custom config or proceed with defaults.")


def render_step2_students():
    """Render Step 2: Students."""
    status = get_workflow_status()
    
    if status["students"]:
        st.header(f"Step 2: Students ✓ ({len(st.session_state.students)} loaded)")
    else:
        st.header("Step 2: Students")
    
    st.markdown("Add your student list by pasting names or uploading a file.")
    
    tab1, tab2 = st.tabs(["📝 Paste Names", "📁 Upload File"])
    
    with tab1:
        students_text = st.text_area(
            "Enter student names (one per line)",
            height=200,
            placeholder="John Doe\nJane Smith\nBob Wilson\n...",
            key="students_text"
        )
        
        if students_text:
            parsed = parse_students_text(students_text)
            st.session_state.students = parsed
    
    with tab2:
        uploaded_file = st.file_uploader(
            "Upload student list (CSV or TXT)",
            type=["csv", "txt"],
            key="students_file"
        )
        
        if uploaded_file is not None:
            content = uploaded_file.read().decode("utf-8")
            if uploaded_file.name.endswith(".csv"):
                try:
                    df = pd.read_csv(io.StringIO(content))
                    if len(df.columns) >= 1:
                        st.session_state.students = df.iloc[:, 0].dropna().tolist()
                except Exception:
                    st.session_state.students = parse_students_text(content)
            else:
                st.session_state.students = parse_students_text(content)
    
    # Display student count and validation
    students = st.session_state.students
    if students:
        st.success(f"✓ {len(students)} students loaded - proceed to Step 3")
        
        # Validate students
        issues = validate_students(students)
        for issue in issues:
            if issue["type"] == "warning":
                st.warning(f"⚠️ {issue['message']}")
            else:
                st.error(f"❌ {issue['message']}")
        
        # Show preview
        with st.expander("Preview students"):
            st.write(students[:10])
            if len(students) > 10:
                st.write(f"... and {len(students) - 10} more")
    else:
        st.info("Enter student names above to continue")


def render_step3_grades():
    """Render Step 3: Grade Entry."""
    config = st.session_state.config
    students = st.session_state.students
    status = get_workflow_status()
    
    # Build header with status
    grades_complete = sum(1 for v in status["grades"].values() if v)
    grades_total = len(status["grades"])
    
    if grades_complete == grades_total and grades_total > 0:
        st.header(f"Step 3: Grade Entry ✓ ({grades_complete}/{grades_total} trimesters)")
    elif grades_complete > 0:
        st.header(f"Step 3: Grade Entry ({grades_complete}/{grades_total} trimesters)")
    else:
        st.header("Step 3: Grade Entry")
    
    if not students:
        st.warning("⚠️ Complete Step 2 first: Add students")
        return
    
    if not config["trimesters"]:
        st.warning("⚠️ No trimesters defined in configuration")
        return
    
    st.markdown("Download the CSV template, fill in grades in Excel/Google Sheets, then upload the completed CSV.")
    
    # Show trimester status
    status_cols = st.columns(len(config["trimesters"]))
    for i, trimester in enumerate(config["trimesters"]):
        with status_cols[i]:
            if status["grades"].get(trimester, False):
                st.success(f"✓ {trimester}")
            else:
                st.info(f"○ {trimester}")
    
    st.divider()
    
    # Create tabs for each trimester
    trimester_tabs = st.tabs(config["trimesters"])
    
    for i, trimester in enumerate(config["trimesters"]):
        with trimester_tabs[i]:
            render_trimester_grade_entry(trimester)


def render_trimester_grade_entry(trimester: str):
    """Render grade entry for a single trimester (CSV-only workflow)."""
    config = st.session_state.config
    students = st.session_state.students
    
    # CSV template download and upload
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📥 Download Template")
        template_df = create_empty_grades_df(students, config)
        csv_template = template_df.to_csv(index=False)
        st.download_button(
            "Download CSV Template",
            data=csv_template,
            file_name=f"{trimester.replace(' ', '_')}_template.csv",
            mime="text/csv",
            key=f"template_{trimester}",
            type="primary"
        )
        st.caption("Fill in grades in Excel or Google Sheets")
    
    with col2:
        st.subheader("📤 Upload Filled CSV")
        uploaded_grades = st.file_uploader(
            "Upload completed grades CSV",
            type=["csv"],
            key=f"grades_upload_{trimester}",
            label_visibility="collapsed"
        )
        
        if uploaded_grades is not None:
            try:
                imported_df = pd.read_csv(uploaded_grades)
                st.session_state.grades[trimester] = imported_df
                st.success("✓ Grades imported successfully!")
            except Exception as e:
                st.error(f"Error importing CSV: {e}")
    
    st.divider()
    
    # Show read-only preview of grades if available
    if trimester in st.session_state.grades:
        df = st.session_state.grades[trimester]
        if df is not None and not df.empty:
            st.subheader("📊 Grades Preview")
            
            # Calculate and show preview with calculated columns
            preview_data = []
            for idx, row in df.iterrows():
                stats = calculate_row_stats(row, config)
                row_dict = {"Student Name": row.get("Student Name", "")}
                # Add original grade columns
                for col in df.columns:
                    if col != "Student Name":
                        row_dict[col] = row[col]
                # Add calculated columns
                row_dict.update(stats)
                preview_data.append(row_dict)
            
            preview_df = pd.DataFrame(preview_data)
            st.dataframe(preview_df, hide_index=True, use_container_width=True)
            
            # Validate grades
            issues = validate_grades(df, config)
            if issues:
                scale = config["scale"]
                st.warning(f"⚠️ {len(issues)} grade(s) outside valid range ({scale['min']}-{scale['max']})")
    else:
        st.info("Upload a CSV file to see grade preview")


def render_step4_charts():
    """Render Step 4: Charts (Optional)."""
    config = st.session_state.config
    grades = st.session_state.grades
    students = st.session_state.students
    status = get_workflow_status()
    
    if status["charts"]:
        st.header("Step 4: Charts (Optional) ✓")
    else:
        st.header("Step 4: Charts (Optional)")
    
    if not students:
        st.info("Complete Step 2 first: Add students")
        return
    
    if not status["grades_any"]:
        st.info("Complete Step 3 first: Upload grades for at least one trimester")
        return
    
    st.markdown("Select columns to visualize. Charts will be included in the generated Excel file.")
    
    # Get available columns for charting
    available_columns = get_available_chart_columns(config)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        selected_columns = st.multiselect(
            "Select columns to include in chart",
            options=available_columns,
            default=st.session_state.chart_config.get("columns", []),
            key="chart_columns_select",
            help="Select one or more columns to visualize"
        )
        st.session_state.chart_config["columns"] = selected_columns
    
    with col2:
        chart_type = st.selectbox(
            "Chart type",
            options=["Bar Chart", "Line Chart", "Area Chart"],
            key="chart_type_select"
        )
        st.session_state.chart_config["chart_type"] = chart_type.lower().replace(" ", "_")
    
    # Preview chart
    if selected_columns:
        st.divider()
        st.subheader("📈 Chart Preview")
        
        chart_df = build_chart_data(grades, config, selected_columns)
        
        if not chart_df.empty and len(chart_df.columns) > 1:
            # Set student as index for better chart display
            chart_df_display = chart_df.set_index("Student")
            
            # Remove columns with all NaN
            chart_df_display = chart_df_display.dropna(axis=1, how='all')
            
            if not chart_df_display.empty:
                if chart_type == "Bar Chart":
                    st.bar_chart(chart_df_display)
                elif chart_type == "Line Chart":
                    st.line_chart(chart_df_display)
                else:  # Area Chart
                    st.area_chart(chart_df_display)
                
                st.caption("This chart will be included in the generated Excel file.")
            else:
                st.warning("No data available for the selected columns")
        else:
            st.warning("No data available to chart. Make sure grades are uploaded.")


def render_step5_generate():
    """Render Step 5: Generate Excel."""
    config = st.session_state.config
    students = st.session_state.students
    grades = st.session_state.grades
    chart_config = st.session_state.chart_config
    status = get_workflow_status()
    
    st.header("Step 5: Generate Excel")
    
    # Validation summary
    st.subheader("Validation Summary")
    
    config_issues = validate_config(config)
    student_issues = validate_students(students)
    
    all_issues = config_issues + student_issues
    
    errors = [i for i in all_issues if i["type"] == "error"]
    warnings = [i for i in all_issues if i["type"] == "warning"]
    
    # Show workflow status
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Workflow Status:**")
        if status["config"]:
            st.markdown("✓ Step 1: Configuration loaded")
        else:
            st.markdown("○ Step 1: Using default configuration")
        
        if status["students"]:
            st.markdown(f"✓ Step 2: {len(students)} students loaded")
        else:
            st.markdown("❌ Step 2: No students loaded")
        
        grades_complete = sum(1 for v in status["grades"].values() if v)
        grades_total = len(status["grades"])
        if grades_complete == grades_total and grades_total > 0:
            st.markdown(f"✓ Step 3: All {grades_total} trimesters complete")
        elif grades_complete > 0:
            st.markdown(f"○ Step 3: {grades_complete}/{grades_total} trimesters complete")
        else:
            st.markdown("❌ Step 3: No grades uploaded")
        
        if status["charts"]:
            st.markdown(f"✓ Step 4: Chart configured ({len(chart_config['columns'])} columns)")
        else:
            st.markdown("○ Step 4: No chart configured (optional)")
    
    with col2:
        if errors:
            st.markdown("**Errors:**")
            for error in errors:
                st.error(f"❌ {error['message']}")
        
        if warnings:
            st.markdown("**Warnings:**")
            for warning in warnings:
                st.warning(f"⚠️ {warning['message']}")
        
        if not errors and not warnings:
            st.success("✓ All validations passed")
    
    can_generate = len(errors) == 0 and len(students) > 0
    
    st.divider()
    
    # Generate button
    if st.button("🚀 Generate Excel", disabled=not can_generate, type="primary", use_container_width=True):
        with st.spinner("Generating Excel file..."):
            # Prepare grades data
            grades_by_trimester = {}
            for trimester, df in grades.items():
                if df is not None and not df.empty:
                    grades_by_trimester[trimester] = df
            
            # Build chart data if columns selected
            chart_data = None
            if chart_config.get("columns"):
                chart_data = build_chart_data(grades, config, chart_config["columns"])
            
            # Generate workbook
            wb = generate_workbook(
                config, 
                students, 
                grades_by_trimester,
                chart_data=chart_data,
                chart_config=chart_config
            )
            
            # Save to bytes buffer
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            
            # Download button
            st.download_button(
                "📥 Download Excel File",
                data=buffer,
                file_name=config.get("output_file", "grades.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            st.success("✓ Excel file generated successfully!")
    
    if not can_generate:
        if not students:
            st.info("Complete Step 2 to enable generation")
        elif errors:
            st.info("Fix errors above to enable generation")


def main():
    """Main application entry point."""
    init_session_state()
    
    st.title("📊 Grading Matrix Generator")
    
    # Sidebar for quick access
    render_sidebar()
    
    # Step 1: Configuration
    render_step1_config()
    
    st.divider()
    
    # Step 2: Students
    render_step2_students()
    
    st.divider()
    
    # Step 3: Grade Entry
    render_step3_grades()
    
    st.divider()
    
    # Step 4: Charts
    render_step4_charts()
    
    st.divider()
    
    # Step 5: Generate
    render_step5_generate()


if __name__ == "__main__":
    main()
