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
    initial_sidebar_state="expanded"
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
    .validation-warning {
        background-color: #fff3cd;
        border: 1px solid #ffc107;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .validation-error {
        background-color: #f8d7da;
        border: 1px solid #dc3545;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialize session state with default values."""
    if "config" not in st.session_state:
        st.session_state.config = get_default_config()
    
    if "students" not in st.session_state:
        st.session_state.students = []
    
    if "grades" not in st.session_state:
        st.session_state.grades = {}


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


def render_sidebar():
    """Render the configuration sidebar."""
    st.sidebar.header("⚙️ Configuration")
    
    # Config import/export
    st.sidebar.subheader("Import/Export Config")
    
    uploaded_config = st.sidebar.file_uploader(
        "Upload config JSON",
        type=["json"],
        key="config_uploader"
    )
    
    if uploaded_config is not None:
        try:
            user_config = json.load(uploaded_config)
            st.session_state.config = merge_config(user_config)
            st.sidebar.success("Config loaded!")
            st.rerun()
        except json.JSONDecodeError:
            st.sidebar.error("Invalid JSON file")
    
    config_json = config_to_json(st.session_state.config)
    st.sidebar.download_button(
        "📥 Download Current Config",
        data=config_json,
        file_name="grading_config.json",
        mime="application/json"
    )
    
    st.sidebar.divider()
    
    # Scale settings
    st.sidebar.subheader("📏 Grade Scale")
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        scale_min = st.number_input(
            "Min",
            value=st.session_state.config["scale"]["min"],
            step=1,
            key="scale_min"
        )
    with col2:
        scale_max = st.number_input(
            "Max",
            value=st.session_state.config["scale"]["max"],
            step=1,
            key="scale_max"
        )
    
    decimal_places = st.sidebar.number_input(
        "Decimal places",
        min_value=0,
        max_value=3,
        value=st.session_state.config["scale"]["decimal_places"],
        key="decimal_places"
    )
    
    st.session_state.config["scale"] = {
        "min": scale_min,
        "max": scale_max,
        "decimal_places": decimal_places
    }
    
    st.sidebar.divider()
    
    # Set A configuration
    st.sidebar.subheader(f"📘 {st.session_state.config['set_a']['name']}")
    
    set_a_name = st.sidebar.text_input(
        "Set A Name",
        value=st.session_state.config["set_a"]["name"],
        key="set_a_name"
    )
    
    set_a_weight = st.sidebar.slider(
        "Set A Weight (%)",
        min_value=0,
        max_value=100,
        value=int(st.session_state.config["set_a"]["weight"] * 100),
        key="set_a_weight"
    )
    
    set_a_projects_text = st.sidebar.text_area(
        "Set A Projects (one per line)",
        value="\n".join(st.session_state.config["set_a"]["projects"]),
        height=100,
        key="set_a_projects"
    )
    set_a_projects = [p.strip() for p in set_a_projects_text.split("\n") if p.strip()]
    
    st.session_state.config["set_a"] = {
        "name": set_a_name,
        "weight": set_a_weight / 100,
        "projects": set_a_projects
    }
    
    st.sidebar.divider()
    
    # Set B configuration
    st.sidebar.subheader(f"📗 {st.session_state.config['set_b']['name']}")
    
    set_b_name = st.sidebar.text_input(
        "Set B Name",
        value=st.session_state.config["set_b"]["name"],
        key="set_b_name"
    )
    
    set_b_weight = st.sidebar.slider(
        "Set B Weight (%)",
        min_value=0,
        max_value=100,
        value=int(st.session_state.config["set_b"]["weight"] * 100),
        key="set_b_weight"
    )
    
    set_b_projects_text = st.sidebar.text_area(
        "Set B Projects (one per line)",
        value="\n".join(st.session_state.config["set_b"]["projects"]),
        height=100,
        key="set_b_projects"
    )
    set_b_projects = [p.strip() for p in set_b_projects_text.split("\n") if p.strip()]
    
    st.session_state.config["set_b"] = {
        "name": set_b_name,
        "weight": set_b_weight / 100,
        "projects": set_b_projects
    }
    
    # Weight validation
    total_weight = set_a_weight + set_b_weight
    if total_weight != 100:
        st.sidebar.warning(f"⚠️ Weights sum to {total_weight}% (should be 100%)")
    
    st.sidebar.divider()
    
    # Trimesters
    st.sidebar.subheader("📅 Trimesters")
    
    trimesters_text = st.sidebar.text_area(
        "Trimesters (one per line)",
        value="\n".join(st.session_state.config["trimesters"]),
        height=100,
        key="trimesters"
    )
    trimesters = [t.strip() for t in trimesters_text.split("\n") if t.strip()]
    st.session_state.config["trimesters"] = trimesters
    
    st.sidebar.divider()
    
    # Qualitative grades
    st.sidebar.subheader("🏆 Qualitative Grades")
    
    qual_grades = st.session_state.config.get("qualitative_grades", [])
    
    # Display as editable dataframe
    qual_df = pd.DataFrame(qual_grades) if qual_grades else pd.DataFrame(columns=["label", "min", "max"])
    
    edited_qual = st.sidebar.data_editor(
        qual_df,
        num_rows="dynamic",
        column_config={
            "label": st.column_config.TextColumn("Label", width="medium"),
            "min": st.column_config.NumberColumn("Min", width="small"),
            "max": st.column_config.NumberColumn("Max", width="small"),
        },
        key="qual_grades_editor",
        hide_index=True
    )
    
    # Update config with edited qualitative grades
    st.session_state.config["qualitative_grades"] = edited_qual.to_dict("records")
    
    # Validate qualitative ranges
    scale = st.session_state.config["scale"]
    for grade in st.session_state.config["qualitative_grades"]:
        if grade.get("min", 0) < scale["min"] or grade.get("max", 100) > scale["max"]:
            st.sidebar.error(f"⚠️ '{grade.get('label', 'Unknown')}' range is outside scale")
            break


def render_students_section():
    """Render the students input section."""
    st.header("👥 Students")
    
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
                # Try to parse as CSV
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
        st.success(f"✓ {len(students)} students loaded")
        
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
        st.info("Enter student names above to begin")


def render_grade_entry():
    """Render the grade entry section with trimester tabs."""
    st.header("📝 Grade Entry")
    
    config = st.session_state.config
    students = st.session_state.students
    
    if not students:
        st.warning("Please add students first")
        return
    
    if not config["trimesters"]:
        st.warning("Please configure trimesters in the sidebar")
        return
    
    # Initialize grades for each trimester if not exists
    for trimester in config["trimesters"]:
        if trimester not in st.session_state.grades:
            st.session_state.grades[trimester] = create_empty_grades_df(students, config)
    
    # Create tabs for each trimester
    trimester_tabs = st.tabs(config["trimesters"])
    
    for i, trimester in enumerate(config["trimesters"]):
        with trimester_tabs[i]:
            render_trimester_grade_entry(trimester)


def render_trimester_grade_entry(trimester: str):
    """Render grade entry for a single trimester."""
    config = st.session_state.config
    students = st.session_state.students
    
    # CSV template download
    col1, col2 = st.columns([1, 3])
    
    with col1:
        template_df = create_empty_grades_df(students, config)
        csv_template = template_df.to_csv(index=False)
        st.download_button(
            "📥 Download CSV Template",
            data=csv_template,
            file_name=f"{trimester.replace(' ', '_')}_template.csv",
            mime="text/csv",
            key=f"template_{trimester}"
        )
    
    with col2:
        uploaded_grades = st.file_uploader(
            "Upload grades CSV",
            type=["csv"],
            key=f"grades_upload_{trimester}"
        )
        
        if uploaded_grades is not None:
            try:
                imported_df = pd.read_csv(uploaded_grades)
                st.session_state.grades[trimester] = imported_df
                st.success("Grades imported!")
            except Exception as e:
                st.error(f"Error importing CSV: {e}")
    
    st.divider()
    
    # Get current grades DataFrame
    current_df = st.session_state.grades.get(trimester)
    if current_df is None:
        current_df = create_empty_grades_df(students, config)
    
    # Ensure DataFrame has correct columns based on current config
    expected_columns = get_grade_columns(config)
    
    # Check if columns match, if not recreate
    if list(current_df.columns) != expected_columns:
        # Try to preserve existing data
        new_df = create_empty_grades_df(students, config)
        for col in current_df.columns:
            if col in new_df.columns:
                new_df[col] = current_df[col].values[:len(students)] if len(current_df) >= len(students) else current_df[col].tolist() + [None] * (len(students) - len(current_df))
        current_df = new_df
    
    # Also check if student list changed
    if len(current_df) != len(students) or list(current_df["Student Name"]) != students:
        new_df = create_empty_grades_df(students, config)
        # Try to match existing grades by student name
        for idx, student in enumerate(students):
            if student in current_df["Student Name"].values:
                old_row = current_df[current_df["Student Name"] == student].iloc[0]
                for col in new_df.columns:
                    if col in old_row.index and col != "Student Name":
                        new_df.at[idx, col] = old_row[col]
        current_df = new_df
    
    # Configure column types for the data editor
    column_config = {
        "Student Name": st.column_config.TextColumn(
            "Student Name",
            disabled=True,
            width="medium"
        )
    }
    
    scale = config["scale"]
    for proj in config["set_a"]["projects"]:
        col_name = f"A_{proj}"
        column_config[col_name] = st.column_config.NumberColumn(
            f"{config['set_a']['name']}: {proj}",
            min_value=scale["min"],
            max_value=scale["max"],
            step=0.1 if scale["decimal_places"] > 0 else 1,
            width="small"
        )
    
    for proj in config["set_b"]["projects"]:
        col_name = f"B_{proj}"
        column_config[col_name] = st.column_config.NumberColumn(
            f"{config['set_b']['name']}: {proj}",
            min_value=scale["min"],
            max_value=scale["max"],
            step=0.1 if scale["decimal_places"] > 0 else 1,
            width="small"
        )
    
    # Editable data grid
    edited_df = st.data_editor(
        current_df,
        column_config=column_config,
        hide_index=True,
        key=f"grades_editor_{trimester}",
        use_container_width=True
    )
    
    # Update session state
    st.session_state.grades[trimester] = edited_df
    
    # Validate grades
    issues = validate_grades(edited_df, config)
    if issues:
        st.warning(f"⚠️ {len(issues)} grade(s) outside valid range ({scale['min']}-{scale['max']})")
    
    # Show calculated columns preview
    st.subheader("📊 Calculated Preview")
    
    preview_data = []
    for idx, row in edited_df.iterrows():
        stats = calculate_row_stats(row, config)
        preview_data.append({
            "Student Name": row["Student Name"],
            **stats
        })
    
    preview_df = pd.DataFrame(preview_data)
    st.dataframe(
        preview_df,
        column_config={
            "Student Name": st.column_config.TextColumn("Student Name", width="medium"),
            "Avg A": st.column_config.NumberColumn(f"Avg {config['set_a']['name']}", width="small"),
            "Avg B": st.column_config.NumberColumn(f"Avg {config['set_b']['name']}", width="small"),
            "Final Grade": st.column_config.NumberColumn("Final Grade", width="small"),
            "Qualitative": st.column_config.TextColumn("Qualitative", width="medium"),
        },
        hide_index=True,
        use_container_width=True
    )


def render_summary():
    """Render the summary section."""
    st.header("📈 Summary")
    
    config = st.session_state.config
    students = st.session_state.students
    grades = st.session_state.grades
    
    if not students:
        st.info("Add students to see summary")
        return
    
    if not grades:
        st.info("Enter grades to see summary")
        return
    
    # Build summary table
    summary_data = []
    for student in students:
        row_data = {"Student Name": student}
        
        trimester_grades = []
        for trimester in config["trimesters"]:
            if trimester in grades:
                df = grades[trimester]
                student_row = df[df["Student Name"] == student]
                if not student_row.empty:
                    stats = calculate_row_stats(student_row.iloc[0], config)
                    row_data[f"{trimester} Grade"] = stats["Final Grade"]
                    row_data[f"{trimester} Qual"] = stats["Qualitative"]
                    if stats["Final Grade"] != "":
                        trimester_grades.append(float(stats["Final Grade"]))
        
        # Calculate year average
        if trimester_grades:
            year_avg = sum(trimester_grades) / len(trimester_grades)
            row_data["Year Average"] = round(year_avg, config["scale"]["decimal_places"])
            
            # Year qualitative
            for grade in sorted(config.get("qualitative_grades", []), key=lambda x: x["min"], reverse=True):
                if year_avg >= grade["min"]:
                    row_data["Year Qualitative"] = grade["label"]
                    break
        else:
            row_data["Year Average"] = ""
            row_data["Year Qualitative"] = ""
        
        summary_data.append(row_data)
    
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, hide_index=True, use_container_width=True)


def render_generate_section():
    """Render the generate Excel section."""
    st.header("📥 Generate Excel")
    
    config = st.session_state.config
    students = st.session_state.students
    grades = st.session_state.grades
    
    # Validation summary
    st.subheader("Validation")
    
    config_issues = validate_config(config)
    student_issues = validate_students(students)
    
    all_issues = config_issues + student_issues
    
    errors = [i for i in all_issues if i["type"] == "error"]
    warnings = [i for i in all_issues if i["type"] == "warning"]
    
    if errors:
        for error in errors:
            st.error(f"❌ {error['message']}")
    
    if warnings:
        for warning in warnings:
            st.warning(f"⚠️ {warning['message']}")
    
    if not errors and not warnings:
        st.success("✓ All validations passed")
    
    can_generate = len(errors) == 0 and len(students) > 0
    
    st.divider()
    
    # Generate button
    if st.button("🚀 Generate Excel", disabled=not can_generate, type="primary"):
        with st.spinner("Generating Excel file..."):
            # Prepare grades data
            grades_by_trimester = {}
            for trimester, df in grades.items():
                if df is not None and not df.empty:
                    grades_by_trimester[trimester] = df
            
            # Generate workbook
            wb = generate_workbook(config, students, grades_by_trimester)
            
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
                type="primary"
            )
            
            st.success("Excel file generated successfully!")
    
    if not can_generate:
        if not students:
            st.info("Add students to enable generation")
        elif errors:
            st.info("Fix errors above to enable generation")


def main():
    """Main application entry point."""
    init_session_state()
    
    st.title("📊 Grading Matrix Generator")
    st.markdown("Create flexible grading matrices with customizable scales, weights, and qualitative grades.")
    
    # Sidebar configuration
    render_sidebar()
    
    # Main content
    render_students_section()
    
    st.divider()
    
    render_grade_entry()
    
    st.divider()
    
    render_summary()
    
    st.divider()
    
    render_generate_section()


if __name__ == "__main__":
    main()

