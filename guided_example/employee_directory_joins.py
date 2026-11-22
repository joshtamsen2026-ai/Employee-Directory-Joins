## Step 1: Import SQLite and CSV support, plus python-docx for the Word report.
## Install the report dependency with: python -m pip install -r requirements.txt
import csv
import os
import sqlite3
import tempfile
from contextlib import closing
from pathlib import Path

from docx import Document

EMPLOYEE_ID_COLUMN = "Employee ID"
DEPARTMENT_ID_COLUMN = "Department ID"


## These helpers keep the console and Word report in sync for every query.
def show_heading(report, title, description, columns):
    """Print a query heading and add its corresponding table to the report."""
    print(f"\n=== {title} ===")
    report.add_heading(title, level=2)
    report.add_paragraph(description)
    table = report.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"
    for cell, column in zip(table.rows[0].cells, columns):
        cell.text = column
        for run in cell.paragraphs[0].runs:
            run.bold = True
    return table


def show_result(table, result, values):
    """Print a result and add the same values to the report table."""
    print(f"  {result}")
    for cell, value in zip(table.add_row().cells, values):
        cell.text = value


def setup_database(connection: sqlite3.Connection) -> None:
    """Create and populate the example database with foreign keys enabled."""
    cursor = connection.cursor()
    cursor.execute("PRAGMA foreign_keys = ON")

    ## Step 3: Create departments first because employees refer to them.
    ## Each department has a unique ID, a name, and a location.
    cursor.execute("""
        CREATE TABLE departments (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            location TEXT NOT NULL
        )
    """)
    ## Step 4: Create employees and link each one to a department through department_id.
    ## The foreign key keeps those links valid; NOT NULL makes each listed field required.
    cursor.execute("""
        CREATE TABLE employees (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            role TEXT NOT NULL,
            salary INTEGER NOT NULL,
            department_id INTEGER NOT NULL,
            FOREIGN KEY (department_id) REFERENCES departments(id)
        )
    """)
    ## Step 5: Create projects and link each project to its lead using employee_id.
    ## An employee can lead multiple projects because many rows can reference the same employee.
    cursor.execute("""
        CREATE TABLE projects (
            id INTEGER PRIMARY KEY,
            title TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            FOREIGN KEY (employee_id) REFERENCES employees(id)
        )
    """)

    ## Step 6: Insert departments using placeholders and executemany to safely add several rows.
    ## Research has no employees, so the department LEFT JOIN can demonstrate an unmatched row.
    cursor.executemany(
        "INSERT INTO departments (id, name, location) VALUES (?, ?, ?)",
        [
            (1, "Engineering", "Seattle"),
            (2, "Marketing", "Austin"),
            (3, "Human Resources", "Denver"),
            (4, "Research", "Boston"),
        ],
    )
    ## Step 7: Add eight employees across three departments.
    ## Their department IDs match existing departments; two departments have three employees each.
    cursor.executemany(
        "INSERT INTO employees (id, name, role, salary, department_id) "
        "VALUES (?, ?, ?, ?, ?)",
        [
            (1, "Alice Park", "Software Engineer", 110000, 1),
            (2, "Bob Martinez", "Software Engineer", 105000, 1),
            (3, "Carol Zhang", "Engineering Manager", 140000, 1),
            (4, "David Okafor", "Marketing Manager", 95000, 2),
            (5, "Emma Garcia", "Designer", 80000, 2),
            (6, "Frank Wilson", "Analyst", 75000, 2),
            (7, "Grace Kim", "HR Manager", 90000, 3),
            (8, "Henry Patel", "Recruiter", 70000, 3),
        ],
    )
    ## Step 8: Add five projects and identify their leads by employee ID.
    ## Alice leads two projects, while several employees lead none, to show both join cases.
    cursor.executemany(
        "INSERT INTO projects (id, title, employee_id) VALUES (?, ?, ?)",
        [
            (1, "Customer Portal", 1),
            (2, "Mobile App", 1),
            (3, "Platform Upgrade", 3),
            (4, "Brand Campaign", 4),
            (5, "Hiring Initiative", 7),
        ],
    )
    ## Step 9: Commit the inserts so all sample data is saved in this connection's database.
    connection.commit()


def main() -> None:
    """Run the joins and create the Word and CSV reports."""
    ## Create the report up front so the same results can be printed and added to Word.
    report = Document()
    report.add_heading("Employee Directory Join Results", 0)
    report.add_heading("Summary", level=1)
    summary = report.add_paragraph(
        "This example creates an in-memory SQLite database with departments, employees, "
        "and projects. Foreign keys connect employees to departments and projects to "
        "their leads, ensuring each reference points to an existing row. Sample data "
        "includes an empty department, employees without projects, and a lead with "
        "multiple projects to show different join outcomes. INNER JOIN matches employees "
        "to departments and links projects through their leads to departments. LEFT JOIN "
        "keeps all departments or employees even when no related row exists, and IS NULL "
        "identifies employees without a project. The query results below show these "
        "relationships before the temporary database is closed."
    )
    report.add_heading("Findings", level=1)
    counts_finding = report.add_paragraph()
    empty_departments_finding = report.add_paragraph()
    unassigned_employees_finding = report.add_paragraph()
    multiple_projects_finding = report.add_paragraph()
    report.add_heading("Fixes and actions", level=1)
    actions = report.add_paragraph(
        "No database fixes were needed or applied. The empty department and employees "
        "without projects are intentional sample cases, not necessarily errors. "
        "Foreign-key checks were enabled to prevent invalid department and project-lead "
        "references, and missing matches are clearly labeled in the results."
    )
    report.add_heading("Query results", level=1)

    ## Step 2: Create a temporary database that is closed after the queries.
    ## Enable foreign keys during setup to reject invalid references.
    with closing(sqlite3.connect(":memory:")) as connection:
        setup_database(connection)
        cursor = connection.cursor()

        ## Collect the original table data, including IDs and relationships, for the CSV.
        source_rows = []
        cursor.execute("SELECT id, name, location FROM departments ORDER BY id")
        for department_id, name, location in cursor.fetchall():
            source_rows.append({
                "Section": "Departments",
                DEPARTMENT_ID_COLUMN: department_id,
                "Department": name,
                "Location": location,
            })
        cursor.execute(
            "SELECT id, name, role, salary, department_id FROM employees ORDER BY id"
        )
        for employee_id, name, role, salary, department_id in cursor.fetchall():
            source_rows.append({
                "Section": "Employees",
                EMPLOYEE_ID_COLUMN: employee_id,
                "Employee": name,
                "Role": role,
                "Salary": salary,
                DEPARTMENT_ID_COLUMN: department_id,
            })
        cursor.execute("SELECT id, title, employee_id FROM projects ORDER BY id")
        for project_id, title, employee_id in cursor.fetchall():
            source_rows.append({
                "Section": "Projects",
                "Project ID": project_id,
                "Project": title,
                EMPLOYEE_ID_COLUMN: employee_id,
            })

        ## Step 10: Use INNER JOIN to match each employee's department_id to a department's id.
        ## This returns only matching employees and departments, ordered by employee name.
        table = show_heading(
            report,
            "Query 1: Employees and Departments (INNER JOIN)",
            "Matches each employee to their department; only matching rows are included.",
            ("Employee", "Department"),
        )
        cursor.execute("""
            SELECT e.name, d.name
            FROM employees e
            INNER JOIN departments d ON e.department_id = d.id
            ORDER BY e.name
        """)
        ## Fetch the selected columns and display each employee and department.
        for name, department in cursor.fetchall():
            show_result(table, f"{name} - {department}", (name, department))

        ## Step 11: LEFT JOIN employees so every department remains visible.
        ## Departments without matches have a NULL employee name, displayed below as "No employees".
        table = show_heading(
            report,
            "Query 2: All Departments and Employees (LEFT JOIN)",
            "Includes every department, even when it has no employees.",
            ("Department", "Employee"),
        )
        cursor.execute("""
            SELECT d.name, e.name
            FROM departments d
            LEFT JOIN employees e ON e.department_id = d.id
            ORDER BY d.name, e.name
        """)
        department_rows = cursor.fetchall()
        for department, name in department_rows:
            employee = name if name is not None else "No employees"
            show_result(table, f"{department} - {employee}", (department, employee))

        ## Step 12: Start from employees and LEFT JOIN projects on the lead's employee ID.
        ## Employees with no project have a NULL title; multiple projects produce multiple rows.
        table = show_heading(
            report,
            "Query 3: Employees and Projects Led (LEFT JOIN)",
            "Includes every employee; an employee can appear once per project "
            "or with no project.",
            ("Employee", "Project led"),
        )
        cursor.execute("""
            SELECT e.name, p.title
            FROM employees e
            LEFT JOIN projects p ON p.employee_id = e.id
            ORDER BY e.name, p.title
        """)
        for name, project in cursor.fetchall():
            title = project if project is not None else "No projects led"
            show_result(table, f"{name} - {title}", (name, title))

        ## Step 13: Keep the employee LEFT JOIN, then filter for a NULL project ID.
        ## A missing project ID means there was no matching project, so that employee leads none.
        table = show_heading(
            report,
            "Query 4: Employees Without Projects (LEFT JOIN + IS NULL)",
            "Finds employees who have no matching project by filtering for a NULL project ID.",
            ("Employee",),
        )
        cursor.execute("""
            SELECT e.name
            FROM employees e
            LEFT JOIN projects p ON p.employee_id = e.id
            WHERE p.id IS NULL
            ORDER BY e.name
        """)
        employees_without_projects = [name for (name,) in cursor.fetchall()]
        for name in employees_without_projects:
            show_result(table, name, (name,))

        ## Step 14: Join projects to their leads, then join those employees to their departments.
        ## INNER JOIN returns the project title, lead name, and department for each matched project.
        table = show_heading(
            report,
            "Query 5: Projects, Leads, and Departments",
            "Joins all three tables to show each project, its lead, and the lead's department.",
            ("Project", "Lead", "Department"),
        )
        cursor.execute("""
            SELECT p.title, e.name, d.name
            FROM projects p
            INNER JOIN employees e ON p.employee_id = e.id
            INNER JOIN departments d ON e.department_id = d.id
            ORDER BY p.title
        """)
        for project, name, department in cursor.fetchall():
            show_result(
                table, f"{project} - {name} ({department})", (project, name, department)
            )

        ## Summarize actual query results and database counts rather than assuming fixed sample data.
        cursor.execute("""
            SELECT (SELECT COUNT(*) FROM departments),
                   (SELECT COUNT(*) FROM employees),
                   (SELECT COUNT(*) FROM projects)
        """)
        department_count, employee_count, project_count = cursor.fetchone()
        counts_finding.text = (
            f"Database contents: {department_count} departments, {employee_count} employees, "
            f"and {project_count} projects."
        )
        empty_departments = [
            department for department, name in department_rows if name is None
        ]
        empty_departments_finding.text = (
            "Departments without employees: "
            + (", ".join(empty_departments) if empty_departments else "None")
            + "."
        )
        unassigned_employees_finding.text = (
            "Employees who do not lead projects: "
            + (", ".join(employees_without_projects) if employees_without_projects else "None")
            + "."
        )
        cursor.execute("""
            SELECT e.name, COUNT(p.id)
            FROM employees e
            INNER JOIN projects p ON p.employee_id = e.id
            GROUP BY e.id, e.name
            HAVING COUNT(p.id) > 1
            ORDER BY e.name, e.id
        """)
        multiple_leads = [f"{name} ({count})" for name, count in cursor.fetchall()]
        multiple_projects_finding.text = (
            "Employees leading multiple projects: "
            + (", ".join(multiple_leads) if multiple_leads else "None")
            + "."
        )

    ## Step 15: The connection is closed above, discarding the in-memory database.

    ## Step 16: Save the Word report beside the script without overwriting earlier reports.
    with tempfile.NamedTemporaryFile(
        mode="w+b",
        prefix="employee_directory_report_",
        suffix=".docx",
        dir=Path(__file__).resolve().parent,
        delete=False,
    ) as report_file:
        report.save(report_file)
        report_path = Path(report_file.name)

    ## Step 17: Write a spreadsheet-friendly CSV with categorized source data, findings,
    ## and the same query results shown in the Word tables.
    csv_fields = (
        "Section", DEPARTMENT_ID_COLUMN, "Department", "Location", EMPLOYEE_ID_COLUMN,
        "Employee", "Role", "Salary", "Project ID", "Project", "Details",
    )
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8-sig",
        newline="",
        prefix="employee_directory_report_",
        suffix=".csv",
        dir=Path(__file__).resolve().parent,
        delete=False,
    ) as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=csv_fields)
        writer.writeheader()
        writer.writerow({"Section": "Summary", "Details": summary.text})
        for finding in (
            counts_finding, empty_departments_finding,
            unassigned_employees_finding, multiple_projects_finding,
        ):
            writer.writerow({"Section": "Findings", "Details": finding.text})
        writer.writerow({"Section": "Fixes and actions", "Details": actions.text})
        writer.writerows(source_rows)

        query_titles = [
            paragraph.text for paragraph in report.paragraphs
            if paragraph.style.name == "Heading 2"
        ]
        column_names = {"Lead": "Employee", "Project led": "Project"}
        for title, table in zip(query_titles, report.tables):
            headers = [
                column_names.get(cell.text, cell.text) for cell in table.rows[0].cells
            ]
            for row in table.rows[1:]:
                values = {header: cell.text for header, cell in zip(headers, row.cells)}
                writer.writerow({"Section": title, **values})
        csv_path = Path(csv_file.name)

    ## Step 18: Open both reports with their default Windows applications.
    print(f"\nReport saved: {report_path}")
    print(f"CSV saved: {csv_path}")
    os.startfile(report_path)
    os.startfile(csv_path)


if __name__ == "__main__":
    main()
