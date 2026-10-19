Automated Data Pipeline: Sales Analytics and Excel Stylizer

A professional, enterprise-grade Python data automation pipeline that bridges data processing and visual business reporting. This project showcases clean software engineering practices, utilizing Pandas for rigorous backend data analytics and OpenPyXL for custom spreadsheet formatting.

Key Features

Data Engineering Automation: Dynamically converts unstructured input dictionaries into clean tabular Pandas DataFrames.

Financial Calculations: Automatically computes row-level metrics (Total_Income) and generates rolling summary aggregates (GRAND TOTAL, AVERAGE PRICE) using optimized index methods (.loc).

Automated Formatting Engine: Leverages OpenPyXL styles to programmatically color-code high-priority metrics (Soft green background headers) and custom-weight summary fonts dynamically.

Robust Failure Resilience: Built with unified error handling structures (try-except) and strict file existence verifications (os.path.exists) to emulate production-ready software behavior.

Tech Stack and Dependencies

Core Engine: Python 3.x
Data Preprocessing: Pandas
Reporting Interface: OpenPyXL
File Systems: OS Module

Sample Run Output

When the centralized pipeline executes, it generates clean terminal traces and outputs two discrete physical assets:

Starting data automation pipeline...
Step 1: Data processed and temporary excel file generated.
Visual styles and clean formatting applied successfully to: stylized_sales_report.xlsx
Project Success: The entire data pipeline completed successfully!

Code Organization

automated_pipeline.py: The production-ready unified codebase executing both the analysis and style architecture flawlessly.
main.py / format_report.py: Original micro-scripts demonstrating step-by-step modular code evolutionary logic.