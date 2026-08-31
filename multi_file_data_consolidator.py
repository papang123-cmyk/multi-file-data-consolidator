from pathlib import Path
from datetime import datetime
import shutil
import re
import logging
import csv

import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

BASE_FOLDER = Path(__file__).resolve().parent

INPUT_FOLDER = BASE_FOLDER / "input"
OUTPUT_FOLDER = BASE_FOLDER / "output"
REVIEW_FOLDER = BASE_FOLDER / "review"
BACKUP_FOLDER = BASE_FOLDER / "backup"
LOG_FOLDER = BASE_FOLDER / "logs"

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

RUN_OUTPUT_FOLDER = OUTPUT_FOLDER / f"run_{TIMESTAMP}"
RUN_REVIEW_FOLDER = REVIEW_FOLDER / f"run_{TIMESTAMP}"
RUN_BACKUP_FOLDER = BACKUP_FOLDER / f"run_{TIMESTAMP}"

LOG_FILE = LOG_FOLDER / f"processing_{TIMESTAMP}.log"


# ============================================================
# LOGGING
# ============================================================

def setup_logging():
    """Create a log file and also display important messages."""

    LOG_FOLDER.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(
                LOG_FILE,
                encoding="utf-8"
            ),
            logging.StreamHandler()
        ]
    )


# ============================================================
# FOLDER MANAGEMENT
# ============================================================

def create_folders():
    """Create all required folders."""

    folders = [
        INPUT_FOLDER,
        OUTPUT_FOLDER,
        REVIEW_FOLDER,
        BACKUP_FOLDER,
        LOG_FOLDER,
        RUN_OUTPUT_FOLDER,
        RUN_REVIEW_FOLDER,
        RUN_BACKUP_FOLDER
    ]

    for folder in folders:
        folder.mkdir(parents=True, exist_ok=True)


# ============================================================
# FILE DETECTION
# ============================================================

def find_files():
    """Find supported CSV and Excel files."""

    files = []

    try:
        for path in INPUT_FOLDER.rglob("*"):

            if not path.is_file():
                continue

            if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                files.append(path)

    except Exception:
        logging.exception("Error while scanning the input folder.")

    return sorted(files)


# ============================================================
# FILE TYPE DETECTION
# ============================================================

def detect_file_type(file_path):
    """Identify whether a file is CSV or Excel."""

    extension = file_path.suffix.lower()

    if extension == ".csv":
        return "csv"

    if extension in {".xlsx", ".xls"}:
        return "excel"

    return "unknown"


# ============================================================
# CSV ENCODING DETECTION
# ============================================================

def detect_csv_encoding(file_path):
    """
    Try several common encodings.

    This avoids requiring another external dependency.
    """

    encodings = [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin1"
    ]

    for encoding in encodings:

        try:
            with open(
                file_path,
                "r",
                encoding=encoding
            ) as file:

                file.read(10000)

            return encoding

        except UnicodeDecodeError:
            continue

        except Exception:
            logging.exception(
                f"Could not inspect encoding: {file_path.name}"
            )
            return None

    return None


# ============================================================
# CSV SEPARATOR DETECTION
# ============================================================

def detect_csv_separator(file_path, encoding):
    """Try to detect the CSV separator."""

    try:

        with open(
            file_path,
            "r",
            encoding=encoding,
            newline=""
        ) as file:

            sample = file.read(10000)

        try:
            dialect = csv.Sniffer().sniff(
                sample,
                delimiters=",;\t|"
            )

            return dialect.delimiter

        except csv.Error:

            logging.warning(
                f"Could not detect separator for {file_path.name}. "
                f"Using comma."
            )

            return ","

    except Exception:
        logging.exception(
            f"Could not detect separator: {file_path.name}"
        )

        return ","


# ============================================================
# LOAD CSV
# ============================================================

def load_csv(file_path):
    """Safely load a CSV file."""

    encoding = detect_csv_encoding(file_path)

    if encoding is None:
        logging.error(
            f"Could not determine encoding: {file_path.name}"
        )
        return None

    separator = detect_csv_separator(
        file_path,
        encoding
    )

    try:

        df = pd.read_csv(
            file_path,
            encoding=encoding,
            sep=separator,
            dtype=object,
            keep_default_na=True
        )

        return df

    except pd.errors.EmptyDataError:

        logging.error(
            f"CSV file is empty: {file_path.name}"
        )

    except pd.errors.ParserError:

        logging.error(
            f"CSV structure could not be parsed: "
            f"{file_path.name}"
        )

    except Exception:

        logging.exception(
            f"Unexpected CSV loading error: "
            f"{file_path.name}"
        )

    return None


# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel(file_path):
    """Safely load the first worksheet of an Excel file."""

    try:

        excel_file = pd.ExcelFile(file_path)

        if not excel_file.sheet_names:

            logging.error(
                f"No worksheets found: {file_path.name}"
            )

            return None

        first_sheet = excel_file.sheet_names[0]

        df = pd.read_excel(
            file_path,
            sheet_name=first_sheet,
            dtype=object
        )

        return df

    except Exception:

        logging.exception(
            f"Could not load Excel file: {file_path.name}"
        )

        return None


# ============================================================
# SAFE FILE LOADER
# ============================================================

def load_file(file_path):
    """Load a supported file safely."""

    file_type = detect_file_type(file_path)

    if file_type == "csv":
        return load_csv(file_path)

    if file_type == "excel":
        return load_excel(file_path)

    logging.warning(
        f"Unsupported file type: {file_path.name}"
    )

    return None


# ============================================================
# EMPTY DATAFRAME CHECK
# ============================================================

def is_valid_dataframe(df, file_name):
    """Check whether a loaded DataFrame is usable."""

    if df is None:
        return False

    if df.empty:
        logging.warning(
            f"File contains no data: {file_name}"
        )
        return False

    if len(df.columns) == 0:
        logging.warning(
            f"File contains no columns: {file_name}"
        )
        return False

    return True


# ============================================================
# COLUMN NAME CLEANING
# ============================================================

def clean_column_name(column):
    """Standardize one column name."""

    column = str(column)

    column = column.strip().lower()

    column = re.sub(
        r"\s+",
        " ",
        column
    )

    column = re.sub(
        r"[^a-z0-9]+",
        "_",
        column
    )

    column = re.sub(
        r"_+",
        "_",
        column
    )

    column = column.strip("_")

    if not column:
        column = "unnamed_column"

    return column


def standardize_column_names(df):
    """Standardize every column name."""

    df = df.copy()

    new_columns = []
    used_names = set()

    for index, column in enumerate(df.columns):

        cleaned = clean_column_name(column)

        original_name = cleaned
        counter = 2

        while cleaned in used_names:

            cleaned = (
                f"{original_name}_{counter}"
            )

            counter += 1

        used_names.add(cleaned)

        new_columns.append(cleaned)

    df.columns = new_columns

    return df


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(value):
    """Remove unnecessary whitespace."""

    if pd.isna(value):
        return value

    if not isinstance(value, str):
        return value

    value = value.strip()

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


def clean_text_columns(df):
    """Clean whitespace throughout the dataset."""

    df = df.copy()

    for column in df.columns:

        try:

            df[column] = df[column].map(
                clean_text
            )

        except Exception:

            logging.warning(
                f"Could not clean column: {column}"
            )

    return df


# ============================================================
# NAME DETECTION
# ============================================================

def looks_like_name_column(column):
    """Determine whether a column probably contains names."""

    name_words = {
        "name",
        "full_name",
        "customer_name",
        "employee_name",
        "first_name",
        "last_name",
        "client_name",
        "student_name"
    }

    return column in name_words or column.endswith("_name")


def standardize_names(df):
    """Convert probable name columns to title case."""

    df = df.copy()

    for column in df.columns:

        if not looks_like_name_column(column):
            continue

        try:

            df[column] = df[column].apply(
                lambda value:
                value.title()
                if isinstance(value, str)
                else value
            )

        except Exception:

            logging.warning(
                f"Could not standardize names: {column}"
            )

    return df


# ============================================================
# EMAIL DETECTION
# ============================================================

def looks_like_email_column(column):
    """Determine whether a column probably contains emails."""

    return (
        "email" in column.lower()
        or "e_mail" in column.lower()
    )


def validate_email(value):
    """Return True when an email looks structurally valid."""

    if pd.isna(value):
        return True

    if not isinstance(value, str):
        return False

    pattern = (
        r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    )

    return bool(
        re.match(pattern, value.strip())
    )


def validate_emails(df):
    """Find invalid email addresses."""

    issues = []

    for column in df.columns:

        if not looks_like_email_column(column):
            continue

        for index, value in df[column].items():

            if pd.isna(value):
                continue

            if not validate_email(value):

                issues.append({
                    "row": index + 2,
                    "column": column,
                    "value": value,
                    "issue": "Possible invalid email"
                })

    return issues


# ============================================================
# DATE DETECTION
# ============================================================

def looks_like_date_column(column):
    """Determine whether a column probably contains dates."""

    date_words = {
        "date",
        "dob",
        "birth_date",
        "start_date",
        "end_date",
        "created_at",
        "updated_at"
    }

    return (
        column in date_words
        or column.endswith("_date")
        or column.endswith("_at")
    )


def standardize_dates(df):
    """Convert recognizable date columns to YYYY-MM-DD."""

    df = df.copy()

    for column in df.columns:

        if not looks_like_date_column(column):
            continue

        try:

            parsed = pd.to_datetime(
                df[column],
                errors="coerce"
            )

            valid_dates = parsed.notna()

            df.loc[
                valid_dates,
                column
            ] = parsed.loc[
                valid_dates
            ].dt.strftime("%Y-%m-%d")

        except Exception:

            logging.warning(
                f"Could not standardize dates: {column}"
            )

    return df


# ============================================================
# NUMERIC COLUMN DETECTION
# ============================================================

def looks_like_numeric_column(column):
    """Determine whether a column probably contains numbers."""

    numeric_words = {
        "age",
        "salary",
        "price",
        "amount",
        "revenue",
        "income",
        "quantity",
        "total",
        "score",
        "id",
        "number"
    }

    return (
        column in numeric_words
        or column.endswith("_amount")
        or column.endswith("_price")
        or column.endswith("_salary")
        or column.endswith("_quantity")
    )


def clean_numeric_values(df):
    """Clean commas and whitespace from numeric-looking columns."""

    df = df.copy()

    for column in df.columns:

        if not looks_like_numeric_column(column):
            continue

        try:

            cleaned = (
                df[column]
                .astype("string")
                .str.replace(
                    ",",
                    "",
                    regex=False
                )
                .str.strip()
            )

            converted = pd.to_numeric(
                cleaned,
                errors="coerce"
            )

            # Only replace the column when
            # at least one value successfully converted.
            if converted.notna().any():

                df[column] = converted

        except Exception:

            logging.warning(
                f"Could not normalize numeric column: {column}"
            )

    return df


# ============================================================
# DUPLICATE DETECTION
# ============================================================

def detect_duplicates(df):
    """Find completely duplicated rows."""

    if df.empty:
        return pd.DataFrame()

    try:

        duplicate_mask = df.duplicated(
            keep=False
        )

        duplicates = df.loc[
            duplicate_mask
        ].copy()

        if not duplicates.empty:

            duplicates.insert(
                0,
                "original_row",
                duplicates.index + 2
            )

        return duplicates

    except Exception:

        logging.exception(
            "Error while detecting duplicates."
        )

        return pd.DataFrame()


# ============================================================
# MISSING VALUE DETECTION
# ============================================================

def detect_missing_values(df, file_name):
    """Create a report of missing values."""

    issues = []

    for column in df.columns:

        missing_count = int(
            df[column].isna().sum()
        )

        if missing_count > 0:

            issues.append({
                "file": file_name,
                "column": column,
                "missing_values": missing_count
            })

    return issues


# ============================================================
# UNUSUAL VALUE DETECTION
# ============================================================

def detect_unusual_values(df, file_name):
    """Detect negative values in common numeric fields."""

    issues = []

    for column in df.columns:

        if not looks_like_numeric_column(column):
            continue

        try:

            numeric_values = pd.to_numeric(
                df[column],
                errors="coerce"
            )

            negative_mask = (
                numeric_values < 0
            )

            for index in df.index[
                negative_mask.fillna(False)
            ]:

                issues.append({
                    "file": file_name,
                    "row": index + 2,
                    "column": column,
                    "value": df.loc[index, column],
                    "issue": "Negative value"
                })

        except Exception:

            logging.warning(
                f"Could not check unusual values: "
                f"{column}"
            )

    return issues


# ============================================================
# DATA TYPE SUMMARY
# ============================================================

def get_data_type_summary(df, file_name):
    """Report detected pandas data types."""

    results = []

    for column in df.columns:

        results.append({
            "file": file_name,
            "column": column,
            "data_type": str(df[column].dtype)
        })

    return results


# ============================================================
# FILE ANALYSIS
# ============================================================

def analyze_file(df, file_path):
    """Analyze one file and return reports."""

    file_name = file_path.name

    analysis = {
        "file": file_name,
        "rows": len(df),
        "columns": len(df.columns),
        "column_names": ", ".join(
            str(column)
            for column in df.columns
        )
    }

    missing = detect_missing_values(
        df,
        file_name
    )

    unusual = detect_unusual_values(
        df,
        file_name
    )

    emails = validate_emails(df)

    for issue in emails:
        issue["file"] = file_name

    duplicates = detect_duplicates(df)

    duplicate_count = len(duplicates)

    analysis["duplicate_rows"] = duplicate_count
    analysis["missing_values"] = sum(
        item["missing_values"]
        for item in missing
    )
    analysis["unusual_values"] = len(
        unusual
    )
    analysis["invalid_emails"] = len(
        emails
    )

    return (
        analysis,
        missing,
        unusual,
        emails,
        duplicates
    )


# ============================================================
# COLUMN COMPARISON
# ============================================================

def compare_columns(file_data):
    """Compare the standardized columns of every file."""

    all_columns = set()

    file_columns = {}

    for file_name, df in file_data.items():

        columns = set(df.columns)

        file_columns[file_name] = columns

        all_columns.update(columns)

    comparison = []

    for file_name, columns in file_columns.items():

        missing = sorted(
            all_columns - columns
        )

        extra = sorted(
            columns - (
                all_columns - set(missing)
            )
        )

        comparison.append({
            "file": file_name,
            "missing_columns": ", ".join(missing),
            "columns_present": ", ".join(
                sorted(columns)
            ),
            "column_count": len(columns),
            "unique_to_file": ", ".join(extra)
        })

    return comparison


# ============================================================
# SAFE BACKUP
# ============================================================

def backup_original(file_path):
    """Copy the original file without modifying it."""

    try:

        destination = (
            RUN_BACKUP_FOLDER
            / file_path.name
        )

        # Avoid accidental overwriting.
        counter = 2

        while destination.exists():

            destination = (
                RUN_BACKUP_FOLDER
                / f"{file_path.stem}_{counter}"
                f"{file_path.suffix}"
            )

            counter += 1

        shutil.copy2(
            file_path,
            destination
        )

        return destination

    except Exception:

        logging.exception(
            f"Could not back up: {file_path.name}"
        )

        return None


# ============================================================
# COMBINE DATASETS
# ============================================================

def combine_datasets(file_data):
    """Combine all processed DataFrames."""

    if not file_data:
        return None

    frames = []

    for file_name, df in file_data.items():

        temp = df.copy()

        temp.insert(
            0,
            "source_file",
            file_name
        )

        frames.append(temp)

    try:

        combined = pd.concat(
            frames,
            ignore_index=True,
            sort=False
        )

        return combined

    except Exception:

        logging.exception(
            "Could not combine datasets."
        )

        return None


# ============================================================
# SORT DATASET
# ============================================================

def sort_dataset(df):
    """Sort by age when an age column exists."""

    df = df.copy()

    if "age" not in df.columns:
        return df

    try:

        numeric_age = pd.to_numeric(
            df["age"],
            errors="coerce"
        )

        df["_sort_age"] = numeric_age

        df = df.sort_values(
            by="_sort_age",
            ascending=True,
            na_position="last"
        )

        df = df.drop(
            columns=["_sort_age"]
        )

        df = df.reset_index(
            drop=True
        )

    except Exception:

        logging.warning(
            "Could not sort dataset by age."
        )

    return df


# ============================================================
# SUMMARY STATISTICS
# ============================================================

def generate_summary_statistics(df):
    """Generate basic numeric statistics."""

    try:

        numeric_df = df.select_dtypes(
            include="number"
        )

        if numeric_df.empty:
            return pd.DataFrame()

        return numeric_df.describe().T

    except Exception:

        logging.exception(
            "Could not generate statistics."
        )

        return pd.DataFrame()


# ============================================================
# SAVE CSV SAFELY
# ============================================================

def save_csv(df, destination):
    """Save a DataFrame to CSV."""

    try:

        df.to_csv(
            destination,
            index=False,
            encoding="utf-8-sig"
        )

        return True

    except Exception:

        logging.exception(
            f"Could not save: {destination.name}"
        )

        return False


# ============================================================
# SAVE REPORT
# ============================================================

def save_report(
    analyses,
    missing_reports,
    unusual_reports,
    email_reports,
    column_comparison,
    processed_files,
    failed_files
):
    """Save the processing report."""

    report_lines = []

    report_lines.append(
        "=" * 70
    )

    report_lines.append(
        "MULTI-FILE DATA CONSOLIDATOR REPORT"
    )

    report_lines.append(
        "=" * 70
    )

    report_lines.append(
        f"Run time: {datetime.now()}"
    )

    report_lines.append(
        f"Files processed: {len(processed_files)}"
    )

    report_lines.append(
        f"Files failed: {len(failed_files)}"
    )

    report_lines.append("")

    report_lines.append(
        "PROCESSED FILES"
    )

    report_lines.append(
        "-" * 70
    )

    for file_name in processed_files:

        report_lines.append(
            f"✓ {file_name}"
        )

    if failed_files:

        report_lines.append("")

        report_lines.append(
            "FAILED FILES"
        )

        report_lines.append(
            "-" * 70
        )

        for file_name in failed_files:

            report_lines.append(
                f"✗ {file_name}"
            )

    report_lines.append("")

    report_lines.append(
        "FILE ANALYSIS"
    )

    report_lines.append(
        "-" * 70
    )

    for analysis in analyses:

        report_lines.append(
            f"\nFile: {analysis['file']}"
        )

        report_lines.append(
            f"Rows: {analysis['rows']}"
        )

        report_lines.append(
            f"Columns: {analysis['columns']}"
        )

        report_lines.append(
            f"Duplicates: "
            f"{analysis['duplicate_rows']}"
        )

        report_lines.append(
            f"Missing values: "
            f"{analysis['missing_values']}"
        )

        report_lines.append(
            f"Unusual values: "
            f"{analysis['unusual_values']}"
        )

        report_lines.append(
            f"Invalid emails: "
            f"{analysis['invalid_emails']}"
        )

    report_lines.append("")

    report_lines.append(
        "COLUMN COMPARISON"
    )

    report_lines.append(
        "-" * 70
    )

    for item in column_comparison:

        report_lines.append(
            f"\n{item['file']}"
        )

        report_lines.append(
            f"Columns: {item['columns_present']}"
        )

        if item["missing_columns"]:

            report_lines.append(
                f"Missing: "
                f"{item['missing_columns']}"
            )

    report_lines.append("")

    report_lines.append(
        "REVIEW ITEMS"
    )

    report_lines.append(
        "-" * 70
    )

    report_lines.append(
        f"Missing-value records: "
        f"{len(missing_reports)}"
    )

    report_lines.append(
        f"Unusual-value records: "
        f"{len(unusual_reports)}"
    )

    report_lines.append(
        f"Invalid-email records: "
        f"{len(email_reports)}"
    )

    report_lines.append("")

    report_lines.append(
        "Original input files were not modified."
    )

    report_lines.append(
        "Backups were created for successfully "
        "processed files."
    )

    report_lines.append("")

    report_lines.append(
        "PROCESSING STATUS"
    )

    report_lines.append(
        "-" * 70
    )

    if failed_files:

        report_lines.append(
            "COMPLETED WITH WARNINGS"
        )

    else:

        report_lines.append(
            "COMPLETED SUCCESSFULLY"
        )

    report_lines.append(
        "=" * 70
    )

    report_path = (
        RUN_OUTPUT_FOLDER
        / "processing_report.txt"
    )

    try:

        report_path.write_text(
            "\n".join(report_lines),
            encoding="utf-8"
        )

        return report_path

    except Exception:

        logging.exception(
            "Could not save processing report."
        )

        return None


# ============================================================
# FINAL VERIFICATION
# ============================================================

def verify_output(
    combined_df,
    output_file,
    processed_count
):
    """Verify that the final output exists and is readable."""

    checks = []

    if output_file.exists():

        checks.append(
            "Output file exists: PASS"
        )

    else:

        checks.append(
            "Output file exists: FAIL"
        )

    if combined_df is not None:

        checks.append(
            f"Rows generated: "
            f"{len(combined_df)}"
        )

        checks.append(
            f"Columns generated: "
            f"{len(combined_df.columns)}"
        )

    else:

        checks.append(
            "Combined dataset exists: FAIL"
        )

    checks.append(
        f"Files processed: {processed_count}"
    )

    verification_path = (
        RUN_OUTPUT_FOLDER
        / "verification.txt"
    )

    try:

        verification_path.write_text(
            "\n".join(checks),
            encoding="utf-8"
        )

        return verification_path

    except Exception:

        logging.exception(
            "Could not save verification report."
        )

        return None


# ============================================================
# MAIN PROCESS
# ============================================================

def process_all_files():

    setup_logging()

    create_folders()

    logging.info(
        "Starting Multi-File Data Consolidator."
    )

    files = find_files()

    if not files:

        logging.warning(
            "No supported CSV or Excel files found."
        )

        print(
            f"\nPut CSV/Excel files inside:\n"
            f"{INPUT_FOLDER}"
        )

        return

    logging.info(
        f"Found {len(files)} supported file(s)."
    )

    file_data = {}

    analyses = []

    missing_reports = []
    unusual_reports = []
    email_reports = []

    processed_files = []
    failed_files = []

    # --------------------------------------------------------
    # PROCESS EACH FILE
    # --------------------------------------------------------

    for file_path in files:

        print("\n" + "=" * 70)

        print(
            f"PROCESSING: {file_path.name}"
        )

        print("=" * 70)

        logging.info(
            f"Processing {file_path.name}"
        )

        # Back up first.
        backup = backup_original(
            file_path
        )

        if backup is None:

            logging.error(
                f"Backup failed: "
                f"{file_path.name}"
            )

            failed_files.append(
                file_path.name
            )

            continue

        # Load.
        df = load_file(file_path)

        if not is_valid_dataframe(
            df,
            file_path.name
        ):

            failed_files.append(
                file_path.name
            )

            continue

        try:

            # Standardization pipeline.
            df = standardize_column_names(df)

            df = clean_text_columns(df)

            df = standardize_names(df)

            df = standardize_dates(df)

            df = clean_numeric_values(df)

            # Analyze after cleaning.
            (
                analysis,
                missing,
                unusual,
                emails,
                duplicates
            ) = analyze_file(
                df,
                file_path
            )

            analyses.append(
                analysis
            )

            missing_reports.extend(
                missing
            )

            unusual_reports.extend(
                unusual
            )

            email_reports.extend(
                emails
            )

            # Save duplicates for manual review.
            if not duplicates.empty:

                duplicate_path = (
                    RUN_REVIEW_FOLDER
                    / f"{file_path.stem}"
                    "_duplicates.csv"
                )

                save_csv(
                    duplicates,
                    duplicate_path
                )

            file_data[
                file_path.name
            ] = df

            processed_files.append(
                file_path.name
            )

            print(
                f"✓ Processed: "
                f"{file_path.name}"
            )

        except Exception:

            logging.exception(
                f"Unexpected processing error: "
                f"{file_path.name}"
            )

            failed_files.append(
                file_path.name
            )

    # --------------------------------------------------------
    # STOP IF NOTHING WAS PROCESSED
    # --------------------------------------------------------

    if not file_data:

        logging.error(
            "No files were successfully processed."
        )

        print(
            "\nNo files could be processed."
        )

        return

    # --------------------------------------------------------
    # COLUMN COMPARISON
    # --------------------------------------------------------

    logging.info(
        "Comparing column structures."
    )

    column_comparison = compare_columns(
        file_data
    )

    # --------------------------------------------------------
    # SAVE REVIEW REPORTS
    # --------------------------------------------------------

    if missing_reports:

        save_csv(
            pd.DataFrame(missing_reports),
            RUN_REVIEW_FOLDER
            / "missing_values.csv"
        )

    if unusual_reports:

        save_csv(
            pd.DataFrame(unusual_reports),
            RUN_REVIEW_FOLDER
            / "unusual_values.csv"
        )

    if email_reports:

        save_csv(
            pd.DataFrame(email_reports),
            RUN_REVIEW_FOLDER
            / "invalid_emails.csv"
        )

    save_csv(
        pd.DataFrame(column_comparison),
        RUN_REVIEW_FOLDER
        / "column_comparison.csv"
    )

    # --------------------------------------------------------
    # COMBINE
    # --------------------------------------------------------

    logging.info(
        "Combining datasets."
    )

    combined_df = combine_datasets(
        file_data
    )

    if combined_df is None:

        logging.error(
            "Could not create combined dataset."
        )

        return

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    combined_df = sort_dataset(
        combined_df
    )

    # --------------------------------------------------------
    # SAVE FINAL DATASET
    # --------------------------------------------------------

    combined_path = (
        RUN_OUTPUT_FOLDER
        / "combined_data.csv"
    )

    if not save_csv(
        combined_df,
        combined_path
    ):

        logging.error(
            "Final dataset could not be saved."
        )

        return

    # --------------------------------------------------------
    # SUMMARY STATISTICS
    # --------------------------------------------------------

    statistics = generate_summary_statistics(
        combined_df
    )

    if not statistics.empty:

        statistics_path = (
            RUN_OUTPUT_FOLDER
            / "summary_statistics.csv"
        )

        try:

            statistics.to_csv(
                statistics_path
            )

        except Exception:

            logging.exception(
                "Could not save statistics."
            )

    # --------------------------------------------------------
    # REPORT
    # --------------------------------------------------------

    save_report(
        analyses,
        missing_reports,
        unusual_reports,
        email_reports,
        column_comparison,
        processed_files,
        failed_files
    )

    # --------------------------------------------------------
    # FINAL VERIFICATION
    # --------------------------------------------------------

    verify_output(
        combined_df,
        combined_path,
        len(processed_files)
    )

    # --------------------------------------------------------
    # FINAL MESSAGE
    # --------------------------------------------------------

    print("\n" + "=" * 70)

    print(
        "PROCESSING COMPLETE"
    )

    print("=" * 70)

    print(
        f"Files found: {len(files)}"
    )

    print(
        f"Files processed: "
        f"{len(processed_files)}"
    )

    print(
        f"Files failed: "
        f"{len(failed_files)}"
    )

    print(
        f"Final rows: "
        f"{len(combined_df)}"
    )

    print(
        f"Final columns: "
        f"{len(combined_df.columns)}"
    )

    print(
        f"\nOutput folder:\n"
        f"{RUN_OUTPUT_FOLDER}"
    )

    print(
        f"\nReview folder:\n"
        f"{RUN_REVIEW_FOLDER}"
    )

    print(
        f"\nBackup folder:\n"
        f"{RUN_BACKUP_FOLDER}"
    )

    print(
        f"\nLog file:\n"
        f"{LOG_FILE}"
    )

    if failed_files:

        print(
            "\n⚠️ Completed with warnings."
        )

    else:

        print(
            "\n✅ All supported files processed successfully."
        )

    print(
        "\nOriginal files were NOT modified."
    )


# ============================================================
# PROGRAM ENTRY POINT
# ============================================================

if __name__ == "__main__":
    try:

        process_all_files()

    except KeyboardInterrupt:

        print(
            "\n\nProgram stopped by user."
        )

        logging.warning(
            "Program interrupted by user."
        )

    except Exception:

        logging.exception(
            "Fatal unexpected error."
        )

        print(
            "\nA serious unexpected error occurred."
        )

        print(
            "Check the log file for details."
        )