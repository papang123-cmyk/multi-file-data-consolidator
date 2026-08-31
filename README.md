# Multi-File Data Consolidator

A Python-based data processing tool that consolidates multiple CSV and Excel files while performing data-quality checks and generating organized reports.

## Features

* Processes multiple supported CSV and Excel files
* Compares column structures across datasets
* Combines processed datasets into a consolidated output
* Detects duplicate records
* Identifies missing values
* Flags unusual data values for review
* Creates backups of original files
* Generates timestamped processing logs
* Keeps original input files unchanged
* Organizes results into separate output and review folders

## Technologies Used

* Python 3
* Pandas
* Python Standard Library

## How It Works

The program scans the input folder for supported data files and processes them automatically.

For each run, it:

1. Finds supported input files.
2. Processes and validates the data.
3. Checks the structure of the datasets.
4. Identifies data-quality issues.
5. Combines the processed datasets.
6. Saves the results in a timestamped output folder.
7. Creates a review folder for data issues.
8. Creates backups of the original files.
9. Records the processing activity in a log file.

Original input files are not modified.

## Example

The test dataset contained:

* **18 rows**
* **6 columns**
* **2 duplicate records**
* **5 missing values**

Missing values were identified in:

| Column | Missing Values |
| ------ | -------------: |
| email  |              1 |
| age    |              2 |
| salary |              2 |

## Output

Each processing run creates timestamped folders containing the generated results, review information, backups, and logs.

Example:

```text
output/
└── run_YYYYMMDD_HHMMSS/

review/
└── run_YYYYMMDD_HHMMSS/

backup/
└── run_YYYYMMDD_HHMMSS/

logs/
└── processing_YYYYMMDD_HHMMSS.log
```

## Requirements

Install the required Python package with:

```bash
python -m pip install -r requirements.txt
```

## Running the Program

Place the files you want to process inside the `input` folder, then run:

```bash
python multi_file_data_consolidator.py
```

The program will process the files and generate the corresponding output, review, backup, and log folders.

## Project Structure

```text
multi-file-data-consolidator/
│
├── multi_file_data_consolidator.py
├── requirements.txt
├── README.md
│
├── input/
├── output/
├── review/
├── backup/
└── logs/
```

## Purpose

This project was built as a practical Python automation project focused on data processing, data-quality analysis, file handling, and automated reporting.
