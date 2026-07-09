# Day 75 — Excel Processor GUI Tool

A simple desktop application (tkinter) that lets any user upload an
Excel file, process it in Python, and download the result — no coding
required.

## Features
- Upload Excel file via file dialog
- Two processing modes:
  - **Add Revenue Column** (Units × Price + total row)
  - **Only Summary** (totals overview)
- Save output to a chosen location
- Live status updates + error handling (message boxes)

## Tech Stack
- Python 3.x
- tkinter (GUI)
- pandas + openpyxl (Excel processing)

## Setup
```bash
pip install pandas openpyxl
