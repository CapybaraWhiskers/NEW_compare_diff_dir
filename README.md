# Directory Comparison Tool

This repository contains a Python tool to compare two directories using Git's diff capabilities with text conversion for Office documents and PDFs.

## Requirements

- Python 3
- Git
- Python packages: `docx2txt`, `openpyxl`, `xlrd`, `python-docx`, `python-pptx`, `PyPDF2`

Install dependencies with:

```bash
pip install python-docx openpyxl xlrd python-pptx PyPDF2 docx2txt
```

## Usage

```
python compare_dirs.py DIR1 DIR2 [--json]
```

The tool outputs changes grouped by status:

1. Added
2. Removed
3. Renamed
4. Modified
5. RenamedAndModified
6. Unchanged

For changed files, unified diff hunks are displayed. Use `--json` to get JSON output.
