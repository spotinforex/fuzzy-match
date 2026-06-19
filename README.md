# fuzzy match

A CLI tool that fuzzy-matches full names across two Excel sheets and outputs a color-coded results spreadsheet with match scores.

## Requirements

- Python 3.8+
- [uv](https://docs.astral.sh/uv/)

## Installation

```bash
uv add thefuzz python-Levenshtein openpyxl pandas
```

Or if you're using a `pyproject.toml`-based project:

```bash
uv init
uv add thefuzz python-Levenshtein openpyxl pandas
```

## Usage

```bash
uv run fuzzy_match.py <file1.xlsx> <file2.xlsx> [OPTIONS]
```

### Options

| Option | Default | Description |
|---|---|---|
| `--col1 NAME` | auto-detect | Column name in file1 |
| `--col2 NAME` | auto-detect | Column name in file2 |
| `--sheet1 NAME` | first sheet | Sheet tab name in file1 |
| `--sheet2 NAME` | first sheet | Sheet tab name in file2 |
| `--threshold INT` | `70` | Minimum match score (0–100) |
| `--output FILE` | `match_results.xlsx` | Output filename |
| `--top N` | `1` | Top N matches per name |

## Examples

**Basic match with defaults:**
```bash
uv run fuzzy_match.py list1.xlsx list2.xlsx
```

**Specify column names:**
```bash
uv run fuzzy_match.py list1.xlsx list2.xlsx --col1 "Full Name" --col2 "Beneficiary Name"
```

**Stricter threshold, top 3 matches per name:**
```bash
uv run fuzzy_match.py list1.xlsx list2.xlsx --threshold 85 --top 3
```

**Specific sheet tabs and custom output file:**
```bash
uv run fuzzy_match.py list1.xlsx list2.xlsx --sheet1 "Staff" --sheet2 "Registrants" --output matched.xlsx
```

## Output

The output workbook contains two sheets:

### Match Results
A filterable, color-coded table with these columns:

| Column | Description |
|---|---|
| Name (Sheet 1) | Original name from file1 |
| Matched Name (Sheet 2) | Best match found in file2 |
| Match Score (%) | Fuzzy similarity score (0–100) |
| Match Quality | Label based on score |

**Match Quality labels:**

| Label | Score Range | Color |
|---|---|---|
| Exact | 100% | Green |
| Very High | 90–99% | Light green |
| High | 80–89% | Yellow |
| Medium | 70–79% | Orange |
| Low | Below threshold | Red |
| Below Threshold | Below threshold | Grey |

### Summary
A quick stats sheet showing total names processed, matches above threshold, exact matches, and the threshold used.

## Notes

- **Column auto-detection** looks for columns named `name`, `full name`, `fullname`, `employee`, `staff`, `person`, or `beneficiary` (case-insensitive). If none match, it falls back to the first text column. Use `--col1`/`--col2` to be explicit.
- Matching uses `token_sort_ratio` from [thefuzz](https://github.com/seatgeek/thefuzz), which handles word-order differences (e.g. "John Doe" vs "Doe John").
- Names are normalized to uppercase and stripped of extra whitespace before matching.
- `python-Levenshtein` is optional but significantly speeds up matching on large lists.