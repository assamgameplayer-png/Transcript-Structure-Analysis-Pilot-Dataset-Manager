# Transcript Structure Analysis: Pilot Dataset Manager

A lightweight Python desktop application (Tkinter + SQLite) for organizing and describing transcript-level structural features of human genes, built for a pilot bioinformatics study.

> **Research question:** *Can transcript-level structural features be used to systematically characterize protein diversity across human genes?*

The program **organizes and describes** values that you copy from NCBI GenBank records. It never invents, estimates, or corrects biological values. Unknown values stay blank.

---

## Table of Contents

1. [Features](#features)
2. [Requirements](#requirements)
3. [Installation and Running](#installation-and-running)
4. [Entering Your Data](#entering-your-data)
5. [Data Dictionary](#data-dictionary)
6. [Application Tour](#application-tour)
7. [How the Comparisons Work](#how-the-comparisons-work)
8. [Data Validation](#data-validation)
9. [Analysis Output](#analysis-output)
10. [Storage and Export](#storage-and-export)
11. [Scientific Rules](#scientific-rules)
12. [Pilot Dataset Design](#pilot-dataset-design)
13. [Code Structure](#code-structure)
14. [Limitations](#limitations)
15. [Roadmap](#roadmap)
16. [License and Citation](#license-and-citation)

---

## Features

- One-row-per-transcript table (`ttk.Treeview`) with vertical and horizontal scrolling
- Add, edit, and delete records through simple dialogs
- Within-gene **Transcript Comparison** table (all transcript pairs of the same gene)
- **Gene Summary** table, calculated from the transcript table
- Descriptive statistics (`Analyze`)
- Input validation with clear warnings (never auto-corrects data)
- **SQLite** storage (`transcript_analysis.db`)
- Import from CSV, export to CSV and Excel (three sheets)
- Single `.py` file, low resource usage, no internet access, no external APIs

## Requirements

| Component | Needed for | Notes |
|---|---|---|
| Python 3.8+ | Everything | |
| `tkinter` | GUI | Included with most Python installs (on some Linux distributions: `sudo apt install python3-tk`) |
| `sqlite3` | Database | Included with Python |
| `pandas`, `openpyxl` | **Export Excel only** | Optional: everything else works without them |

Install the optional Excel dependencies:

```bash
pip install pandas openpyxl
```



## Entering Your Data

There are three ways to get records into the program.

**Option A: Paste into `PILOT_DATA`.** Open `transcript_analysis_app.py` and add one dictionary per transcript to the `PILOT_DATA` list near the top of the file. Use only values copied from your GenBank records, and leave a key out (or `""`) when a value was not recorded. `PILOT_DATA` is used only while the database is empty. If you already ran the program, delete `transcript_analysis.db` first.

**Option B: CSV template.** Click **Export CSV** on the empty table to get a header-only template. Fill it in Excel or any editor, save it as CSV, and use **Load Dataset → CSV**. The CSV must contain the exact column names listed in the [Data Dictionary](#data-dictionary). After loading from CSV, click **Save Dataset** to store the data in SQLite.

**Option C: Manual entry.** Use **Add Record** for one row at a time.

## Data Dictionary

There is **one row per transcript record**. Column names are exact and are the same in the GUI, SQLite, CSV, and Excel.

| Column | Type | Meaning |
|---|---|---|
| `Gene` | text | Official gene symbol (e.g. BRCA1, TP53, EGFR). Required. |
| `GeneID` | integer | NCBI Gene ID (e.g. 672 for BRCA1). Not the same as `Transcript_ID`. |
| `Transcript_ID` | text | RefSeq mRNA accession, normally beginning with `NM_`. |
| `Variant` | text | Transcript variant, only when explicitly given in the record. |
| `Transcript_Length_nt` | integer | Total nucleotide length of the transcript. |
| `Exon_Count` | integer | Number of exon features in that particular GenBank transcript record. |
| `Exon_Coordinates` | text | Exon coordinates such as `1–183; 184–282; 283–336`. |
| `CDS_Start` | integer | CDS start coordinate from the GenBank record. |
| `CDS_End` | integer | CDS end coordinate from the GenBank record. |
| `CDS_Length_nt` | integer | Nucleotide length of the CDS (different from transcript length). |
| `Protein_ID` | text | RefSeq protein accession, normally beginning with `NP_`. Not the same as `Transcript_ID`. |
| `Protein_Length_aa` | integer | Protein length in amino acids. |
| `Protein_Product` | text | Product description from the GenBank record. |
| `Status` | text | `Curated` or `Predicted`. |

Blank values are stored as `NULL` in SQLite and as empty cells in exports.

## Application Tour

**Top buttons**

| Button | Action |
|---|---|
| Add Record | Open a form to add a transcript |
| Edit Record | Edit the selected row |
| Delete Record | Delete the selected row (asks for confirmation) |
| Save Dataset | Validate, then write the table to SQLite |
| Load Dataset | Load from the SQLite database or from a CSV file |
| Export CSV | Export the transcript table to CSV (UTF-8 with BOM, so Excel shows en dashes correctly) |
| Export Excel | Export three sheets: `Transcripts`, `Transcript_Comparison`, `Gene_Summary` |
| Analyze | Validate, then show descriptive statistics in a results window |
| Compare Transcripts | Switch to the Transcript Comparison tab |

**Tabs**

1. **Transcripts**: the main dataset.
2. **Transcript Comparison**: one row for each pair of transcripts belonging to the same gene.
3. **Gene Summary**: one row per gene.

The comparison and summary tables are always recalculated from the transcript table. They are never typed in by hand.

The program asks whether to save unsaved changes when you close the window.

## How the Comparisons Work

### Transcript Comparison columns

| Column | Calculation |
|---|---|
| `Gene` | Gene shared by both transcripts |
| `Transcript_A`, `Transcript_B` | `Transcript_ID` (or `(no Transcript_ID; row N)` if blank) |
| `Transcript_Length_Difference_nt` | Absolute difference |
| `Exon_Count_Different` | `Yes` / `No` |
| `Exon_Structure_Different` | Compares stored exon coordinates segment by segment |
| `CDS_Length_Difference_nt` | Absolute difference |
| `Protein_Length_Difference_aa` | Absolute difference |
| `Protein_ID_Different` | `Yes` / `No` |
| `Protein_Different` | See below |

Hyphens and en/em dashes are treated the same when comparing exon coordinates. The stored text is never modified.

### `Protein_Different`: careful wording

Protein sequences are **not** stored in this dataset, so the program never claims two proteins are identical.

| Situation | Result |
|---|---|
| Protein lengths differ | `Different` |
| Protein lengths are equal | `Same length; sequence not compared` |
| A protein length is missing | `Not determined` |

### Missing values: `Not determined` and `N/A`

If a value needed for a comparison is blank, the result is `Not determined` instead of a guessed `Yes` or `No`. In the Gene Summary, genes with only one transcript show `N/A` because there is nothing to compare.

### Gene Summary columns

| Column | Meaning |
|---|---|
| `Gene`, `GeneID` | Gene symbol and its GeneID |
| `Transcript_Count` | Number of transcript records for the gene |
| `Different_Transcript_Length` | `Yes` if recorded lengths differ, `No` if all recorded and identical |
| `Different_Exon_Structure` | Same logic, using exon coordinates |
| `Different_CDS_Length` | Same logic, using CDS length |
| `Different_Protein_Length` | Same logic, using protein length |
| `Protein_Sequence_Comparison` | Summary of the pairwise `Protein_Different` results (`Mixed - see Transcript Comparison` for genes with 3+ transcripts and differing pair results) |

## Data Validation

Validation runs when you save, analyze, or export, and when you press OK in the Add/Edit dialog.

**Errors (block the action):**

- `Gene` is empty
- A numeric column (`GeneID`, `Transcript_Length_nt`, `Exon_Count`, `CDS_Start`, `CDS_End`, `CDS_Length_nt`, `Protein_Length_aa`) contains something other than a whole number
- `Status` is not `Curated` or `Predicted`

**Warnings (you can choose to continue):**

- `GeneID` is blank
- `Transcript_ID` does not begin with `NM_`
- `Protein_ID` does not begin with `NP_` (extra note for `XP_`, which are predicted proteins)
- `Status` is `Predicted`
- Duplicate `Transcript_ID` values
- One gene with several GeneIDs, or one GeneID used by several gene symbols
- `CDS_Length_nt` differs from `CDS_End - CDS_Start + 1`
- `Exon_Count` differs from the number of `Exon_Coordinates` segments
- CDS length or CDS end is larger than the transcript length

The program only **reports** these problems. It never changes your data.

## Analysis Output

**Analyze** shows:

- Total genes, total transcripts, average transcripts per gene
- Minimum, maximum, and mean for transcript length, CDS length, and protein length (using only records that have a value, with the count shown)
- Number of genes with multiple transcripts
- Number of genes with different transcript lengths, exon structures, CDS lengths, and protein lengths (with the number of genes that could not be determined because of missing values)
- A **pilot design check** comparing your data with the planned 10 genes / 18 transcripts and GeneIDs (reports mismatches only)

## Storage and Export

- **Database:** `transcript_analysis.db` (SQLite), created next to the script.
- **Table:** `transcripts`, with an auto-increment `record_id` plus all 14 columns. Numeric columns are `INTEGER`, all others `TEXT`.
- **Save Dataset** replaces the table contents with the current on-screen data.
- **CSV export:** transcript table only, UTF-8 with BOM.
- **Excel export:** requires `pandas` and `openpyxl`. Numbers are written as numbers and blanks as empty cells.

Suggested `.gitignore` entry, if you do not want to publish the database file itself:

```
transcript_analysis.db
__pycache__/
```

For reproducibility, consider committing your dataset as a CSV export (for example `data/pilot_transcripts.csv`) instead of the binary `.db` file.

## Scientific Rules

The program follows these rules, and they are important when interpreting results:

1. A gene can have multiple transcripts.
2. One transcript corresponds to one transcript record.
3. One transcript has at most one annotated protein product in this dataset.
4. Transcript differences do not automatically mean protein differences.
5. Equal protein length does not prove identical protein sequence.
6. Exon count refers to the particular transcript record, not necessarily the whole gene.
7. The CDS is the coding portion of the transcript.
8. Transcript length and CDS length are different measurements.
9. `NM_` = RefSeq mRNA transcript.
10. `NP_` = RefSeq protein.
11. `GeneID` = NCBI Gene identifier.
12. `XP_` records are predicted proteins and are excluded from the curated main dataset.
13. Missing biological information is not inferred.
14. Original NCBI-derived values are preserved.

## Pilot Dataset Design

Data source: NCBI GenBank / RefSeq records, collected manually.

| Gene | GeneID | Curated transcripts |
|---|---|---|
| BRCA1 | 672 | 2 |
| TP53 | 7157 | 2 |
| EGFR | 1956 | 2 |
| BRCA2 | 675 | 2 |
| CFTR | 1080 | 1 |
| APOE | 348 | 2 |
| F8 | 2157 | 2 |
| LDLR | 3949 | 2 |
| DMD | 1756 | 1 |
| ACE2 | 59272 | 2 |

**Total: 10 genes, 18 curated transcript records.** The predicted DMD `XP_006724531.1` record is intentionally excluded.

## Code Structure

Everything is in `transcript_analysis_app.py`, organized in numbered sections:

| Section | Contents |
|---|---|
| 1. Constants | Column lists, `EXPECTED_GENES`, `PILOT_DATA` |
| 2. Helpers | `make_record`, `to_int`, `parse_exon_structure`, `transcript_label` |
| 3. Database | `init_database`, `save_records_to_db`, `load_records_from_db`, `load_records_from_csv` |
| 4. Validation | `validate_record`, `validate_dataset` |
| 5. Comparison logic | `compare_pair`, `build_comparison_rows`, `summarize_difference`, `build_gene_summary_rows` |
| 6. Analysis | `describe_column`, `build_analysis_text` |
| 7. Dialog | `RecordDialog` (add/edit form) |
| 8. Main window | `TranscriptApp` class and `main()` |

Records are held in memory as a list of dictionaries (one dictionary per transcript, all values stored as strings). All comparison and summary functions take that list, so they are easy to reuse or extend.

## Limitations

- Protein sequences are not stored, so protein identity cannot be assessed, only protein length.
- Numeric columns accept whole numbers only (non-negative integers).
- Data entry is manual. The program does not connect to NCBI.
- Only descriptive statistics are provided (no inferential or advanced statistics yet).
- Gene symbols are matched exactly (after trimming spaces), so `brca1` and `BRCA1` would count as different genes.

## Roadmap

- Store protein sequences so that `Protein_Different` can report actual sequence identity
- Additional summary plots and exports
- Further statistical analysis, once the descriptive pilot results are reviewed

## License and Citation

Its open source unt MIT license
