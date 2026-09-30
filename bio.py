"""
Transcript Structure Analysis - pilot dataset manager
=====================================================

Research question:
    "Can transcript-level structural features be used to systematically
     characterize protein diversity across human genes?"

This program only ORGANIZES and DESCRIBES the values you enter.
It never invents, estimates or corrects biological values.
Blank stays blank.

Requirements : Python 3, tkinter (built in), sqlite3 (built in)
Optional     : pandas + openpyxl  (only needed for "Export Excel")
               pip install pandas openpyxl

HOW TO ADD YOUR 18 PILOT RECORDS
--------------------------------
Option A: paste them into PILOT_DATA below (only used while the database is
          empty; if you already ran the program, delete transcript_analysis.db).
Option B: click "Export CSV" on the empty table to get a header-only template,
          fill it in Excel, save as CSV, then use "Load Dataset" -> CSV.
Option C: use "Add Record" one row at a time.

Scientific rules built into the comparisons
-------------------------------------------
* Transcript length (nt) and CDS length (nt) are different measurements.
* Exon count refers to the particular transcript record.
* Different transcripts do not automatically mean different proteins.
* Equal protein length does NOT prove identical protein sequence, so the
  program says "Same length; sequence not compared" (never "identical").
"""

import csv
import itertools
import os
import re
import sqlite3
import statistics
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# pandas/openpyxl are optional: only Excel export needs them.
try:
    import pandas as pd
    import openpyxl  # noqa: F401  (imported only to check it is installed)
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# =====================================================================
# 1. CONSTANTS
# =====================================================================

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(APP_DIR, "transcript_analysis.db")

# Exact column names of the main transcript table (Table 1)
COLS = [
    "Gene", "GeneID", "Transcript_ID", "Variant", "Transcript_Length_nt",
    "Exon_Count", "Exon_Coordinates", "CDS_Start", "CDS_End",
    "CDS_Length_nt", "Protein_ID", "Protein_Length_aa", "Protein_Product",
    "Status",
]

# Columns that must contain whole numbers when they are filled in
NUMERIC_COLS = [
    "GeneID", "Transcript_Length_nt", "Exon_Count", "CDS_Start", "CDS_End",
    "CDS_Length_nt", "Protein_Length_aa",
]

STATUS_VALUES = ("Curated", "Predicted")

# Columns of the Transcript Comparison table
COMPARISON_COLS = [
    "Gene", "Transcript_A", "Transcript_B",
    "Transcript_Length_Difference_nt", "Exon_Count_Different",
    "Exon_Structure_Different", "CDS_Length_Difference_nt",
    "Protein_Length_Difference_aa", "Protein_ID_Different",
    "Protein_Different",
]

# Columns of the Gene Summary table
SUMMARY_COLS = [
    "Gene", "GeneID", "Transcript_Count", "Different_Transcript_Length",
    "Different_Exon_Structure", "Different_CDS_Length",
    "Different_Protein_Length", "Protein_Sequence_Comparison",
]

# The planned pilot design (from your description). Used ONLY as a check in
# the Analyze window; it never adds or changes records.
# gene symbol: (GeneID, expected number of curated transcripts)
EXPECTED_GENES = {
    "BRCA1": (672, 2),
    "TP53": (7157, 2),
    "EGFR": (1956, 2),
    "BRCA2": (675, 2),
    "CFTR": (1080, 1),
    "APOE": (348, 2),
    "F8": (2157, 2),
    "LDLR": (3949, 2),
    "DMD": (1756, 1),
    "ACE2": (59272, 2),
}

# ---------------------------------------------------------------------
# PILOT DATA: paste your 18 records here, one dict per transcript.
# Use ONLY values copied from your NCBI GenBank records. Leave a key out
# (or use "") when a value was not recorded.
# Format illustration (the angle-bracket parts are placeholders, not data):
#
# {"Gene": "<symbol>", "GeneID": "<number>", "Transcript_ID": "NM_<...>",
#  "Variant": "<...>", "Transcript_Length_nt": "<...>", "Exon_Count": "<...>",
#  "Exon_Coordinates": "<start>-<end>; <start>-<end>", "CDS_Start": "<...>",
#  "CDS_End": "<...>", "CDS_Length_nt": "<...>", "Protein_ID": "NP_<...>",
#  "Protein_Length_aa": "<...>", "Protein_Product": "<...>",
#  "Status": "Curated"},
# ---------------------------------------------------------------------
PILOT_DATA = [
]


# =====================================================================
# 2. SMALL HELPER FUNCTIONS
# =====================================================================

def make_record(**values):
    """Return a record dict containing every column (blank by default)."""
    record = {col: "" for col in COLS}
    for key, value in values.items():
        if key not in record:
            raise ValueError("Unknown column name in record: " + key)
        record[key] = "" if value is None else str(value).strip()
    return record


def to_int(text):
    """Convert text to int, or return None if blank / not a whole number."""
    text = str(text).strip()
    return int(text) if re.fullmatch(r"\d+", text) else None


def parse_exon_structure(text):
    """
    Turn 'a-b; c-d; ...' into a tuple of cleaned segments so two exon
    structures can be compared. Hyphens and en/em dashes are treated the
    same for comparison only; the stored text is never changed.
    Returns None when nothing was recorded.
    """
    text = str(text).strip().replace("\u2013", "-").replace("\u2014", "-")
    parts = [re.sub(r"\s+", "", p) for p in text.split(";") if p.strip()]
    return tuple(parts) if parts else None


def transcript_label(record, index):
    """Label used in the comparison table (Transcript_ID, or row number)."""
    return record["Transcript_ID"].strip() or "(no Transcript_ID; row {})".format(index + 1)


# =====================================================================
# 3. DATABASE (SQLite)
# =====================================================================

def _create_table_sql():
    col_defs = ", ".join(
        "{} {}".format(c, "INTEGER" if c in NUMERIC_COLS else "TEXT") for c in COLS
    )
    return ("CREATE TABLE IF NOT EXISTS transcripts ("
            "record_id INTEGER PRIMARY KEY AUTOINCREMENT, " + col_defs + ")")


INSERT_SQL = "INSERT INTO transcripts ({}) VALUES ({})".format(
    ", ".join(COLS), ", ".join("?" for _ in COLS))


def init_database():
    """Create transcript_analysis.db and the transcripts table if needed."""
    conn = sqlite3.connect(DB_FILE)
    try:
        with conn:
            conn.execute(_create_table_sql())
    finally:
        conn.close()


def save_records_to_db(records):
    """Replace the contents of the transcripts table with `records`."""
    def to_db_value(col, text):
        text = str(text).strip()
        if text == "":
            return None                      # blank stays blank (NULL)
        return int(text) if col in NUMERIC_COLS else text

    conn = sqlite3.connect(DB_FILE)
    try:
        with conn:  # commits on success, rolls back on error
            conn.execute("DELETE FROM transcripts")
            conn.executemany(
                INSERT_SQL,
                [[to_db_value(c, r[c]) for c in COLS] for r in records])
    finally:
        conn.close()


def load_records_from_db():
    """Read all records from the database, in the order they were saved."""
    conn = sqlite3.connect(DB_FILE)
    try:
        rows = conn.execute(
            "SELECT {} FROM transcripts ORDER BY record_id".format(", ".join(COLS))
        ).fetchall()
    finally:
        conn.close()
    return [make_record(**dict(zip(COLS, row))) for row in rows]


def load_records_from_csv(path):
    """Read records from a CSV that has the exact column names."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        missing = [c for c in COLS if c not in (reader.fieldnames or [])]
        if missing:
            raise ValueError("The CSV is missing these required columns:\n" + ", ".join(missing))
        records = []
        for row in reader:
            values = {c: (row.get(c) or "") for c in COLS}
            if any(v.strip() for v in values.values()):   # skip fully blank lines
                records.append(make_record(**values))
    return records


# =====================================================================
# 4. VALIDATION  (never changes data, only reports problems)
# =====================================================================

def validate_record(rec):
    """
    Check one record. Returns (errors, warnings).
    errors   = must be fixed before saving/analyzing
    warnings = unusual, but allowed if you confirm
    """
    errors, warnings = [], []

    if not rec["Gene"].strip():
        errors.append("Gene cannot be empty.")

    for col in NUMERIC_COLS:
        value = rec[col].strip()
        if value and not re.fullmatch(r"\d+", value):
            errors.append("{} must be a whole number (found '{}').".format(col, value))

    if rec["Status"].strip() and rec["Status"].strip() not in STATUS_VALUES:
        errors.append("Status must be 'Curated' or 'Predicted' (found '{}').".format(rec["Status"]))

    if not rec["GeneID"].strip():
        warnings.append("GeneID is blank.")

    tid = rec["Transcript_ID"].strip()
    if tid and not tid.startswith("NM_"):
        warnings.append("Transcript_ID '{}' does not begin with NM_.".format(tid))

    pid = rec["Protein_ID"].strip()
    if pid and not pid.startswith("NP_"):
        note = " (XP_ = predicted protein; excluded from the curated main dataset)" if pid.startswith("XP_") else ""
        warnings.append("Protein_ID '{}' does not begin with NP_{}.".format(pid, note))

    if rec["Status"].strip() == "Predicted":
        warnings.append("Status is 'Predicted'; the main dataset is meant to contain curated NM_/NP_ records.")

    # Consistency checks: they only flag, they never correct anything.
    tx_len = to_int(rec["Transcript_Length_nt"])
    cds_start, cds_end = to_int(rec["CDS_Start"]), to_int(rec["CDS_End"])
    cds_len = to_int(rec["CDS_Length_nt"])
    exon_count = to_int(rec["Exon_Count"])
    exons = parse_exon_structure(rec["Exon_Coordinates"])

    if None not in (cds_start, cds_end, cds_len) and cds_end - cds_start + 1 != cds_len:
        warnings.append("CDS_Length_nt ({}) differs from CDS_End - CDS_Start + 1 ({}). Please re-check the GenBank record.".format(
            cds_len, cds_end - cds_start + 1))
    if exons is not None and exon_count is not None and len(exons) != exon_count:
        warnings.append("Exon_Count ({}) differs from the number of Exon_Coordinates segments ({}).".format(
            exon_count, len(exons)))
    if tx_len is not None:
        if cds_len is not None and cds_len > tx_len:
            warnings.append("CDS_Length_nt is larger than Transcript_Length_nt.")
        if cds_end is not None and cds_end > tx_len:
            warnings.append("CDS_End is beyond Transcript_Length_nt.")

    return errors, warnings


def validate_dataset(records):
    """Validate every record plus cross-record checks. Returns (errors, warnings)."""
    errors, warnings = [], []

    for i, rec in enumerate(records):
        who = "Row {} ({} {})".format(i + 1, rec["Gene"] or "no gene", rec["Transcript_ID"]).strip()
        rec_errors, rec_warnings = validate_record(rec)
        errors += ["{}: {}".format(who, m) for m in rec_errors]
        warnings += ["{}: {}".format(who, m) for m in rec_warnings]

    # Duplicate transcript accessions
    seen = {}
    for i, rec in enumerate(records):
        tid = rec["Transcript_ID"].strip()
        if tid:
            seen.setdefault(tid, []).append(i + 1)
    for tid, rows in seen.items():
        if len(rows) > 1:
            warnings.append("Transcript_ID {} appears in rows {}.".format(tid, ", ".join(map(str, rows))))

    # GeneID / Gene symbol consistency
    gene_to_ids, id_to_genes = {}, {}
    for rec in records:
        gene, gid = rec["Gene"].strip(), rec["GeneID"].strip()
        if gene and gid:
            gene_to_ids.setdefault(gene, set()).add(gid)
            id_to_genes.setdefault(gid, set()).add(gene)
    for gene, ids in gene_to_ids.items():
        if len(ids) > 1:
            warnings.append("Gene {} has more than one GeneID: {}.".format(gene, ", ".join(sorted(ids))))
    for gid, genes in id_to_genes.items():
        if len(genes) > 1:
            warnings.append("GeneID {} is used by more than one gene symbol: {}.".format(gid, ", ".join(sorted(genes))))

    return errors, warnings


# =====================================================================
# 5. COMPARISON LOGIC
# =====================================================================

def group_by_gene(records):
    """Return {gene: [(row_index, record), ...]} in order of first appearance."""
    groups = {}
    for idx, rec in enumerate(records):
        gene = rec["Gene"].strip()
        if gene:
            groups.setdefault(gene, []).append((idx, rec))
    return groups


def abs_diff(a_text, b_text):
    """Absolute numeric difference, or '' if either value is missing."""
    a, b = to_int(a_text), to_int(b_text)
    return "" if a is None or b is None else abs(a - b)


def compare_flag(a, b):
    """'Yes'/'No' if two parsed values differ; 'Not determined' if one is missing."""
    if a is None or b is None or a == "" or b == "":
        return "Not determined"
    return "Yes" if a != b else "No"


def protein_different(rec_a, rec_b):
    """
    Careful protein comparison (no sequences are stored):
      * different lengths  -> the proteins must be different -> 'Different'
      * same length        -> NOT proof of identity -> 'Same length; sequence not compared'
      * a length is missing -> 'Not determined'
    """
    len_a, len_b = to_int(rec_a["Protein_Length_aa"]), to_int(rec_b["Protein_Length_aa"])
    if len_a is None or len_b is None:
        return "Not determined"
    if len_a != len_b:
        return "Different"
    return "Same length; sequence not compared"


def compare_pair(gene, idx_a, a, idx_b, b):
    """Build one Transcript Comparison row (tuple in COMPARISON_COLS order)."""
    return (
        gene,
        transcript_label(a, idx_a),
        transcript_label(b, idx_b),
        abs_diff(a["Transcript_Length_nt"], b["Transcript_Length_nt"]),
        compare_flag(to_int(a["Exon_Count"]), to_int(b["Exon_Count"])),
        compare_flag(parse_exon_structure(a["Exon_Coordinates"]),
                     parse_exon_structure(b["Exon_Coordinates"])),
        abs_diff(a["CDS_Length_nt"], b["CDS_Length_nt"]),
        abs_diff(a["Protein_Length_aa"], b["Protein_Length_aa"]),
        compare_flag(a["Protein_ID"].strip() or None, b["Protein_ID"].strip() or None),
        protein_different(a, b),
    )


def build_comparison_rows(records):
    """Every pair of transcripts that belong to the same gene."""
    rows = []
    for gene, members in group_by_gene(records).items():
        for (i, a), (j, b) in itertools.combinations(members, 2):
            rows.append(compare_pair(gene, i, a, j, b))
    return rows


def summarize_difference(values):
    """
    values = one parsed value per transcript of a gene (None = not recorded).
      'N/A'            only one transcript, nothing to compare
      'Yes'            at least two different recorded values
      'Not determined' values are missing so 'No' cannot be claimed
      'No'             all values recorded and identical
    """
    if len(values) < 2:
        return "N/A"
    present = [v for v in values if v is not None]
    if len(set(present)) > 1:
        return "Yes"
    if len(present) < len(values):
        return "Not determined"
    return "No"


def build_gene_summary_rows(records):
    """One row per gene, calculated from the transcript table."""
    rows = []
    for gene, members in group_by_gene(records).items():
        recs = [r for _, r in members]

        gene_ids = [r["GeneID"].strip() for r in recs if r["GeneID"].strip()]
        gene_id = gene_ids[0] if gene_ids else ""
        if to_int(gene_id) is not None:
            gene_id = to_int(gene_id)

        # Protein sequence comparison, summarised from the pairwise results
        pair_results = [protein_different(a, b) for a, b in itertools.combinations(recs, 2)]
        if not pair_results:
            protein_seq = "N/A"
        elif len(set(pair_results)) == 1:
            protein_seq = pair_results[0]
        else:
            protein_seq = "Mixed - see Transcript Comparison"

        rows.append((
            gene,
            gene_id,
            len(recs),
            summarize_difference([to_int(r["Transcript_Length_nt"]) for r in recs]),
            summarize_difference([parse_exon_structure(r["Exon_Coordinates"]) for r in recs]),
            summarize_difference([to_int(r["CDS_Length_nt"]) for r in recs]),
            summarize_difference([to_int(r["Protein_Length_aa"]) for r in recs]),
            protein_seq,
        ))
    return rows


# =====================================================================
# 6. ANALYSIS TEXT
# =====================================================================

def describe_column(records, column, label):
    """Min / max / mean of one numeric column (only records with a value)."""
    values = [v for v in (to_int(r[column]) for r in records) if v is not None]
    lines = ["{} (values recorded: {} of {} records)".format(label, len(values), len(records))]
    if values:
        lines.append("  Minimum: {}".format(min(values)))
        lines.append("  Maximum: {}".format(max(values)))
        lines.append("  Mean:    {:.2f}".format(statistics.mean(values)))
    else:
        lines.append("  No values recorded")
    return lines


def build_analysis_text(records):
    groups = group_by_gene(records)
    n_genes = len(groups)
    n_tx = sum(len(m) for m in groups.values())
    summary = build_gene_summary_rows(records)

    def count_line(text, col_index):
        yes = sum(1 for row in summary if row[col_index] == "Yes")
        undetermined = sum(1 for row in summary if row[col_index] == "Not determined")
        extra = "   ({} gene(s) not determined: missing values)".format(undetermined) if undetermined else ""
        return "  {}: {}{}".format(text, yes, extra)

    lines = ["DESCRIPTIVE STATISTICS", "=" * 60,
             "Total genes: {}".format(n_genes),
             "Total transcripts: {}".format(n_tx),
             "Average transcripts per gene: {:.2f}".format(n_tx / n_genes), ""]

    lines += describe_column(records, "Transcript_Length_nt", "Transcript length (nt)") + [""]
    lines += describe_column(records, "CDS_Length_nt", "CDS length (nt)") + [""]
    lines += describe_column(records, "Protein_Length_aa", "Protein length (aa)") + [""]

    lines += ["WITHIN-GENE COMPARISON COUNTS", "-" * 60,
              "  Genes with multiple transcripts: {}".format(
                  sum(1 for m in groups.values() if len(m) > 1)),
              count_line("Genes with different transcript lengths", 3),
              count_line("Genes with different exon structures", 4),
              count_line("Genes with different CDS lengths", 5),
              count_line("Genes with different protein lengths", 6), ""]

    # Check against the planned design (report only; nothing is changed)
    lines += ["PILOT DESIGN CHECK (planned: 10 genes / 18 transcripts)", "-" * 60]
    problems = []
    for gene, (gid, expected_n) in EXPECTED_GENES.items():
        members = groups.get(gene, [])
        if len(members) != expected_n:
            problems.append("{}: expected {} transcript(s), found {}".format(gene, expected_n, len(members)))
        for _, rec in members:
            if rec["GeneID"].strip() and to_int(rec["GeneID"]) != gid:
                problems.append("{}: GeneID {} differs from planned {}".format(gene, rec["GeneID"], gid))
                break
    for gene in groups:
        if gene not in EXPECTED_GENES:
            problems.append("{}: not in the planned gene list".format(gene))
    lines += ["  " + p for p in problems] if problems else ["  All 10 genes present with the planned transcript counts."]

    lines += ["", "NOTE: protein sequences are not stored, so equal protein length is",
              "never reported as identical sequence."]
    return "\n".join(lines)


# =====================================================================
# 7. ADD / EDIT DIALOG
# =====================================================================

class RecordDialog(tk.Toplevel):
    """Simple modal form for adding or editing one transcript record."""

    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.result = None

        initial = initial or make_record()
        self.vars = {}
        for row, col in enumerate(COLS):
            ttk.Label(self, text=col).grid(row=row, column=0, sticky="e", padx=6, pady=2)
            var = tk.StringVar(value=initial[col])
            if col == "Status":
                widget = ttk.Combobox(self, textvariable=var, values=("",) + STATUS_VALUES,
                                      state="readonly", width=20)
            else:
                widget = ttk.Entry(self, textvariable=var, width=60)
            widget.grid(row=row, column=1, sticky="w", padx=6, pady=2)
            self.vars[col] = var

        buttons = ttk.Frame(self)
        buttons.grid(row=len(COLS), column=0, columnspan=2, pady=8)
        ttk.Button(buttons, text="OK", command=self.on_ok).pack(side="left", padx=5)
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="left", padx=5)

    def on_ok(self):
        record = {col: var.get().strip() for col, var in self.vars.items()}
        errors, warnings = validate_record(record)
        if errors:
            messagebox.showwarning("Invalid record", "\n".join("• " + e for e in errors), parent=self)
            return
        if warnings:
            text = "\n".join("• " + w for w in warnings) + "\n\nSave this record anyway?"
            if not messagebox.askyesno("Please check", text, parent=self):
                return
        self.result = record
        self.destroy()

    def show(self):
        """Open the dialog and return the record (or None if cancelled)."""
        self.wait_visibility()
        self.grab_set()
        self.wait_window()
        return self.result


# =====================================================================
# 8. MAIN APPLICATION WINDOW
# =====================================================================

MAIN_WIDTHS = {
    "Gene": 70, "GeneID": 70, "Transcript_ID": 120, "Variant": 80,
    "Transcript_Length_nt": 140, "Exon_Count": 90, "Exon_Coordinates": 300,
    "CDS_Start": 80, "CDS_End": 80, "CDS_Length_nt": 110, "Protein_ID": 120,
    "Protein_Length_aa": 120, "Protein_Product": 300, "Status": 80,
}


def make_tree(parent, columns, widths=None):
    """Create a Treeview with vertical + horizontal scrollbars."""
    frame = ttk.Frame(parent)
    tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
    ysb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    xsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    for col in columns:
        width = (widths or {}).get(col, max(90, len(col) * 9))
        tree.heading(col, text=col)
        tree.column(col, width=width, minwidth=60, stretch=False, anchor="w")
    tree.grid(row=0, column=0, sticky="nsew")
    ysb.grid(row=0, column=1, sticky="ns")
    xsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    return frame, tree


def format_messages(messages, limit=12):
    """Bullet list for message boxes (long lists are shortened)."""
    text = "\n".join("• " + m for m in messages[:limit])
    if len(messages) > limit:
        text += "\n... and {} more.".format(len(messages) - limit)
    return text


class TranscriptApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Transcript Structure Analysis - Pilot Dataset")
        self.root.geometry("1200x620")
        self.records = []      # list of record dicts (all values are strings)
        self.dirty = False     # True when there are unsaved changes
        self.status_var = tk.StringVar()

        self.build_widgets()
        self.load_startup_data()
        self.refresh_all()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- layout ----------
    def build_widgets(self):
        # Top section: buttons
        top = ttk.Frame(self.root, padding=6)
        top.pack(side="top", fill="x")
        buttons = [
            ("Add Record", self.add_record),
            ("Edit Record", self.edit_record),
            ("Delete Record", self.delete_record),
            ("Save Dataset", self.save_dataset),
            ("Load Dataset", self.load_dataset),
            ("Export CSV", self.export_csv),
            ("Export Excel", self.export_excel),
            ("Analyze", self.analyze),
            ("Compare Transcripts", self.compare_transcripts),
        ]
        for text, command in buttons:
            ttk.Button(top, text=text, command=command).pack(side="left", padx=2)

        # Status bar
        ttk.Label(self.root, textvariable=self.status_var, anchor="w",
                  relief="sunken").pack(side="bottom", fill="x")

        # Main section: three tables in tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(side="top", fill="both", expand=True, padx=6, pady=4)

        frame1, self.tree = make_tree(self.notebook, COLS, MAIN_WIDTHS)
        frame2, self.cmp_tree = make_tree(self.notebook, COMPARISON_COLS)
        frame3, self.sum_tree = make_tree(self.notebook, SUMMARY_COLS)
        self.notebook.add(frame1, text="Transcripts")
        self.notebook.add(frame2, text="Transcript Comparison")
        self.notebook.add(frame3, text="Gene Summary")

    # ---------- refreshing the tables ----------
    def refresh_all(self):
        """Redraw all three tables from self.records."""
        for tree in (self.tree, self.cmp_tree, self.sum_tree):
            tree.delete(*tree.get_children())
        for idx, rec in enumerate(self.records):
            self.tree.insert("", "end", iid=str(idx), values=[rec[c] for c in COLS])
        for row in build_comparison_rows(self.records):
            self.cmp_tree.insert("", "end", values=list(row))
        for row in build_gene_summary_rows(self.records):
            self.sum_tree.insert("", "end", values=list(row))

    def set_status(self, text):
        n_genes = len(group_by_gene(self.records))
        suffix = "  [unsaved changes]" if self.dirty else ""
        self.status_var.set("{}   |   {} transcripts, {} genes{}".format(
            text, len(self.records), n_genes, suffix))

    def mark_changed(self, text):
        self.dirty = True
        self.refresh_all()
        self.set_status(text)

    # ---------- start-up ----------
    def load_startup_data(self):
        try:
            init_database()
            rows = load_records_from_db()
        except sqlite3.Error as err:
            messagebox.showerror("Database error", str(err))
            rows = []
        if rows:
            self.records = rows
            self.set_status("Loaded from transcript_analysis.db")
        elif PILOT_DATA:
            self.records = [make_record(**r) for r in PILOT_DATA]
            self.dirty = True
            self.set_status("Pilot data loaded from PILOT_DATA - click Save Dataset")
        else:
            self.set_status("Empty dataset - add records or use Load Dataset")

    # ---------- validation gate ----------
    def run_validation(self, action):
        """Validate the whole table. Returns True if it is OK to continue."""
        errors, warnings = validate_dataset(self.records)
        if errors:
            messagebox.showwarning(
                "Validation errors",
                "Cannot {} until these are fixed:\n\n{}".format(action, format_messages(errors)))
            return False
        if warnings:
            return messagebox.askyesno(
                "Validation warnings",
                "Please review these warnings:\n\n{}\n\nContinue to {} anyway?".format(
                    format_messages(warnings), action))
        return True

    # ---------- record buttons ----------
    def get_selected_index(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("No row selected", "Select a row in the Transcripts table first.")
            self.notebook.select(0)
            return None
        return int(selection[0])

    def add_record(self):
        record = RecordDialog(self.root, "Add Record").show()
        if record:
            self.records.append(record)
            self.mark_changed("Record added")

    def edit_record(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        record = RecordDialog(self.root, "Edit Record", initial=self.records[idx]).show()
        if record:
            self.records[idx] = record
            self.mark_changed("Record updated")

    def delete_record(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        rec = self.records[idx]
        if messagebox.askyesno("Delete record",
                               "Delete {} {}?".format(rec["Gene"], rec["Transcript_ID"]).strip()):
            del self.records[idx]
            self.mark_changed("Record deleted")

    # ---------- save / load ----------
    def save_dataset(self):
        """Validate, then write the table to SQLite. Returns True on success."""
        if not self.run_validation("save"):
            return False
        try:
            save_records_to_db(self.records)
        except sqlite3.Error as err:
            messagebox.showerror("Database error", str(err))
            return False
        self.dirty = False
        self.set_status("Saved to transcript_analysis.db")
        return True

    def load_dataset(self):
        if self.dirty and not messagebox.askyesno(
                "Unsaved changes", "Loading will replace the current table (unsaved changes are lost). Continue?"):
            return
        choice = messagebox.askyesnocancel(
            "Load Dataset",
            "Yes = load from the SQLite database (transcript_analysis.db)\n"
            "No = load from a CSV file\nCancel = do nothing")
        if choice is None:
            return
        try:
            if choice:
                records = load_records_from_db()
                source = "database"
            else:
                path = filedialog.askopenfilename(title="Choose a CSV file",
                                                  filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
                if not path:
                    return
                records = load_records_from_csv(path)
                source = "CSV"
        except (sqlite3.Error, OSError, ValueError, csv.Error) as err:
            messagebox.showerror("Load failed", str(err))
            return
        self.records = records
        self.dirty = (source == "CSV")   # CSV data is not in the database until you save
        self.refresh_all()
        self.set_status("Loaded from " + source)

    # ---------- export ----------
    def export_csv(self):
        if not self.run_validation("export"):
            return
        path = filedialog.asksaveasfilename(title="Export CSV", defaultextension=".csv",
                                            initialfile="transcripts.csv",
                                            filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            # utf-8-sig so Excel shows en dashes correctly
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(COLS)
                for rec in self.records:
                    writer.writerow([rec[c] for c in COLS])
        except OSError as err:
            messagebox.showerror("Export failed", str(err))
            return
        self.set_status("Exported CSV: " + path)

    def export_excel(self):
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Excel export unavailable",
                                 "Excel export needs pandas and openpyxl.\n\nInstall with:\npip install pandas openpyxl")
            return
        if not self.run_validation("export"):
            return
        path = filedialog.asksaveasfilename(title="Export Excel", defaultextension=".xlsx",
                                            initialfile="transcripts.xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        def clean(row):
            """Blank text -> empty cell (None)."""
            return [None if v == "" else v for v in row]

        tx_rows = []
        for rec in self.records:
            tx_rows.append([to_int(rec[c]) if c in NUMERIC_COLS else (rec[c].strip() or None)
                            for c in COLS])
        try:
            # dtype=object keeps whole numbers as whole numbers (no 5.0)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                pd.DataFrame(tx_rows, columns=COLS, dtype=object).to_excel(
                    writer, sheet_name="Transcripts", index=False)
                pd.DataFrame([clean(r) for r in build_comparison_rows(self.records)],
                             columns=COMPARISON_COLS, dtype=object).to_excel(
                    writer, sheet_name="Transcript_Comparison", index=False)
                pd.DataFrame([clean(r) for r in build_gene_summary_rows(self.records)],
                             columns=SUMMARY_COLS, dtype=object).to_excel(
                    writer, sheet_name="Gene_Summary", index=False)
        except (OSError, ValueError) as err:
            messagebox.showerror("Export failed", str(err))
            return
        self.set_status("Exported Excel: " + path)

    # ---------- analysis ----------
    def compare_transcripts(self):
        self.refresh_all()
        self.notebook.select(1)
        if not self.cmp_tree.get_children():
            messagebox.showinfo("Transcript Comparison",
                                "No gene has two or more transcripts yet, so there is nothing to compare.")

    def analyze(self):
        if not self.records:
            messagebox.showinfo("Analyze", "The dataset is empty.")
            return
        if not self.run_validation("analyze"):
            return
        self.refresh_all()
        self.show_text_window("Analysis results", build_analysis_text(self.records))

    def show_text_window(self, title, text):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("700x600")
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=6, pady=6)
        box = tk.Text(frame, wrap="word", font=("Courier", 10))
        scroll = ttk.Scrollbar(frame, orient="vertical", command=box.yview)
        box.configure(yscrollcommand=scroll.set)
        box.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        box.insert("1.0", text)
        box.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=4)

    # ---------- closing ----------
    def on_close(self):
        if self.dirty:
            answer = messagebox.askyesnocancel("Unsaved changes", "Save changes before closing?")
            if answer is None:
                return
            if answer and not self.save_dataset():
                return
        self.root.destroy()


def main():
    root = tk.Tk()
    TranscriptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()