import os
import sys
import json
import shutil
import subprocess
import threading
import traceback
from collections import defaultdict
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pyodbc
except ImportError:
    pyodbc = None


# ============================================================
# CONFIG
# ============================================================

ALLOWED_NON_MWL_CODES = {
    "ABP", "ACB", "ACOH", "ACOM", "ACOP", "ACOS", "ADP", "AEP",
    "AJB", "ALT", "AM", "AMH", "AML", "AOC", "ARV", "ATC",
    "AW", "AWA", "AWD", "AWW", "AZ", "MSA"
}

LACP_REQUIRED_FIELDS = [
    "InspectionID",
    "Surveyed_By",
    "Certificate_Number",
    "Inspection_Date",
    "Direction",
    "Inspection_Status",
    "Lateral_Segment_Reference",
    "Street",
    "City",
    "Size",
    "Material",
    "IsImperial",
]

PACP_REQUIRED_FIELDS = [
    "InspectionID",
    "Surveyed_By",
    "Certificate_Number",
    "Inspection_Date",
    "Direction",
    "Inspection_Status",
    "Street",
    "City",
    "Material",
    "IsImperial",
]


# ============================================================
# REPORT WRITER
# ============================================================

class ReportWriter:
    def __init__(self, callback=None):
        self.lines = []
        self.callback = callback

    def write(self, text=""):
        self.lines.append(text)
        if self.callback:
            self.callback(text)

    def section(self, title):
        self.write("")
        self.write("=" * 88)
        self.write(f">>> {title}")
        self.write("=" * 88)

    def pass_line(self, text):
        self.write(f"[PASS] {text}")

    def fail_line(self, text):
        self.write(f"[FAIL] {text}")

    def warn_line(self, text):
        self.write(f"[WARN] {text}")

    def info_line(self, text):
        self.write(f"       {text}")

    def get_text(self):
        return "\n".join(self.lines)


# ============================================================
# HELPERS
# ============================================================

def build_connection_string(db_path):
    return (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={db_path};"
    )


def connect_db(db_path):
    conn_str = build_connection_string(db_path)
    return pyodbc.connect(conn_str, autocommit=True)


def get_table_columns(cursor, table_name):
    """
    Safer than cursor.columns(...) for some Access files.
    Pull column names from cursor.description using a no-row query.
    """
    sql = f"SELECT * FROM [{table_name}] WHERE 1=0"
    cursor.execute(sql)
    if cursor.description is None:
        return []
    return [col[0] for col in cursor.description]


def is_blank(value):
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def safe_join(root, subpath):
    """
    Resolve a media folder path stored in the database.

    This handles the common deliverable layout:

        Project Folder/
            Pictures/
            Video/

    The GUI may still ask for the Video folder, but image records in the
    database often point to Pictures. So this function tries the selected
    root first, then also tries the parent folder as a media root.

    It also fixes database values like \\Pictures. A leading slash in Access
    should not be treated as a Windows absolute path here.
    """
    root = os.path.normpath(str(root).strip())

    if subpath is None:
        return root

    raw_subpath = str(subpath).strip()
    if raw_subpath == "":
        return root

    # Normalize slashes, but intentionally strip leading slashes first.
    # Access values like \Pictures mean "Pictures folder", not C:\Pictures.
    cleaned = raw_subpath.replace("/", os.sep).replace("\\", os.sep)
    cleaned = cleaned.lstrip(os.sep).strip()

    if cleaned == "":
        return root

    # If the DB contains a real absolute path with a drive letter, use it.
    # Example: C:\SomeFolder\Pictures
    drive, _ = os.path.splitdrive(cleaned)
    if drive and os.path.isabs(cleaned):
        return os.path.normpath(cleaned)

    cleaned = os.path.normpath(cleaned)

    root_parts = [p for p in root.split(os.sep) if p]
    sub_parts = [p for p in cleaned.split(os.sep) if p]

    if root_parts and sub_parts:
        # Avoid .../Video/Video when the DB path already starts with Video.
        if root_parts[-1].lower() == sub_parts[0].lower():
            sub_parts = sub_parts[1:]

    if not sub_parts:
        return root

    # Candidate 1: selected root + database path.
    candidate = os.path.normpath(os.path.join(root, *sub_parts))
    if os.path.isdir(candidate):
        return candidate

    # Candidate 2: parent of selected root + database path.
    # This fixes Project/Video selected as root while DB points to Pictures.
    parent_root = os.path.dirname(root)
    parent_candidate = os.path.normpath(os.path.join(parent_root, *sub_parts))
    if os.path.isdir(parent_candidate):
        return parent_candidate

    # Candidate 3: if selected root is Video and DB path is Pictures, force sibling.
    if root_parts and root_parts[-1].lower() == "video" and sub_parts[0].lower() == "pictures":
        sibling_candidate = os.path.normpath(os.path.join(parent_root, *sub_parts))
        return sibling_candidate

    # Return the original root-based guess if nothing exists yet, so the report
    # still shows a sensible expected path.
    return candidate


def list_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        return []
    try:
        return [
            name for name in os.listdir(folder_path)
            if os.path.isfile(os.path.join(folder_path, name))
        ]
    except Exception:
        return []


def count_jpg_files(file_list):
    return sum(1 for f in file_list if f.lower().endswith((".jpg", ".jpeg")))


def get_table_name_map(cursor):
    table_map = {}
    for row in cursor.tables():
        if row.table_type and row.table_type.upper() == "TABLE":
            table_map[row.table_name.lower()] = row.table_name
    return table_map


def resolve_table_name(cursor, requested_table_name):
    table_map = get_table_name_map(cursor)
    return table_map.get(requested_table_name.lower(), requested_table_name)


# ============================================================
# CHECKS
# ============================================================

def check_condition_table(cursor, table_name, report):
    actual_table_name = resolve_table_name(cursor, table_name)
    report.section(f"CHECKING CONDITION CODES: {actual_table_name}")

    try:
        columns = get_table_columns(cursor, actual_table_name)
    except Exception as e:
        report.write(f"ERROR reading columns for {actual_table_name}: {e}")
        return

    colmap = {c.lower(): c for c in columns}

    inspection_col = colmap.get("inspectionid")
    if not inspection_col:
        report.write(f"ERROR: {actual_table_name} does not contain InspectionID.")
        return

    code_col = None
    for candidate in ["LACP_Code", "PACP_Code", "Code"]:
        if candidate.lower() in colmap:
            code_col = colmap[candidate.lower()]
            break

    if not code_col:
        report.write(f"ERROR: {actual_table_name} does not contain LACP_Code, PACP_Code, or Code.")
        return

    sql = f"""
        SELECT [{inspection_col}] AS InspID,
               [{code_col}] AS CondCode
        FROM [{actual_table_name}]
        WHERE [{inspection_col}] IS NOT NULL
    """

    rows = cursor.execute(sql).fetchall()

    grouped = defaultdict(list)
    for row in rows:
        inspection_id = row.InspID
        code = normalize_text(row.CondCode).upper()
        if code:
            grouped[inspection_id].append(code)

    issues = []

    for inspection_id, codes in grouped.items():
        if len(codes) < 3:
            issues.append((inspection_id, "Less than 3 code entries", codes))
            continue

        if "MWL" not in codes:
            issues.append((inspection_id, "MWL is missing", codes))
            continue

        allowed_non_mwl_count = sum(1 for code in codes if code in ALLOWED_NON_MWL_CODES)

        if allowed_non_mwl_count < 2:
            issues.append(
                (
                    inspection_id,
                    "Fewer than 2 allowed non-MWL codes present",
                    codes
                )
            )
            continue

    report.write(f"Total InspectionIDs checked: {len(grouped)}")
    report.write(f"InspectionIDs with issues: {len(issues)}")

    if issues:
        report.write("")
        report.write("Problem InspectionIDs:")
        for inspection_id, reason, codes in issues:
            report.write(f"  InspectionID={inspection_id} | {reason} | Codes={codes}")
    else:
        report.write(f"{actual_table_name}: all condition code groups passed")


def check_media_table(cursor, table_name, folder_field, file_field, root_folder, jpg_only, report):
    actual_table_name = resolve_table_name(cursor, table_name)
    report.section(f"CHECKING MEDIA: {actual_table_name}")

    try:
        columns = get_table_columns(cursor, actual_table_name)
    except Exception as e:
        report.write(f"ERROR reading columns for {actual_table_name}: {e}")
        return

    colmap = {c.lower(): c for c in columns}

    if folder_field.lower() not in colmap:
        report.write(f"ERROR: {actual_table_name} is missing field: {folder_field}")
        return

    if file_field.lower() not in colmap:
        report.write(f"ERROR: {actual_table_name} is missing field: {file_field}")
        return

    folder_col = colmap[folder_field.lower()]
    file_col = colmap[file_field.lower()]

    sql = f"""
        SELECT [{folder_col}] AS DbFolder,
               [{file_col}] AS DbFile
        FROM [{actual_table_name}]
    """

    rows = cursor.execute(sql).fetchall()

    db_refs_by_folder = defaultdict(list)
    total_db_refs = 0

    for row in rows:
        db_folder = normalize_text(row.DbFolder)
        db_file = normalize_text(row.DbFile)

        if is_blank(db_folder) or is_blank(db_file):
            continue

        if jpg_only and not db_file.lower().endswith((".jpg", ".jpeg")):
            continue

        db_refs_by_folder[db_folder].append(db_file)
        total_db_refs += 1

    report.write(f"Folders referenced in database: {len(db_refs_by_folder)}")
    report.write(f"Total file references in database: {total_db_refs}")

    missing_folders = []
    missing_files = []
    folder_summaries = []

    for db_folder, db_files in db_refs_by_folder.items():
        actual_folder = safe_join(root_folder, db_folder)

        if not os.path.isdir(actual_folder):
            missing_folders.append((db_folder, actual_folder, len(db_files)))
            for db_file in db_files:
                missing_files.append((db_folder, db_file, actual_folder, "Folder does not exist"))
            folder_summaries.append((db_folder, actual_folder, 0, len(db_files), False))
            continue

        actual_files = list_files_in_folder(actual_folder)

        if jpg_only:
            actual_count = count_jpg_files(actual_files)
            actual_file_set = {
                f.lower() for f in actual_files
                if f.lower().endswith((".jpg", ".jpeg"))
            }
        else:
            actual_count = len(actual_files)
            actual_file_set = {f.lower() for f in actual_files}

        db_count = len(db_files)
        folder_summaries.append((db_folder, actual_folder, actual_count, db_count, True))

        for db_file in db_files:
            if db_file.lower() not in actual_file_set:
                missing_files.append((db_folder, db_file, actual_folder, "File missing"))

    report.write("")
    report.write("Folder Counts:")
    for db_folder, actual_folder, actual_count, db_count, exists in folder_summaries:
        exists_text = "YES" if exists else "NO"
        report.write(
            f"  DB Folder={db_folder} | Exists={exists_text} | "
            f"Folder Count={actual_count} | DB Count={db_count} | Path={actual_folder}"
        )

    report.write("")
    report.write("Summary:")
    report.write(f"  Missing folders: {len(missing_folders)}")
    report.write(f"  Missing files: {len(missing_files)}")

    if missing_folders:
        report.write("")
        report.write("Missing Folders:")
        for db_folder, actual_folder, ref_count in missing_folders:
            report.write(
                f"  Folder from DB={db_folder} | Expected Path={actual_folder} | DB References={ref_count}"
            )
    else:
        report.write(f"{actual_table_name}: all referenced folders are present")

    if missing_files:
        report.write("")
        report.write("Missing Files:")
        for db_folder, db_file, actual_folder, reason in missing_files:
            report.write(
                f"  Image_Path/Video_Location={db_folder} | "
                f"Image_Reference/Video_Name={db_file} | "
                f"Expected Folder={actual_folder} | Reason={reason}"
            )
    else:
        report.write(f"{actual_table_name}: all referenced files are present")


def check_required_fields(cursor, table_name, required_fields, report):
    actual_table_name = resolve_table_name(cursor, table_name)
    report.section(f"CHECKING REQUIRED FIELDS: {actual_table_name}")

    try:
        columns = get_table_columns(cursor, actual_table_name)
    except Exception as e:
        report.write(f"ERROR reading columns for {actual_table_name}: {e}")
        return

    colmap = {c.lower(): c for c in columns}

    missing_columns = [f for f in required_fields if f.lower() not in colmap]
    if missing_columns:
        report.write("WARNING: These required columns do not exist in the table:")
        for col in missing_columns:
            report.write(f"  {col}")

    if "inspectionid" not in colmap:
        report.write("ERROR: InspectionID column is missing, cannot continue this check.")
        return

    inspection_col = colmap["inspectionid"]

    select_parts = [f"[{inspection_col}] AS InspID"]
    selected_actual_cols = []

    for field in required_fields:
        if field.lower() == "inspectionid":
            continue
        if field.lower() in colmap:
            actual_col = colmap[field.lower()]
            selected_actual_cols.append(actual_col)
            select_parts.append(f"[{actual_col}]")

    sql = f"SELECT {', '.join(select_parts)} FROM [{actual_table_name}]"
    rows = cursor.execute(sql).fetchall()

    issues = []

    for row in rows:
        inspection_id = row.InspID

        if is_blank(inspection_id):
            issues.append((actual_table_name, inspection_id, "InspectionID"))

        for actual_col in selected_actual_cols:
            value = getattr(row, actual_col)
            if is_blank(value):
                issues.append((actual_table_name, inspection_id, actual_col))

    report.write(f"Rows checked: {len(rows)}")
    report.write(f"Missing values: {len(issues)}")

    if issues:
        report.write("")
        report.fail_line("Missing required values were found:")
        for table, inspection_id, field in issues:
            report.write(
                f"  Table={table} | InspectionID={inspection_id} | Missing Field={field}"
            )
    else:
        report.pass_line(f"{actual_table_name}: all required fields are populated.")



# ============================================================
# PIPE SEGMENT REFERENCE (PSR)
# ============================================================

def check_pipe_segment_reference(cursor, table_name, report):
    """Explicitly validate Pipe_Segment_Reference for one inspection table."""
    actual_table_name = resolve_table_name(cursor, table_name)
    report.section(f"PIPE SEGMENT REFERENCE (PSR) CHECK - {actual_table_name}")
    report.write(f"Table: {actual_table_name}")
    report.write("Field: Pipe_Segment_Reference")
    report.write("Rule : Value must not be NULL, empty, or whitespace only")
    report.write("")

    try:
        columns = get_table_columns(cursor, actual_table_name)
    except Exception as exc:
        report.fail_line(f"Could not read columns from {actual_table_name}: {exc}")
        return

    colmap = {c.lower(): c for c in columns}
    if "pipe_segment_reference" not in colmap:
        report.fail_line(f"{actual_table_name}.Pipe_Segment_Reference column does not exist.")
        return
    if "inspectionid" not in colmap:
        report.fail_line(f"{actual_table_name}.InspectionID column does not exist, so affected rows cannot be identified.")
        return

    psr_col = colmap["pipe_segment_reference"]
    inspection_col = colmap["inspectionid"]
    sql = f"SELECT [{inspection_col}] AS InspID, [{psr_col}] AS PSR FROM [{actual_table_name}]"
    rows = cursor.execute(sql).fetchall()

    blank_rows = []
    for row in rows:
        if is_blank(row.PSR):
            blank_rows.append(row.InspID)

    report.write(f"Rows checked : {len(rows)}")
    report.write(f"Blank PSRs   : {len(blank_rows)}")
    report.write("")

    if blank_rows:
        report.fail_line(f"{actual_table_name} contains {len(blank_rows)} blank Pipe_Segment_Reference value(s).")
        report.write("Affected rows:")
        for inspection_id in blank_rows:
            report.write(f"   - InspectionID: {inspection_id} | Pipe_Segment_Reference: BLANK")
    else:
        report.pass_line(f"{actual_table_name}.Pipe_Segment_Reference has no blank values.")

# ============================================================
# VIDEO INTEGRITY / FFPROBE
# ============================================================

VIDEO_EXTENSIONS = {".mp4", ".avi", ".mov", ".mkv", ".mpeg", ".mpg", ".m4v", ".wmv", ".ts", ".mts", ".m2ts"}

def get_ffprobe_path():
    """Find bundled ffprobe first, then fall back to PATH."""
    candidates = []
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        candidates.extend([
            os.path.join(base, "ffprobe.exe"),
            os.path.join(os.path.dirname(sys.executable), "ffprobe.exe"),
        ])
    else:
        candidates.extend([
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "ffprobe.exe"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin", "ffprobe.exe"),
        ])
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return candidate
    return shutil.which("ffprobe") or shutil.which("ffprobe.exe")


def validate_video_with_ffprobe(video_path, ffprobe_path):
    """Return (is_valid, reason) based on ffprobe's ability to read a real video stream."""
    cmd = [
        ffprobe_path, "-v", "error",
        "-show_entries", "format=duration:stream=codec_type",
        "-of", "json", video_path
    ]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except subprocess.TimeoutExpired:
        return False, "FFprobe timed out"
    except Exception as exc:
        return False, f"FFprobe error: {exc}"

    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "FFprobe could not read the file").strip()
        return False, detail.splitlines()[0][:220]

    try:
        data = json.loads(result.stdout or "{}")
    except json.JSONDecodeError:
        return False, "FFprobe returned invalid metadata"

    streams = data.get("streams") or []
    if not any(stream.get("codec_type") == "video" for stream in streams):
        return False, "No valid video stream found"

    raw_duration = (data.get("format") or {}).get("duration")
    try:
        duration = float(raw_duration)
    except (TypeError, ValueError):
        return False, "Video duration is missing or invalid"

    if duration <= 0:
        return False, "Video duration is 0 seconds"

    return True, f"{duration:.1f} sec"


def check_video_integrity(video_root, report):
    report.section("VIDEO INTEGRITY (FFPROBE)")
    ffprobe_path = get_ffprobe_path()
    if not ffprobe_path:
        report.write("FFprobe was not found. Video existence checks still ran, but video integrity could not be validated.")
        report.write("Place ffprobe.exe beside the app/source or rebuild with ffprobe.exe included.")
        return

    video_files = []
    for root, _, filenames in os.walk(video_root):
        for filename in filenames:
            if os.path.splitext(filename)[1].lower() in VIDEO_EXTENSIONS:
                video_files.append(os.path.join(root, filename))

    report.write(f"Video files found for integrity validation: {len(video_files)}")
    invalid = []
    for index, path in enumerate(video_files, 1):
        ok, reason = validate_video_with_ffprobe(path, ffprobe_path)
        if not ok:
            invalid.append((path, reason))
        if index % 25 == 0:
            report.write(f"  FFprobe progress: {index}/{len(video_files)}")

    report.write(f"Invalid/corrupt videos found: {len(invalid)}")
    if invalid:
        report.write("")
        report.write("Invalid / Corrupt Videos:")
        for path, reason in invalid:
            report.write(f"  File={path} | Reason={reason}")
    else:
        report.write("All discovered video files passed FFprobe validation.")


# ============================================================
# MAIN CHECK RUNNER
# ============================================================

def run_all_checks(db_path, video_root, pictures_root, report_callback=None):
    report = ReportWriter(callback=report_callback)

    if pyodbc is None:
        raise RuntimeError("pyodbc is not installed. Run: pip install pyodbc")

    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Database file not found: {db_path}")

    if not os.path.isdir(video_root):
        raise FileNotFoundError(f"Video folder not found: {video_root}")

    if not os.path.isdir(pictures_root):
        raise FileNotFoundError(f"Pictures folder not found: {pictures_root}")

    report.section("STARTING CHECKS")
    report.write(f"Database: {db_path}")
    report.write(f"Selected Video Folder: {video_root}")
    report.write(f"Selected Pictures Folder: {pictures_root}")
    report.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    conn = None
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()

        check_condition_table(cursor, "LACP_Conditions", report)
        check_condition_table(cursor, "PACP_Conditions", report)

        check_media_table(
            cursor,
            table_name="LACP_Media_Conditions",
            folder_field="Image_Path",
            file_field="Image_Reference",
            root_folder=pictures_root,
            jpg_only=True,
            report=report
        )

        check_media_table(
            cursor,
            table_name="LACP_Media_Inspections",
            folder_field="Video_Location",
            file_field="Video_Name",
            root_folder=video_root,
            jpg_only=False,
            report=report
        )

        check_media_table(
            cursor,
            table_name="PACP_Media_Conditions",
            folder_field="Image_Path",
            file_field="Image_Reference",
            root_folder=pictures_root,
            jpg_only=True,
            report=report
        )

        check_media_table(
            cursor,
            table_name="PACP_Media_Inspections",
            folder_field="Video_Location",
            file_field="Video_Name",
            root_folder=video_root,
            jpg_only=False,
            report=report
        )

        check_video_integrity(video_root, report)

        check_pipe_segment_reference(cursor, "PACP_Inspections", report)
        check_pipe_segment_reference(cursor, "LACP_Inspections", report)

        check_required_fields(cursor, "LACP_Inspections", LACP_REQUIRED_FIELDS, report)
        check_required_fields(cursor, "PACP_Inspections", PACP_REQUIRED_FIELDS, report)

        report.section("DONE")
        report.write(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    finally:
        if conn is not None:
            conn.close()

    return report.get_text()


# ============================================================
# GUI
# ============================================================

class DatabaseCheckerGUI:
    BG = "#B8D8C0"
    CARD = "#D7E9DB"
    NAVY = "#0F2B46"
    BLUE = "#1473E6"
    BLUE_HOVER = "#0B5FC6"
    BLUE_SOFT = "#EAF3FF"
    BORDER = "#82B391"
    TEXT = "#17324D"
    MUTED = "#62778C"
    SUCCESS = "#138A5B"

    def __init__(self, root):
        self.root = root
        self.root.title("PACP / LACP Database Checker")
        self.root.geometry("1180x800")
        self.root.minsize(980, 680)
        self.root.configure(bg=self.BG)

        self.db_path_var = tk.StringVar()
        self.video_root_var = tk.StringVar()
        self.pictures_root_var = tk.StringVar()
        self.report_text_cache = ""

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("App.TFrame", background=self.BG)
        style.configure("Card.TFrame", background=self.CARD)
        style.configure("Header.TFrame", background=self.NAVY)

        style.configure(
            "Title.TLabel",
            background=self.NAVY,
            foreground="white",
            font=("Segoe UI Semibold", 22)
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.NAVY,
            foreground="#C9DDF3",
            font=("Segoe UI", 10)
        )
        style.configure(
            "Section.TLabel",
            background=self.CARD,
            foreground=self.NAVY,
            font=("Segoe UI Semibold", 12)
        )
        style.configure(
            "Field.TLabel",
            background=self.CARD,
            foreground=self.TEXT,
            font=("Segoe UI Semibold", 10)
        )
        style.configure(
            "Hint.TLabel",
            background=self.CARD,
            foreground=self.MUTED,
            font=("Segoe UI", 9)
        )
        style.configure(
            "Status.TLabel",
            background=self.BG,
            foreground=self.MUTED,
            font=("Segoe UI Semibold", 10)
        )

        style.configure(
            "Modern.TEntry",
            fieldbackground="#FFFFFF",
            foreground=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            padding=8
        )
        style.map(
            "Modern.TEntry",
            bordercolor=[("focus", self.BLUE)],
            lightcolor=[("focus", self.BLUE)],
            darkcolor=[("focus", self.BLUE)]
        )

        style.configure(
            "Browse.TButton",
            background=self.BLUE_SOFT,
            foreground=self.BLUE,
            borderwidth=0,
            padding=(14, 8),
            font=("Segoe UI Semibold", 9)
        )
        style.map(
            "Browse.TButton",
            background=[("active", "#DCEBFF"), ("pressed", "#CCE2FF")],
            foreground=[("active", self.BLUE_HOVER)]
        )

        style.configure(
            "Primary.TButton",
            background=self.BLUE,
            foreground="white",
            borderwidth=0,
            padding=(18, 10),
            font=("Segoe UI Semibold", 10)
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.BLUE_HOVER), ("pressed", "#084E9D"), ("disabled", "#9ABBE2")],
            foreground=[("disabled", "#EAF1F8")]
        )

        style.configure(
            "Secondary.TButton",
            background=self.CARD,
            foreground=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            borderwidth=1,
            padding=(14, 9),
            font=("Segoe UI Semibold", 9)
        )
        style.map(
            "Secondary.TButton",
            background=[("active", self.BLUE_SOFT), ("pressed", "#DCEBFF")],
            foreground=[("active", self.BLUE)]
        )

    def _build_ui(self):
        # Header
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(28, 20))
        header.pack(fill="x")
        ttk.Label(header, text="PACP / LACP Deliverable Checker", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Validate database fields, video deliverables, and inspection images in one pass.",
            style="Subtitle.TLabel"
        ).pack(anchor="w", pady=(4, 0))

        main = ttk.Frame(self.root, style="App.TFrame", padding=(24, 20, 24, 24))
        main.pack(fill="both", expand=True)

        # Input card
        card = tk.Frame(main, bg=self.CARD, highlightbackground=self.BORDER, highlightthickness=1, bd=0)
        card.pack(fill="x", pady=(0, 16))

        card_inner = ttk.Frame(card, style="Card.TFrame", padding=(20, 16))
        card_inner.pack(fill="both", expand=True)
        ttk.Label(card_inner, text="Deliverable Locations", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            card_inner,
            text="Choose the database and the folders containing its related media.",
            style="Hint.TLabel"
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 14))

        fields = [
            ("Access Database", "Select the .mdb or .accdb file to validate", self.db_path_var, self.browse_db),
            ("Video Folder", "Folder containing inspection video files", self.video_root_var, self.browse_video_root),
            ("Pictures Folder", "Folder containing inspection image files", self.pictures_root_var, self.browse_pictures_root),
        ]

        row = 2
        for label, hint, variable, command in fields:
            label_frame = ttk.Frame(card_inner, style="Card.TFrame")
            label_frame.grid(row=row, column=0, sticky="nw", padx=(0, 16), pady=8)
            ttk.Label(label_frame, text=label, style="Field.TLabel").pack(anchor="w")
            ttk.Label(label_frame, text=hint, style="Hint.TLabel").pack(anchor="w", pady=(2, 0))

            ttk.Entry(card_inner, textvariable=variable, style="Modern.TEntry").grid(
                row=row, column=1, sticky="ew", pady=8
            )
            ttk.Button(card_inner, text="Browse…", command=command, style="Browse.TButton").grid(
                row=row, column=2, padx=(10, 0), pady=8
            )
            row += 1

        card_inner.columnconfigure(1, weight=1)

        # Action bar
        action_frame = ttk.Frame(main, style="App.TFrame")
        action_frame.pack(fill="x", pady=(0, 16))

        self.run_button = ttk.Button(
            action_frame,
            text="▶  Run Checks",
            command=self.start_run,
            style="Primary.TButton"
        )
        self.run_button.pack(side="left")

        ttk.Button(
            action_frame,
            text="Save Report",
            command=self.save_report,
            style="Secondary.TButton"
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            action_frame,
            text="Clear Output",
            command=self.clear_output,
            style="Secondary.TButton"
        ).pack(side="left", padx=(10, 0))

        self.status_var = tk.StringVar(value="Ready")
        self.status_dot = tk.Label(
            action_frame,
            text="●",
            bg=self.BG,
            fg=self.SUCCESS,
            font=("Segoe UI", 11)
        )
        self.status_dot.pack(side="right", padx=(8, 0))
        ttk.Label(action_frame, textvariable=self.status_var, style="Status.TLabel").pack(side="right")

        # Report card
        output_card = tk.Frame(main, bg=self.CARD, highlightbackground=self.BORDER, highlightthickness=1, bd=0)
        output_card.pack(fill="both", expand=True)

        output_header = tk.Frame(output_card, bg=self.CARD)
        output_header.pack(fill="x", padx=18, pady=(14, 10))
        tk.Label(
            output_header,
            text="Validation Report",
            bg=self.CARD,
            fg=self.NAVY,
            font=("Segoe UI Semibold", 12)
        ).pack(side="left")
        tk.Label(
            output_header,
            text="Results appear here as checks complete",
            bg=self.CARD,
            fg=self.MUTED,
            font=("Segoe UI", 9)
        ).pack(side="right")

        text_container = tk.Frame(output_card, bg=self.CARD)
        text_container.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        self.text = tk.Text(
            text_container,
            wrap="none",
            font=("Consolas", 10),
            bg="#0D1B2A",
            fg="#D9E7F5",
            insertbackground="white",
            selectbackground=self.BLUE,
            selectforeground="white",
            relief="flat",
            bd=0,
            padx=14,
            pady=12
        )
        self.text.tag_configure("section", foreground="#7DC4FF", font=("Consolas", 10, "bold"))
        self.text.tag_configure("pass", foreground="#72E6A1", font=("Consolas", 10, "bold"))
        self.text.tag_configure("fail", foreground="#FF8A8A", font=("Consolas", 10, "bold"))
        self.text.tag_configure("warn", foreground="#FFD166", font=("Consolas", 10, "bold"))
        self.text.tag_configure("label", foreground="#B8D8FF")
        self.text.pack(side="left", fill="both", expand=True)

        y_scroll = ttk.Scrollbar(text_container, orient="vertical", command=self.text.yview)
        y_scroll.pack(side="right", fill="y")
        self.text.configure(yscrollcommand=y_scroll.set)

    def _set_status(self, text, color=None):
        self.status_var.set(text)
        if color:
            self.status_dot.configure(fg=color)

    def browse_db(self):
        path = filedialog.askopenfilename(
            title="Select Access Database",
            filetypes=[
                ("Access Database", "*.mdb *.accdb"),
                ("MDB Files", "*.mdb"),
                ("ACCDB Files", "*.accdb"),
                ("All Files", "*.*"),
            ]
        )
        if path:
            self.db_path_var.set(path)

    def browse_video_root(self):
        path = filedialog.askdirectory(title="Select Video Folder")
        if path:
            self.video_root_var.set(path)

    def browse_pictures_root(self):
        path = filedialog.askdirectory(title="Select Pictures Folder")
        if path:
            self.pictures_root_var.set(path)

    def append_output(self, line):
        def _append():
            tag = None
            stripped = line.strip()
            if stripped.startswith(">>>") or (stripped and set(stripped) == {"="}):
                tag = "section"
            elif stripped.startswith("[PASS]"):
                tag = "pass"
            elif stripped.startswith("[FAIL]") or stripped.startswith("ERROR"):
                tag = "fail"
            elif stripped.startswith("[WARN]") or stripped.startswith("WARNING"):
                tag = "warn"
            elif stripped.startswith(("Table:", "Field:", "Rule :", "Rows checked", "Blank PSRs", "Summary:")):
                tag = "label"

            if tag:
                self.text.insert("end", line + "\n", tag)
            else:
                self.text.insert("end", line + "\n")
            self.text.see("end")
        self.root.after(0, _append)

    def clear_output(self):
        self.text.delete("1.0", "end")
        self.report_text_cache = ""
        self._set_status("Ready", self.SUCCESS)

    def save_report(self):
        content = self.text.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("No Report", "There is no report to save yet.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Report",
            defaultextension=".txt",
            filetypes=[
                ("Text Files", "*.txt"),
                ("All Files", "*.*"),
            ]
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Saved", f"Report saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def start_run(self):
        db_path = self.db_path_var.get().strip()
        video_root = self.video_root_var.get().strip()
        pictures_root = self.pictures_root_var.get().strip()

        if not db_path:
            messagebox.showerror("Missing Database", "Please select the Access database file.")
            return

        if not video_root:
            messagebox.showerror("Missing Video Folder", "Please select the Video folder.")
            return

        if not pictures_root:
            messagebox.showerror("Missing Pictures Folder", "Please select the Pictures folder.")
            return

        self.clear_output()
        self.run_button.config(state="disabled")
        self._set_status("Running checks…", self.BLUE)

        worker = threading.Thread(
            target=self._run_checks_thread,
            args=(db_path, video_root, pictures_root),
            daemon=True
        )
        worker.start()

    def _run_checks_thread(self, db_path, video_root, pictures_root):
        try:
            report_text = run_all_checks(
                db_path=db_path,
                video_root=video_root,
                pictures_root=pictures_root,
                report_callback=self.append_output
            )
            self.report_text_cache = report_text
            self.root.after(0, lambda: self._set_status("Finished", self.SUCCESS))
        except Exception as e:
            error_text = f"ERROR: {str(e)}\n\n{traceback.format_exc()}"
            self.append_output(error_text)
            self.root.after(0, lambda: self._set_status("Failed", "#D64545"))
        finally:
            self.root.after(0, lambda: self.run_button.config(state="normal"))


# ============================================================
# ENTRY POINT
# ============================================================

def main():
    root = tk.Tk()
    DatabaseCheckerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()