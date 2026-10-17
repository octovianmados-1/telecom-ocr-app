import os

# Prevent each OCR worker from spawning another large OpenMP thread pool.
os.environ.setdefault("OMP_THREAD_LIMIT", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")

import re
import csv
import zipfile
import subprocess
import tempfile
import shutil
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
import numpy as np
try:
    import tkinter as tk
    from tkinter import filedialog, ttk
except ImportError:
    tk = None

# ============================================================
# OPTIONAL PYTHON MODULES
# ============================================================

try:
    import cv2
except ImportError:
    cv2 = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    import rarfile
except ImportError:
    rarfile = None

# ============================================================
# CONFIGURATION
# ============================================================

TESSERACT_PATH = None
OUTPUT_CSV = "ocr_result.csv"
# Default to up to 10 workers; override with OCR_MAX_WORKERS=4/6/8/10 to benchmark.
_DEFAULT_WORKERS = min(10, max(1, os.cpu_count() or 1))
try:
    MAX_WORKERS = max(1, min(10, int(os.environ.get("OCR_MAX_WORKERS", _DEFAULT_WORKERS))))
except (TypeError, ValueError):
    MAX_WORKERS = _DEFAULT_WORKERS
DEBUG_MODE = False
DEBUG_DIR_NAME = "ocr_debug"

# ============================================================
# FILE TYPES
# ============================================================

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}

VIDEO_EXTENSIONS = {
    ".mp4",
    ".avi",
    ".mov",
    ".mkv",
    ".3gp",
    ".webm",
}

# ============================================================
# BASIC HELPERS
# ============================================================

def normalize_text(text):
    if not text:
        return ""

    text = text.replace("\r", "\n")
    replacements = {
        "：": ":",
        "＝": "=",
        "–": "-",
        "—": "-",
        "−": "-",
        "’": "'",
        "`": "'",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def clean_number(value):
    if value is None:
        return ""

    value = str(value).strip()
    value = value.replace(",", ".")
    value = value.replace("O", "0").replace("o", "0")
    value = re.sub(r"[^0-9.\-+]", "", value)

    parts = value.split(".")
    if len(parts) > 2:
        value = parts[0] + "." + "".join(parts[1:])
    return value


def add_excel_apostrophe(value):
    if value is None:
        return ""

    value = str(value).strip()
    if not value:
        return ""
    if value.startswith("'"):
        return value
    return "'" + value


# ============================================================
# CLEANING HELPERS (FOR G-NETTRACK)
# ============================================================

def clean_radio(value):
    if not value:
        return ""
    value = str(value).strip()
    value = value.replace(" ", "")
    value = value.replace(",", ".")
    value = value.replace("|", "/").replace("\\", "/").replace("l", "/").replace("I", "/")
    value = value.replace("O", "0").replace("o", "0")
    for dash in ["–", "—", "−", "~"]:
        value = value.replace(dash, "-")
    value = re.sub(r"/+", "/", value)
    value = re.sub(r"[^0-9.\-+/]", "", value)
    return value


def clean_coord(value):
    if not value:
        return ""
    value = str(value).strip()
    for dash in ["–", "—", "−", "~"]:
        value = value.replace(dash, "-")
    value = value.replace("O", "0").replace("o", "0")
    value = value.replace(",", ".")
    value = value.replace(" ", "")
    match = re.search(r"[-+]?\d+(?:\.\d+)?", value)
    return match.group(0) if match else ""


def clean_integer(value):
    if not value:
        return ""
    value = str(value).strip()
    value = value.replace(" ", "")
    value = value.replace("O", "0").replace("o", "0")
    value = value.replace("|", "/").replace("\\", "/")
    return value


# ============================================================
# SELECT OCR WORKER COUNT (FIRST STARTUP DIALOG)
# ============================================================

def select_worker_count():
    """Show the OCR_MAX_WORKERS selector immediately when the script starts."""
    global MAX_WORKERS

    choices = (4, 6, 8, 10)
    default_choice = min(choices, key=lambda value: abs(value - MAX_WORKERS))
    selected = {"value": default_choice}

    root = tk.Tk()
    root.title("OCR Performance Settings")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    width, height = 390, 185
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = max(0, (screen_width - width) // 2)
    y = max(0, (screen_height - height) // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    container = ttk.Frame(root, padding=16)
    container.pack(fill="both", expand=True)

    ttk.Label(
        container,
        text="Select OCR_MAX_WORKERS",
        font=("Segoe UI", 11, "bold")
    ).pack(anchor="w")

    ttk.Label(
        container,
        text="Choose the number of archive-processing workers before starting."
    ).pack(anchor="w", pady=(5, 10))

    variable = tk.StringVar(value=str(default_choice))
    dropdown = ttk.Combobox(
        container,
        textvariable=variable,
        values=[str(value) for value in choices],
        state="readonly",
        width=12
    )
    dropdown.pack(anchor="w")

    def continue_processing():
        try:
            selected_value = int(variable.get())
        except (TypeError, ValueError):
            selected_value = default_choice

        if selected_value not in choices:
            selected_value = default_choice

        selected["value"] = selected_value
        root.destroy()

    buttons = ttk.Frame(container)
    buttons.pack(fill="x", pady=(14, 0))
    ttk.Button(
        buttons,
        text="Continue",
        command=continue_processing
    ).pack(side="right")

    # Closing the dialog keeps the current/default worker count and continues.
    root.protocol("WM_DELETE_WINDOW", continue_processing)
    dropdown.focus_set()
    root.mainloop()

    MAX_WORKERS = int(selected["value"])
    os.environ["OCR_MAX_WORKERS"] = str(MAX_WORKERS)
    return MAX_WORKERS


# ============================================================
# SELECT FOLDER
# ============================================================

def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(
        title="Select folder containing RAR / ZIP files"
    )
    root.destroy()
    return folder


# ============================================================
# FIND 7-ZIP
# ============================================================

def find_7zip():
    possible_paths = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ]

    for path in possible_paths:
        if os.path.isfile(path):
            return path

    try:
        result = subprocess.run(
            ["where", "7z"],
            capture_output=True,
            text=True,
            shell=True
        )
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                line = line.strip()
                if os.path.isfile(line):
                    return line
    except Exception:
        pass

    return None


# ============================================================
# FIND ARCHIVES
# ============================================================

def find_archives(folder):
    archives = []

    for filename in os.listdir(folder):
        path = os.path.join(folder, filename)

        if not os.path.isfile(path):
            continue

        extension = os.path.splitext(filename)[1].lower()

        if extension in {".rar", ".zip"}:
            archives.append(filename)

    archives.sort(key=lambda x: x.lower())

    return archives


# ============================================================
# TESSERACT
# ============================================================

def configure_tesseract():
    global TESSERACT_PATH

    if pytesseract is None:
        print(
            "\nERROR: pytesseract is not installed.\n\n"
            "Run:\npython -m pip install pytesseract"
        )
        return False

    if TESSERACT_PATH:
        if os.path.isfile(TESSERACT_PATH):
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
            return True
        return False

    possible_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expandvars(
            r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"
        ),
    ]

    for path in possible_paths:
        if os.path.isfile(path):
            pytesseract.pytesseract.tesseract_cmd = path
            TESSERACT_PATH = path
            return True

    try:
        result = subprocess.run(
            ["where", "tesseract"],
            capture_output=True,
            text=True,
            shell=True
        )

        if result.returncode == 0:
            for line in result.stdout.splitlines():
                line = line.strip()

                if os.path.isfile(line):
                    pytesseract.pytesseract.tesseract_cmd = line
                    TESSERACT_PATH = line
                    return True

    except Exception:
        pass

    return False


# ============================================================
# IMAGE OCR (IN-MEMORY SUPPORT)
# ============================================================

def ocr_image_original(image):
    try:
        if image is None:
            return ""

        text = pytesseract.image_to_string(
            image,
            config="--psm 6"
        )

        return normalize_text(text)

    except Exception as e:
        print(f"    OCR error: {e}")
        return ""


def preprocess_for_ocr(image):
    if image is None:
        return None

    gray = cv2.cvtColor(
        image,
        cv2.COLOR_BGR2GRAY
    )

    # 1. Dark Mode Fix
    if np.mean(gray) < 127:
        gray = cv2.bitwise_not(gray)

    if max(gray.shape) < 1500:
        gray = cv2.resize(
            gray,
            None,
            fx=2,
            fy=2,
            interpolation=cv2.INTER_CUBIC
        )

    _, threshold = cv2.threshold(
        gray,
        0,
        255,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU
    )

    return threshold


def ocr_image_processed(image):
    try:
        if image is None:
            return ""

        processed = preprocess_for_ocr(image)

        # PSM 11 - sparse text
        text = pytesseract.image_to_string(
            processed,
            config="--psm 11"
        )

        return normalize_text(text)

    except Exception as e:
        print(f"    Processed OCR error: {e}")
        return ""


def ocr_gnet_targeted(image):
    """Read the compact G-NetTrack information panel at a legible scale.

    Phone recordings are often only about 296 px wide. Full-frame OCR at the
    native size can turn ``eNB`` into ``@NB`` and distort coordinate digits.
    This pass crops the top 64% (where the G-NetTrack radio/GPS fields reside),
    enlarges low-resolution captures, and preserves the original dark/light
    colors instead of applying a destructive global threshold.
    """
    try:
        if image is None or cv2 is None or pytesseract is None:
            return ""

        h, w = image.shape[:2]
        if h < 1 or w < 1:
            return ""

        crop = image[:max(1, int(h * 0.64)), :]
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)

        if w < 450:
            scale = 4.0
        elif w < 700:
            scale = 3.0
        elif w < 1000:
            scale = 2.0
        else:
            scale = 1.0

        if scale != 1.0:
            gray = cv2.resize(
                gray,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_CUBIC,
            )

        # PSM 6 keeps the panel's label/value pairs together and was notably
        # more accurate than native full-frame OCR for this recording.
        return normalize_text(
            pytesseract.image_to_string(
                gray,
                config="--oem 3 --psm 6",
            )
        )

    except Exception as e:
        print(f"    Targeted G-NetTrack OCR error: {e}")
        return ""


# ============================================================
# G-NET FIELD HELPER
# ============================================================

def _extract_after_label_same_line(text, label_regex, value_regex):
    """Extract a value after a label without allowing the label suffix to become the value."""
    if not text:
        return ""

    for line in normalize_text(text).splitlines():
        line = line.strip()

        if not line:
            continue

        match = re.search(
            label_regex +
            r"[^:=-]*[:=]?\s*(" +
            value_regex +
            r")",
            line,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return ""


def find_label_same_or_next_line(text, labels, value_pattern):
    """Generic fallback parser. Telecom-critical fields use specialized parsers below."""

    if not text:
        return ""

    lines = [
        line.strip()
        for line in normalize_text(text).splitlines()
        if line.strip()
    ]

    for index, line in enumerate(lines):

        for label in labels:

            if re.search(
                rf"(?:^|\s){label}(?:\s|$|[:=\[\-])",
                line,
                re.IGNORECASE
            ):

                tail = re.split(
                    label,
                    line,
                    maxsplit=1,
                    flags=re.IGNORECASE
                )[-1]

                tail = re.sub(
                    r"^\s*\[[^\]]*\]\s*[:=]?\s*",
                    " ",
                    tail
                )

                match = re.search(
                    rf"({value_pattern})",
                    tail,
                    re.IGNORECASE
                )

                if match:
                    return match.group(1).strip()

                for offset in range(1, 4):

                    if index + offset >= len(lines):
                        break

                    match = re.search(
                        rf"({value_pattern})",
                        lines[index + offset],
                        re.IGNORECASE
                    )

                    if match:
                        return match.group(1).strip()

    return ""


# ============================================================
# TELECOM-SAFE PARSING / VALIDATION (V2)
# ============================================================

RADIO_COMPONENT = (
    r"[-+~–—−]?\s*"
    r"[0-9Oo]+"
    r"(?:\s*[.,]\s*[0-9Oo]{1,2})?"
)

RADIO_PAIR = (
    rf"({RADIO_COMPONENT}"
    rf"(?:\s*[/|\\lI]\s*{RADIO_COMPONENT})?)"
)


def _normalize_radio_pair(value):
    value = clean_radio(value)

    if not value:
        return ""

    return value


def _valid_numeric_parts(value, low, high):
    if not value:
        return False

    try:
        parts = value.split("/")

        return all(
            low <= float(part) <= high
            for part in parts
        )

    except (ValueError, TypeError):
        return False


def valid_rsrp(value):
    return _valid_numeric_parts(
        _normalize_radio_pair(value),
        -160,
        -30
    )


def valid_sinr(value):
    return _valid_numeric_parts(
        _normalize_radio_pair(value),
        -40,
        50
    )


def valid_pci(value):
    if not value:
        return False

    try:
        return all(
            0 <= int(part) <= 503
            for part in value.split("/")
        )

    except (ValueError, TypeError):
        return False


def valid_mcc(value):
    return bool(
        re.fullmatch(
            r"\d{3}",
            value or ""
        )
    )


def valid_mnc(value):
    return bool(
        re.fullmatch(
            r"\d{2,3}",
            value or ""
        )
    )


def valid_coordinate(value, minimum, maximum):
    try:
        return minimum <= float(value) <= maximum

    except (ValueError, TypeError):
        return False


def _specialized_radio_value(text, field):
    """Extract RSRP/SINR while explicitly consuming [5G/4G] before the number."""

    if field == "rsrp":
        labels = r"RSRP"

    else:
        labels = (
            r"(?:SINR|SNR|S1NR|SINA|SIMR|5INR|5NR|SMR)"
        )

    pattern = (
        rf"\b{labels}\b"
        rf"\s*(?:\[[^\]]*\])?"
        rf"\s*[:=]?\s*"
        rf"{RADIO_PAIR}"
    )

    for match in re.finditer(
        pattern,
        normalize_text(text),
        re.IGNORECASE
    ):

        candidate = _normalize_radio_pair(
            match.group(1)
        )

        if (
            field == "rsrp"
            and valid_rsrp(candidate)
        ):
            return candidate

        if (
            field == "sinr"
            and valid_sinr(candidate)
        ):
            return candidate

    return ""


def _parse_gnet_confidence(data):
    checks = {
        "longitude": valid_coordinate(
            data.get("longitude"),
            -180,
            180
        ),
        "latitude": valid_coordinate(
            data.get("latitude"),
            -90,
            90
        ),
        "mcc": valid_mcc(
            data.get("mcc")
        ),
        "mnc": valid_mnc(
            data.get("mnc")
        ),
        "pci": valid_pci(
            data.get("pci")
        ),
        "rsrp": valid_rsrp(
            data.get("rsrp")
        ),
        "sinr": valid_sinr(
            data.get("sinr")
        ),
    }

    valid_count = sum(
        checks.values()
    )

    confidence = round(
        100.0 * valid_count / len(checks),
        1
    )

    return confidence, checks


def _normalize_gnet_identifier(value):
    """Normalize OCR output for integer radio identifiers without changing digit order."""
    if value is None:
        return ""
    value = str(value).strip().replace("O", "0").replace("o", "0")
    value = re.sub(r"[^0-9]", "", value)
    return value


def _extract_gnet_cid(text, serving_type="", int_raw=None):
    """Extract the value explicitly labelled CID/CI, even on an NR screen.

    Some G-NetTrack screenshots show both the LTE anchor's ``CID`` and a
    separate NR field, for example ``eNB:471384CID4 5G508``.  The NR value
    must never overwrite the explicitly labelled CID: in this example CID
    is 4, while 5G508 is a separate radio field.
    """
    normalized = normalize_text(text or "")

    # Look for explicit CID/Cell ID labels first. CID is often concatenated
    # directly to the preceding eNB number and followed immediately by digits.
    patterns = (
        r"(?<![A-Za-z0-9])CELL\s*ID\s*[:=]?\s*([0-9Oo]{1,10})(?![0-9Oo])",
        r"(?<![A-Za-z])CID\s*[:=\-]*\s*([0-9Oo]{1,10})(?![0-9Oo])",
        # Common Tesseract confusions for CID on small/dark G-NetTrack screens.
        # OCR may render the value delimiter as a sequence such as ":-4".
        r"(?<![A-Za-z])C[I1l]?B\s*[:=\-]*\s*([0-9Oo]{1,10})(?![0-9Oo])",
        r"(?<![A-Za-z0-9])CI\s*[:=]\s*([0-9Oo]{1,10})(?![0-9Oo])",
    )
    for pattern in patterns:
        match = re.search(pattern, normalized, re.IGNORECASE)
        if match:
            candidate = _normalize_gnet_identifier(match.group(1))
            if candidate:
                return candidate

    # Compatibility fallback: only ask the generic parser for CID. Do not
    # treat the independent ``5G...`` field as CID when CID itself is absent.
    if int_raw:
        fallback = find_label_same_or_next_line(normalized, [r"CID"], int_raw)
        return _normalize_gnet_identifier(fallback)
    return ""


def _extract_plmn_code(text, field):
    """Extract MCC/MNC only from a complete numeric token after its own label.

    The generic label parser used to accept the first 3 digits from a longer
    TAC (for example MCC:43660 became MCC=436). This parser requires token
    boundaries and never borrows a number from a later field when a labelled
    line already contains an invalid/overlong value.
    """
    normalized = normalize_text(text or "")
    if not normalized or field not in ("mcc", "mnc"):
        return ""

    expected_lengths = {"mcc": {3}, "mnc": {2, 3}}[field]
    label_pattern = re.compile(
        rf"(?<![A-Za-z0-9]){field}(?![A-Za-z0-9])",
        re.IGNORECASE,
    )
    numeric_pattern = re.compile(r"(?<![0-9Oo])[0-9Oo]+(?![0-9Oo])")
    other_field_pattern = re.compile(
        r"\b(?:MCC|MNC|TAC|PCI|CID|RSRP|RSRQ|SINR|SNR|ENB|GNODEB|TYPE|LONGITUDE|LATITUDE)\b",
        re.IGNORECASE,
    )

    lines = [line.strip() for line in normalized.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        for label_match in label_pattern.finditer(line):
            tail = line[label_match.end():]
            number_match = numeric_pattern.search(tail)

            if number_match:
                # If another recognized field label occurs before the number,
                # this isn't an unambiguous value for the current label.
                prefix = tail[:number_match.start()]
                if other_field_pattern.search(prefix):
                    continue

                candidate = _normalize_gnet_identifier(number_match.group(0))
                if len(candidate) in expected_lengths:
                    return candidate

                # A number exists on the labelled line but has the wrong
                # length (often the five-digit TAC). Do not scan following
                # lines and accidentally assign PCI/another field as MCC/MNC.
                continue

            # Some OCR layouts split a label and its value into separate
            # lines. Accept only a bare numeric next line, not a line that
            # contains another field label or unrelated text.
            if index + 1 < len(lines):
                next_line = lines[index + 1].strip()
                if re.fullmatch(r"[\s:=\[\]()\-]*[0-9Oo]+[\s,.;]*", next_line):
                    candidate_match = numeric_pattern.search(next_line)
                    if candidate_match:
                        candidate = _normalize_gnet_identifier(candidate_match.group(0))
                        if len(candidate) in expected_lengths:
                            return candidate

    return ""


def parse_gnet(text):
    text = normalize_text(text)

    data = {
        "longitude": "",
        "latitude": "",
        "mcc": "",
        "mnc": "",
        "serving_type": "",
        "enodebid": "",
        "pci": "",
        "cid": "",
        "rsrp": "",
        "sinr": "",
        "confidence": 0.0,
        "validation": {},
    }

    coord_raw = (
        r"[-+~–—−]?\s*"
        r"[0-9Oo]+"
        r"(?:\s*[.,]\s*[0-9Oo]+)?"
    )

    int_raw = (
        r"[0-9Oo]+"
        r"(?![0-9Oo]|\s*[.,]\s*[0-9Oo])"
    )

    pci_raw = (
        r"[0-9Oo]+"
        r"(?:\s*[/|\\]\s*[0-9Oo]+)?"
        r"(?![0-9Oo]|\s*[.,]\s*[0-9Oo])"
    )

    data["longitude"] = clean_coord(
        find_label_same_or_next_line(
            text,
            [r"longitude", r"long", r"bujur"],
            coord_raw
        )
    )

    data["latitude"] = clean_coord(
        find_label_same_or_next_line(
            text,
            [
                r"latitude",
                # Frequent OCR spellings when the thin characters are blurred.
                r"lathtude",
                r"lathude",
                r"lat(?:i|h)?tude",
                r"lat",
                r"lintang",
            ],
            coord_raw
        )
    )

    # PLMN codes use exact numeric tokens. This prevents a five-digit TAC
    # from being truncated and accidentally reported as MCC or MNC.
    data["mcc"] = _extract_plmn_code(text, "mcc")
    data["mnc"] = _extract_plmn_code(text, "mnc")

    data["serving_type"] = find_label_same_or_next_line(
        text,
        [
            r"serving\s*type",
            r"servingtype",
            r"jenis",
            r"type"
        ],
        r"[A-Za-z0-9 +/_-]+"
    ).strip()

    data["enodebid"] = clean_integer(
        find_label_same_or_next_line(
            text,
            [
                r"eNodeB\s*ID",
                r"eNodeBID",
                r"eNB",
                r"enb",
                r"gNodeB\s*ID",
                r"gNodeBID",
                # OCR often misreads the lowercase 'g' in 'gNB' as 'Q'
                # (e.g. this screenshot is recognized as "QNB: 32767").
                # Accept common OCR substitutions while keeping label matching local.
                r"[gq9]n[b8]",
                r"QNB",
                r"9NB",
                r"0NB"
            ],
            int_raw
        )
    )

    data["pci"] = clean_integer(
        find_label_same_or_next_line(
            text,
            [r"PCI"],
            pci_raw
        )
    )

    data["cid"] = _extract_gnet_cid(
        text,
        data.get("serving_type", ""),
        int_raw
    )

    data["rsrp"] = _specialized_radio_value(
        text,
        "rsrp"
    )

    data["sinr"] = _specialized_radio_value(
        text,
        "sinr"
    )

    confidence, checks = _parse_gnet_confidence(
        data
    )

    data["confidence"] = confidence
    data["validation"] = checks

    return data


# ============================================================
# G-NETTRACK DETECTION
# ============================================================

def gnet_keyword_score(text):
    text_lower = text.lower()

    keywords = [
        "rsrp",
        "sinr",
        "snr",
        "enodeb",
        "enb",
        "gnodeb",
        "gnb",
        "pci",
        "cid",
        "mcc",
        "mnc",
        "longitude",
        "latitude",
        "bujur",
        "lintang",
        "serving type",
        "servingtype",
        "jenis",
    ]

    score = 0

    for keyword in keywords:
        if keyword in text_lower:
            score += 1

    return score


# ============================================================
# SPEEDTEST DETECTION
# ============================================================

SPEED_UNIT_DETECT_REGEX = (
    r"(?<![a-z])"
    r"m\s*[b0oh]\s*"
    r"(?:it)?\s*"
    r"[p/o0]\s*[s5]"
    r"(?![a-z])"
)

SPEED_UNIT_STRIP_REGEX = (
    r"[mgk]\s*"
    r"(?:bit|b|0|o|h)?\s*"
    r"[p/o0]\s*[s5]"
)


def speedtest_keyword_score(text):
    if not text:
        return 0

    text_lower = text.lower()

    score = 0

    if (
        re.search(
            r"d[o0]wn\s*[l1i|]?\s*[o0]a[d0]",
            text_lower
        )
        or "download" in text_lower
    ):
        score += 2

    if (
        re.search(
            r"up\s*[l1i|]\s*[o0][ae][d0]",
            text_lower
        )
        or "upload" in text_lower
    ):
        score += 2

    if "unduh" in text_lower or "undu" in text_lower:
        score += 2

    if "unggah" in text_lower or "ungga" in text_lower:
        score += 2

    if "speedtest" in text_lower or "ookla" in text_lower:
        score += 2

    if re.search(r"\bping\b", text_lower) or "jitter" in text_lower:
        score += 1

    mbps_count = max(
        len(
            re.findall(
                r"m\s*b\s*p\s*s",
                text_lower,
                re.IGNORECASE
            )
        ),
        len(
            re.findall(
                SPEED_UNIT_DETECT_REGEX,
                text_lower,
                re.IGNORECASE
            )
        )
    )

    if mbps_count >= 1:
        score += 3

    if mbps_count >= 2:
        score += 3

    return score


# ============================================================
# G-NET MERGE
# ============================================================

def count_gnet_fields(data):
    score = 0

    for field, value in data.items():

        if value:
            score += 1

            if field == "sinr":
                score += 20

            elif field == "rsrp":
                score += 10

    return score


def merge_missing_gnet_fields(best, new):
    for key in best:

        if not best.get(key) and new.get(key):
            best[key] = new[key]

    return best


# ============================================================
# RADIO FORMATTING
# ============================================================

def format_radio_value(value, serving_type):
    if not value:
        return ""

    value = clean_radio(value)

    serving = (
        serving_type or ""
    ).lower()

    is_5g = (
        "5g" in serving
        or "nr" in serving
    )

    is_lte = (
        "lte" in serving
        or "4g" in serving
    )

    if "/" in value:

        parts = value.split("/")

        if len(parts) >= 2:
            return parts[0] + "/" + parts[1]

    if is_5g:
        return value + "/-"

    if is_lte:
        return "-/" + value

    return value


# ============================================================
# SPEEDTEST OCR
# ============================================================

def _is_dark_image(image):
    """True when the screenshot has a mostly dark background."""

    try:

        if image is None:
            return False

        gray = cv2.cvtColor(
            image,
            cv2.COLOR_BGR2GRAY
        )

        return float(
            np.mean(gray)
        ) < 127

    except Exception:
        return False


def ocr_speedtest_data(image, config, invert=False):
    try:

        if image is None:
            return [], 0, 0

        h_orig, w_orig = image.shape[:2]

        gray = cv2.cvtColor(
            image,
            cv2.COLOR_BGR2GRAY
        )

        if invert:
            gray = cv2.bitwise_not(gray)

        if w_orig < 600:
            scale = 2.0

        elif w_orig < 1000:
            scale = 1.5

        else:
            scale = 1.0

        if scale != 1.0:

            gray = cv2.resize(
                gray,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_CUBIC
            )

        gray = cv2.normalize(
            gray,
            None,
            0,
            255,
            cv2.NORM_MINMAX
        )

        data = pytesseract.image_to_data(
            gray,
            config=config,
            output_type=pytesseract.Output.DICT
        )

        results = []

        count = len(
            data["text"]
        )

        for i in range(count):

            raw = data["text"][i].strip()

            if not raw:
                continue

            try:
                confidence = float(
                    data["conf"][i]
                )

            except Exception:
                confidence = 0.0

            if confidence < 5:
                continue

            x = data["left"][i] / scale
            y = data["top"][i] / scale
            w = data["width"][i] / scale
            h = data["height"][i] / scale

            results.append({
                "text": raw,
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "cx": x + w / 2,
                "cy": y + h / 2,
                "confidence": confidence,
                "block_num": data["block_num"][i],
                "par_num": data["par_num"][i],
                "line_num": data["line_num"][i],
            })

        return (
            results,
            w_orig,
            h_orig
        )

    except Exception as e:
        print(
            f"    Speedtest OCR error: {e}"
        )

        return [], 0, 0


def normalize_speed_number(value):
    if not value:
        return ""

    value = str(value).strip()

    value = value.replace(",", ".")

    value = value.replace(
        "O",
        "0"
    ).replace(
        "o",
        "0"
    )

    value = re.sub(
        r"[^0-9.]",
        "",
        value
    )

    parts = value.split(".")

    if len(parts) > 2:
        value = (
            parts[0]
            + "."
            + "".join(parts[1:])
        )

    match = re.search(
        r"\d+(?:\.\d+)?",
        value
    )

    return (
        match.group(0)
        if match
        else ""
    )


def _speed_unit_multiplier(text):
    if not text:
        return False, 1.0

    raw = (
        str(text)
        .lower()
        .replace(" ", "")
    )

    raw = (
        raw.replace(
            "mbit/s",
            "mbps"
        )
        .replace(
            "mb/s",
            "mbps"
        )
        .replace(
            "gbit/s",
            "gbps"
        )
        .replace(
            "gb/s",
            "gbps"
        )
        .replace(
            "kbit/s",
            "kbps"
        )
        .replace(
            "kb/s",
            "kbps"
        )
        .replace(
            "mbp5",
            "mbps"
        )
        .replace(
            "gbp5",
            "gbps"
        )
        .replace(
            "kbp5",
            "kbps"
        )
    )

    raw = (
        raw.replace(
            "mb0s",
            "mbps"
        )
        .replace(
            "m0ps",
            "mbps"
        )
        .replace(
            "mops",
            "mbps"
        )
        .replace(
            "mbos",
            "mbps"
        )
        .replace(
            "mhps",
            "mbps"
        )
        .replace(
            "gb0s",
            "gbps"
        )
        .replace(
            "g0ps",
            "gbps"
        )
        .replace(
            "kb0s",
            "kbps"
        )
        .replace(
            "k0ps",
            "kbps"
        )
    )

    if (
        re.search(
            r"(?:^|[^a-z])gbps(?:$|[^a-z])",
            raw
        )
        or "gbps" in raw
    ):
        return True, 1000.0

    if "mbps" in raw:
        return True, 1.0

    if "kbps" in raw:
        return True, 0.001

    letters = re.sub(
        r"[^a-z]",
        "",
        raw
    )

    if any(
        x in letters
        for x in (
            "gbps",
            "gbpss"
        )
    ):
        return True, 1000.0

    if any(
        x in letters
        for x in (
            "mbps",
            "mops",
            "mbos",
            "mhps"
        )
    ):
        return True, 1.0

    if "kbps" in letters:
        return True, 0.001

    return False, 1.0


def get_speed_numeric_candidates(items):
    candidates = []

    for item in items:

        raw = item["text"].strip()

        if not raw:
            continue

        cleaned = (
            raw.lower()
            .replace(",", ".")
        )

        cleaned = re.sub(
            SPEED_UNIT_STRIP_REGEX,
            "",
            cleaned,
            flags=re.IGNORECASE
        )

        if any(
            ch in cleaned
            for ch in (
                ":",
                "%",
                "/"
            )
        ):
            continue

        if re.search(
            r"[a-z]",
            re.sub(
                r"[olisz]",
                "",
                cleaned
            )
        ):
            continue

        if (
            not re.search(
                r"\d",
                cleaned
            )
            and not re.fullmatch(
                r"[ol|]{2,}(?:\.[ol|]+)?",
                cleaned
            )
        ):
            continue

        cleaned = (
            cleaned.replace(
                "o",
                "0"
            )
            .replace(
                "l",
                "1"
            )
            .replace(
                "i",
                "1"
            )
            .replace(
                "|",
                "1"
            )
            .replace(
                "s",
                "5"
            )
            .replace(
                "z",
                "2"
            )
        )

        cleaned = re.sub(
            r"[^0-9.]",
            "",
            cleaned
        )

        if cleaned.count(".") > 1:

            first = cleaned.find(
                "."
            )

            cleaned = (
                cleaned[:first + 1]
                + cleaned[first + 1:].replace(
                    ".",
                    ""
                )
            )

        if not cleaned:
            continue

        try:
            numeric = float(cleaned)

        except Exception:
            continue

        if numeric <= 0.0 or numeric > 10000:
            continue

        candidates.append({
            "value": cleaned,
            "numeric": numeric,
            "raw": raw,
            "x": item["x"],
            "y": item["y"],
            "cx": item["cx"],
            "cy": item["cy"],
            "w": item["w"],
            "h": item["h"],
            "confidence": item["confidence"],
            "block_num": item.get("block_num", 0),
            "par_num": item.get("par_num", 0),
            "line_num": item.get("line_num", 0),
        })

    return candidates


def find_speed_mbps_items(items):
    result = []

    for item in items:

        recognized, multiplier = _speed_unit_multiplier(
            item["text"]
        )

        if recognized:

            copy = dict(item)

            copy["multiplier"] = multiplier

            result.append(copy)

    return result


def find_speed_labels(items):
    download = []
    upload = []

    for item in items:

        raw = item["text"].lower()

        text = re.sub(
            r"[^a-z]",
            "",
            raw
        )

        if any(
            kw in text
            for kw in (
                "download",
                "downioad",
                "downoad",
                "downld",
                "downlo",
                "dwnl",
                "unduh",
                "undu",
                "nduh"
            )
        ):
            download.append(item)

        if any(
            kw in text
            for kw in (
                "upload",
                "upioad",
                "uploed",
                "upld",
                "uploa",
                "uplad",
                "unggah",
                "ungga",
                "nggah",
                "ungah"
            )
        ):
            upload.append(item)

    return download, upload


def _distance(a, b):
    dx = a["cx"] - b["cx"]
    dy = a["cy"] - b["cy"]

    return (
        dx * dx
        + dy * dy
    ) ** 0.5


def _relative_label_score(
    number,
    label,
    w_orig,
    h_orig
):
    dx = (
        abs(
            number["cx"]
            - label["cx"]
        )
        / max(w_orig, 1)
    )

    dy = (
        abs(
            number["cy"]
            - label["cy"]
        )
        / max(h_orig, 1)
    )

    if dx > 0.45 or dy > 0.35:
        return -1000.0

    score = 0.0

    if dx < 0.12:
        score += 35

    elif dx < 0.22:
        score += 22

    elif dx < 0.35:
        score += 8

    if dy < 0.10:
        score += 35

    elif dy < 0.18:
        score += 22

    elif dy < 0.30:
        score += 8

    if (
        number["cy"]
        >= label["cy"]
        - 0.03 * h_orig
    ):
        score += 15

    if (
        number.get("line_num")
        == label.get("line_num")
    ):
        score += 8

    if (
        number.get("block_num")
        == label.get("block_num")
    ):
        score += 4

    return score


def _unit_score(
    number,
    unit,
    w_orig,
    h_orig
):
    dx = (
        abs(
            number["cx"]
            - unit["cx"]
        )
        / max(w_orig, 1)
    )

    dy = (
        abs(
            number["cy"]
            - unit["cy"]
        )
        / max(h_orig, 1)
    )

    if dx > 0.40 or dy > 0.25:
        return -500.0

    score = 0.0

    if dy < 0.045:
        score += 30

    elif dy < 0.09:
        score += 22

    elif dy < 0.16:
        score += 10

    if dx < 0.08:
        score += 25

    elif dx < 0.16:
        score += 16

    elif dx < 0.28:
        score += 8

    if (
        number.get("line_num")
        == unit.get("line_num")
    ):
        score += 12

    if (
        number.get("block_num")
        == unit.get("block_num")
    ):
        score += 5

    return score


def _raw_token_has_speed_unit(raw):
    recognized, multiplier = _speed_unit_multiplier(
        raw
    )

    if not recognized:
        return False, 1.0

    numeric = normalize_speed_number(
        raw
    )

    return bool(numeric), multiplier


def _score_speed_candidate(
    candidate,
    labels,
    units,
    w_orig,
    h_orig
):
    score = 0.0

    dl_score = 0.0
    ul_score = 0.0

    score += (
        min(
            candidate["confidence"],
            100.0
        )
        * 0.15
    )

    score += (
        min(
            candidate["h"],
            180.0
        )
        * 0.25
    )

    if 0.01 <= candidate["numeric"] <= 5000:
        score += 5

    raw_has_unit, raw_multiplier = _raw_token_has_speed_unit(
        candidate["raw"]
    )

    if raw_has_unit:
        score += 55

    best_unit_score = -1e9

    best_multiplier = (
        raw_multiplier
        if raw_has_unit
        else 1.0
    )

    for unit in units:

        s = _unit_score(
            candidate,
            unit,
            w_orig,
            h_orig
        )

        if s > best_unit_score:

            best_unit_score = s

            best_multiplier = unit.get(
                "multiplier",
                1.0
            )

    if (
        units
        and best_unit_score > -500
    ):
        score += best_unit_score

    for label in labels[0]:

        s = _relative_label_score(
            candidate,
            label,
            w_orig,
            h_orig
        )

        dl_score = max(
            dl_score,
            s
        )

    for label in labels[1]:

        s = _relative_label_score(
            candidate,
            label,
            w_orig,
            h_orig
        )

        ul_score = max(
            ul_score,
            s
        )

    score += max(
        dl_score,
        ul_score
    )

    return {
        "score": score,
        "download_score": dl_score,
        "upload_score": ul_score,
        "multiplier": best_multiplier,
        "has_unit": (
            raw_has_unit
            or bool(units)
        ),
    }


def _pick_speed_value(
    candidates,
    labels,
    units,
    w_orig,
    h_orig,
    label_index
):
    if not candidates:
        return None

    target_labels = labels[label_index]
    other_labels = labels[1 - label_index]

    scored = []

    for candidate in candidates:

        meta = _score_speed_candidate(
            candidate,
            labels,
            units,
            w_orig,
            h_orig
        )

        label_score = (
            meta["download_score"]
            if label_index == 0
            else meta["upload_score"]
        )

        other_score = (
            meta["upload_score"]
            if label_index == 0
            else meta["download_score"]
        )

        final_score = (
            label_score * 3.0
            + meta["score"]
        )

        if (
            target_labels
            and other_labels
            and other_score > label_score
        ):
            final_score -= 150

        if not target_labels:

            if not meta["has_unit"]:
                final_score -= 80

            else:
                final_score += 5

        scored.append(
            (
                final_score,
                candidate,
                meta
            )
        )

    scored.sort(
        key=lambda x: (
            x[0],
            x[2]["has_unit"],
            x[1]["confidence"],
            x[1]["h"]
        ),
        reverse=True
    )

    best_score, best_candidate, best_meta = scored[0]

    target_score = (
        best_meta["download_score"]
        if label_index == 0
        else best_meta["upload_score"]
    )

    if target_labels:

        if target_score < 5:
            return None

    else:

        if (
            not best_meta["has_unit"]
            or best_score < 40
        ):
            return None

    value_mbps = (
        best_candidate["numeric"]
        * best_meta["multiplier"]
    )

    if (
        value_mbps <= 0
        or value_mbps > 10000
    ):
        return None

    return {
        "candidate": best_candidate,
        "value_mbps": value_mbps,
        "score": best_score,
        "multiplier": best_meta["multiplier"],
    }


def _format_speed_mbps(value_mbps):
    if value_mbps is None:
        return ""

    if abs(
        value_mbps
        - round(value_mbps)
    ) < 1e-9:

        return str(
            int(
                round(value_mbps)
            )
        )

    return (
        f"{value_mbps:.3f}"
        .rstrip("0")
        .rstrip(".")
    )


def _format_detected_speed(number, value_mbps, multiplier=1.0):
    """Preserve a real decimal digit from OCR when the unit is already Mbps."""
    if number and abs(float(multiplier or 1.0) - 1.0) < 1e-9:
        raw = str(number.get("value", "")).strip().replace(",", ".")
        if "." in raw:
            try:
                if abs(float(raw) - float(value_mbps)) < 0.01:
                    return raw.rstrip(".")
            except (TypeError, ValueError):
                pass
    return _format_speed_mbps(value_mbps)


def parse_speedtest_ocr_pass(
    image,
    config,
    invert=False
):
    items, w_orig, h_orig = ocr_speedtest_data(
        image,
        config,
        invert
    )

    if not items or w_orig == 0:
        return {
            "download": "",
            "upload": "",
            "score": 0
        }

    candidates = get_speed_numeric_candidates(
        items
    )

    candidates = [
        c for c in candidates
        if c["cy"] > h_orig * 0.03
    ]

    units = find_speed_mbps_items(
        items
    )

    dl_labels, ul_labels = find_speed_labels(
        items
    )

    labels = (
        dl_labels,
        ul_labels
    )

    if not candidates:
        return {
            "download": "",
            "upload": "",
            "score": 0
        }

    dl = _pick_speed_value(
        candidates,
        labels,
        units,
        w_orig,
        h_orig,
        0
    )

    ul = _pick_speed_value(
        candidates,
        labels,
        units,
        w_orig,
        h_orig,
        1
    )

    if dl and ul:

        same = (
            dl["candidate"]
            is ul["candidate"]
        )

        if same:

            if dl["score"] >= ul["score"]:
                ul = None

            else:
                dl = None

        if dl and ul:

            same_position = (
                abs(
                    dl["candidate"]["cx"]
                    - ul["candidate"]["cx"]
                )
                < 0.01 * w_orig
                and
                abs(
                    dl["candidate"]["cy"]
                    - ul["candidate"]["cy"]
                )
                < 0.03 * h_orig
            )

            if same_position:

                if dl["score"] >= ul["score"]:
                    ul = None

                else:
                    dl = None

    download = (
        _format_speed_mbps(
            dl["value_mbps"]
        )
        if dl
        else ""
    )

    upload = (
        _format_speed_mbps(
            ul["value_mbps"]
        )
        if ul
        else ""
    )

    score = 0.0

    if dl:
        score += max(
            0.0,
            dl["score"]
        )

    if ul:
        score += max(
            0.0,
            ul["score"]
        )

    if dl_labels:
        score += 30

    if ul_labels:
        score += 30

    if len(units) >= 1:
        score += 20

    if len(units) >= 2:
        score += 20

    return {
        "download": download,
        "upload": upload,
        "score": score
    }


def _speed_pair_quality(
    download,
    upload
):
    if not download and not upload:
        return -1000

    values = []

    for value in (
        download,
        upload
    ):

        if value:

            try:
                values.append(
                    float(value)
                )

            except Exception:
                pass

    if not values:
        return -1000

    score = 0.0

    if all(
        0 < v <= 10000
        for v in values
    ):
        score += 10

    if download and upload:

        score += 40

        dl = float(download)
        ul = float(upload)

        ratio = (
            max(dl, ul)
            / max(
                min(dl, ul),
                0.001
            )
        )

        if ratio <= 100:
            score += 10

        elif ratio <= 1000:
            score += 2

    return score


# ============================================================
# SPEEDTEST CARD-LEVEL OCR
# ============================================================

def _speed_card_crop_safe(
    image,
    rect,
    pad=0
):
    h, w = image.shape[:2]

    x, y, rw, rh = rect

    x1 = max(
        0,
        int(x - pad)
    )

    y1 = max(
        0,
        int(y - pad)
    )

    x2 = min(
        w,
        int(x + rw + pad)
    )

    y2 = min(
        h,
        int(y + rh + pad)
    )

    return image[
        y1:y2,
        x1:x2
    ]


def _detect_speedtest_metric_cards(image):
    if image is None:
        return []

    h, w = image.shape[:2]

    gray = cv2.cvtColor(
        image,
        cv2.COLOR_BGR2GRAY
    )

    y0 = int(
        h * 0.04
    )

    y1 = int(
        h * 0.42
    )

    top = gray[
        y0:y1,
        :
    ]

    all_rects = []

    for low, high in (
        (25, 90),
        (40, 120),
        (60, 160),
        (80, 200)
    ):

        edges = cv2.Canny(
            top,
            low,
            high
        )

        kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT,
            (3, 3)
        )

        closed = cv2.morphologyEx(
            edges,
            cv2.MORPH_CLOSE,
            kernel,
            iterations=1
        )

        closed = cv2.dilate(
            closed,
            cv2.getStructuringElement(
                cv2.MORPH_RECT,
                (3, 3)
            ),
            iterations=1
        )

        contours, _ = cv2.findContours(
            closed,
            cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE
        )

        for contour in contours:

            x, y, rw, rh = cv2.boundingRect(
                contour
            )

            y += y0

            if (
                rw < w * 0.20
                or rw > w * 0.96
            ):
                continue

            if (
                rh < h * 0.055
                or rh > h * 0.22
            ):
                continue

            aspect = (
                rw
                / max(rh, 1)
            )

            if (
                aspect < 1.35
                or aspect > 4.2
            ):
                continue

            if (
                y < h * 0.06
                or y > h * 0.38
            ):
                continue

            all_rects.append(
                (
                    x,
                    y,
                    rw,
                    rh
                )
            )

    rects = []

    for rect in sorted(
        all_rects,
        key=lambda r: r[2] * r[3],
        reverse=True
    ):

        x, y, rw, rh = rect

        duplicate = False

        for x2, y2, w2, h2 in rects:

            ix1 = max(
                x,
                x2
            )

            iy1 = max(
                y,
                y2
            )

            ix2 = min(
                x + rw,
                x2 + w2
            )

            iy2 = min(
                y + rh,
                y2 + h2
            )

            iw = max(
                0,
                ix2 - ix1
            )

            ih = max(
                0,
                iy2 - iy1
            )

            inter = (
                iw * ih
            )

            union = (
                rw * rh
                + w2 * h2
                - inter
            )

            iou = (
                inter / union
                if union
                else 0
            )

            if (
                iou > 0.75
                or (
                    abs(x - x2) < 8
                    and abs(y - y2) < 8
                    and abs(rw - w2) < 15
                    and abs(rh - h2) < 15
                )
            ):
                duplicate = True
                break

        if not duplicate:
            rects.append(
                rect
            )

    if not rects:
        return []

    pair_candidates = []

    for i in range(
        len(rects)
    ):

        for j in range(
            i + 1,
            len(rects)
        ):

            a = rects[i]
            b = rects[j]

            x1, y1_, w1, h1 = a
            x2, y2_, w2, h2 = b

            size_ratio = (
                min(
                    w1 * h1,
                    w2 * h2
                )
                / max(
                    w1 * h1,
                    w2 * h2
                )
            )

            if size_ratio < 0.65:
                continue

            y_diff = (
                abs(
                    y1_ - y2_
                )
                / max(
                    h,
                    1
                )
            )

            h_diff = (
                abs(
                    h1 - h2
                )
                / max(
                    h,
                    1
                )
            )

            horizontal = (
                abs(
                    (x1 + w1 / 2)
                    - (x2 + w2 / 2)
                )
                > min(w1, w2) * 0.45
            )

            score = 0.0

            score += (
                size_ratio * 50
            )

            score += max(
                0.0,
                35 - y_diff * 500
            )

            score += max(
                0.0,
                20 - h_diff * 500
            )

            if horizontal:
                score += 25

            pair_candidates.append(
                (
                    score,
                    a,
                    b
                )
            )

    if pair_candidates:

        _, a, b = max(
            pair_candidates,
            key=lambda z: z[0]
        )

        return sorted(
            [a, b],
            key=lambda r: (
                r[1],
                r[0]
            )
        )

    # ========================================================
    # FIX:
    # If the pair heuristic fails, keep TWO likely cards
    # rather than only the largest rectangle.
    # ========================================================

    if len(rects) >= 2:

        selected = sorted(
            rects,
            key=lambda r: r[2] * r[3],
            reverse=True
        )[:2]

        return sorted(
            selected,
            key=lambda r: (
                r[1],
                r[0]
            )
        )

    best = max(
        rects,
        key=lambda r: r[2] * r[3]
    )

    return [best]


def _speed_digit_vote(
    gray,
    component,
    fast=False
):
    cx, cy, cw, ch, _ = component

    y1 = max(
        0,
        cy - 4
    )

    y2 = min(
        gray.shape[0],
        cy + ch + 4
    )

    x1 = max(
        0,
        cx - 4
    )

    x2 = min(
        gray.shape[1],
        cx + cw + 4
    )

    roi = gray[
        y1:y2,
        x1:x2
    ]

    if roi.size == 0:
        return "", 0.0

    if fast:
        # The focused card crop already isolates large, high-contrast
        # digits. Two variants are usually enough and avoid dozens of
        # repeated Tesseract launches per value.
        variants = [
            ("gray", roi),
            (
                "thr140",
                cv2.threshold(
                    roi,
                    140,
                    255,
                    cv2.THRESH_BINARY
                )[1]
            ),
        ]
        psms = (8,)
    else:
        variants = [
            ("gray", roi)
        ]
        psms = (8, 13)

    for threshold in (() if fast else (
        100,
        120,
        140,
        160,
        180,
        200
    )):

        variants.append(
            (
                f"thr{threshold}",
                cv2.threshold(
                    roi,
                    threshold,
                    255,
                    cv2.THRESH_BINARY
                )[1]
            )
        )

    votes = []

    for _, variant in variants:

        padded = cv2.copyMakeBorder(
            variant,
            8,
            8,
            8,
            8,
            cv2.BORDER_CONSTANT,
            value=0
        )

        padded = cv2.resize(
            padded,
            None,
            fx=8,
            fy=8,
            interpolation=cv2.INTER_CUBIC
        )

        for psm in psms:

            text = pytesseract.image_to_string(
                padded,
                config=(
                    f"--psm {psm} "
                    f"-c tessedit_char_whitelist=0123456789"
                )
            ).strip()

            if re.fullmatch(
                r"\d",
                text
            ):
                votes.append(text)

    if not votes:
        return "", 0.0

    counts = Counter(
        votes
    )

    digit, count = counts.most_common(
        1
    )[0]

    confidence = (
        count
        / len(votes)
    )

    return digit, confidence


# ============================================================
# NEW: SPEEDTEST DECIMAL DETECTION
# ============================================================

def _normalize_speed_ocr_text(text):
    if not text:
        return ""

    text = str(text).strip()

    replacements = {
        "O": "0",
        "o": "0",
        "Q": "0",
        "D": "0",
        "l": "1",
        "I": "1",
        "|": "1",
        "i": "1",
        "S": "5",
        "s": "5",
        "Z": "2",
        "z": "2",
        ",": ".",
    }

    for old, new in replacements.items():
        text = text.replace(
            old,
            new
        )

    text = re.sub(
        r"\s+",
        "",
        text
    )

    return text


def _speed_ocr_decimal_votes(
    gray,
    expected_digit_count,
    fast=False
):
    """
    Detect decimal position by OCRing ONLY the large number region.

    Returns:
        Counter where key = decimal position after digit index.
        Example:
            87.42 -> index 2
            123.4 -> index 3
    """

    votes = Counter()

    if gray is None or gray.size == 0:
        return votes

    h, w = gray.shape[:2]

    # Number normally occupies the central/lower area of the card.
    x1 = max(
        0,
        int(w * 0.16)
    )

    x2 = min(
        w,
        int(w * 0.98)
    )

    y1 = max(
        0,
        int(h * 0.18)
    )

    y2 = min(
        h,
        int(h * 0.82)
    )

    roi = gray[
        y1:y2,
        x1:x2
    ]

    if roi.size == 0:
        return votes

    if fast:
        # For an already isolated result-card crop, use a small OCR vote set.
        # Difficult values fall through to the full card parser afterward.
        variants = [
            roi,
            cv2.threshold(roi, 140, 255, cv2.THRESH_BINARY)[1],
            cv2.threshold(roi, 140, 255, cv2.THRESH_BINARY_INV)[1],
        ]
        psms = (7, 8)
    else:
        variants = [roi]
        for threshold in (100, 120, 140, 160, 180, 200):
            variants.append(
                cv2.threshold(
                    roi, threshold, 255, cv2.THRESH_BINARY
                )[1]
            )
        for threshold in (100, 140, 180):
            variants.append(
                cv2.threshold(
                    roi, threshold, 255, cv2.THRESH_BINARY_INV
                )[1]
            )
        psms = (6, 7, 8, 13)

    for variant in variants:

        enlarged = cv2.resize(
            variant,
            None,
            fx=5,
            fy=5,
            interpolation=cv2.INTER_CUBIC
        )

        for psm in psms:

            try:
                text = pytesseract.image_to_string(
                    enlarged,
                    config=(
                        f"--psm {psm} "
                        f"-c tessedit_char_whitelist=0123456789.,"
                    )
                )

            except Exception:
                continue

            text = _normalize_speed_ocr_text(
                text
            )

            if not text:
                continue

            matches = re.findall(
                r"(\d{1,6})[.,](\d{1,4})",
                text
            )

            for left, right in matches:

                total_digits = (
                    len(left)
                    + len(right)
                )

                decimal_index = len(left)

                if (
                    total_digits
                    == expected_digit_count
                ):
                    score = 1.0

                    # Two decimal places are common in Speedtest
                    # values such as 87.42 / 18.31.
                    if len(right) == 2:
                        score += 0.5

                    votes[
                        decimal_index
                    ] += score

                elif abs(
                    total_digits
                    - expected_digit_count
                ) == 1:

                    # Weaker evidence when OCR missed/added
                    # one digit.
                    votes[
                        decimal_index
                    ] += 0.25

    return votes


def _speed_visual_decimal_votes(
    gray,
    components
):
    """
    Detect decimal position visually.

    The decimal is normally:
      - between two digit components
      - much smaller than a digit
      - positioned in the lower portion of the digit line

    One vote is added per threshold variant so that the
    same decimal detected repeatedly becomes strong evidence.
    """

    votes = Counter()

    if (
        gray is None
        or gray.size == 0
        or len(components) < 2
    ):
        return votes

    median_h = float(
        np.median(
            [
                c[3]
                for c in components
            ]
        )
    )

    median_w = float(
        np.median(
            [
                c[2]
                for c in components
            ]
        )
    )

    if median_h <= 0:
        return votes

    thresholds = (
        90,
        110,
        130,
        150,
        170,
        190,
        210
    )

    for threshold in thresholds:

        bw_normal = cv2.threshold(
            gray,
            threshold,
            255,
            cv2.THRESH_BINARY
        )[1]

        bw_inverse = cv2.threshold(
            gray,
            threshold,
            255,
            cv2.THRESH_BINARY_INV
        )[1]

        threshold_best = {}

        for bw in (
            bw_normal,
            bw_inverse
        ):

            count, _, stats, _ = cv2.connectedComponentsWithStats(
                bw,
                8
            )

            small_components = []

            for i in range(
                1,
                count
            ):

                sx, sy, sw, sh, area = stats[i]

                if area < 4 or area > 180:
                    continue

                if sw < 1 or sh < 1:
                    continue

                # A decimal point should be noticeably smaller
                # than the main digit components.
                if sw > max(
                    12,
                    median_w * 1.20
                ):
                    continue

                if sh > max(
                    20,
                    median_h * 0.60
                ):
                    continue

                small_components.append(
                    (
                        sx,
                        sy,
                        sw,
                        sh,
                        area
                    )
                )

            if not small_components:
                continue

            for index in range(
                len(components) - 1
            ):

                left_digit = components[
                    index
                ]

                right_digit = components[
                    index + 1
                ]

                gap_left = (
                    left_digit[0]
                    + left_digit[2]
                )

                gap_right = right_digit[0]

                gap_center = (
                    gap_left
                    + gap_right
                ) / 2.0

                digit_top = min(
                    left_digit[1],
                    right_digit[1]
                )

                digit_bottom = max(
                    left_digit[1]
                    + left_digit[3],
                    right_digit[1]
                    + right_digit[3]
                )

                best_local_score = 0.0

                for small in small_components:

                    sx, sy, sw, sh, area = small

                    scx = (
                        sx
                        + sw / 2.0
                    )

                    scy = (
                        sy
                        + sh / 2.0
                    )

                    horizontal_distance = abs(
                        scx
                        - gap_center
                    )

                    if horizontal_distance > max(
                        8.0,
                        median_w * 0.85
                    ):
                        continue

                    # Decimal usually sits in the lower part of
                    # the digit line.
                    relative_y = (
                        scy
                        - digit_top
                    ) / max(
                        median_h,
                        1.0
                    )

                    if (
                        relative_y < 0.48
                        or relative_y > 1.08
                    ):
                        continue

                    if (
                        scy < digit_top
                        or scy > digit_bottom + 0.20 * median_h
                    ):
                        continue

                    score = 0.0

                    # Horizontal alignment.
                    if horizontal_distance <= 2:
                        score += 3.0

                    elif horizontal_distance <= 4:
                        score += 2.0

                    elif horizontal_distance <= 7:
                        score += 1.0

                    # Vertical placement.
                    vertical_distance = abs(
                        relative_y - 0.76
                    )

                    if vertical_distance < 0.10:
                        score += 3.0

                    elif vertical_distance < 0.20:
                        score += 2.0

                    elif vertical_distance < 0.32:
                        score += 1.0

                    # Size.
                    size_ratio = (
                        sh
                        / max(
                            median_h,
                            1.0
                        )
                    )

                    if 0.08 <= size_ratio <= 0.35:
                        score += 2.0

                    elif 0.35 < size_ratio <= 0.55:
                        score += 1.0

                    # A decimal should normally fit between the
                    # two digits.
                    if (
                        gap_left - 3
                        <= scx
                        <= gap_right + 3
                    ):
                        score += 2.0

                    if score > best_local_score:
                        best_local_score = score

                if best_local_score >= 5.0:
                    # `index` identifies the gap between digit[index]
                    # and digit[index + 1].  A decimal position is
                    # counted AFTER the digit to its left, so the
                    # correct decimal index is index + 1.
                    decimal_position = index + 1
                    threshold_best[
                        decimal_position
                    ] = best_local_score

        # One vote per threshold for each gap.
        # This prevents normal/inverse threshold pairs from
        # artificially doubling the evidence.
        for index, score in threshold_best.items():

            if score >= 7.0:
                votes[index] += 1.5

            elif score >= 5.0:
                votes[index] += 1.0

    return votes


def _detect_speed_decimal_index(
    gray,
    components,
    fast=False
):
    """
    Combines OCR and visual evidence to determine where the
    decimal separator belongs.

    Returns:
        int decimal position, or None
    """

    if (
        gray is None
        or not components
        or len(components) < 2
    ):
        return None

    expected_digit_count = len(
        components
    )

    visual_votes = _speed_visual_decimal_votes(
        gray,
        components
    )

    # On a tightly cropped large-font card, visual punctuation geometry
    # is fast and reliable. Only run the much heavier OCR decimal voting
    # when the visual evidence is not decisive.
    if fast and visual_votes:
        fast_visual_best, fast_visual_score = visual_votes.most_common(1)[0]
        if fast_visual_score >= 3.0:
            return int(fast_visual_best)

    ocr_votes = _speed_ocr_decimal_votes(
        gray,
        expected_digit_count,
        fast=fast
    )

    if (
        not ocr_votes
        and not visual_votes
    ):
        return None

    ocr_best = None
    ocr_score = 0.0

    if ocr_votes:

        ocr_best, ocr_score = ocr_votes.most_common(
            1
        )[0]

    visual_best = None
    visual_score = 0.0

    if visual_votes:

        visual_best, visual_score = visual_votes.most_common(
            1
        )[0]

    # ========================================================
    # Strongest case:
    # OCR and visual geometry agree.
    # ========================================================
    if (
        ocr_best is not None
        and visual_best is not None
        and ocr_best == visual_best
    ):
        if (
            ocr_score >= 1.5
            or visual_score >= 2
        ):
            return int(
                ocr_best
            )

    # ========================================================
    # OCR consensus is strong.
    # ========================================================
    if (
        ocr_best is not None
        and ocr_score >= 2.0
    ):

        # If visual strongly contradicts it, only switch when
        # visual evidence is clearly stronger.
        if (
            visual_best is not None
            and visual_best != ocr_best
            and visual_score >= ocr_score + 2.0
            and visual_score >= 4
        ):
            return int(
                visual_best
            )

        return int(
            ocr_best
        )

    # ========================================================
    # Visual consensus fallback.
    # ========================================================
    if (
        visual_best is not None
        and visual_score >= 3.0
    ):
        return int(
            visual_best
        )

    # ========================================================
    # Last fallback: whichever evidence is available.
    # ========================================================
    if ocr_best is not None:
        return int(
            ocr_best
        )

    if visual_best is not None:
        return int(
            visual_best
        )

    return None


def _speed_card_number(
    image,
    card,
    fast=False
):
    x, y, w, h = card

    card_img = _speed_card_crop_safe(
        image,
        card
    )

    if (
        card_img is None
        or card_img.size == 0
    ):
        return None

    gray = cv2.cvtColor(
        card_img,
        cv2.COLOR_BGR2GRAY
    )

    groups = []

    for threshold in (
        100,
        120,
        140,
        160,
        180,
        200
    ):

        bw = cv2.threshold(
            gray,
            threshold,
            255,
            cv2.THRESH_BINARY
        )[1]

        count, labels, stats, _ = cv2.connectedComponentsWithStats(
            bw,
            8
        )

        components = []

        for i in range(
            1,
            count
        ):

            cx, cy, cw, ch, area = stats[i]

            if area < 18:
                continue

            if (
                ch < max(
                    16,
                    h * 0.16
                )
                or
                ch > h * 0.55
            ):
                continue

            if (
                cw < max(
                    3,
                    w * 0.022
                )
                or
                cw > w * 0.12
            ):
                continue

            if (
                cy < h * 0.23
                or
                cy > h * 0.76
            ):
                continue

            if (
                cx + cw / 2
                < w * 0.24
            ):
                continue

            components.append(
                (
                    cx,
                    cy,
                    cw,
                    ch,
                    area
                )
            )

        components.sort(
            key=lambda c: c[0]
        )

        if not components:
            continue

        selected = []

        median_h = float(
            np.median(
                [
                    c[3]
                    for c in components
                ]
            )
        )

        anchor = max(
            components,
            key=lambda c: c[3]
        )

        for c in components:

            if (
                c[3]
                < median_h * 0.78
            ):
                continue

            if (
                abs(
                    c[1]
                    - anchor[1]
                )
                > h * 0.07
            ):
                continue

            selected.append(
                c
            )

        selected.sort(
            key=lambda c: c[0]
        )

        if len(selected) > 6:

            best_window = None

            for start in range(
                len(selected) - 5
            ):

                window = selected[
                    start:start + 6
                ]

                span = (
                    window[-1][0]
                    - window[0][0]
                )

                if (
                    best_window is None
                    or
                    span > best_window[0]
                ):
                    best_window = (
                        span,
                        window
                    )

            selected = (
                list(
                    best_window[1]
                )
                if best_window
                else selected[:6]
            )

        if selected:

            heights = [
                c[3]
                for c in selected
            ]

            height_cv = (
                np.std(heights)
                / max(
                    np.mean(heights),
                    1.0
                )
            ) if heights else 99

            score = (
                min(
                    len(selected),
                    6
                ) * 100
                - height_cv * 100
                - abs(
                    threshold - 140
                ) * 0.15
            )

            groups.append(
                (
                    score,
                    threshold,
                    selected,
                    bw
                )
            )

    if not groups:
        return None

    _, threshold, components, bw = max(
        groups,
        key=lambda g: g[0]
    )

    digits = []
    digit_confidences = []

    for component in components:

        digit, confidence = _speed_digit_vote(
            gray,
            component,
            fast=fast
        )

        if not digit:
            return None

        digits.append(
            digit
        )

        digit_confidences.append(
            confidence
        )

    # ========================================================
    # IMPROVED DECIMAL DETECTION
    #
    # Previously the code relied almost entirely on tiny
    # connected components between digits.
    #
    # Now:
    #   1. OCR votes on decimal punctuation.
    #   2. Visual threshold votes locate a small dot.
    #   3. Both are combined.
    # ========================================================

    decimal_index = _detect_speed_decimal_index(
        gray,
        components,
        fast=fast
    )

    value = ""

    for index, digit in enumerate(
        digits
    ):

        value += digit

        if (
            decimal_index is not None
            and
            index + 1
            == decimal_index
        ):
            value += "."

    # ========================================================
    # SAFETY FALLBACK
    #
    # Preserve the old visual separator detection as a weak
    # fallback only when the new detector found nothing.
    # ========================================================

    if decimal_index is None:

        small_components = []

        count, labels, stats, _ = cv2.connectedComponentsWithStats(
            bw,
            8
        )

        for i in range(
            1,
            count
        ):

            cx, cy, cw, ch, area = stats[i]

            if not (
                1 <= cw <= 10
                and 3 <= ch <= 16
            ):
                continue

            if (
                area < 7
                or area > 130
            ):
                continue

            small_components.append(
                (
                    cx,
                    cy,
                    cw,
                    ch,
                    area
                )
            )

        separator_indices = set()

        if len(components) >= 2:

            for index in range(
                len(components) - 1
            ):

                a = components[index]
                b = components[index + 1]

                gap_left = (
                    a[0]
                    + a[2]
                )

                gap_right = b[0]

                digit_bottom = max(
                    a[1] + a[3],
                    b[1] + b[3]
                )

                digit_top = min(
                    a[1],
                    b[1]
                )

                candidates = []

                for small in small_components:

                    scx, scy, scw, sch, sarea = small

                    if (
                        scx < gap_left - 3
                        or
                        scx > gap_right + 3
                    ):
                        continue

                    if (
                        scy
                        < digit_top + a[3] * 0.50
                        or
                        scy
                        > digit_bottom + a[3] * 0.25
                    ):
                        continue

                    candidates.append(
                        small
                    )

                if candidates:

                    mid = (
                        gap_left
                        + gap_right
                    ) / 2

                    candidates.sort(
                        key=lambda s:
                        abs(
                            s[0]
                            - mid
                        )
                    )

                    separator_indices.add(
                        index
                    )

        if separator_indices:

            decimal_index = (
                min(separator_indices)
                + 1
            )

            value = ""

            for index, digit in enumerate(
                digits
            ):

                value += digit

                if (
                    decimal_index is not None
                    and
                    index + 1
                    == decimal_index
                ):
                    value += "."

    if not value:
        return None

    numeric_value = normalize_speed_number(
        value
    )

    if not numeric_value:
        return None

    try:
        numeric = float(
            numeric_value
        )

    except Exception:
        return None

    confidence = float(
        np.mean(
            digit_confidences
        )
    )

    return {
        "value": value,
        "numeric": numeric,
        "confidence": confidence,
        "digit_count": len(digits),
        "threshold": threshold,
        "components": components,
        "decimal_index": decimal_index,
    }


def _speed_card_unit(
    card_image
):
    if (
        card_image is None
        or card_image.size == 0
    ):
        return False, 1.0

    h, w = card_image.shape[:2]

    gray = cv2.cvtColor(
        card_image,
        cv2.COLOR_BGR2GRAY
    )

    roi = gray[
        int(h * 0.67):int(h * 0.98),
        int(w * 0.35):w
    ]

    variants = [
        roi,
        cv2.resize(
            roi,
            None,
            fx=5,
            fy=5,
            interpolation=cv2.INTER_CUBIC
        )
    ]

    texts = []

    for variant in variants:

        if variant is roi:

            variant = cv2.resize(
                variant,
                None,
                fx=5,
                fy=5,
                interpolation=cv2.INTER_CUBIC
            )

        for psm in (
            6,
            7,
            11,
            13
        ):

            text = pytesseract.image_to_string(
                variant,
                config=f"--psm {psm}"
            ).strip()

            if text:
                texts.append(
                    text
                )

            recognized, multiplier = _speed_unit_multiplier(
                text
            )

            if recognized:
                return True, multiplier

    joined = (
        " ".join(texts)
        .lower()
        .replace(" ", "")
    )

    if (
        "gbps" in joined
        or "gb0s" in joined
        or "gops" in joined
    ):
        return True, 1000.0

    if (
        "kbps" in joined
        or "kb0s" in joined
        or "kops" in joined
    ):
        return True, 0.001

    if any(
        k in joined
        for k in (
            "mbps",
            "mops",
            "mbos",
            "mhps",
            "mbp"
        )
    ):
        return True, 1.0

    return False, 1.0


def _speed_card_label(
    card_image
):
    if (
        card_image is None
        or card_image.size == 0
    ):
        return ""

    h, w = card_image.shape[:2]

    gray = cv2.cvtColor(
        card_image,
        cv2.COLOR_BGR2GRAY
    )

    roi = gray[
        int(h * 0.04):int(h * 0.38),
        :int(w * 0.95)
    ]

    roi = cv2.resize(
        roi,
        None,
        fx=5,
        fy=5,
        interpolation=cv2.INTER_CUBIC
    )

    text_candidates = []

    for psm in (
        6,
        7,
        11,
        12
    ):

        text = pytesseract.image_to_string(
            roi,
            config=f"--psm {psm}"
        ).strip().lower()

        if text:
            text_candidates.append(
                re.sub(
                    r"[^a-z]",
                    "",
                    text
                )
            )

    joined = " ".join(
        text_candidates
    )

    if any(
        key in joined
        for key in (
            "download",
            "downioad",
            "downoad",
            "downlo",
            "dwnl",
            "unduh",
            "undu",
            "nduh"
        )
    ):
        return "download"

    if any(
        key in joined
        for key in (
            "upload",
            "upioad",
            "uploed",
            "uploa",
            "uplad",
            "unggah",
            "ungga",
            "nggah",
            "ungah"
        )
    ):
        return "upload"

    if any(
        key in joined
        for key in (
            "ping",
            "pinq",
            "jitter",
            "jiter",
            "latency",
            "laten",
            "packet",
            "loss"
        )
    ):
        return "other"

    return ""


def _speed_card_ocr_number(card_image, visual_seed=None):
    """OCR just the large number area of one metric card as a rescue path.

    This deliberately runs only when connected-component recognition is empty
    or collapses a likely speed into one digit (e.g. 7). Multiple PSM results
    must agree on the digit sequence so unrelated small numbers do not win.
    """
    if card_image is None or card_image.size == 0:
        return None

    try:
        h, w = card_image.shape[:2]
        gray = cv2.cvtColor(card_image, cv2.COLOR_BGR2GRAY)
        # Use two slightly different tight windows: the decimal mark in
        # Speedtest's thin font can disappear when the crop is too wide, while
        # a tighter crop may trim the first digit on some screen scales.
        roi_specs = (
            (0.25, 0.78, 0.20, 0.97),
            (0.28, 0.74, 0.30, 1.00),
        )
        candidates = []
        for top_frac, bottom_frac, left_frac, right_frac in roi_specs:
            roi = gray[
                int(h * top_frac):int(h * bottom_frac),
                int(w * left_frac):int(w * right_frac),
            ]
            if roi.size == 0:
                continue
            roi = cv2.resize(roi, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
            variants = [roi]
            try:
                _, otsu = cv2.threshold(roi, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                variants.append(otsu)
            except Exception:
                pass

            for variant in variants:
                for psm in (7, 13, 6):
                    try:
                        raw = pytesseract.image_to_string(
                            variant,
                            config=(
                                f"--psm {psm} "
                                "-c tessedit_char_whitelist=0123456789.,"
                            ),
                        ).strip()
                    except Exception:
                        continue
                    if not raw:
                        continue
                    raw = raw.replace(",", ".")
                    match = re.search(r"(?<!\d)(\d+(?:\.\d+)?)(?!\d)", raw)
                    if not match:
                        continue
                    token = match.group(1)
                    digits = re.sub(r"\D", "", token)
                    if not digits:
                        continue
                    try:
                        numeric = float(token)
                    except ValueError:
                        continue
                    if not (0 < numeric <= 10000):
                        continue
                    decimal_index = len(token.split(".", 1)[0]) if "." in token else None
                    candidates.append({
                        "token": token,
                        "digits": digits,
                        "numeric": numeric,
                        "decimal_index": decimal_index,
                    })

        if not candidates:
            return None

        # Vote first on the exact sequence of digits, then on punctuation.
        from collections import Counter as _Counter
        digit_counts = _Counter(item["digits"] for item in candidates)
        best_digits, best_count = digit_counts.most_common(1)[0]
        if best_count < 2:
            return None
        matching = [item for item in candidates if item["digits"] == best_digits]
        decimal_votes = _Counter(
            item["decimal_index"] for item in matching
            if item["decimal_index"] is not None
        )
        decimal_index = decimal_votes.most_common(1)[0][0] if decimal_votes else None

        # If OCR drops the tiny comma/dot but the component-based reader saw
        # its position on the exact same digit sequence, keep that evidence.
        if decimal_index is None and visual_seed:
            seed_value = str(visual_seed.get("value", ""))
            seed_digits = re.sub(r"\D", "", seed_value)
            seed_decimal = visual_seed.get("decimal_index")
            if seed_digits == best_digits and seed_decimal is not None:
                decimal_index = seed_decimal
            elif len(seed_digits) <= 1 and len(best_digits) == 3:
                # If segmentation collapsed a large 3-digit display to one
                # digit, the full-card OCR often sees the digits but loses the
                # faint decimal. For this Speedtest card style, make the best
                # guess at a trailing decimal instead of returning 595/110.
                decimal_index = 2

        # Speedtest result cards commonly show one decimal place. Use the
        # detected punctuation when available; otherwise the targeted fallback
        # above applies only when segmentation clearly lost most of the value.
        if decimal_index is not None and 0 < decimal_index < len(best_digits):
            token = best_digits[:decimal_index] + "." + best_digits[decimal_index:]
        else:
            token = best_digits
        try:
            numeric = float(token)
        except ValueError:
            return None
        if not (0 < numeric <= 10000):
            return None
        return {
            "value": token,
            "numeric": numeric,
            "confidence": min(1.0, best_count / 3.0),
            "digit_count": len(best_digits),
            "decimal_index": decimal_index,
            "ocr_votes": best_count,
            "source": "focused-card-ocr",
        }
    except Exception:
        return None


def _speedtest_cards_are_side_by_side(cards):
    """Avoid interpreting vertically stacked or unrelated rectangles as DL/UL cards."""
    if not cards or len(cards) < 2:
        return False
    a, b = cards[0], cards[1]
    x1, y1, w1, h1 = a
    x2, y2, w2, h2 = b
    center_y1 = y1 + h1 / 2.0
    center_y2 = y2 + h2 / 2.0
    center_x1 = x1 + w1 / 2.0
    center_x2 = x2 + w2 / 2.0
    return (
        abs(center_y1 - center_y2) <= max(h1, h2) * 0.50
        and abs(center_x1 - center_x2) >= min(w1, w2) * 0.60
    )


def _parse_speedtest_cards(
    image,
    fast=False,
    cards=None
):
    if image is None:
        return None

    if cards is None:
        cards = _detect_speedtest_metric_cards(image)

    if len(cards) < 2 or not _speedtest_cards_are_side_by_side(cards):
        return None

    parsed = []

    for card in cards[:4]:

        card_img = _speed_card_crop_safe(
            image,
            card
        )

        label = _speed_card_label(
            card_img
        )

        if label == "other":
            continue

        number = _speed_card_number(
            image,
            card,
            fast=fast
        )

        # A one-digit result is often an OCR/segmentation failure (e.g. the
        # speed is read as 7 while the large glyphs actually show 59,5/11,0).
        # Read only the large value region and require a digit-sequence vote.
        if (
            not number
            or number.get("digit_count", 0) <= 1
            or float(number.get("numeric", 0.0) or 0.0) < 10.0
        ):
            focused_number = _speed_card_ocr_number(card_img, visual_seed=number)
            if focused_number and (
                not number
                or focused_number.get("digit_count", 0) > number.get("digit_count", 0)
            ):
                number = focused_number

        if not number:
            continue

        recognized_unit, multiplier = _speed_card_unit(
            card_img
        )

        value_mbps = (
            number["numeric"]
            * multiplier
        )

        if (
            value_mbps <= 0
            or value_mbps > 10000
        ):
            continue

        parsed.append({
            "card": card,
            "label": label,
            "value": number["value"],
            "numeric": value_mbps,
            "multiplier": multiplier,
            "unit_detected": recognized_unit,
            "confidence": number["confidence"],
            "decimal_index": number.get(
                "decimal_index"
            ),
        })

    if len(parsed) < 2:
        return None

    download = None
    upload = None

    for item in parsed:

        if (
            item["label"] == "download"
            and download is None
        ):
            download = item

        elif (
            item["label"] == "upload"
            and upload is None
        ):
            upload = item

    if (
        download is None
        or upload is None
    ):

        if (
            download is None
            and upload is None
        ):

            # No label could be read:
            # left card = download
            # right card = upload
            ordered = sorted(
                parsed,
                key=lambda item: (
                    item["card"][1],
                    item["card"][0]
                )
            )

            if len(ordered) >= 2:

                first, second = sorted(
                    ordered[:2],
                    key=lambda item:
                    item["card"][0]
                )

                download = first
                upload = second

        else:

            # One label was read:
            # remaining card gets missing role.
            remaining = [
                item
                for item in parsed
                if item is not download
                and item is not upload
            ]

            remaining.sort(
                key=lambda item: (
                    item["card"][1],
                    item["card"][0]
                )
            )

            if remaining:

                if download is None:
                    download = remaining[0]

                else:
                    upload = remaining[0]

    if (
        download is None
        or upload is None
        or download is upload
    ):
        return None

    result = {
        "download": _format_detected_speed(
            {"value": download.get("value", ""), "decimal_index": download.get("decimal_index")},
            download["numeric"],
            download["multiplier"],
        ),
        "upload": _format_detected_speed(
            {"value": upload.get("value", ""), "decimal_index": upload.get("decimal_index")},
            upload["numeric"],
            upload["multiplier"],
        ),
        "score": (
            5000
            + download["confidence"] * 100
            + upload["confidence"] * 100
        ),
    }

    return result


# ============================================================
# STACKED SPEEDTEST RESULT-CARD FALLBACK
# ============================================================

def _parse_speedtest_stacked_label_cards(image, ocr_result=None):
    """Read full-width, vertically stacked Unduh/Unggah result cards.

    Some Indonesian Speedtest result-detail screenshots stack the two
    result cards vertically and render the large speed digits in a font
    that whole-screen OCR can misread.  Use the OCR label positions to
    crop each card separately, then let the existing visual digit/decimal
    recognizer read only that card's large value.  Side-by-side cards are
    left to the existing card parser.
    """
    if image is None:
        return None

    try:
        h, w = image.shape[:2]
        if h < 200 or w < 160:
            return None

        # Reuse the initial screen OCR when available. Only rerun it with
        # inversion if the labels cannot be located on a dark screenshot.
        if ocr_result is None:
            ocr_result = ocr_speedtest_data(
                image,
                "--psm 6",
                invert=False
            )
        items, w_orig, h_orig = ocr_result
        if not items or not w_orig or not h_orig:
            return None

        dl_labels, ul_labels = find_speed_labels(items)

        def top_result_label(labels):
            # The lower latency section also contains words "Unduh" and
            # "Unggah". Restrict labels to the upper result-card region.
            eligible = [
                item for item in labels
                if item.get("cy", h_orig) < h_orig * 0.72
            ]
            if not eligible:
                return None
            return min(eligible, key=lambda item: item["cy"])

        dl_label = top_result_label(dl_labels)
        ul_label = top_result_label(ul_labels)

        if dl_label is None or ul_label is None:
            if _is_dark_image(image):
                items, w_orig, h_orig = ocr_speedtest_data(
                    image,
                    "--psm 6",
                    invert=True
                )
                dl_labels, ul_labels = find_speed_labels(items)
                dl_label = top_result_label(dl_labels)
                ul_label = top_result_label(ul_labels)

        if dl_label is None or ul_label is None:
            return None

        # If both labels are on the same row, these are not stacked cards.
        if abs(dl_label["cy"] - ul_label["cy"]) < h_orig * 0.09:
            return None

        card_h = max(100, int(round(h_orig * 0.16)))
        card_x = 0
        card_w = w_orig
        crop_offset = int(round(h_orig * 0.022))

        # Reuse the already recognized Mbps unit tokens to preserve the
        # script's existing kbps/mbps/gbps conversion behavior.
        units = [
            item for item in find_speed_mbps_items(items)
            if item.get("cy", h_orig) < h_orig * 0.76
        ]

        def read_value(label, expected_role):
            cy = label["cy"]
            card_y = max(
                0,
                min(h_orig - card_h, int(round(cy - crop_offset)))
            )
            card = (card_x, card_y, card_w, card_h)
            number = _speed_card_number(image, card, fast=True)
            if not number:
                return None

            numeric = float(number["numeric"])
            if not 0 < numeric <= 10000:
                return None

            # Match the Mbps-family unit vertically to this result label.
            nearest_unit = None
            if units:
                nearest_unit = min(
                    units,
                    key=lambda item: abs(item["cy"] - cy)
                )
            multiplier = 1.0
            if nearest_unit is not None:
                recognized, found_multiplier = _speed_unit_multiplier(
                    nearest_unit["text"]
                )
                if (
                    recognized
                    and abs(nearest_unit["cy"] - cy) <= h_orig * 0.18
                ):
                    multiplier = found_multiplier

            value_mbps = numeric * multiplier
            if not 0 < value_mbps <= 10000:
                return None

            return {
                "role": expected_role,
                "value": _format_detected_speed(number, value_mbps, multiplier),
                "numeric": value_mbps,
                "confidence": number.get("confidence", 0.0),
                "digit_count": number.get("digit_count", 0),
            }

        download = read_value(dl_label, "download")
        upload = read_value(ul_label, "upload")

        if not download or not upload:
            return None

        # Reject an accidental reuse of the same card crop/value.
        if (
            download["value"] == upload["value"]
            and abs(dl_label["cy"] - ul_label["cy"]) < h_orig * 0.15
        ):
            return None

        return {
            "download": download["value"],
            "upload": upload["value"],
            "score": (
                6000
                + download["confidence"] * 100
                + upload["confidence"] * 100
            ),
        }

    except Exception:
        # Keep all pre-existing parser paths available as fallbacks.
        return None


# ============================================================
# SPEEDTEST UNIT-ANCHORED VALUE EXTRACTION
# ============================================================

def _parse_speedtest_unit_anchored(image, ocr_result=None):
    """Read the large result directly above its Mbps unit.

    This handles mobile Speedtest screenshots where the numeric
    result is right-aligned far from the left-side Unduh/Unggah
    labels.  The nearby Mbps token is a much stronger anchor than
    arbitrary numbers elsewhere on the screen (for example ping,
    jitter, or data-usage figures).

    Return None when the geometry is not clear, so the existing
    OCR and card-based fallbacks can continue unchanged.
    """
    try:
        if ocr_result is None:
            ocr_result = ocr_speedtest_data(
                image,
                "--psm 6",
                invert=False
            )
        items, w_orig, h_orig = ocr_result

        if not items or not w_orig or not h_orig:
            return None

        candidates = get_speed_numeric_candidates(items)

        # Main speed readings use a large font. This filters out
        # small figures such as data used, ping, and jitter values.
        minimum_digit_height = max(
            12.0,
            h_orig * 0.015
        )
        candidates = [
            candidate for candidate in candidates
            if candidate["h"] >= minimum_digit_height
            and candidate["confidence"] >= 20
            and candidate["cy"] < h_orig * 0.72
        ]

        units = [
            unit for unit in find_speed_mbps_items(items)
            if unit["cy"] < h_orig * 0.76
        ]

        if len(candidates) < 2 or len(units) < 2:
            return None

        matches = []
        used_candidates = set()

        # A result is typically just above its Mbps label. Match each
        # unit only to a large, nearby numeric token, not to any number
        # that happens to be somewhere on the same screenshot.
        for unit in sorted(
            units,
            key=lambda item: (item["cy"], item["cx"])
        ):
            recognized, multiplier = _speed_unit_multiplier(
                unit["text"]
            )
            if not recognized:
                continue

            possible = []
            for candidate in candidates:
                if id(candidate) in used_candidates:
                    continue

                dx = abs(
                    candidate["cx"] - unit["cx"]
                ) / max(w_orig, 1)
                dy = (
                    unit["cy"] - candidate["cy"]
                ) / max(h_orig, 1)

                # Unit may be directly below the value, or aligned
                # almost level with it. It must not be far away or
                # above a data-usage figure beneath the result card.
                if dx > 0.30 or dy < -0.005 or dy > 0.10:
                    continue

                value_mbps = candidate["numeric"] * multiplier
                if not 0 < value_mbps <= 10000:
                    continue

                score = (
                    candidate["confidence"] * 0.8
                    + min(candidate["h"], h_orig * 0.08) * 1.5
                    - dx * 100
                    - max(dy, 0) * 120
                )
                possible.append((
                    score,
                    candidate,
                    value_mbps,
                    multiplier,
                    unit
                ))

            if possible:
                possible.sort(
                    key=lambda result: result[0],
                    reverse=True
                )
                best = possible[0]
                used_candidates.add(id(best[1]))
                matches.append(best)

        if len(matches) < 2:
            return None

        # Mobile cards stack vertically. Desktop/landscape cards are
        # often side-by-side, so use left-to-right order when their
        # units share approximately the same row.
        matches.sort(
            key=lambda result: (
                result[4]["cy"],
                result[4]["cx"]
            )
        )
        matches = matches[:2]

        if abs(
            matches[0][4]["cy"] - matches[1][4]["cy"]
        ) <= h_orig * 0.035:
            matches.sort(
                key=lambda result: result[4]["cx"]
            )

        download = matches[0]
        upload = matches[1]

        return {
            "download": _format_detected_speed(download[1], download[2], download[3]),
            "upload": _format_detected_speed(upload[1], upload[2], upload[3]),
            # Strong geometry + nearby Mbps-unit agreement should win
            # over generic whole-screen OCR candidates.
            "score": 1000.0 + download[0] + upload[0],
        }

    except Exception:
        # Do not prevent the established parser from running if this
        # specialized method encounters an unfamiliar screenshot.
        return None


# ============================================================
# SPEEDTEST IMAGE PARSER
# ============================================================

def parse_speedtest_image(
    image
):
    if image is None:
        return {
            "download": "",
            "upload": "",
            "score": 0
        }

    # One base OCR pass is shared by the stacked-card and unit-anchored
    # recognizers. This avoids paying for the same PSM-6 OCR twice.
    base_ocr = ocr_speedtest_data(
        image,
        "--psm 6",
        invert=False
    )

    # Stage 1: for the standard side-by-side Speedtest layout, prioritize the
    # isolated download/upload result cards over whole-screen numeric candidates.
    # This prevents small values such as jitter/notification digits from being
    # mistaken for the large DL/UL readings.
    detected_cards = _detect_speedtest_metric_cards(image)
    if _speedtest_cards_are_side_by_side(detected_cards):
        side_by_side_result = _parse_speedtest_cards(
            image,
            fast=True,
            cards=detected_cards,
        )
        if (
            side_by_side_result
            and side_by_side_result.get("download")
            and side_by_side_result.get("upload")
        ):
            return side_by_side_result

    # Stage 2: stacked-result and unit-anchored recognizers remain available
    # for screenshots whose card geometry is not side-by-side.
    stacked_result = _parse_speedtest_stacked_label_cards(
        image,
        ocr_result=base_ocr
    )
    if (
        stacked_result
        and stacked_result.get("download")
        and stacked_result.get("upload")
    ):
        return stacked_result

    unit_anchored_result = _parse_speedtest_unit_anchored(
        image,
        ocr_result=base_ocr
    )
    if (
        unit_anchored_result
        and unit_anchored_result.get("download")
        and unit_anchored_result.get("upload")
    ):
        return unit_anchored_result

    # Stage 3: fast visual card parse. The digit crops are isolated, so use
    # the two-variant recognition path first. Retry full card OCR only when
    # the fast attempt fails, returns a partial pair, or has weak digit votes.
    fast_card_result = _parse_speedtest_cards(image, fast=True)
    if (
        fast_card_result
        and fast_card_result.get("download")
        and fast_card_result.get("upload")
        and float(fast_card_result.get("score", 0.0)) >= 5100.0
    ):
        return fast_card_result

    # Stage 3b: accuracy fallback for difficult fonts, low-contrast cards, or
    # uncertain decimal marks. This expensive path runs only when needed.
    card_result = _parse_speedtest_cards(image, fast=False)
    if (
        card_result
        and card_result.get("download")
        and card_result.get("upload")
    ):
        return card_result

    # Stage 4: expensive multi-PSM fallback only when the focused recognizers
    # failed to produce a complete pair. Start with one likely configuration
    # and stop immediately when labels, units, and both values agree strongly.
    dark = _is_dark_image(image)
    if dark:
        passes = [
            ("--psm 6", True),
            ("--psm 11", True),
            ("--psm 6", False),
            ("--psm 12", True),
            ("--psm 11", False),
            ("--psm 3", False),
        ]
    else:
        passes = [
            ("--psm 6", False),
            ("--psm 11", False),
            ("--psm 12", False),
            ("--psm 6", True),
            ("--psm 3", False),
            ("--psm 4", False),
        ]

    results = []
    pair_counts = Counter()
    pair_values = {}

    for config, invert in passes:
        result = parse_speedtest_ocr_pass(
            image,
            config,
            invert
        )

        download = result.get("download", "")
        upload = result.get("upload", "")
        if not (download or upload):
            continue

        pair_quality = _speed_pair_quality(download, upload)
        result["pair_quality"] = pair_quality
        results.append(result)

        if download and upload:
            try:
                key = (round(float(download), 2), round(float(upload), 2))
            except (TypeError, ValueError):
                key = None
            if key is not None:
                pair_counts[key] += 1
                pair_values[key] = (download, upload)

            # These per-pass scores strongly reward label and Mbps-unit
            # proximity. A complete, high-scoring pair is safe to accept
            # without running the remaining Tesseract configurations.
            if result.get("score", 0.0) >= 120.0 and pair_quality >= 50.0:
                return {
                    "download": download,
                    "upload": upload,
                    "score": float(result.get("score", 0.0)) + pair_quality,
                }

        # If two independent passes agree on a complete pair, stop rather
        # than running every remaining configuration.
        if pair_counts:
            agreed = [key for key, count in pair_counts.items() if count >= 2]
            if agreed:
                best_key = max(
                    agreed,
                    key=lambda key: (
                        pair_counts[key],
                        _speed_pair_quality(*pair_values[key])
                    )
                )
                best_dl, best_ul = pair_values[best_key]
                return {
                    "download": best_dl,
                    "upload": best_ul,
                    "score": 200.0 + pair_counts[best_key] * 100.0
                        + _speed_pair_quality(best_dl, best_ul),
                }

    if not results:
        return {
            "download": "",
            "upload": "",
            "score": 0
        }

    # Vote on complete pairs first so DL and UL are not accidentally
    # combined from unrelated OCR passes.
    complete_results = [
        r for r in results
        if r.get("download") and r.get("upload")
    ]
    pair_groups = {}

    for result in complete_results:
        try:
            pair_key = (
                round(float(result["download"]), 2),
                round(float(result["upload"]), 2),
            )
        except (TypeError, ValueError):
            continue

        group = pair_groups.setdefault(pair_key, {
            "download": result["download"],
            "upload": result["upload"],
            "count": 0,
            "score": 0.0,
            "best_score": 0.0,
            "best_pair_quality": -1000.0,
        })
        group["count"] += 1
        group["score"] += (
            result.get("score", 0.0)
            + result.get("pair_quality", 0.0)
        )
        group["best_score"] = max(
            group["best_score"], result.get("score", 0.0)
        )
        group["best_pair_quality"] = max(
            group["best_pair_quality"], result.get("pair_quality", 0.0)
        )

    if pair_groups:
        best_group = max(
            pair_groups.values(),
            key=lambda group: (
                group["count"], group["score"],
                group["best_pair_quality"], group["best_score"]
            )
        )
        return {
            "download": best_group["download"],
            "upload": best_group["upload"],
            "score": (
                best_group["count"] * 100
                + best_group["score"]
                + best_group["best_pair_quality"]
            ),
        }

    # Last fallback for incomplete OCR: vote each field separately, but only
    # after all complete-pair attempts failed.
    def vote_metric(metric):
        groups = {}
        for result in results:
            value = result.get(metric, "")
            if not value:
                continue
            try:
                key = round(float(value), 2)
            except (TypeError, ValueError):
                continue
            group = groups.setdefault(key, {"value": value, "count": 0, "score": 0.0})
            group["count"] += 1
            group["score"] += (
                result.get("score", 0.0)
                + result.get("pair_quality", 0.0)
            )
        if not groups:
            return ""
        return max(
            groups.values(),
            key=lambda group: (group["count"], group["score"])
        )["value"]

    best_download = vote_metric("download")
    best_upload = vote_metric("upload")
    final_score = _speed_pair_quality(best_download, best_upload)
    final_score += max(result.get("score", 0.0) for result in results)
    return {
        "download": best_download,
        "upload": best_upload,
        "score": final_score,
    }

def speed_download_number(data):
    value = data.get(
        "download",
        ""
    )

    if not value:
        return -1.0

    try:
        return float(
            clean_number(value)
        )

    except Exception:
        return -1.0


def speed_upload_number(data):
    value = data.get(
        "upload",
        ""
    )

    if not value:
        return -1.0

    try:
        return float(
            clean_number(value)
        )

    except Exception:
        return -1.0


# ============================================================
# ANALYZE IMAGE (IN-MEMORY)
# ============================================================

def analyze_image_mat(image):
    """Run full-frame OCR once and preprocess only when G-NetTrack fields need recovery."""
    text_orig = ocr_image_original(image)
    gnet_kw_orig = gnet_keyword_score(text_orig)
    speed_kw_orig = speedtest_keyword_score(text_orig)

    # Parsing text is inexpensive compared with launching another Tesseract
    # process, so use it to decide whether targeted preprocessing is needed.
    gnet_orig = parse_gnet(text_orig)
    original_field_count = count_gnet_fields(gnet_orig)
    original_confidence = float(gnet_orig.get("confidence", 0.0) or 0.0)

    lower_text = (text_orig or "").lower()
    gnet_evidence = (
        gnet_kw_orig >= 1
        or original_field_count > 0
        or "g-nettrack" in lower_text
        or "gnettrack" in lower_text
    )

    # Most archive images are Speedtest, gallery, or unrelated screens. Do
    # not run the extra processed G-NetTrack OCR pass unless the original
    # actually resembles a G-NetTrack capture.
    if not gnet_evidence:
        return text_orig, gnet_kw_orig, speed_kw_orig

    # When all core fields including gNB/eNodeBID are already present, another
    # preprocessing OCR pass is unlikely to improve the result.
    core_fields = (
        "longitude", "latitude", "mcc", "mnc", "enodebid",
        "pci", "rsrp", "sinr"
    )
    has_core_fields = all(gnet_orig.get(field) for field in core_fields)
    original_is_adequate = (
        has_core_fields
        and original_confidence >= 85.0
    )
    if original_is_adequate:
        return text_orig, gnet_kw_orig, speed_kw_orig

    # Use a dedicated enlarged pass on the G-NetTrack statistics panel.
    # The previous 2x processed crop was too small for narrow phone videos and
    # often corrupted eNB/CID labels and coordinate digits.
    text_proc = ocr_gnet_targeted(image)
    gnet_kw_proc = gnet_keyword_score(text_proc)
    gnet_proc = parse_gnet(text_proc)

    orig_quality = (
        original_confidence,
        original_field_count
    )
    proc_quality = (
        float(gnet_proc.get("confidence", 0.0) or 0.0),
        count_gnet_fields(gnet_proc)
    )

    if (
        proc_quality > orig_quality
        or gnet_kw_proc > gnet_kw_orig
        or (
            not gnet_orig.get("enodebid")
            and gnet_proc.get("enodebid")
        )
    ):
        combined_text = text_proc + "\n" + text_orig
        return (
            combined_text,
            max(gnet_kw_orig, gnet_kw_proc),
            speed_kw_orig
        )

    return text_orig, gnet_kw_orig, speed_kw_orig


# ============================================================
# VIDEO FRAME EXTRACTION (MULTI-SAMPLE)
# ============================================================

def extract_best_video_frame(
    video_bytes,
    timestamps=(10, 20),
    return_analysis=False
):
    """Sample at 10s and 20s, stopping early if a usable G-NetTrack frame is found."""
    if cv2 is None or not video_bytes:
        return (None, None) if return_analysis else None

    temp_video = None
    cap = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".mp4", delete=False) as handle:
            handle.write(video_bytes)
            temp_video = handle.name

        cap = cv2.VideoCapture(temp_video)
        if not cap.isOpened():
            return (None, None) if return_analysis else None

        best_frame = None
        best_analysis = None
        best_score = -1.0

        for sec in timestamps:
            cap.set(cv2.CAP_PROP_POS_MSEC, sec * 1000)
            ret, frame = cap.read()
            if not ret or frame is None:
                continue

            analysis = analyze_image_mat(frame)
            text, gnet_score, speed_score = analysis
            gnet_data = parse_gnet(text)
            fields_score = count_gnet_fields(gnet_data)
            confidence = float(gnet_data.get("confidence", 0.0) or 0.0)

            # Prefer actual extracted telecom fields, not just a frame with
            # many generic OCR keywords.
            total_score = (
                float(gnet_score) * 3.0
                + fields_score * 10.0
                + confidence
                + (20.0 if gnet_data.get("enodebid") else 0.0)
                + (5.0 if gnet_data.get("rsrp") else 0.0)
                + (5.0 if gnet_data.get("sinr") else 0.0)
            )

            if total_score > best_score:
                best_score = total_score
                best_frame = frame
                best_analysis = analysis

            # Stop sampling once the frame contains a credible ID and enough
            # other validated fields; otherwise try the next sample timestamp.
            if (
                gnet_data.get("enodebid")
                and gnet_data.get("longitude")
                and gnet_data.get("latitude")
                and gnet_data.get("rsrp")
                and gnet_data.get("sinr")
                and confidence >= 60.0
            ):
                break

        if return_analysis:
            return best_frame, best_analysis
        return best_frame

    except Exception:
        return (None, None) if return_analysis else None

    finally:
        if cap is not None:
            try:
                cap.release()
            except Exception:
                pass
        if temp_video and os.path.exists(temp_video):
            try:
                os.remove(temp_video)
            except Exception:
                pass


# ============================================================
# IN-MEMORY ARCHIVE EXTRACTION & PROCESSOR
# ============================================================

def stream_archive_entries(
    archive_path,
    seven_zip=None
):
    """Yields (filename, is_video, bytes_content) directly from memory, prioritizing images first."""

    ext = os.path.splitext(
        archive_path
    )[1].lower()

    def sort_priority(filename):

        file_ext = os.path.splitext(
            filename
        )[1].lower()

        if file_ext in IMAGE_EXTENSIONS:
            return 0

        if file_ext in VIDEO_EXTENSIONS:
            return 1

        return 2

    if ext == ".zip":

        with zipfile.ZipFile(
            archive_path,
            "r"
        ) as z:

            sorted_info = sorted(
                z.infolist(),
                key=lambda info:
                sort_priority(
                    info.filename
                )
            )

            for info in sorted_info:

                if info.is_dir():
                    continue

                name_lower = info.filename.lower()

                is_img = any(
                    name_lower.endswith(e)
                    for e in IMAGE_EXTENSIONS
                )

                is_vid = any(
                    name_lower.endswith(e)
                    for e in VIDEO_EXTENSIONS
                )

                if is_img or is_vid:

                    yield (
                        os.path.basename(
                            info.filename
                        ),
                        is_vid,
                        z.read(
                            info.filename
                        )
                    )

    elif (
        ext == ".rar"
        and rarfile is not None
    ):

        with rarfile.RarFile(
            archive_path,
            "r"
        ) as r:

            sorted_info = sorted(
                r.infolist(),
                key=lambda info:
                sort_priority(
                    info.filename
                )
            )

            for info in sorted_info:

                if info.isdir():
                    continue

                name_lower = info.filename.lower()

                is_img = any(
                    name_lower.endswith(e)
                    for e in IMAGE_EXTENSIONS
                )

                is_vid = any(
                    name_lower.endswith(e)
                    for e in VIDEO_EXTENSIONS
                )

                if is_img or is_vid:

                    yield (
                        os.path.basename(
                            info.filename
                        ),
                        is_vid,
                        r.read(
                            info.filename
                        )
                    )

    else:

        if not seven_zip:
            return

        temp_dir = tempfile.mkdtemp(
            prefix="fallback_rar_"
        )

        try:

            subprocess.run(
                [
                    seven_zip,
                    "x",
                    archive_path,
                    f"-o{temp_dir}",
                    "-y"
                ],
                capture_output=True
            )

            all_files = []

            for root, _, files in os.walk(
                temp_dir
            ):

                for f in files:
                    all_files.append(
                        os.path.join(
                            root,
                            f
                        )
                    )

            all_files.sort(
                key=sort_priority
            )

            for p in all_files:

                f = os.path.basename(
                    p
                )

                name_lower = f.lower()

                is_img = any(
                    name_lower.endswith(e)
                    for e in IMAGE_EXTENSIONS
                )

                is_vid = any(
                    name_lower.endswith(e)
                    for e in VIDEO_EXTENSIONS
                )

                if is_img or is_vid:

                    with open(
                        p,
                        "rb"
                    ) as fh:

                        yield (
                            f,
                            is_vid,
                            fh.read()
                        )

        finally:

            shutil.rmtree(
                temp_dir,
                ignore_errors=True
            )


def classify_speedtest_image(image, text="", speed_score=0):
    """Classify an image using independent text and layout evidence.

    A single word such as "download", a single Mbps token, or two arbitrary
    rectangles is not sufficient. Accept the image only when the evidence
    describes a plausible pair of Speedtest result readings.
    """
    lower = normalize_text(text or "").lower()

    has_brand = bool(re.search(r"\b(?:speedtest|ookla)\b", lower))
    has_download_label = bool(re.search(
        r"download|downioad|downoad|downld|downlo|dwnl|unduh|undu|nduh",
        lower,
        re.IGNORECASE
    ))
    has_upload_label = bool(re.search(
        r"upload|upioad|uploed|upld|uploa|uplad|unggah|ungga|nggah|ungah",
        lower,
        re.IGNORECASE
    ))
    has_both_labels = has_download_label and has_upload_label

    # Count distinct unit occurrences, not just whether any unit-like token
    # exists. The paired readings normally have two Mbps-family unit labels.
    unit_pattern = (
        r"(?<![a-z])(?:m|g|k)\s*(?:bit|b|0|o|h)?\s*"
        r"(?:p|/|o|0)\s*(?:s|5)(?![a-z])"
    )
    unit_count = len(re.findall(unit_pattern, lower, re.IGNORECASE))
    if unit_count == 0:
        unit_count = len(re.findall(r"\b(?:mbps|gbps|kbps)\b", lower, re.IGNORECASE))

    has_latency_context = bool(re.search(r"\bping\b|jitter", lower, re.IGNORECASE))

    cards = []
    side_by_side_cards = False
    try:
        if image is not None and cv2 is not None:
            cards = _detect_speedtest_metric_cards(image)
            side_by_side_cards = _speedtest_cards_are_side_by_side(cards)
    except Exception:
        cards = []
        side_by_side_cards = False

    # Higher score means more independent evidence. `speed_score` is retained
    # for telemetry/compatibility but deliberately cannot trigger acceptance
    # on its own: it previously allowed a single weak keyword to pass.
    score = 0.0
    if has_brand:
        score += 4.0
    if has_download_label:
        score += 1.5
    if has_upload_label:
        score += 1.5
    if has_both_labels:
        score += 1.5
    if unit_count == 1:
        score += 1.25
    elif unit_count >= 2:
        score += 3.0
    if side_by_side_cards:
        score += 2.5
    if has_latency_context:
        score += 0.5

    # A brand is useful only with evidence of result data. Without branding,
    # require both direction labels plus either two unit labels or a genuine
    # side-by-side card arrangement. This prevents G-NetTrack's "UL/DL: kbps",
    # unrelated dashboards, and generic card layouts from being selected.
    branded_result = has_brand and (
        has_both_labels or unit_count >= 2 or side_by_side_cards
    )
    unbranded_result = has_both_labels and (
        unit_count >= 2 or side_by_side_cards
    )
    accepted = bool(branded_result or unbranded_result)

    return {
        "accepted": accepted,
        "score": score,
        "brand": has_brand,
        "download_label": has_download_label,
        "upload_label": has_upload_label,
        "unit_count": unit_count,
        "side_by_side_cards": side_by_side_cards,
        "latency_context": has_latency_context,
        "generic_speed_score": speed_score,
    }


def _is_likely_speedtest_image(image, text="", speed_score=0):
    """Compatibility wrapper: use the stricter multi-evidence classifier."""
    return classify_speedtest_image(image, text, speed_score).get("accepted", False)


def process_archive_task(
    task_data
):
    archive_path, tesseract_cmd, seven_zip = task_data

    if (
        tesseract_cmd
        and pytesseract
    ):
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    archive_id = os.path.splitext(
        os.path.basename(
            archive_path
        )
    )[0]

    best_gnet = {
        "longitude": "",
        "latitude": "",
        "mcc": "",
        "mnc": "",
        "serving_type": "",
        "enodebid": "",
        "pci": "",
        "cid": "",
        "rsrp": "",
        "sinr": "",
    }

    best_gnet_score = 0

    best_speed = {
        "download": "",
        "upload": ""
    }

    best_speed_priority = -1.0
    best_speed_source = ""

    for filename, is_video, file_bytes in stream_archive_entries(
        archive_path,
        seven_zip
    ):

        frame = None
        precomputed_analysis = None

        if is_video:

            frame, precomputed_analysis = extract_best_video_frame(
                file_bytes,
                return_analysis=True
            )

        else:

            nparr = np.frombuffer(
                file_bytes,
                np.uint8
            )

            frame = cv2.imdecode(
                nparr,
                cv2.IMREAD_COLOR
            )

        if frame is None:
            continue

        if precomputed_analysis is None:
            text, gnet_score, speed_score = analyze_image_mat(frame)
        else:
            text, gnet_score, speed_score = precomputed_analysis

        # ====================================================
        # G-NETTRACK
        # Includes the NR/5G-aware cell identifier parser.
        # ====================================================

        if gnet_score >= 2:

            gnet = parse_gnet(
                text
            )

            actual_score = count_gnet_fields(
                gnet
            )

            actual_score += int(
                gnet.get(
                    "confidence",
                    0.0
                )
                / 10.0
            )

            if actual_score > best_gnet_score:

                best_gnet = gnet
                best_gnet_score = actual_score

            elif actual_score == best_gnet_score:

                best_gnet = merge_missing_gnet_fields(
                    best_gnet,
                    gnet
                )

        # ====================================================
        # SPEEDTEST
        #
        # FIX 1:
        # ALWAYS run specialized Speedtest parser on images.
        #
        # Previously:
        #     speed_score >= 2
        #
        # could prevent Speedtest OCR from running when generic
        # OCR failed to recognize "Download", "Upload", etc.
        #
        # FIX 2:
        # Do not favor a larger numeric value.
        # ====================================================

        speedtest_classification = {
            "accepted": False,
            "score": 0.0,
        }
        if not is_video:
            speedtest_classification = classify_speedtest_image(
                frame,
                text,
                speed_score,
            )

        if (
            not is_video
            and speedtest_classification.get("accepted", False)
        ):

            speed = parse_speedtest_image(
                frame
            )

            current_dl = speed_download_number(
                speed
            )

            current_ul = speed_upload_number(
                speed
            )

            pair_quality = _speed_pair_quality(
                speed.get(
                    "download",
                    ""
                ),
                speed.get(
                    "upload",
                    ""
                )
            )

            has_download = (
                current_dl > 0
            )

            has_upload = (
                current_ul > 0
            )

            has_complete_pair = (
                has_download
                and has_upload
            )

            # Reject an OCR output that contains no valid positive speed.
            # Classification and selection work together: first require
            # credible Speedtest evidence, then prefer a complete DL/UL pair,
            # then compare classification strength and OCR/pair quality.
            if not (has_download or has_upload):
                continue

            priority = 0.0
            if has_complete_pair:
                priority += 1_000_000.0
            else:
                priority += 100_000.0

            # Classification strength outweighs noisy OCR scores (which can
            # be artificially large for unrelated images with many tokens).
            priority += float(speedtest_classification.get("score", 0.0)) * 10_000.0

            # Cap parser score so it cannot overwhelm classification quality.
            parser_score = max(0.0, float(speed.get("score", 0.0)))
            priority += min(parser_score, 9_999.0)
            priority += max(0.0, pair_quality * 10.0)

            if priority > best_speed_priority:
                best_speed = speed
                best_speed_priority = priority
                best_speed_source = filename

    serving_type = best_gnet.get(
        "serving_type",
        ""
    )

    rsrp = format_radio_value(
        best_gnet.get(
            "rsrp",
            ""
        ),
        serving_type
    )

    sinr = format_radio_value(
        best_gnet.get(
            "sinr",
            ""
        ),
        serving_type
    )

    if rsrp:
        rsrp = add_excel_apostrophe(
            rsrp
        )

    if sinr:
        sinr = add_excel_apostrophe(
            sinr
        )

    download = best_speed.get(
        "download",
        ""
    )

    upload = best_speed.get(
        "upload",
        ""
    )

    dl_ul = (
        f"{download or '-'}/{upload or '-'} Mbps"
        if (
            download
            or upload
        )
        else ""
    )

    result = {
        "RAR ID": archive_id,
        "Longitude": best_gnet.get(
            "longitude",
            ""
        ),
        "Latitude": best_gnet.get(
            "latitude",
            ""
        ),
        "MCC": best_gnet.get(
            "mcc",
            ""
        ),
        "MNC": best_gnet.get(
            "mnc",
            ""
        ),
        "Serving Type": serving_type,
        "eNodeBID": best_gnet.get(
            "enodebid",
            ""
        ),
        "PCI": best_gnet.get(
            "pci",
            ""
        ),
        "CID": best_gnet.get(
            "cid",
            ""
        ),
        "RSRP": rsrp,
        "SINR": sinr,
        "DL/UL speed": dl_ul,
        "OCR G-NetTrack Score": best_gnet_score,
        "G-NetTrack Confidence (%)": best_gnet.get(
            "confidence",
            0.0
        ),
        "RSRP Valid": (
            "YES"
            if valid_rsrp(
                best_gnet.get(
                    "rsrp",
                    ""
                )
            )
            else "NO"
        ),
        "SINR Valid": (
            "YES"
            if valid_sinr(
                best_gnet.get(
                    "sinr",
                    ""
                )
            )
            else "NO"
        ),
        "OCR Speedtest Score": (
            1
            if best_speed_source
            else 0
        ),
    }

    return (
        archive_id,
        result
    )


# ============================================================
# WRITE CSV
# ============================================================

def write_csv(
    output_path,
    results
):
    if not results:
        return

    columns = [
        "RAR ID",
        "Longitude",
        "Latitude",
        "MCC",
        "MNC",
        "Serving Type",
        "eNodeBID",
        "PCI",
        "CID",
        "RSRP",
        "SINR",
        "DL/UL speed",
        "OCR G-NetTrack Score",
        "G-NetTrack Confidence (%)",
        "RSRP Valid",
        "SINR Valid",
        "OCR Speedtest Score",
    ]

    with open(
        output_path,
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=columns
        )

        writer.writeheader()

        for result in results:
            writer.writerow(
                result
            )


# ============================================================
# MAIN
# ============================================================

def main():
    # Show worker selection as the first UI when the script opens.
    select_worker_count()

    print(
        "=" * 70
    )

    print(
        "PARALLEL IN-MEMORY OCR PROCESSOR"
    )

    print(
        "RAR / ZIP -> G-NetTrack + Speedtest -> CSV"
    )

    print(
        "=" * 70
    )

    if (
        cv2 is None
        or pytesseract is None
    ):

        print(
            "\nERROR: OpenCV or pytesseract is missing."
        )

        return

    if not configure_tesseract():

        print(
            "\nERROR: Tesseract OCR executable not found."
        )

        input(
            "\nPress Enter to exit..."
        )

        return

    seven_zip = find_7zip()

    if (
        rarfile is not None
        and seven_zip
    ):
        rarfile.UNRAR_TOOL = seven_zip

    folder = select_folder()

    if not folder:

        print(
            "No folder selected."
        )

        return

    archives = find_archives(
        folder
    )

    worker_count = min(MAX_WORKERS, len(archives))

    print(
        f"\nFound {len(archives)} archive(s). "
        f"Processing with {worker_count} worker(s). "
        f"(Selected OCR_MAX_WORKERS={MAX_WORKERS}.)\n"
    )

    if not archives:

        print(
            "No RAR/ZIP archives to process."
        )

        input(
            "\nPress Enter to exit..."
        )

        return

    tasks = [
        (
            os.path.join(
                folder,
                archive_name
            ),
            TESSERACT_PATH,
            seven_zip
        )
        for archive_name in archives
    ]

    results = []

    with ProcessPoolExecutor(
        max_workers=worker_count
    ) as executor:

        futures = {
            executor.submit(
                process_archive_task,
                task
            ): task[0]
            for task in tasks
        }

        completed_count = 0

        for future in as_completed(
            futures
        ):

            archive_path = futures[
                future
            ]

            try:

                archive_id, res = future.result()

                results.append(
                    res
                )

                completed_count += 1

                print(
                    f"[{completed_count}/{len(archives)}] "
                    f"Completed: {archive_id}"
                )

            except Exception as e:

                print(
                    f"Error processing "
                    f"{os.path.basename(archive_path)}: {e}"
                )

    output_path = os.path.join(
        folder,
        OUTPUT_CSV
    )

    write_csv(
        output_path,
        results
    )

    print(
        "\n"
        + "=" * 70
    )

    print(
        "PROCESS COMPLETE"
    )

    print(
        f"Results written directly to: {output_path}"
    )

    print(
        "=" * 70
    )

    input(
        "\nPress Enter to exit..."
    )


if __name__ == "__main__":
    main()