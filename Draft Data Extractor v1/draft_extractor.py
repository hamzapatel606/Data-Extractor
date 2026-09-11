# -*- coding: utf-8 -*-
"""
Draft Data Extractor V1 -> Excel
Incremental / safe batch processing
=========================================================
The extractor detects voter boxes from the actual PDF layout and OCRs each
box independently.

Key design points:
  - Column gutters are detected from the page image instead of assuming thirds.
  - Box boundaries are detected from horizontal separator lines.
  - The box header (Serial + EPIC) is OCRed separately with multiple PSM modes.
  - The box body (Name / Relative Name / Age / Gender) is OCRed separately.
  - Physical grid position (page, row, column) is preserved so the authoritative
    serial order follows the PDF even when OCR cannot read the printed serial.
"""

import gc
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
try:
    import numpy as np
    import pandas as pd
    import pytesseract
    from pdf2image import convert_from_path, pdfinfo_from_path
    from PIL import Image, ImageOps
except ImportError as e:
    # On a new PC, required libraries may not be installed. A common problem is that
    # pip installs a package into a different Python environment. The command below
    # uses the Python interpreter that is running this script.
    import sys
    raise SystemExit(
        f"Required Python library is missing: {e.name}\n\n"
        f"Copy and run this complete command in the terminal:\n\n"
        f'  & "{sys.executable}" -m pip install numpy pandas pytesseract pdf2image pillow openpyxl\n\n'
        f"(-m pip is required; otherwise pip may install the package into a different Python environment)")


# ------------------------------------------------------------------ SETTINGS
# Tesseract and Poppler may be installed in different locations on different PCs.
# The script searches common locations automatically. If you want to specify paths
# manually, fill in the settings below; manual paths take priority.
MANUAL_TESSERACT = r''      # example: r'C:\Program Files\Tesseract-OCR\tesseract.exe'
MANUAL_POPPLER   = r''      # example: r'C:\poppler-26.02.0\Library\bin'
MANUAL_FOLDER    = r''      # optional; defaults to the folder containing this script


def _pehla_mojood(rasta_list):
    for p in rasta_list:
        if p and os.path.exists(p):
            return p
    return None


HOME = os.path.expanduser('~')
# The folder containing this script is checked first.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def dhoondo_poppler(manual=''):
    """
    Find the Poppler bin folder.

    Checking only the common Poppler bin path is not sufficient. Depending on how the archive is extracted,
    the path may be nested differently, and users may also place it under Downloads or Desktop.
    The script therefore checks several locations and finally searches recursively for pdftoppm.
    """
    def theek(d):
        if not d:
            return None
        for exe in ('pdftoppm.exe', 'pdftoppm'):
            if os.path.exists(os.path.join(d, exe)):
                return d
        return None

    if manual:
        return theek(manual)

    got = shutil.which('pdftoppm')          # if it is already available on PATH
    if got:
        return os.path.dirname(got)

    roots = ['C:\\', 'D:\\', 'E:\\',
             r'C:\Program Files', r'C:\Program Files (x86)',
             HOME,
             os.path.join(HOME, 'Downloads'),
             os.path.join(HOME, 'Desktop'),
             os.path.join(HOME, 'OneDrive', 'Desktop'),
             os.path.join(HOME, 'Documents'),
             SCRIPT_DIR]

    pats = [os.path.join('poppler*', 'Library', 'bin'),
            os.path.join('poppler*', 'poppler*', 'Library', 'bin'),
            os.path.join('poppler*', 'bin'),
            os.path.join('*', 'poppler*', 'Library', 'bin')]
    for root in roots:
        if not os.path.isdir(root):
            continue
        for pat in pats:
            for d in sorted(glob.glob(os.path.join(root, pat))):
                if theek(d):
                    return d

    # Final fallback: recursively search inside folders whose names start with
    # "poppler" until pdftoppm is found.
    for root in roots:
        if not os.path.isdir(root):
            continue
        try:
            top = [d for d in os.listdir(root)
                   if d.lower().startswith('poppler')
                   and os.path.isdir(os.path.join(root, d))]
        except OSError:
            continue
        for d in top:
            for cur, _dirs, files in os.walk(os.path.join(root, d)):
                if 'pdftoppm.exe' in files or 'pdftoppm' in files:
                    return cur
    return None


POPPLER_HELP = (
    "Poppler was not found.\n"
    "  1) If Poppler is already installed on another PC, copy the complete folder\n"
    "     (for example, C:\\poppler-26.02.0) copy it to C:\\ on this PC.\n"
    "  2) Or download it from: github.com/oschwartz10612/poppler-windows -> Releases\n"
    "     Extract the ZIP to C:\\ so that C:\\poppler-xx\\Library\\bin is created.\n"
    "  3) Or set MANUAL_POPPLER near the top of this script to the complete\n"
    "     path of the bin folder containing pdftoppm.exe.")


TESSERACT_CMD = MANUAL_TESSERACT or _pehla_mojood([
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    os.path.join(HOME, r'AppData\Local\Tesseract-OCR\tesseract.exe'),
    os.path.join(HOME, r'AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
    shutil.which('tesseract'),
])

POPPLER_PATH = dhoondo_poppler(MANUAL_POPPLER)

# On some PCs the Desktop is inside OneDrive, while on others it is not.
# Both common locations are checked.
INPUT_FOLDER = MANUAL_FOLDER or (os.path.join(SCRIPT_DIR, 'input') if os.path.isdir(os.path.join(SCRIPT_DIR, 'input')) else SCRIPT_DIR)
FOLDER_PATH = INPUT_FOLDER


def setup_check():
    """Check that all required dependencies and folders are available before starting."""
    masail = []
    print("--- Setup check ---")
    print(f"  Tesseract : {TESSERACT_CMD or 'NOT FOUND'}")
    if not TESSERACT_CMD:
        masail.append("Tesseract was not found. Install it or set MANUAL_TESSERACT.")
    print(f"  Poppler   : {POPPLER_PATH or 'NOT FOUND'}")
    if POPPLER_PATH is None:
        masail.append("Poppler was not found. Install it or set MANUAL_POPPLER.")
    print(f"  PDF folder: {FOLDER_PATH or 'NOT FOUND'}")
    if not FOLDER_PATH:
        masail.append("PDF input folder was not found. Set MANUAL_FOLDER or create the input folder.")
    if masail:
        print()
        for x in masail:
            print("  !! " + x)
        raise SystemExit(1)
    print("  All setup checks passed.\n")


if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
# Always write results to a sibling output folder beside the input folder.
PROJECT_DIR = SCRIPT_DIR
OUTPUT_DIR = os.path.join(PROJECT_DIR, 'output')
OUTPUT_PATH = os.path.join(OUTPUT_DIR, 'Master_Voter_List.xlsx')
STATE_PATH = os.path.join(OUTPUT_DIR, '.processed_pdfs.json')

TARGET_WIDTH_PX  = 2000   # target page width in pixels
BOX_UPSCALE      = 2      # box upscaling factor before OCR
MAX_BOX_FRAC     = 0.30   # larger bands are rejected as non-voter regions such as tables/images
WORKERS          = max(1, (os.cpu_count() or 2) - 1)   # number of CPU workers
# NOTE: OCR is CPU-bound and each worker processes one page at a time. Increasing
# WORKERS can improve throughput. If multiprocessing causes problems on Windows,
# set WORKERS = 1.
USE_FULLPAGE_EPIC = True  # extract EPIC values from full-page OCR (more reliable and faster)
BATCH_OCR        = True   # process multiple images in one Tesseract call (see below)
BODY_ARGS        = ['--oem', '3', '--psm', '6', '-c', 'preserve_interword_spaces=1']
BODY_CONFIG      = r'--oem 3 --psm 6 -c preserve_interword_spaces=1'
HEADER_PSMS      = (11, 6, 7)      # tried in this order
# NOTE: this roll contains multiple EPIC prefixes (RSC, HFL, AFZ, WYY, YTR, UFG,
# NHL, URB, ZHS, NLF, YDE, FTL...) therefore no prefix whitelist is enforced.

# ------------------------------------------------------------------ PATTERNS
# OCR may convert ":" into characters such as #, !, ;, ., ,, =, or quotes; all are included here
SEP = r'[\s:;#!*~.,=+<>"\u201c\u201d\'\u2018\u2019\-–—]{0,5}'
RE_EPIC    = re.compile(r'\b([A-Z]{2,4}\d{6,8})\b')
RE_EPIC_OK = re.compile(r'^[A-Z]{3}\d{7}$')
RE_NAME    = re.compile(r'^[^A-Za-z]{0,3}Name' + SEP + r'(.*)$', re.I)
RE_FATHER  = re.compile(r"Father[’'`\s]*s?\s*Name" + SEP + r'(.*)$', re.I)
RE_HUSBAND = re.compile(r"Husband[’'`\s]*s?\s*Name" + SEP + r'(.*)$', re.I)
RE_MOTHER  = re.compile(r"Mother[’'`\s]*s?\s*Name" + SEP + r'(.*)$', re.I)
RE_OTHER   = re.compile(r"Other[’'`\s]*s?\s*(?:Name)?" + SEP + r'(.*)$', re.I)
RE_HOUSE   = re.compile(r'H[o0]use\s*N[uo]?m?b?e?r?' + SEP + r'(.*)$', re.I)
RE_AGE     = re.compile(r'Age\D{0,4}(\d{1,3})', re.I)
RE_GENDER  = re.compile(r'Gender\D{0,4}(Male|Female|Other|M|F)\b', re.I)
RE_PART    = re.compile(r'Part\s*No\.?\D{0,4}(\d{1,4})', re.I)
RE_AC      = re.compile(r'Assembly\s+Constituency.*?Name' + SEP + r'(.+)$', re.I)
RE_PAGE    = re.compile(r'Page\s*(\d+)\s*$', re.I)
JUNK       = re.compile(r'\b(photo|available|is\s*available)\b', re.I)


def clean(t):
    t = JUNK.sub(' ', t or '')
    t = re.sub(r'[\|\[\]\{\}_]+', ' ', t)
    t = re.sub(r'\s{2,}', ' ', t).strip(' :;#!.,=+<>-')
    # WARNING: Excel treats cells beginning with "=" or "+" as formulas and may
    # evaluate or clear them. Text such as "Name =: nasrin gilani" is therefore
    # sanitized before it is written to Excel.
    while t[:1] in ('=', '+', '@'):
        t = t[1:].lstrip(' :;.,-')
    return t


def find_column_cuts(page):
    """Find the actual blank gutters between the three columns."""
    ink = (np.array(page) < 128).sum(axis=0)
    w = len(ink)
    cuts = [0]
    for frac in (1 / 3.0, 2 / 3.0):
        c, win = int(w * frac), int(w * 0.07)
        lo, hi = max(0, c - win), min(w, c + win)
        cuts.append(lo + int(np.argmin(ink[lo:hi])))
    cuts.append(w)
    return cuts


def find_box_bands(strip, min_h):
    """Detect the upper and lower boundaries of each voter box within a column."""
    ink = (np.array(strip) < 128).sum(axis=1)
    blank = ink <= 1
    bands, start = [], None
    for i, b in enumerate(blank):
        if not b and start is None:
            start = i
        elif b and start is not None:
            if i - start >= min_h:
                bands.append((start, i))
            start = None
    if start is not None and len(blank) - start >= min_h:
        bands.append((start, len(blank)))
    return bands


def read_header(box):
    """
    Extract the EPIC and serial number from the first line of a voter box.

    A single OCR read is not trusted. The EPIC is read from two different crops.
    If both results agree, the result is accepted; if they differ, (raazi=False)
    is returned and the caller runs read_header_hard(). This is necessary because
    OCR can confuse 5 with 9 or C with O, producing a value that looks valid but is wrong.

    Returns: (epic, serial, raazi)
    """
    h = box.height
    votes = []
    for frac in (0.22, 0.26, 0.18, 0.32):
        crop = box.crop((0, 0, box.width, max(12, int(h * frac))))
        crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
        for psm in HEADER_PSMS:
            txt = pytesseract.image_to_string(crop, config=f'--oem 3 --psm {psm}')
            m = RE_EPIC.search(txt.replace(' ', '').upper())
            if m and RE_EPIC_OK.match(m.group(1)):
                votes.append(m.group(1))
                break
        if len(votes) >= 2:
            break

    if len(votes) >= 2:
        epic, raazi = votes[0], votes[0] == votes[1]
    elif votes:
        epic, raazi = votes[0], False      # only one reading is available - verification is required
    else:
        epic, raazi = '', False

    # the serial number is on the right side of the small header area
    H = max(12, int(h * 0.22))
    sc = box.crop((int(box.width * 0.03), int(H * 0.12),
                   int(box.width * 0.40), int(H * 0.95)))
    sc = ImageOps.expand(sc.resize((sc.width * 4, sc.height * 4), Image.LANCZOS), 25, 255)
    t = pytesseract.image_to_string(
        sc, config='--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789')
    d = re.findall(r'\d{1,4}', t)
    serial = max(d, key=len) if d else ''
    return epic, serial, raazi


def read_header_hard(box):
    """
    When the normal EPIC extraction is not reliable, the right side of the header
    is enlarged and OCRed repeatedly, both with and without a whitelist.

    The whitelist mode usually reads letters (RSC / AFZ / YTR...) well but can
    confuse 5 with 9. The non-whitelist mode often reads digits better but can
    confuse letters (for example, RSC as RGC). Both are therefore used:
      - if both results agree -> return immediately (the common fast path)
      - if they disagree but the first three letters agree -> do not guess; flag the record
    """
    from collections import Counter
    H = max(12, int(box.height * 0.22))
    wl, nw = Counter(), Counter()

    for x0f, up in ((0.35, 4), (0.45, 4), (0.35, 6), (0.45, 6), (0.30, 4)):
        c = box.crop((int(box.width * x0f), 0, box.width, H))
        c = ImageOps.expand(c.resize((c.width * up, c.height * up), Image.LANCZOS), 20, 255)
        for psm in (7, 13, 11, 8):
            for white, bag in ((True, wl), (False, nw)):
                cfg = (' -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
                       if white else '')
                t = pytesseract.image_to_string(c, config=f'--oem 3 --psm {psm}' + cfg)
                mm = RE_EPIC.search(t.replace(' ', '').upper())
                if mm and RE_EPIC_OK.match(mm.group(1)):
                    bag[mm.group(1)] += 1

    if not wl:
        return (nw.most_common(1)[0][0], False) if nw else ('', False)
    best = wl.most_common(1)[0][0]
    if nw:
        alt = nw.most_common(1)[0][0]
        if alt != best and alt[:3] == best[:3]:
            return best, True          # the two readings disagree -> verification is required
    return best, False


def ocr_batch(images, args):
    """
    Process multiple images in a single Tesseract call.

    Why: pytesseract starts a new tesseract.exe process and reloads language data
    on every call. A page can require 100-200 calls, so process startup alone can
    become expensive. Tesseract accepts an image list and separates outputs with
    form-feed characters, allowing the same images and settings to be processed
    in a single Tesseract process.

    If a batch operation fails to return the complete result set, the function
    the function falls back to the previous method so the result is not lost.
    """
    if not images:
        return []
    cfg = ' '.join(args)
    if not BATCH_OCR or len(images) == 1:
        return [pytesseract.image_to_string(im, config=cfg) for im in images]

    tmp = tempfile.mkdtemp(prefix='ocrbatch_')
    try:
        paths = []
        for i, im in enumerate(images):
            fp = os.path.join(tmp, f'{i:04d}.png')
            im.save(fp)
            paths.append(fp)
        lst = os.path.join(tmp, 'list.txt')
        with open(lst, 'w', encoding='utf-8') as f:
            f.write('\n'.join(paths))
        cmd = [pytesseract.pytesseract.tesseract_cmd, lst, 'stdout'] + args
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding='utf-8', errors='replace')
        parts = r.stdout.split('\f')
        if len(parts) == len(images) + 1 and not parts[-1].strip():
            parts.pop()
        if len(parts) != len(images):
            raise RuntimeError('Batch OCR did not return the complete result set')
        return parts
    except Exception:
        return [pytesseract.image_to_string(im, config=cfg) for im in images]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def ocr_batch_multi(images, arg_sets):
    """
    Process the same image set with multiple settings while writing the images to disk
    only once and reusing the list file. There is one Tesseract call per setting.
    """
    if not images:
        return [[] for _ in arg_sets]
    if not BATCH_OCR:
        return [[pytesseract.image_to_string(im, config=' '.join(a)) for im in images]
                for a in arg_sets]

    tmp = tempfile.mkdtemp(prefix='ocrmulti_')
    try:
        paths = []
        for i, im in enumerate(images):
            fp = os.path.join(tmp, f'{i:05d}.png')
            im.save(fp)
            paths.append(fp)
        lst = os.path.join(tmp, 'list.txt')
        with open(lst, 'w', encoding='utf-8') as f:
            f.write('\n'.join(paths))

        out = []
        for args in arg_sets:
            try:
                r = subprocess.run(
                    [pytesseract.pytesseract.tesseract_cmd, lst, 'stdout'] + args,
                    capture_output=True, text=True, encoding='utf-8', errors='replace')
                parts = r.stdout.split('\f')
                if len(parts) == len(images) + 1 and not parts[-1].strip():
                    parts.pop()
                if len(parts) != len(images):
                    raise RuntimeError('Incomplete result set')
                out.append(parts)
            except Exception:
                out.append([pytesseract.image_to_string(im, config=' '.join(args))
                            for im in images])
        return out
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def read_headers_batch(boxes):
    """
    Read EPIC + serial for all boxes in one batch. The decision logic is the same as
    read_header(): two crop readings must agree, but now each crop/PSM uses one call.
    """
    n = len(boxes)
    votes = [[] for _ in range(n)]

    for frac in (0.22, 0.26, 0.18, 0.32):
        pending = [i for i in range(n) if len(votes[i]) < 2]
        if not pending:
            break
        # boxes without a result at this fraction
        need = list(pending)
        crops = {}
        for i in need:
            b = boxes[i]
            c = b.crop((0, 0, b.width, max(12, int(b.height * frac))))
            crops[i] = c.resize((c.width * 3, c.height * 3), Image.LANCZOS)

        got = set()
        for psm in HEADER_PSMS:
            todo = [i for i in need if i not in got]
            if not todo:
                break
            texts = ocr_batch([crops[i] for i in todo],
                              ['--oem', '3', '--psm', str(psm)])
            for i, txt in zip(todo, texts):
                mm = RE_EPIC.search(txt.replace(' ', '').upper())
                if mm and RE_EPIC_OK.match(mm.group(1)):
                    votes[i].append(mm.group(1))
                    got.add(i)

    out = []
    for i in range(n):
        v = votes[i]
        if len(v) >= 2:
            out.append((v[0], v[0] == v[1]))
        elif v:
            out.append((v[0], False))
        else:
            out.append(('', False))

    # small serial-header crops - all in one call
    sc_imgs = []
    for b in boxes:
        H = max(12, int(b.height * 0.22))
        sc = b.crop((int(b.width * 0.03), int(H * 0.12),
                     int(b.width * 0.40), int(H * 0.95)))
        sc_imgs.append(ImageOps.expand(
            sc.resize((sc.width * 4, sc.height * 4), Image.LANCZOS), 25, 255))
    serial_txt = ocr_batch(sc_imgs, ['--oem', '3', '--psm', '7',
                                     '-c', 'tessedit_char_whitelist=0123456789'])
    serials = []
    for t in serial_txt:
        d = re.findall(r'\d{1,4}', t)
        serials.append(max(d, key=len) if d else '')

    return [(e, s, r) for (e, r), s in zip(out, serials)]


def _hard_pass(boxes, variants):
    """Vote across whitelist and non-whitelist OCR results for the supplied crop variants."""
    from collections import Counter
    n = len(boxes)
    imgs = []
    for x0f, up in variants:
        for b in boxes:
            H = max(12, int(b.height * 0.22))
            c = b.crop((int(b.width * x0f), 0, b.width, H))
            imgs.append(ImageOps.expand(
                c.resize((c.width * up, c.height * up), Image.LANCZOS), 20, 255))

    WHITE = ['-c', 'tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789']
    arg_sets, is_white = [], []
    for psm in (7, 13, 11, 8):
        for white in (True, False):
            arg_sets.append(['--oem', '3', '--psm', str(psm)] + (WHITE if white else []))
            is_white.append(white)

    wl = [Counter() for _ in range(n)]
    nw = [Counter() for _ in range(n)]
    for texts, white in zip(ocr_batch_multi(imgs, arg_sets), is_white):
        bags = wl if white else nw
        for k, t in enumerate(texts):
            mm = RE_EPIC.search(t.replace(' ', '').upper())
            if mm and RE_EPIC_OK.match(mm.group(1)):
                bags[k % n][mm.group(1)] += 1
    return wl, nw


def _hard_decide(wlc, nwc):
    if not wlc:
        return (nwc.most_common(1)[0][0], False) if nwc else ('', False)
    best = wlc.most_common(1)[0][0]
    if nwc:
        alt = nwc.most_common(1)[0][0]
        if alt != best and alt[:3] == best[:3]:
            return best, True        # the two readings disagree -> verification is required
    return best, False


def read_headers_hard_batch(boxes):
    """
    Batch version of read_header_hard() with the same decision logic, but faster.

    Two stages:
      1. Run both methods on two crops. Boxes where whitelist and non-whitelist
         results agree are accepted.
      2. Run three additional crops only for boxes whose results disagree.

    This usually resolves most boxes with 16 reads instead of all 40. Extra effort is
    spent only on uncertain boxes, preserving accuracy without unnecessary OCR.
    """
    from collections import Counter
    n = len(boxes)
    STAGE1 = ((0.35, 4), (0.45, 4))
    STAGE2 = ((0.35, 6), (0.45, 6), (0.30, 4))

    wl, nw = _hard_pass(boxes, STAGE1)

    baqi = []
    for i in range(n):
        a = wl[i].most_common(1)[0][0] if wl[i] else None
        b = nw[i].most_common(1)[0][0] if nw[i] else None
        if not (a and b and a == b):
            baqi.append(i)

    if baqi:
        w2, n2 = _hard_pass([boxes[i] for i in baqi], STAGE2)
        for j, i in enumerate(baqi):
            wl[i].update(w2[j])
            nw[i].update(n2[j])

    return [_hard_decide(wl[i], nw[i]) for i in range(n)]




def page_epics(page, cuts, bands_by_col, dpi):
    """
    Extract all EPIC values with one full-page OCR pass and associate each value
    with its voter box using its position on the page.

    Why: isolated boxes give Tesseract less surrounding layout context and can cause
    digit confusion. Full-page OCR preserves the page layout. In the benchmark,
    full-page matching was more accurate than box-only OCR while requiring one OCR call.
    """
    tmp = tempfile.mkdtemp(prefix='fullpage_')
    try:
        fp = os.path.join(tmp, 'page.png')
        page.save(fp, 'PNG', dpi=(dpi, dpi))
        d = pytesseract.image_to_data(fp, output_type=pytesseract.Output.DICT)
    except Exception:
        return {}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    out = {}
    for i, w in enumerate(d['text']):
        w = (w or '').strip().upper()
        if not RE_EPIC_OK.match(w):
            continue
        x = d['left'][i] + d['width'][i] // 2
        y = d['top'][i] + d['height'][i] // 2
        col = 0 if x < cuts[1] else (1 if x < cuts[2] else 2)
        for r, (y0, y1) in enumerate(bands_by_col[col]):
            if y0 <= y <= y1:
                out[(r, col)] = w
                break
    return out


def parse_body(text):
    """
    Read the lines inside a voter box. In ECI PDFs, a long name or address may
    continue onto the next line (example: "ARSHAD MOHAMMAD MUKHATAR / SHAIKH").
    An unlabeled continuation line is appended to the previous field so the value
    is not truncated.
    """
    rec = {'Name': '', 'Relation Type': '', 'Relative Name': '',
           'House Number': '', 'Age': '', 'Gender': ''}
    last = None
    for raw in text.split('\n'):
        line = raw.strip()
        if not line:
            continue

        hit = False
        for rx, rel in ((RE_FATHER, 'Father'), (RE_HUSBAND, 'Husband'),
                        (RE_MOTHER, 'Mother'), (RE_OTHER, 'Other')):
            mm = rx.search(line)
            if mm:
                rec['Relation Type'], rec['Relative Name'] = rel, clean(mm.group(1))
                last, hit = 'Relative Name', True
                break
        if hit:
            continue

        mm = RE_HOUSE.search(line)
        if mm:
            rec['House Number'] = clean(mm.group(1))
            last = 'House Number'
            continue

        mm = RE_NAME.match(line)
        if mm and not rec['Name']:
            rec['Name'] = clean(mm.group(1))
            last = 'Name'
            continue

        if re.search(r'age|gender', line, re.I):
            a, g = RE_AGE.search(line), RE_GENDER.search(line)
            if a:
                rec['Age'] = a.group(1)
            if g:
                gv = g.group(1).upper()
                rec['Gender'] = {'M': 'Male', 'F': 'Female'}.get(gv, g.group(1).title())
            last = None
            continue

        # an unlabeled line is treated as a continuation of the previous field
        tail = clean(line)
        if last and tail and re.search(r'[A-Za-z0-9]', tail) and len(tail) > 1:
            rec[last] = (rec[last] + ' ' + tail).strip()

    return rec


def read_page_context(page, ctx):
    top = page.crop((0, 0, page.width, int(page.height * 0.09)))
    bot = page.crop((0, int(page.height * 0.93), page.width, page.height))
    for txt in ocr_batch([top, bot], BODY_ARGS):
        for line in txt.split('\n'):
            for rx, key in ((RE_PART, 'Part No'), (RE_AC, 'Assembly Constituency'),
                            (RE_PAGE, 'Page No')):
                m = rx.search(line.strip())
                if m:
                    val = clean(m.group(1))
                    val = re.split(r'Part\s*No', val, flags=re.I)[0].strip(' .:')
                    ctx[key] = val
    return ctx


def assign_serials(df):
    """Assign authoritative Sr No from physical PDF order, never from OCR."""
    df = df.copy()
    df['Serial (OCR)'] = df['Serial No'].astype(str)
    df = df.sort_values(['Source File', 'Part No', 'PDF Page', 'Grid Row', 'Grid Col'])
    df['Sr No'] = ''
    for key, g in df.groupby(['Source File', 'Part No'], sort=False):
        for n, (i, _) in enumerate(g.iterrows(), start=1):
            df.at[i, 'Sr No'] = str(n)
    return df


def audit(r):
    """Field-level validation. Returns (status, reasons, remark)."""
    reasons = []
    e = str(r.get('EPIC / Voter ID', '') or '').strip()
    if not e:
        reasons.append('EPIC missing')
    elif not RE_EPIC_OK.match(e):
        reasons.append('EPIC format suspicious')

    nm = str(r.get('Name') or '').strip()
    if not nm or not re.search(r'[A-Za-z]{2}', nm):
        reasons.append('Name missing or unreadable')

    if not str(r.get('Relative Name') or '').strip():
        reasons.append('Relative Name missing')

    a = str(r.get('Age', '') or '').strip()
    if not a.isdigit():
        reasons.append('Age missing or unreadable')
    elif not (18 <= int(a) <= 120):
        reasons.append('Age outside expected range')

    if r.get('Gender') not in ('Male', 'Female', 'Other'):
        reasons.append('Gender missing or unreadable')

    # EPIC Check is a real validation result, not a note about which
    # OCR source supplied the EPIC.  A full-page EPIC extraction is accepted
    # when it matches the expected EPIC pattern; source/method notes belong
    # in Remark, not in the validation field.
    epic_check = str(r.get('EPIC Check', '') or '').strip()
    if epic_check == 'Suspicious':
        reasons.append('EPIC requires verification')

    # Sr No is authoritative from physical PDF/grid order.  OCR serial is
    # diagnostic only and must NEVER make an otherwise valid voter
    # Suspicious.  This prevents cases such as calculated 802 vs OCR 809
    # from corrupting the authoritative serial or creating false reviews.

    if reasons:
        return 'Suspicious', '; '.join(reasons), 'Manual verification recommended.'
    return 'OK', '', ''


def process_page(job):  # noqa
    """One page per worker. Collect all boxes on the page first, then OCR them in batches."""
    path, pdf_name, pno, dpi = job
    t0 = time.time()
    page = convert_from_path(path, dpi=dpi, first_page=pno, last_page=pno,
                             poppler_path=POPPLER_PATH or None, grayscale=True)[0]
    ctx = {'Part No': '', 'Assembly Constituency': '', 'Page No': ''}
    ctx = read_page_context(page, ctx)
    cuts = find_column_cuts(page)
    min_h = int(page.height * 0.02)

    # ---- 1) collect all boxes on the page first (no OCR yet)
    boxes, meta = [], []
    bands_by_col = {}
    for col in range(3):
        strip = page.crop((cuts[col], 0, cuts[col + 1], page.height))
        keep_bands = [(y0, y1) for y0, y1 in find_box_bands(strip, min_h)
                      if (y1 - y0) <= page.height * MAX_BOX_FRAC]
        bands_by_col[col] = keep_bands
        for row, (y0, y1) in enumerate(keep_bands):
            b = strip.crop((0, y0, strip.width, y1))
            if BOX_UPSCALE > 1:
                b = b.resize((b.width * BOX_UPSCALE, b.height * BOX_UPSCALE), Image.LANCZOS)
            boxes.append(b)
            meta.append((row, col))

    # full-page EPIC extraction (one OCR call) - the most reliable source
    page_size = (page.width, page.height)
    pg_epic = page_epics(page, cuts, bands_by_col, dpi) if USE_FULLPAGE_EPIC else {}

    # ---- 2) OCR all text in batches
    bodies = [parse_body(t) for t in ocr_batch(boxes, BODY_ARGS)]
    heads = read_headers_batch(boxes)

    # ---- 3) keep only genuine voter boxes (non-voter regions are removed here)
    keep = []
    for i, (body, (epic, serial, raazi)) in enumerate(zip(bodies, heads)):
        box_ok = bool(RE_EPIC_OK.match(epic or '')) and raazi
        if not body['Name'] and not box_ok and meta[i] not in pg_epic:
            continue

        # combine the full-page result with the box-level result
        full = pg_epic.get(meta[i], '')
        note = ''
        if full and epic and full != epic:
            epic, note = full, 'taken from full-page OCR (different from box OCR)'
            epic_ok = True
        elif full:
            epic, epic_ok = full, True
        else:
            epic_ok = box_ok
            if box_ok:
                note = 'not found in full-page OCR - taken from box OCR'
        keep.append((i, body, epic, serial, epic_ok, note))

    # ---- 4) run the hard EPIC pass only for boxes without a reliable result
    shaki = [k for k in keep if not k[4]]
    flags = {}
    if shaki:
        for (i, *_), (e2, shak) in zip(
                shaki, read_headers_hard_batch([boxes[k[0]] for k in shaki])):
            flags[i] = (e2, shak)

    rows = []
    for i, body, epic, serial, epic_ok, note in keep:
        # Human-facing EPIC Check is strictly OK/Suspicious.
        # OCR-source details are retained in Remark.
        rec_flag = 'OK' if epic_ok else 'Suspicious'
        method_note = note
        if not epic_ok:
            e2, shak = flags.get(i, ('', False))
            if e2:
                epic = e2
                rec_flag = 'Suspicious'
                method_note = ('EPIC uncertain - verify against the PDF' if shak
                               else 'EPIC reread - verification required')
            else:
                method_note = 'EPIC could not be read'
        row, col = meta[i]
        # Normalized box coordinates: independent of rendered PDF resolution.
        # This is required later for exact PDF highlighting.
        x0 = cuts[col]
        x1 = cuts[col + 1]
        y0, y1 = bands_by_col[col][row]
        # NOTE: the box bands are measured in the rendered image.
        # Normalization keeps the coordinates in 0..1 so the viewer can scale them.
        page_w = float(page_size[0])
        page_h = float(page_size[1])
        pdf_box = [round(x0 / page_w, 6), round(y0 / page_h, 6),
                   round(x1 / page_w, 6), round(y1 / page_h, 6)]
        rec = {'Source File': pdf_name, 'Part No': ctx['Part No'],
               'Assembly Constituency': ctx['Assembly Constituency'],
               'PDF Page': pno, 'PDF Box': str(pdf_box),
               'Grid Row': row, 'Grid Col': col,
               'Serial No': serial, 'EPIC / Voter ID': epic}
        rec.update(body)
        rec['EPIC Check'] = rec_flag
        rows.append(rec)

    del boxes, page
    gc.collect()
    return pno, rows, time.time() - t0


# ------------------------------------------------------------------ INCREMENTAL PROCESSING

def file_sha256(path, chunk_size=1024 * 1024):
    """Fingerprint of the actual PDF file. If a PDF changes while keeping the same filename,
    the previous result must not be reused silently."""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def load_processing_state():
    """PDF fingerprints from previous successful runs."""
    if not os.path.exists(STATE_PATH):
        return {}
    try:
        with open(STATE_PATH, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        print(f"Could not read processing state ({e}) - a new state will be created.")
        return {}


def save_processing_state(state):
    """Processing state is saved only after the Excel workbook is written successfully."""
    tmp = STATE_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2, sort_keys=True)
    os.replace(tmp, STATE_PATH)


def load_old_excel():
    """Preserve manual corrections already present in the existing workbook."""
    if not os.path.exists(OUTPUT_PATH):
        return None
    try:
        old = pd.read_excel(OUTPUT_PATH, sheet_name='All Voters')
        old.columns = [str(c).strip() for c in old.columns]
        return old
    except Exception as e:
        print(f"Could not read existing Excel ({e}) - a new Excel workbook will be created.")
        return None


# ------------------------------------------------------------------ MAIN
if __name__ == '__main__':
    setup_check()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_pdfs = sorted(f for f in os.listdir(FOLDER_PATH) if f.lower().endswith('.pdf'))
    if not all_pdfs:
        raise SystemExit("No PDF files were found in the input folder!")

    old_df = load_old_excel()
    state = load_processing_state()

    # Calculate each PDF hash once. If a PDF is replaced while keeping the same filename,
    # the changed fingerprint will trigger a new OCR run.
    fingerprints = {}
    for pdf in all_pdfs:
        fingerprints[pdf] = file_sha256(os.path.join(FOLDER_PATH, pdf))

    done_pdfs = []
    new_pdfs = []
    for pdf in all_pdfs:
        old = state.get(pdf, {})
        if isinstance(old, dict) and old.get('sha256') == fingerprints[pdf]:
            done_pdfs.append(pdf)
        else:
            new_pdfs.append(pdf)

    if old_df is not None and not state:
        # When upgrading from V1.1 to V1.2 for the first time, the state file may not exist.
        # Re-OCRing the existing Excel data is unnecessary and could overwrite manual
        # corrections, so existing Part No values are treated as already complete.
        existing_parts = set(
            old_df.get('Part No', pd.Series(dtype=str)).dropna().astype(str).str.strip()
        )
        bootstrap = []
        still_new = []
        for pdf in new_pdfs:
            part_hint = re.search(r'(?:part[_ -]?)?(\d{1,4})(?:\.pdf)?$', pdf, re.I)
            # Also handle Part numbers that appear after the language text in ECI filenames.
            m = re.search(r'(?:^|[-_])(\d{1,4})(?:[-_](?:WI|FINAL|DRAFT))?\.pdf$', pdf, re.I)
            part = m.group(1) if m else (part_hint.group(1) if part_hint else '')
            if part and part in existing_parts:
                bootstrap.append(pdf)
            else:
                still_new.append(pdf)
        for pdf in bootstrap:
            state[pdf] = {'sha256': fingerprints[pdf], 'source': 'bootstrap-existing-excel'}
        done_pdfs.extend(bootstrap)
        new_pdfs = still_new
        if bootstrap:
            print(f"{len(bootstrap)} PDFs were treated as already complete based on Part No values in the existing Excel.")

    if done_pdfs:
        print(f"{len(done_pdfs)} PDFs are already processed and will be skipped.")
    if not new_pdfs:
        # The state file can be saved after bootstrapping; the existing Excel is left untouched.
        save_processing_state(state)
        print("No new PDFs found. Existing Master_Voter_List.xlsx was not changed.")
        print(f"File: {OUTPUT_PATH}")
        raise SystemExit(0)

    pdfs = new_pdfs
    print(f"{len(pdfs)} NEW PDFs will be OCR processed. {WORKERS} CPU workers are in use.\n")

    jobs = []
    for pdf in pdfs:
        path = os.path.join(FOLDER_PATH, pdf)
        info = pdfinfo_from_path(path, poppler_path=POPPLER_PATH or None)
        w_pt = float(str(info['Page size']).split()[0])
        dpi = max(72, min(300, round(TARGET_WIDTH_PX * 72 / w_pt)))
        print(f"{pdf}: {info['Pages']} pages, DPI={dpi}")
        jobs += [(path, pdf, n, dpi) for n in range(1, info['Pages'] + 1)]

    def mmss(sec):
        return f"{int(sec) // 60}m {int(sec) % 60:02d}s"

    rows, done = [], 0
    t_start = time.time()
    previous_total = 0.0
    with ProcessPoolExecutor(max_workers=WORKERS) as pool:
        futures = [pool.submit(process_page, j) for j in jobs]
        for fut in as_completed(futures):
            pno, page_rows, page_sec = fut.result()
            rows += page_rows
            done += 1
            beeta = time.time() - t_start
            interval_sec = beeta - previous_total
            previous_total = beeta
            baqi = (beeta / done) * (len(jobs) - done)
            print(f"   [{done}/{len(jobs)}] Page {pno:>3} -> {len(page_rows):>3} voters"
                  f" | On this page: {interval_sec:5.1f}s"
                  f" | Elapsed: {mmss(beeta)} | Estimated remaining: {mmss(baqi)}", flush=True)
    t_ocr = time.time() - t_start

    df_new = pd.DataFrame(rows)
    if df_new.empty:
        raise SystemExit('No voter boxes were extracted. Existing Excel/state was not changed.')

    df_new = assign_serials(df_new)
    audits = df_new.apply(audit, axis=1, result_type='expand')
    audits.columns = ['Status Mark', 'Suspicious Reason', 'Remark']
    df_new = pd.concat([df_new, audits], axis=1)

    # Existing rows are authoritative: manual corrections must not be overwritten by
    # another OCR/audit pass. Only rows from new PDFs are generated fresh.
    if old_df is not None and len(old_df):
        df = pd.concat([old_df, df_new], ignore_index=True, sort=False)
    else:
        df = df_new

    # Internal columns remain in the workbook but are hidden from the user.
    # This keeps Source File available for future incremental runs while keeping the
    # visible worksheet clean.
    for c, default in {
        'Source File': '', 'Grid Row': 0, 'Grid Col': 0, 'Serial (OCR)': ''
    }.items():
        if c not in df.columns:
            df[c] = default

    df['_p'] = pd.to_numeric(df['Part No'], errors='coerce').fillna(0)
    df['_pg'] = pd.to_numeric(df['PDF Page'], errors='coerce').fillna(0)
    df = df.sort_values(['_p', 'Source File', '_pg', 'Grid Row', 'Grid Col'])
    df = df.drop(columns=['_p', '_pg'])
    # Rebuild the display row number instead of inserting a duplicate column.
    # Existing workbooks may already contain a Row column from an earlier run.
    if 'Row' in df.columns:
        df = df.drop(columns=['Row'])
    df.insert(0, 'Row', range(1, len(df) + 1))

    cols = ['Row', 'Sr No', 'EPIC / Voter ID', 'EPIC Check', 'Name',
            'Relation Type', 'Relative Name', 'House Number', 'Age', 'Gender',
            'Status Mark', 'Suspicious Reason', 'Remark', 'Part No',
            'PDF Page', 'PDF Box', 'Source File', 'Grid Row', 'Grid Col', 'Serial (OCR)']
    df = df[[c for c in cols if c in df.columns]]

    # Temporary workbook first. Existing Excel is replaced only after the complete
    # workbook has been written successfully. This prevents a failed run from
    # destroying the previous good workbook.
    tmp_xlsx = OUTPUT_PATH + '.tmp.xlsx'
    with pd.ExcelWriter(tmp_xlsx, engine='openpyxl') as xl:
        df.to_excel(xl, sheet_name='All Voters', index=False)
        bad = df[df['Status Mark'] == 'Suspicious']
        bad.to_excel(xl, sheet_name='Needs Review', index=False)

    # Hide internal machine columns, but keep them in the workbook so future runs
    # can continue incrementally without exposing implementation details to the user.
    from openpyxl import load_workbook
    wb = load_workbook(tmp_xlsx)
    ws = wb['All Voters']
    headers = {cell.value: cell.column for cell in ws[1]}
    for name in ('Source File', 'Grid Row', 'Grid Col', 'Serial (OCR)'):
        if name in headers:
            ws.column_dimensions[ws.cell(1, headers[name]).column_letter].hidden = True
    ws2 = wb['Needs Review']
    headers2 = {cell.value: cell.column for cell in ws2[1]}
    for name in ('Source File', 'Grid Row', 'Grid Col', 'Serial (OCR)'):
        if name in headers2:
            ws2.column_dimensions[ws2.cell(1, headers2[name]).column_letter].hidden = True
    wb.save(tmp_xlsx)
    os.replace(tmp_xlsx, OUTPUT_PATH)

    # Mark these PDFs as successfully processed only after the workbook replacement succeeds.
    for pdf in pdfs:
        state[pdf] = {'sha256': fingerprints[pdf], 'source': 'processed'}
    save_processing_state(state)

    suspicious_n = int((df['Status Mark'] == 'Suspicious').sum())
    ok_n = len(df) - suspicious_n
    kul = time.time() - t_start
    print(f"\nTotal {len(df)} voters | OK {ok_n} | Suspicious {suspicious_n}")
    print(f"New PDF: {len(pdfs)} | already processed: {len(done_pdfs)}")
    print(f"Timing: OCR {mmss(t_ocr)} | total {mmss(kul)}"
          f" | {kul / max(1, len(jobs)):.1f}s per page")
    print(f"File: {OUTPUT_PATH}")
