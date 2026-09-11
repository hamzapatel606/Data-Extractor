# -*- coding: utf-8 -*-
"""
ASDD Data Extractor V1
Purpose-built for ECI ASDD PDFs with the fixed table structure:
S.No | Serial No. | EPIC Number | Elector Name | Relative Details | DOB/Age | Uncollectable Reason

Accuracy-first design:
- Uses the PDF's native text layer and physical word coordinates.
- Uses ASDD-specific portrait-page column zones (595 x 842 pt).
- Reconstructs wrapped fields using visual line positions.
- Handles records whose final wrapped line continues onto the next PDF page.
- Preserves the full "Already enrolled (EPIC)" text from the source.
- Stores a normalized PDF bounding box for every record's starting page.
- Validates every record and writes a Needs Review sheet.
- Incremental SHA-256 processing replaces rows when a source PDF changes.
"""
import os
import re
import json
import hashlib
import time
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed

try:
    import pymupdf as fitz
    import pandas as pd
    from openpyxl import load_workbook
except ImportError as e:
    raise SystemExit(
        f"Missing Python package: {e.name}\n\n"
        f'Install with:\n  "{sys.executable}" -m pip install pymupdf pandas openpyxl')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(SCRIPT_DIR, 'input')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
OUTPUT_PATH = os.path.join(OUTPUT_DIR, 'ASDD_Master.xlsx')
STATE_PATH = os.path.join(OUTPUT_DIR, '.processed_pdfs.json')

# ASDD PDF is portrait A4: 595.28 x 841.89 pt.
# These boundaries are based on the stable x-coordinates of the native PDF text.
COLS = {
    'S.No': (20, 50),
    'Serial No': (50, 80),
    'EPIC Number': (80, 145),
    'Elector Name': (145, 295),
    'Relative Details': (295, 445),
    'DOB/Age': (445, 490),
    'Uncollectable Reason': (490, 595),
}

# The sample contains standard 3-letter + 7-digit EPICs, plus source records
# such as DXHO582668 and AJNPA1159L. All are 10-character uppercase alphanumerics.
RE_EPIC = re.compile(r'^[A-Z0-9]{10}$')
RE_INT = re.compile(r'^\d{1,4}$')
RE_AGE = re.compile(r'^\d{1,3}$')
RE_HEADER = re.compile(
    r'AC\s*:\s*(\d+)\s*-\s*([^;]+?)\s*;\s*Part\s*:\s*(\d+)\s*-\s*(.+)$',
    re.I,
)

OUTPUT_COLUMNS = [
    'S.No', 'Serial No', 'EPIC Number', 'Elector Name', 'Relative Details',
    'DOB/Age', 'Uncollectable Reason', 'Part No', 'PDF Page', 'PDF Box',
    'Source File', 'EPIC Check', 'Status Mark', 'Suspicious Reason', 'Remark'
]

TABLE_TOP = 50.0
TABLE_BOTTOM = 805.0
ROW_MARGIN = 2.0
LINE_BUCKET = 0.8


def clean(s):
    s = re.sub(r'\s+', ' ', str(s or '')).strip()
    # Prevent Excel formula injection while preserving normal voter text.
    while s[:1] in ('=', '+', '@'):
        s = s[1:].lstrip()
    return s


def sha256(path, chunk=1024 * 1024):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def load_state():
    if not os.path.exists(STATE_PATH):
        return {}
    try:
        with open(STATE_PATH, encoding='utf-8') as f:
            x = json.load(f)
        return x if isinstance(x, dict) else {}
    except Exception:
        return {}


def save_state(state):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    tmp = STATE_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2, sort_keys=True)
    os.replace(tmp, STATE_PATH)


def col_for_x(cx):
    for name, (lo, hi) in COLS.items():
        if lo <= cx < hi:
            return name
    return None


def line_text(words):
    """Return words in visual line order, grouped by native y-position."""
    lines = {}
    for w in words:
        key = round(w[1] / LINE_BUCKET) * LINE_BUCKET
        lines.setdefault(key, []).append(w)
    parts = []
    for _, ws in sorted(lines.items()):
        parts.append(' '.join(clean(w[4]) for w in sorted(ws, key=lambda z: z[0])))
    return clean(' '.join(parts))


def extract_metadata(words):
    ac_no = ac_name = part_no = part_name = ''
    # Header is on page 1 only in the sample, but parsing every page is harmless.
    top = [w for w in words if w[1] < 50]
    lines = {}
    for w in top:
        key = round(w[1], 1)
        lines.setdefault(key, []).append(w)
    for _, ws in sorted(lines.items()):
        txt = clean(' '.join(w[4] for w in sorted(ws, key=lambda z: z[0])))
        m = RE_HEADER.search(txt)
        if m:
            ac_no, ac_name, part_no, part_name = m.group(1), clean(m.group(2)), m.group(3), clean(m.group(4))
            break
    return {'ac_no': ac_no, 'ac_name': ac_name, 'part_no': part_no, 'part_name': part_name}


def find_starts(words):
    starts = []
    for w in words:
        x0, y0, x1, y1, t, *_ = w
        t = clean(t)
        cx = (x0 + x1) / 2
        # Only the first column can begin a record. Exclude footer page numbers.
        if 20 <= cx < 50 and TABLE_TOP < y0 < TABLE_BOTTOM and RE_INT.fullmatch(t):
            starts.append(w)
    uniq = []
    seen = set()
    for w in sorted(starts, key=lambda z: (z[1], z[0])):
        k = (round(w[0], 2), round(w[1], 2), w[4])
        if k not in seen:
            seen.add(k)
            uniq.append(w)
    return uniq


def words_to_record(rw, pno, page):
    by_col = {k: [] for k in COLS}
    for w in rw:
        cx = (w[0] + w[2]) / 2
        c = col_for_x(cx)
        if c:
            by_col[c].append(w)

    rec = {k: line_text(v) for k, v in by_col.items()}
    rec['EPIC Number'] = rec['EPIC Number'].replace(' ', '').upper()
    rec['DOB/Age'] = rec['DOB/Age'].strip('() ')
    rec['PDF Page'] = pno

    useful = [w for w in rw if col_for_x((w[0] + w[2]) / 2)]
    if useful:
        x0 = min(w[0] for w in useful)
        y0 = min(w[1] for w in useful)
        x1 = max(w[2] for w in useful)
        y1 = max(w[3] for w in useful)
        x0 = max(0, x0 - ROW_MARGIN)
        y0 = max(0, y0 - ROW_MARGIN)
        x1 = min(page.rect.width, x1 + ROW_MARGIN)
        y1 = min(page.rect.height, y1 + ROW_MARGIN)
        rec['PDF Box'] = str([
            round(x0 / page.rect.width, 6),
            round(y0 / page.rect.height, 6),
            round(x1 / page.rect.width, 6),
            round(y1 / page.rect.height, 6),
        ])
    else:
        rec['PDF Box'] = ''
    return rec


def parse_page(page, pno):
    words = [w for w in page.get_text('words', sort=True) if clean(w[4])]
    starts = find_starts(words)
    rows = []

    for idx, sw in enumerate(starts):
        y_start = sw[1] - ROW_MARGIN
        y_end = starts[idx + 1][1] - ROW_MARGIN if idx + 1 < len(starts) else TABLE_BOTTOM
        rw = [w for w in words if y_start <= w[1] < y_end and w[1] < TABLE_BOTTOM]
        rows.append(words_to_record(rw, pno, page))

    # Some wrapped rows end exactly at a page boundary. Their remaining words
    # appear above the first numbered row on the next page, e.g. S.No 104 -> 105,
    # S.No 247 -> 248 in the supplied ASDD PDF. Return only genuine table content,
    # never the repeated header.
    continuation = {k: '' for k in COLS}
    if starts:
        first_y = starts[0][1]
        pre = [w for w in words if TABLE_TOP < w[1] < first_y - 1]
        for w in pre:
            c = col_for_x((w[0] + w[2]) / 2)
            if c:
                continuation[c] = clean((continuation[c] + ' ' + w[4]).strip())
        # Header words sit at y 33-43, below TABLE_TOP, so they are excluded.
    return rows, continuation


def merge_continuation(previous, continuation):
    """Merge page-leading continuation text into the previous record."""
    for field in ('Elector Name', 'Relative Details', 'DOB/Age', 'Uncollectable Reason'):
        extra = clean(continuation.get(field, ''))
        if extra:
            previous[field] = clean(f"{previous.get(field, '')} {extra}")
    previous['EPIC Number'] = clean(previous.get('EPIC Number', '')).replace(' ', '').upper()
    previous['DOB/Age'] = previous.get('DOB/Age', '').strip('() ')


def process_pdf(job):
    path, filename = job
    doc = fitz.open(path)
    rows = []
    meta = {'ac_no': '', 'ac_name': '', 'part_no': '', 'part_name': ''}
    pending_previous = None

    try:
        for pno, page in enumerate(doc, 1):
            words = [w for w in page.get_text('words', sort=True) if clean(w[4])]
            page_meta = extract_metadata(words)
            for k in meta:
                if page_meta.get(k) and not meta[k]:
                    meta[k] = page_meta[k]

            page_rows, continuation = parse_page(page, pno)

            # Merge continuation from this page into the final row from the prior page.
            if continuation and any(continuation[k] for k in ('Elector Name', 'Relative Details', 'DOB/Age', 'Uncollectable Reason')):
                if rows:
                    merge_continuation(rows[-1], continuation)
                else:
                    # A continuation without a previous row is an extraction anomaly.
                    pass

            for r in page_rows:
                r['Part No'] = meta['part_no']
                r['Source File'] = filename
            rows.extend(page_rows)
    finally:
        doc.close()

    return filename, rows, meta


def validate(rec):
    reasons = []
    epic = rec['EPIC Number']
    if not epic:
        reasons.append('EPIC missing')
        epic_check = 'Suspicious'
    elif not RE_EPIC.fullmatch(epic):
        reasons.append('EPIC format suspicious')
        epic_check = 'Suspicious'
    else:
        epic_check = 'OK'

    if not rec['S.No'] or not RE_INT.fullmatch(rec['S.No']):
        reasons.append('S.No missing or invalid')
    if not rec['Serial No'] or not RE_INT.fullmatch(rec['Serial No']):
        reasons.append('Serial No missing or invalid')
    if not rec['Elector Name'] or len(re.sub(r'[^A-Za-z]', '', rec['Elector Name'])) < 2:
        reasons.append('Elector Name missing or unreadable')
    if not rec['Relative Details'] or len(re.sub(r'[^A-Za-z]', '', rec['Relative Details'])) < 2:
        reasons.append('Relative Details missing or unreadable')
    if not rec['DOB/Age'] or not RE_AGE.fullmatch(rec['DOB/Age']):
        reasons.append('DOB/Age missing or invalid')
    elif not 18 <= int(rec['DOB/Age']) <= 120:
        reasons.append('Age outside expected range')
    if not rec['Uncollectable Reason']:
        reasons.append('Uncollectable Reason missing')
    if not rec['Part No']:
        reasons.append('Part No missing')
    if not rec['PDF Box']:
        reasons.append('PDF Box missing')

    status = 'Suspicious' if reasons else 'OK'
    suspicious = '; '.join(reasons)
    remark = 'Manual verification recommended.' if reasons else ''
    return epic_check, status, suspicious, remark


def read_existing():
    if not os.path.exists(OUTPUT_PATH):
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        return pd.read_excel(OUTPUT_PATH, sheet_name='All ASDD')
    except Exception:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def main():
    os.makedirs(INPUT_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    pdfs = sorted(f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith('.pdf'))
    if not pdfs:
        raise SystemExit('No PDF files found in input folder.')

    state = load_state()
    fingerprints = {f: sha256(os.path.join(INPUT_FOLDER, f)) for f in pdfs}
    new = [f for f in pdfs if state.get(f, {}).get('sha256') != fingerprints[f]]
    done = [f for f in pdfs if f not in new]

    if not new:
        print(f'No new/changed PDFs. Existing output preserved: {OUTPUT_PATH}')
        return

    old = read_existing()
    if not old.empty and 'Source File' in old.columns:
        # If a source PDF changed, remove its old rows before appending the fresh extraction.
        old = old[~old['Source File'].isin(new)].copy()

    jobs = [(os.path.join(INPUT_FOLDER, f), f) for f in new]
    workers = max(1, min(4, (os.cpu_count() or 2) - 1))
    all_rows = []
    t0 = time.time()

    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(process_pdf, j) for j in jobs]
        for fut in as_completed(futures):
            filename, rows, meta = fut.result()
            print(f'{filename}: extracted {len(rows)} records | AC {meta.get("ac_no", "")} | Part {meta.get("part_no", "")}')
            all_rows.extend(rows)

    if not all_rows:
        raise SystemExit('No records extracted. Existing workbook/state was not changed.')

    new_df = pd.DataFrame(all_rows)
    audits = new_df.apply(validate, axis=1, result_type='expand')
    audits.columns = ['EPIC Check', 'Status Mark', 'Suspicious Reason', 'Remark']
    new_df = pd.concat([new_df, audits], axis=1)

    # File-level S.No validation. This catches missed/duplicated rows even when
    # every individual row is syntactically valid.
    for source, idxs in new_df.groupby('Source File').groups.items():
        nums = pd.to_numeric(new_df.loc[idxs, 'S.No'], errors='coerce').dropna().astype(int).tolist()
        expected = list(range(1, len(nums) + 1))
        if nums != expected:
            msg = 'S.No sequence has a gap, duplicate, or ordering error'
            for i in idxs:
                new_df.at[i, 'Status Mark'] = 'Suspicious'
                prior = str(new_df.at[i, 'Suspicious Reason'] or '').strip()
                new_df.at[i, 'Suspicious Reason'] = f'{prior}; {msg}' if prior else msg
                new_df.at[i, 'Remark'] = 'Manual verification recommended.'

    new_df = new_df[OUTPUT_COLUMNS]
    df = pd.concat([old, new_df], ignore_index=True, sort=False)

    # Preserve part/source/page/S.No ordering.
    df['_part'] = pd.to_numeric(df['Part No'], errors='coerce').fillna(0)
    df['_page'] = pd.to_numeric(df['PDF Page'], errors='coerce').fillna(0)
    df['_sno'] = pd.to_numeric(df['S.No'], errors='coerce').fillna(0)
    df = df.sort_values(['_part', 'Source File', '_page', '_sno'], kind='stable')
    df = df[OUTPUT_COLUMNS]

    tmp = OUTPUT_PATH + '.tmp.xlsx'
    with pd.ExcelWriter(tmp, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='All ASDD', index=False)
        df[df['Status Mark'].eq('Suspicious')].to_excel(writer, sheet_name='Needs Review', index=False)

    wb = load_workbook(tmp)
    for ws in wb.worksheets:
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = min(60, max(len(str(c.value or '')) for c in col) + 2)
            ws.column_dimensions[col[0].column_letter].width = max_len
    wb.save(tmp)
    os.replace(tmp, OUTPUT_PATH)

    for f in new:
        state[f] = {'sha256': fingerprints[f], 'source': 'processed'}
    save_state(state)

    suspicious = int(df['Status Mark'].eq('Suspicious').sum())
    print(f'\nTotal records: {len(df)} | OK: {len(df) - suspicious} | Suspicious: {suspicious}')
    print(f'New PDFs: {len(new)} | Already processed: {len(done)}')
    print(f'Output: {OUTPUT_PATH}')
    print(f'Time: {time.time() - t0:.2f}s')


if __name__ == '__main__':
    main()
