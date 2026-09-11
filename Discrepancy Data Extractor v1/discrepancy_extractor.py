# -*- coding: utf-8 -*-
"""
Discrepancy Data Extractor V1
Purpose-built for ECI Discrepancy PDFs with the fixed table structure:
S.No | Part Serial Number | EPIC Number | Elector Name | Age | Gender | Reason for discrepancy

Accuracy-first design:
- Uses the PDF's native text layer first (iText-generated PDFs expose exact text + coordinates).
- Does NOT use Draft-list parsing logic.
- Uses physical word coordinates to reconstruct each record and its wrapped fields.
- Stores a normalized PDF bounding box for every record.
- Validates every record and writes a Needs Review sheet.
- Incremental SHA-256 processing prevents silent reuse after a source PDF changes.
"""
import os, re, json, hashlib, time, shutil, glob, sys
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
OUTPUT_PATH = os.path.join(OUTPUT_DIR, 'Discrepancy_Master.xlsx')
STATE_PATH = os.path.join(OUTPUT_DIR, '.processed_pdfs.json')

# The uploaded PDF is A4 landscape (842 x 595 pt). The structure is fixed by the
# Discrepancy PDF format, so these are intentionally explicit column zones.
# Boundaries are chosen between the stable header columns, not from OCR guesses.
COLS = {
    'S.No': (55, 125),
    'Part Serial Number': (125, 265),
    'EPIC Number': (265, 365),
    'Elector Name': (365, 505),
    'Age': (505, 585),
    'Gender': (585, 685),
    'Reason for discrepancy': (685, 842),
}

RE_EPIC = re.compile(r'^[A-Z]{3}\d{7}$')
RE_INT = re.compile(r'^\d{1,4}$')
RE_GENDER = re.compile(r'^[MF]$')
RE_PART_HEADER = re.compile(r'Part\s+No\s+and\s+Name\s*:\s*(\d+)\s*-\s*(.+)$', re.I)
RE_AC_HEADER = re.compile(r'AC\s+No\s+and\s+Name\s*:\s*(\d+)\s*-\s*(.+)$', re.I)

OUTPUT_COLUMNS = [
    'S.No', 'Part Serial Number', 'EPIC Number', 'Elector Name', 'Age',
    'Gender', 'Reason for discrepancy', 'Part No', 'PDF Page', 'PDF Box',
    'Source File', 'EPIC Check', 'Status Mark', 'Suspicious Reason', 'Remark'
]


def clean(s):
    s = re.sub(r'\s+', ' ', str(s or '')).strip()
    while s[:1] in ('=', '+', '@'):
        s = s[1:].lstrip()
    return s


def sha256(path, chunk=1024*1024):
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


def parse_page(page, pno, source_file):
    """Parse one PDF page using native text coordinates."""
    words = page.get_text('words', sort=True)
    # tuple: x0,y0,x1,y1,text,block_no,line_no,word_no
    words = [w for w in words if clean(w[4])]

    # Extract page metadata from the first ~100 pt, where headers occur.
    part_no = ''
    ac_no = ''
    ac_name = ''
    part_name = ''
    top_words = [w for w in words if w[1] < 100]
    top_lines = {}
    for w in top_words:
        key = round(w[1], 1)
        top_lines.setdefault(key, []).append(w)
    for _, ws in sorted(top_lines.items()):
        txt = clean(' '.join(w[4] for w in sorted(ws, key=lambda z: z[0])))
        m = RE_PART_HEADER.search(txt)
        if m:
            part_no, part_name = m.group(1), clean(m.group(2))
        m = RE_AC_HEADER.search(txt)
        if m:
            ac_no, ac_name = m.group(1), clean(m.group(2))

    # Find record starts: numeric S.No in the first column. We deliberately use
    # coordinate + numeric shape, so numeric values in other columns cannot start rows.
    starts = []
    for w in words:
        x0,y0,x1,y1,t,*_ = w
        t = clean(t)
        cx = (x0+x1)/2
        if 70 <= cx <= 115 and RE_INT.fullmatch(t) and y0 > 60:
            starts.append(w)
    # de-duplicate if PDF word stream ever repeats a word.
    uniq = []
    seen = set()
    for w in sorted(starts, key=lambda z: (z[1], z[0])):
        k = (round(w[0],2), round(w[1],2), w[4])
        if k not in seen:
            seen.add(k); uniq.append(w)
    starts = uniq

    rows = []
    for idx, sw in enumerate(starts):
        y_start = sw[1] - 2.0
        y_end = starts[idx+1][1] - 2.0 if idx+1 < len(starts) else page.rect.height - 20
        # Include all text belonging to this record until the next S.No.
        rw = [w for w in words if w[1] >= y_start and w[1] < y_end]
        # Prevent accidental header capture and keep only table columns.
        by_col = {k: [] for k in COLS}
        for w in rw:
            cx = (w[0]+w[2])/2
            c = col_for_x(cx)
            if c:
                by_col[c].append(w)

        def text_col(name):
            ws = by_col[name]
            if not ws:
                return ''
            # Preserve visual line order. Within a line, preserve x order.
            lines = {}
            for w in ws:
                # Native PDF line positions differ slightly between columns; y rounding
                # at 0.8 pt is enough to merge words on the same visual line.
                key = round(w[1] / 0.8) * 0.8
                lines.setdefault(key, []).append(w)
            return clean(' '.join(clean(' '.join(x[4] for x in sorted(v, key=lambda z:z[0])))
                          for _, v in sorted(lines.items())))

        rec = {
            'S.No': clean(text_col('S.No')),
            'Part Serial Number': clean(text_col('Part Serial Number')),
            'EPIC Number': clean(text_col('EPIC Number')).replace(' ', ''),
            'Elector Name': text_col('Elector Name'),
            'Age': clean(text_col('Age')),
            'Gender': clean(text_col('Gender')).upper(),
            'Reason for discrepancy': text_col('Reason for discrepancy'),
            'Part No': part_no,
            'PDF Page': pno,
            'Source File': source_file,
        }

        # Bounding box of the complete record's table content. Normalize to 0..1.
        # We use the actual record words, not guessed row heights.
        useful = [w for w in rw if col_for_x((w[0]+w[2])/2)]
        if useful:
            x0 = min(w[0] for w in useful); y0 = min(w[1] for w in useful)
            x1 = max(w[2] for w in useful); y1 = max(w[3] for w in useful)
            # Add a small margin so the box visually covers the whole record row.
            x0 = max(0, x0 - 2); y0 = max(0, y0 - 2)
            x1 = min(page.rect.width, x1 + 2); y1 = min(page.rect.height, y1 + 2)
            rec['PDF Box'] = str([round(x0/page.rect.width, 6), round(y0/page.rect.height, 6),
                                  round(x1/page.rect.width, 6), round(y1/page.rect.height, 6)])
        else:
            rec['PDF Box'] = ''
        rec['_y'] = sw[1]
        rows.append(rec)

    return rows, {'part_no': part_no, 'ac_no': ac_no, 'ac_name': ac_name, 'part_name': part_name}


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
    if not rec['Part Serial Number'] or not RE_INT.fullmatch(rec['Part Serial Number']):
        reasons.append('Part Serial Number missing or invalid')
    if not rec['Elector Name'] or len(re.sub(r'[^A-Za-z]', '', rec['Elector Name'])) < 2:
        reasons.append('Elector Name missing or unreadable')
    if not rec['Age'] or not re.fullmatch(r'\d{1,3}', rec['Age']):
        reasons.append('Age missing or invalid')
    elif not 18 <= int(rec['Age']) <= 120:
        reasons.append('Age outside expected range')
    if not RE_GENDER.fullmatch(rec['Gender']):
        reasons.append('Gender missing or invalid')
    if not rec['Reason for discrepancy']:
        reasons.append('Reason for discrepancy missing')
    if not rec['Part No']:
        reasons.append('Part No missing')
    if not rec['PDF Box']:
        reasons.append('PDF Box missing')

    status = 'Suspicious' if reasons else 'OK'
    suspicious = '; '.join(reasons)
    remark = 'Manual verification recommended.' if reasons else ''
    return epic_check, status, suspicious, remark


def process_pdf(job):
    path, filename = job
    doc = fitz.open(path)
    rows = []
    meta = {'part_no':'','ac_no':'','ac_name':'','part_name':''}
    for pno, page in enumerate(doc, 1):
        page_rows, page_meta = parse_page(page, pno, filename)
        # Metadata normally appears only on page 1. Carry it forward to every
        # record so Part No is populated on continuation pages.
        for k in meta:
            if page_meta.get(k) and not meta[k]:
                meta[k] = page_meta[k]
        for r in page_rows:
            if not r.get('Part No'):
                r['Part No'] = meta['part_no']
        rows.extend(page_rows)
    doc.close()
    return filename, rows, meta


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

    if done and os.path.exists(OUTPUT_PATH):
        old = pd.read_excel(OUTPUT_PATH, sheet_name='All Discrepancies')
        # Keep already-processed rows. New PDFs are appended below.
    else:
        old = pd.DataFrame(columns=OUTPUT_COLUMNS)

    if not new:
        print(f'No new/changed PDFs. Existing output preserved: {OUTPUT_PATH}')
        return

    jobs = [(os.path.join(INPUT_FOLDER, f), f) for f in new]
    workers = max(1, min(4, (os.cpu_count() or 2)-1))
    all_rows = []
    t0 = time.time()
    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(process_pdf, j) for j in jobs]
        for fut in as_completed(futures):
            filename, rows, meta = fut.result()
            print(f'{filename}: extracted {len(rows)} records | Part {meta.get("part_no","")}')
            all_rows.extend(rows)

    if not all_rows:
        raise SystemExit('No records extracted. Existing workbook/state was not changed.')

    new_df = pd.DataFrame(all_rows)
    audits = new_df.apply(validate, axis=1, result_type='expand')
    audits.columns = ['EPIC Check','Status Mark','Suspicious Reason','Remark']
    new_df = pd.concat([new_df, audits], axis=1)

    # File-level structural validation. A missing/duplicate S.No is an extraction
    # failure even if every individual row looks syntactically valid.
    for source, idxs in new_df.groupby('Source File').groups.items():
        nums = pd.to_numeric(new_df.loc[idxs, 'S.No'], errors='coerce').dropna().astype(int).tolist()
        expected = list(range(1, len(nums) + 1))
        if nums != expected:
            msg = 'S.No sequence has a gap, duplicate, or ordering error'
            for i in idxs:
                new_df.at[i, 'Status Mark'] = 'Suspicious'
                new_df.at[i, 'Suspicious Reason'] = (
                    (str(new_df.at[i, 'Suspicious Reason']).strip() + '; ' if str(new_df.at[i, 'Suspicious Reason']).strip() else '') + msg
                )
                new_df.at[i, 'Remark'] = 'Manual verification recommended.'

    # Remove parser internals and enforce exact user-facing column order.
    new_df = new_df[OUTPUT_COLUMNS]
    df = pd.concat([old, new_df], ignore_index=True, sort=False)

    # Preserve source ordering: Part No, source file, PDF page, then S.No.
    df['_part'] = pd.to_numeric(df['Part No'], errors='coerce').fillna(0)
    df['_page'] = pd.to_numeric(df['PDF Page'], errors='coerce').fillna(0)
    df['_sno'] = pd.to_numeric(df['S.No'], errors='coerce').fillna(0)
    df = df.sort_values(['_part','Source File','_page','_sno'], kind='stable')
    df = df[OUTPUT_COLUMNS]

    tmp = OUTPUT_PATH + '.tmp.xlsx'
    with pd.ExcelWriter(tmp, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='All Discrepancies', index=False)
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
    print(f'\nTotal records: {len(df)} | OK: {len(df)-suspicious} | Suspicious: {suspicious}')
    print(f'New PDFs: {len(new)} | Already processed: {len(done)}')
    print(f'Output: {OUTPUT_PATH}')
    print(f'Time: {time.time()-t0:.2f}s')

if __name__ == '__main__':
    main()
