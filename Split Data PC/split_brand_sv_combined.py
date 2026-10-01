"""
PISAH PER BRAND (GMV + NON GMV DIGABUNG)  [SHORT VIDEO - TAP VID PC]
================================================================================
Baca tab "Custom report" (spreadsheet TAP VID PC), lalu bikin 1 tab per brand
yang isinya video UNIK. GMV dan NON GMV ada di tab yang sama, dibedakan lewat
kolom "Jenis" di paling kanan. Padanan SV dari split_brand_live_combined.py.

ATURAN:
  - Semua baris di "Custom report" diambil, TANPA filter bulan/tanggal.
  - Brand dinormalisasi ke HURUF BESAR (Pepsodent == PEPSODENT).
  - Unik per (brand, Video ID). Satu Video ID bisa muncul di beberapa brand ->
    tiap brand dapat barisnya sendiri.
  - Kolom yang DIJUMLAH untuk baris dengan Video ID sama dalam 1 brand:
      Affiliate video-attributed GMV, Creator video-attributed orders,
      Creator-attributed items sold.
  - Kolom lain diambil dari baris pertama Video ID itu, KECUALI 'Date' yang
    diambil tanggal paling awal.
  - Jenis = GMV (total GMV > 0) / NON GMV (total GMV = 0).
  - Baris diurutkan dari GMV terbesar, jadi yang GMV ada di atas.
  - Baris yang kolomnya bergeser (kolom 'Post time' bukan format
    YYYY-MM-DD HH:MM:SS) DILEWATI dan nomor barisnya ditampilkan, supaya angka
    yang salah kolom tidak ikut dijumlah.

NAMA TAB: "<BRAND><TAB_SUFFIX>", default "PEPSODENT - UNIK". Suffix dipakai
supaya tidak bentrok dengan tab brand yang sudah ada di spreadsheet. Tab yang
sudah ada TIDAK ditimpa: kalau ada yang bentrok, script berhenti dan menyebut
tab mana.

CARA PAKAI:
  python split_brand_sv_combined.py            # DRY RUN: cuma baca & ringkas
  python split_brand_sv_combined.py --write    # bikin tab & tulis (butuh Editor)
"""

import re
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============== KONFIGURASI ==============
CREDENTIALS_FILE = 'credentials.json'
SPREADSHEET_ID = '1fQLKulAy4PGEzaBCMN6RM4BEIW7HTwDMmra01skVRuA'  # TAP VID PC 24-30 SEP 2026
SOURCE_SHEET = 'Custom report'
TAB_SUFFIX = ' - UNIK'
MERGE_BRANDS = {}  # misal {'DOVE DEO': 'DOVE', 'DOVE SCL': 'DOVE'} kalau mau digabung
WRITE_CHUNK_ROWS = 1500   # baris per request tulis (biar payload tidak kebesaran)
MAX_CELLS = 10_000_000    # batas sel per spreadsheet di Google Sheets
# ===========================================

WRITE = '--write' in sys.argv
SCOPES = ['https://www.googleapis.com/auth/spreadsheets' if WRITE
          else 'https://www.googleapis.com/auth/spreadsheets.readonly']

COL_ID = 'Video ID'
COL_BRAND = 'brand'
COL_DATE = 'Date'
COL_POST = 'Post time'
COL_GMV = 'Affiliate video-attributed GMV'
SUM_COLS = [COL_GMV, 'Creator video-attributed orders', 'Creator-attributed items sold']
JENIS_HEADER = 'Jenis'

POST_RE = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$')


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def to_number(v):
    if isinstance(v, (int, float)):
        return v
    s = str(v or '').strip().replace('Rp', '').replace(',', '')
    try:
        return float(s) if '.' in s else int(s)
    except ValueError:
        return 0


def cell(row, i):
    return row[i] if i < len(row) else ''


def read_source(service):
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{SOURCE_SHEET}'!A1:AB",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    vals = r.get('values', [])
    if not vals:
        raise ValueError(f'Tab "{SOURCE_SHEET}" kosong.')
    return vals[0], vals[1:]


def aggregate(header, rows):
    """Return (groups {(brand, video_id): row}, accepted_rows, skipped_blank, skipped_shifted)."""
    def idx(name):
        if name not in header:
            raise ValueError(f'Kolom "{name}" gak ada di header {SOURCE_SHEET}.')
        return header.index(name)

    i_id, i_brand, i_date, i_post = idx(COL_ID), idx(COL_BRAND), idx(COL_DATE), idx(COL_POST)
    sum_idx = [idx(c) for c in SUM_COLS]

    groups, accepted = {}, []
    skipped_blank, skipped_shifted = 0, []
    for n, row in enumerate(rows, start=2):  # n = nomor baris di sheet
        vid = str(cell(row, i_id)).strip()
        brand = str(cell(row, i_brand)).strip().upper()
        brand = MERGE_BRANDS.get(brand, brand)
        if not vid or not brand:
            skipped_blank += 1
            continue
        if not POST_RE.match(str(cell(row, i_post)).strip()):
            skipped_shifted.append(n)
            continue
        accepted.append(row)
        key = (brand, vid)
        if key not in groups:
            base = [cell(row, i) for i in range(len(header))]
            base[i_brand] = brand
            for i in sum_idx:
                base[i] = 0
            groups[key] = base
        base = groups[key]
        for i in sum_idx:
            base[i] += to_number(cell(row, i))
        d = str(cell(row, i_date)).strip()
        if d and (not str(base[i_date]).strip() or d < str(base[i_date]).strip()):
            base[i_date] = d
    return groups, accepted, skipped_blank, skipped_shifted


def build_tabs(groups, i_gmv):
    tabs = {}
    for (brand, _), row in groups.items():
        jenis = 'GMV' if row[i_gmv] > 0 else 'NON GMV'
        tabs.setdefault(f'{brand}{TAB_SUFFIX}', []).append(row + [jenis])
    for title in tabs:
        tabs[title].sort(key=lambda r: r[i_gmv], reverse=True)
    return dict(sorted(tabs.items()))


def grid_cells(service):
    meta = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID, fields='sheets(properties(title,gridProperties))').execute()
    titles = {s['properties']['title'] for s in meta['sheets']}
    total = sum(s['properties']['gridProperties']['rowCount'] * s['properties']['gridProperties']['columnCount']
                for s in meta['sheets'])
    return titles, total


def write_tabs(service, header, tabs):
    out_header = header + [JENIS_HEADER]
    service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': [
        {'addSheet': {'properties': {
            'title': title,
            'gridProperties': {'rowCount': max(len(rows) + 1, 2), 'columnCount': len(out_header),
                               'frozenRowCount': 1}}}}
        for title, rows in tabs.items()]}).execute()

    for title, rows in tabs.items():
        values = [out_header] + rows
        for start in range(0, len(values), WRITE_CHUNK_ROWS):
            part = values[start:start + WRITE_CHUNK_ROWS]
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A{start + 1}",
                valueInputOption='RAW', body={'values': part}).execute()
        print(f'  ... "{title}" selesai ({len(rows)} baris)')


def main():
    service = get_sheets_service()
    print(f'Mode: {"WRITE" if WRITE else "DRY RUN (tidak ada yang ditulis)"}')

    header, rows = read_source(service)
    print(f'{len(rows)} baris di "{SOURCE_SHEET}" (semua diambil, tanpa filter bulan).')

    groups, accepted, skipped_blank, skipped_shifted = aggregate(header, rows)
    print(f'{len(groups)} baris unik (brand + Video ID).')
    print(f'{skipped_blank} baris dilewati (Video ID/brand kosong).')
    if skipped_shifted:
        print(f'{len(skipped_shifted)} baris dilewati karena kolom bergeser (Post time bukan tanggal-jam): '
              f'baris sheet {", ".join(map(str, skipped_shifted[:10]))}'
              f'{" ..." if len(skipped_shifted) > 10 else ""}')

    i_gmv = header.index(COL_GMV)
    tabs = build_tabs(groups, i_gmv)
    print(f'\n{len(tabs)} tab akan dibuat:')
    for title, trows in tabs.items():
        n_gmv = sum(1 for r in trows if r[-1] == 'GMV')
        print(f'  {title:<22} {len(trows):>6} baris  (GMV: {n_gmv:>4}, NON GMV: {len(trows) - n_gmv:>6})   '
              f'GMV total: {sum(r[i_gmv] for r in trows):>14,}')

    for name in SUM_COLS:
        i = header.index(name)
        src = sum(to_number(cell(r, i)) for r in accepted)
        out = sum(r[i] for r in groups.values())
        print(f'  cek total {name}: sumber {src:,} | hasil {out:,} | {"OK" if src == out else "BEDA!"}')

    titles, existing_cells = grid_cells(service)
    new_cells = sum(max(len(r) + 1, 2) * (len(header) + 1) for r in tabs.values())
    print(f'\nSel sekarang {existing_cells:,} + sel baru {new_cells:,} = {existing_cells + new_cells:,} '
          f'(batas {MAX_CELLS:,}).')
    clash = sorted(set(tabs) & titles)
    if clash:
        raise SystemExit(f'Tab sudah ada, tidak ditimpa. Hapus/rename dulu atau ganti TAB_SUFFIX: {", ".join(clash)}')
    if existing_cells + new_cells > MAX_CELLS:
        raise SystemExit('Total sel melebihi batas Google Sheets, tab tidak dibuat.')

    if not WRITE:
        print('\nDRY RUN selesai. Jalankan dengan --write untuk membuat tab.')
        return
    try:
        write_tabs(service, header, tabs)
    except Exception:
        print('\nGAGAL di tengah jalan. Sebagian tab mungkin sudah terbuat: hapus tab "'
              f'*{TAB_SUFFIX}" di spreadsheet sebelum menjalankan ulang.')
        raise
    print(f'\nSELESAI. {len(tabs)} tab dibuat.')


if __name__ == '__main__':
    main()
