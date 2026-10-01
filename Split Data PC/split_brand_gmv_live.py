"""
PISAH PER BRAND x GMV / NON GMV  [TAP LS PC 24-30 SEP 2026]
================================================================================
Baca tab "Custom report" (spreadsheet TAP LS PC 24-30 SEP 2026), lalu bikin tab
baru per brand dan per jenis, misal "PEPSODENT - GMV" & "PEPSODENT - NON GMV".

ATURAN:
  - Brand dinormalisasi ke HURUF BESAR (Pepsodent == PEPSODENT).
  - Unik per (brand, Livestream room ID). Satu room ID bisa muncul di beberapa
    brand -> tiap brand dapat barisnya sendiri.
  - Kolom yang DIJUMLAH untuk baris dengan room ID sama dalam 1 brand:
      Creator LIVE-attributed GMV, Creator LIVE-attributed orders,
      Creator-attributed items sold.
  - Kolom lain diambil dari baris pertama room ID itu, KECUALI 'Date' yang
    diambil tanggal paling awal.
  - GMV = total Creator LIVE-attributed GMV > 0. NON GMV = total = 0.
  - Tambahan 3 kolom di kanan, dipecah dari 'LIVE time info':
      Start Live (tanggal+jam awal), End Live (tanggal+jam akhir),
      Date Live (tanggal awal saja, tanpa jam).

TAB YANG SUDAH ADA TIDAK DITIMPA: kalau salah satu nama tab tujuan sudah ada,
script berhenti dan menyebutkan tab mana.

CARA PAKAI:
  python split_brand_gmv_live.py            # DRY RUN: cuma baca & tampilkan ringkasan
  python split_brand_gmv_live.py --write    # bikin tab & tulis (butuh akses Editor)
"""

import re
import sys
from collections import Counter

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============== KONFIGURASI ==============
CREDENTIALS_FILE = 'credentials.json'
SPREADSHEET_ID = '1EkvCjequi3sDnEBBpVQLtDcY18bYBbECa3DaztE2IW0'  # TAP LS PC 24-30 SEP 2026
SOURCE_SHEET = 'Custom report'
# ===========================================

WRITE = '--write' in sys.argv
SCOPES = ['https://www.googleapis.com/auth/spreadsheets' if WRITE
          else 'https://www.googleapis.com/auth/spreadsheets.readonly']

COL_ROOM = 'Livestream room ID'
COL_BRAND = 'brand'
COL_DATE = 'Date'
COL_TIME = 'LIVE time info'
SUM_COLS = ['Creator LIVE-attributed GMV', 'Creator LIVE-attributed orders',
            'Creator-attributed items sold']
GMV_COL = 'Creator LIVE-attributed GMV'
EXTRA_HEADERS = ['Start Live', 'End Live', 'Date Live']

TIME_RE = re.compile(r'^\s*(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\s*-\s*(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\s*$')


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
    """Return {(brand, room_id): row_gabungan}, jumlah baris dilewati."""
    def idx(name):
        if name not in header:
            raise ValueError(f'Kolom "{name}" gak ada di header {SOURCE_SHEET}.')
        return header.index(name)

    i_room, i_brand, i_date, i_time = idx(COL_ROOM), idx(COL_BRAND), idx(COL_DATE), idx(COL_TIME)
    sum_idx = [idx(c) for c in SUM_COLS]

    groups = {}
    skipped = 0
    for row in rows:
        room = str(cell(row, i_room)).strip()
        brand = str(cell(row, i_brand)).strip().upper()
        if not room or not brand:
            skipped += 1
            continue
        key = (brand, room)
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

    bad_time = 0
    for base in groups.values():
        m = TIME_RE.match(str(base[i_time]))
        if m:
            base.extend([m.group(1), m.group(2), m.group(1)[:10]])
        else:
            bad_time += 1
            base.extend(['', '', ''])
    return groups, skipped, bad_time, header.index(GMV_COL), i_brand


def split_tabs(groups, i_gmv):
    tabs = {}
    for (brand, _), row in groups.items():
        kind = 'GMV' if row[i_gmv] > 0 else 'NON GMV'
        tabs.setdefault(f'{brand} - {kind}', []).append(row)
    return dict(sorted(tabs.items()))


def write_tabs(service, header, tabs):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields='sheets(properties(title))').execute()
    existing = {s['properties']['title'] for s in meta['sheets']}
    clash = sorted(set(tabs) & existing)
    if clash:
        raise SystemExit(f'Tab sudah ada, tidak ditimpa. Hapus/rename dulu: {", ".join(clash)}')

    out_header = header + EXTRA_HEADERS
    service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': [
        {'addSheet': {'properties': {
            'title': title,
            'gridProperties': {'rowCount': max(len(rows) + 1, 2), 'columnCount': len(out_header),
                               'frozenRowCount': 1}}}}
        for title, rows in tabs.items()]}).execute()

    data = [{'range': f"'{title}'!A1", 'values': [out_header] + rows} for title, rows in tabs.items()]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body={'valueInputOption': 'RAW', 'data': data}).execute()


def main():
    service = get_sheets_service()
    print(f'Mode: {"WRITE" if WRITE else "DRY RUN (tidak ada yang ditulis)"}')

    header, rows = read_source(service)
    print(f'{len(rows)} baris di "{SOURCE_SHEET}".')

    groups, skipped, bad_time, i_gmv, _ = aggregate(header, rows)
    print(f'{len(groups)} baris unik (brand + room ID). {skipped} baris dilewati (room ID/brand kosong). '
          f'{bad_time} baris format LIVE time info tidak terbaca.')

    tabs = split_tabs(groups, i_gmv)
    print(f'\n{len(tabs)} tab akan dibuat:')
    for title, trows in tabs.items():
        print(f'  {title:<28} {len(trows):>5} baris   GMV total: {sum(r[i_gmv] for r in trows):>14,}')

    # cek: total tiap kolom jumlah harus sama dengan sumber
    for name in SUM_COLS:
        i = header.index(name)
        src = sum(to_number(cell(r, i)) for r in rows if str(cell(r, header.index(COL_ROOM))).strip()
                  and str(cell(r, header.index(COL_BRAND))).strip())
        out = sum(r[i] for r in groups.values())
        print(f'  cek total {name}: sumber {src:,} | hasil {out:,} | {"OK" if src == out else "BEDA!"}')

    if not WRITE:
        print('\nDRY RUN selesai. Jalankan dengan --write untuk membuat tab.')
        return
    write_tabs(service, header, tabs)
    print(f'\nSELESAI. {len(tabs)} tab dibuat.')


if __name__ == '__main__':
    main()
