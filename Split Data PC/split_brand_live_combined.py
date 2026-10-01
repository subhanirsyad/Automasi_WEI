"""
PISAH PER BRAND (GMV + NON GMV DIGABUNG)  [TAP LS PC 24-30 SEP 2026]
================================================================================
Versi lain dari split_brand_gmv_live.py (file itu TIDAK diubah, fungsi bacanya
dipakai ulang). Beda: 1 tab per brand, GMV dan NON GMV ada di tab yang sama dan
dibedakan lewat kolom "Jenis" di paling kanan.

ATURAN (sama dengan split_brand_gmv_live.py kecuali yang disebut beda):
  - Semua baris di "Custom report" diambil, TANPA filter bulan/tanggal.
  - Brand huruf besar; 1 tab per brand (DOVE DEO dan DOVE SCL terpisah).
    Brand di MERGE_BRANDS (default kosong) bisa digabung jadi satu tab.
  - Unik per (brand, Livestream room ID); GMV, orders, items sold dijumlah.
  - Kolom lain dari baris pertama, 'Date' = tanggal paling awal.
  - Kolom tambahan: Start Live, End Live, Date Live (dari 'LIVE time info'),
    lalu Jenis = GMV (total GMV > 0) / NON GMV (total GMV = 0).
  - Baris diurutkan dari GMV terbesar, jadi yang GMV ada di atas.

Tab yang sudah ada tidak ditimpa: script berhenti dan menyebut tab mana.

CARA PAKAI:
  python split_brand_live_combined.py            # DRY RUN: cuma baca & ringkas
  python split_brand_live_combined.py --write    # bikin tab & tulis (butuh Editor)
"""

import sys

from split_brand_gmv_live import (
    COL_BRAND, EXTRA_HEADERS, GMV_COL, SOURCE_SHEET, SPREADSHEET_ID, SUM_COLS, WRITE,
    aggregate, cell, get_sheets_service, read_source, to_number,
)

# ============== KONFIGURASI ==============
MERGE_BRANDS = {}  # isi {'DOVE DEO': 'DOVE', 'DOVE SCL': 'DOVE'} kalau mau digabung
# ===========================================

JENIS_HEADER = 'Jenis'


def merge_brand_names(header, rows):
    i = header.index(COL_BRAND)
    out = []
    for row in rows:
        row = list(row)
        if i < len(row):
            name = str(row[i]).strip().upper()
            row[i] = MERGE_BRANDS.get(name, name)
        out.append(row)
    return out


def build_tabs(groups, i_gmv):
    tabs = {}
    for (brand, _), row in groups.items():
        jenis = 'GMV' if row[i_gmv] > 0 else 'NON GMV'
        tabs.setdefault(brand, []).append(row + [jenis])
    for brand in tabs:
        tabs[brand].sort(key=lambda r: r[i_gmv], reverse=True)
    return dict(sorted(tabs.items()))


def write_tabs(service, header, tabs):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields='sheets(properties(title))').execute()
    existing = {s['properties']['title'] for s in meta['sheets']}
    clash = sorted(set(tabs) & existing)
    if clash:
        raise SystemExit(f'Tab sudah ada, tidak ditimpa. Hapus/rename dulu: {", ".join(clash)}')

    out_header = header + EXTRA_HEADERS + [JENIS_HEADER]
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
    print(f'{len(rows)} baris di "{SOURCE_SHEET}" (semua diambil, tanpa filter bulan).')
    rows = merge_brand_names(header, rows)

    groups, skipped, bad_time, i_gmv, _ = aggregate(header, rows)
    print(f'{len(groups)} baris unik (brand + room ID). {skipped} baris dilewati (room ID/brand kosong). '
          f'{bad_time} baris format LIVE time info tidak terbaca.')

    tabs = build_tabs(groups, i_gmv)
    print(f'\n{len(tabs)} tab akan dibuat:')
    for title, trows in tabs.items():
        n_gmv = sum(1 for r in trows if r[-1] == 'GMV')
        print(f'  {title:<12} {len(trows):>5} baris  (GMV: {n_gmv:>3}, NON GMV: {len(trows) - n_gmv:>4})   '
              f'GMV total: {sum(r[i_gmv] for r in trows):>14,}')

    i_room = header.index('Livestream room ID')
    i_brand = header.index(COL_BRAND)
    for name in SUM_COLS:
        i = header.index(name)
        src = sum(to_number(cell(r, i)) for r in rows
                  if str(cell(r, i_room)).strip() and str(cell(r, i_brand)).strip())
        out = sum(r[i] for r in groups.values())
        print(f'  cek total {name}: sumber {src:,} | hasil {out:,} | {"OK" if src == out else "BEDA!"}')

    if not WRITE:
        print('\nDRY RUN selesai. Jalankan dengan --write untuk membuat tab.')
        return
    write_tabs(service, header, tabs)
    print(f'\nSELESAI. {len(tabs)} tab dibuat.')


if __name__ == '__main__':
    main()
