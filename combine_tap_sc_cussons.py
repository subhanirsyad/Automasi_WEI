"""
GABUNGIN TAP Cussons + SC Cussons -> "Gabungan TAP&SC2"  [CUSSONS SEPTEMBER]
================================================================================
Baca semua baris di tab "TAP Cussons (1-30 September)" & "SC Cussons (1-30
September)" (spreadsheet master Cussons), gabungin jadi 1 baris per
Creator+Product+Date+Source di tab "Gabungan TAP&SC2", dengan GMV & Qty
dipecah ke 3 kolom channel (LS/SL/SV) -- pola PERSIS ditiru dari tab lama
"Gabungan TAP&SC" (Agustus), sudah diverifikasi angkanya cocok ke source asli.

ALOKASI DARI TAP (1 baris TAP = irisan beberapa channel dalam 1 baris) --
SEMUA langsung dari kolom yang udah ada di sheet TAP Cussons, GAK ada hitungan
proporsional/alokasi lagi di script ini:
  GMV LS = 'Creator LIVE-attributed GMV' (kolom L)
  GMV SV = 'Affiliate video-attributed GMV' (kolom K)
  GMV SL = 'GMV SL' (kolom M)
  QTY LS = 'Item Sold LS'
  QTY SV = 'Item Sold SV'
  QTY SL = 'Item Sold SL'
  (Kolom 'Item Sold LS/SV/SL' ini belum ada di sheet TAP Cussons September saat
  script ini ditulis -- kalau belum ditambahin, script bakal error jelas minta
  dicek headernya, bukan diam-diam salah baca.)

ALOKASI DARI SC (1 baris SC = 1 order = 1 channel pasti, gak perlu dipecah).
FILTER (baris SC yang GAK memenuhi ini DIBUANG, gak dimasukin ke Gabungan):
  - 'Cek Creator' != #N/A  (creator gak dikenal / belum ada di roster dibuang)
  - 'Order Status' != Ineligible (order yang gak eligible dibuang)
  Channel dari kolom 'Content Type' (M) di sheet SC Cussons -- tapi kita baca
  hasil formula Channel (kolom N: "LIVE STREAMING"/"SHORT VIDEO"/"SHARELINK")
  yang udah kehitung di situ, GMV ('Est. Commission Base', kolom P -- BUKAN
  Payment Amount) & Qty (kolom G) masuk PENUH ke 1 kolom channel yang sesuai,
  2 kolom channel lain dikosongin.

KOLOM YANG SENGAJA TIDAK DITULIS SCRIPT INI (rumus, harus udah di-drag ke
bawah duluan di tab tujuan -- persis kayak di "Gabungan TAP&SC" lama):
  - Cek (B), CAT (E), Product Fokus (F)  -> kosong / rumus manual
  - Total GMV (O) = SUM(I;K;M), Kuantitas Sold (P) = SUM(J;L;N) -> rumus manual
Kolom WEEK (H) DIAMBIL dari nilai yang udah kehitung di tab sumber (bukan
dihitung ulang di sini), biar konsisten sama rumus Week yang baru (SEP W1
(1-6) dst) yang sudah di-set di tab TAP/SC Cussons September itu sendiri.

DEDUP (count-based, BUKAN cuma "udah ada / belum"): kalau ada 2+ baris dengan
kombinasi Creator+Product+Date+Source yang PERSIS SAMA (misal 2 order beda di
hari yang sama), SEMUANYA tetap ditulis -- bukan cuma yang pertama. Yang
di-skip cuma kalau kombinasi itu sudah muncul di Gabungan SEBANYAK ATAU LEBIH
BANYAK dibanding di data sumber sekarang (jadi tetap aman/idempotent kalau
script ini di-run ulang tanpa ada data baru).

Kalau tab "Gabungan TAP&SC2" belum ada, script ini BIKIN tab baru + isi
header row-nya doang -- kamu masih perlu drag manual rumus Cek/CAT/Product
Fokus/Total GMV/Kuantitas Sold dari tab "Gabungan TAP&SC" (Agustus) ke tab
baru ini sebelum baris data kebaca bener oleh rumus-rumus itu.

PILIH WEEK: setelah baca semua baris TAP+SC, script nampilin daftar minggu
yang ketemu di data (dari kolom 'week' yang udah ada, format "SEP W<n> (..)")
dan minta kamu pilih mau proses WEEK berapa -- cuma baris minggu itu yang
di-dedup & ditulis, sisanya diabaikan buat run ini. Kosongin inputnya kalau
mau proses SEMUA minggu sekaligus (perilaku lama).

CARA PAKAI:
  python combine_tap_sc_cussons.py
"""

import datetime
import re
from collections import Counter

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============== KONFIGURASI ==============
CREDENTIALS_FILE = 'credentials.json'
SPREADSHEET_ID = ''  # [INT] PZ Cussons
TAP_SHEET = 'TAP Cussons (1-30 September)'
SC_SHEET = 'SC Cussons (1-30 September)'
GABUNGAN_SHEET = 'Gabungan TAP&SC2'
# ===========================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

GABUNGAN_HEADERS = [
    'Creator Username', 'Cek', 'Product ID', 'Produk Name', 'CAT', 'Product Fokus',
    'Date', 'WEEK', 'GMV LS', 'QYT LS', 'GMV SL', 'QYT SL', 'GMV SV', 'QYT SV',
    'Total GMV', 'Kuantitas Sold', 'SOURCE',
]
GABUNGAN_DATA_START_ROW = 3  # row 1 = filter/subtotal area, row 2 = header (samain kayak tab lama)

# Kolom yang DITULIS langsung (bukan rumus): map nama field -> huruf kolom di Gabungan
WRITE_COLS = {
    'creator': 'A', 'product_id': 'C', 'product_name': 'D', 'date': 'G', 'week': 'H',
    'gmv_ls': 'I', 'qty_ls': 'J', 'gmv_sl': 'K', 'qty_sl': 'L', 'gmv_sv': 'M', 'qty_sv': 'N',
    'source': 'Q',
}


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def serial_to_date_str(serial):
    if serial in (None, ''):
        return ''
    base = datetime.date(1899, 12, 30)
    return (base + datetime.timedelta(days=float(serial))).strftime('%Y-%m-%d')


def to_number(v):
    if v in (None, ''):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('Rp', '').replace('.', '').replace(',', '.').strip()
    if s in ('', '-', '--'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def split_header(vals, marker):
    """Cari baris header (baris pertama yang ada kolom `marker`) -- header gak selalu
    di baris 1 (SC Cussons sekarang header-nya di baris 2, baris 1 kosong)."""
    for i, row in enumerate(vals[:10]):
        if marker in row:
            return row, vals[i + 1:]
    return vals[0], vals[1:]


def read_tap_rows(service):
    """Baca semua baris TAP Cussons September, hasil: list of dict per (creator, product, date)."""
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{TAP_SHEET}'!A1:AZ",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    vals = r.get('values', [])
    if not vals:
        return []
    header, data_rows = split_header(vals, 'Product ID')

    def idx(name):
        return header.index(name) if name in header else -1

    i_date = idx('Date')
    i_week = idx('WEEK')
    i_creator = idx('Creator name')
    i_pid = idx('Product ID')
    i_pname = idx('Product name')
    i_k = idx('Affiliate video-attributed GMV')
    i_l = idx('Creator LIVE-attributed GMV')
    i_gmv_sl = idx('GMV SL')        # kolom M, sudah ada di sheet TAP Cussons
    i_qty_ls = idx('Item Sold LS')  # langsung, BUKAN dihitung dari rasio orders lagi
    i_qty_sv = idx('Item Sold SV')
    i_qty_sl = idx('Item Sold SL')

    required = [i_date, i_week, i_creator, i_pid, i_pname, i_k, i_l, i_gmv_sl,
                i_qty_ls, i_qty_sv, i_qty_sl]
    if -1 in required:
        raise ValueError(f'Header TAP Cussons gak lengkap, cek lagi (idx={required}).')

    def cell(row, i):
        return row[i] if i < len(row) else None

    records = []
    for row in data_rows:
        if not row or not cell(row, i_creator):
            continue
        gmv_ls = to_number(cell(row, i_l))
        gmv_sv = to_number(cell(row, i_k))
        gmv_sl = to_number(cell(row, i_gmv_sl))
        qty_ls = to_number(cell(row, i_qty_ls))
        qty_sv = to_number(cell(row, i_qty_sv))
        qty_sl = to_number(cell(row, i_qty_sl))

        records.append({
            'creator': str(cell(row, i_creator)).strip(),
            'product_id': str(cell(row, i_pid) or '').strip(),
            'product_name': cell(row, i_pname) or '',
            'date': str(cell(row, i_date) or '').strip(),
            'week': cell(row, i_week) or '',
            'gmv_ls': gmv_ls, 'qty_ls': qty_ls,
            'gmv_sl': gmv_sl, 'qty_sl': qty_sl,
            'gmv_sv': gmv_sv, 'qty_sv': qty_sv,
            'source': 'TAP',
        })
    return records


def read_sc_rows(service):
    """Baca semua baris SC Cussons September, hasil: list of dict per order (1 order = 1 channel)."""
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{SC_SHEET}'!A1:T",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    vals = r.get('values', [])
    if not vals:
        return []
    header, data_rows = split_header(vals, 'Order ID')

    def idx(name):
        return header.index(name) if name in header else -1

    i_pid = idx('Product ID')
    i_pname = idx('Product Name')
    i_gmv = idx('Est. Commission Base')  # kolom P -- BUKAN Payment Amount
    i_qty = idx('Quantity')
    i_creator = idx('Creator Username')
    i_channel = idx('Channels')  # header rumus channel di SC Cussons (kolom N)
    if i_channel == -1:
        i_channel = idx('Channel')
    i_date = idx('Date')
    i_week = idx('WEEK')
    i_cek_creator = idx('Cek Creator')
    i_order_status = idx('Order Status')

    required = [i_pid, i_pname, i_gmv, i_qty, i_creator, i_channel, i_date, i_week,
                i_cek_creator, i_order_status]
    if -1 in required:
        raise ValueError(f'Header SC Cussons gak lengkap, cek lagi (idx={required}).')

    def cell(row, i):
        return row[i] if i < len(row) else None

    records = []
    skipped_na = 0
    skipped_ineligible = 0
    for row in data_rows:
        if not row or not cell(row, i_creator):
            continue

        cek_creator = str(cell(row, i_cek_creator) or '').strip().upper()
        if cek_creator.startswith('#N/A'):  # error VLOOKUP-nya berupa pesan panjang, bukan "#N/A" polos
            skipped_na += 1
            continue

        order_status = str(cell(row, i_order_status) or '').strip().upper()
        if order_status == 'INELIGIBLE':
            skipped_ineligible += 1
            continue

        gmv = to_number(cell(row, i_gmv))
        qty = to_number(cell(row, i_qty))
        channel = str(cell(row, i_channel) or '').strip().upper()

        gmv_ls = gmv_sl = gmv_sv = 0.0
        qty_ls = qty_sl = qty_sv = 0.0
        if channel == 'LIVE STREAMING':
            gmv_ls, qty_ls = gmv, qty
        elif channel == 'SHORT VIDEO':
            gmv_sv, qty_sv = gmv, qty
        else:  # SHARELINK / lainnya
            gmv_sl, qty_sl = gmv, qty

        date_val = cell(row, i_date)
        date_str = serial_to_date_str(date_val) if isinstance(date_val, (int, float)) else str(date_val or '')

        records.append({
            'creator': str(cell(row, i_creator)).strip(),
            'product_id': str(cell(row, i_pid) or '').strip(),
            'product_name': cell(row, i_pname) or '',
            'date': date_str,
            'week': cell(row, i_week) or '',
            'gmv_ls': gmv_ls, 'qty_ls': qty_ls,
            'gmv_sl': gmv_sl, 'qty_sl': qty_sl,
            'gmv_sv': gmv_sv, 'qty_sv': qty_sv,
            'source': 'SC',
        })

    print(f'  SC: {skipped_na} baris dilewati (Cek Creator = #N/A), '
          f'{skipped_ineligible} baris dilewati (Order Status = Ineligible).')
    return records


def ensure_gabungan_sheet(service):
    """Bikin tab Gabungan TAP&SC2 + header kalau belum ada. Return True kalau baru dibikin."""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields='sheets(properties(title))').execute()
    existing = [s['properties']['title'] for s in meta['sheets']]
    if GABUNGAN_SHEET in existing:
        return False

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': [{'addSheet': {'properties': {'title': GABUNGAN_SHEET}}}]}
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{GABUNGAN_SHEET}'!A2",
        valueInputOption='USER_ENTERED', body={'values': [GABUNGAN_HEADERS]}
    ).execute()
    print(f'Tab "{GABUNGAN_SHEET}" baru dibikin + header ditulis di row 2.')
    print('PENTING: drag manual rumus Cek/CAT/Product Fokus/Total GMV/Kuantitas Sold')
    print(f'dari tab "Gabungan TAP&SC" (Agustus) ke tab "{GABUNGAN_SHEET}" ini dulu')
    print('sebelum baris data yang ditulis script ini kebaca bener.')
    return True


def get_existing_key_counts_and_next_row(service):
    """Hitung BERAPA KALI tiap kombinasi (Creator+Product+Date+Source) udah muncul
    di Gabungan TAP&SC2. Dipakai buat dedup count-based: kalau ada 2 baris SC/TAP
    yang kuncinya sama persis (misal 2 order beda di hari yang sama), DUA-DUANYA
    tetap ditulis -- bukan cuma yang pertama. Baris baru cuma di-skip kalau
    kombinasi itu SUDAH muncul sebanyak (atau lebih) di Gabungan dibanding di data
    sumber yang lagi diproses sekarang (jadi tetap aman/idempotent kalau di-run
    ulang tanpa data baru)."""
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{GABUNGAN_SHEET}'!A{GABUNGAN_DATA_START_ROW}:Q",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    vals = r.get('values', [])
    counts = Counter()
    next_row = GABUNGAN_DATA_START_ROW
    for i, row in enumerate(vals):
        creator = row[0] if len(row) > 0 else ''
        product_id = row[2] if len(row) > 2 else ''
        date = row[6] if len(row) > 6 else ''
        source = row[16] if len(row) > 16 else ''
        if creator or product_id or date:
            next_row = GABUNGAN_DATA_START_ROW + i + 1
        if creator and date:
            key = f'{str(creator).strip()}|{str(product_id).strip()}|{str(date).strip()}|{str(source).strip()}'
            counts[key] += 1
    return counts, next_row


def get_sheet_id_and_row_count(service, sheet_name):
    meta = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID, fields='sheets(properties(sheetId,title,gridProperties))'
    ).execute()
    for sh in meta['sheets']:
        if sh['properties']['title'] == sheet_name:
            return sh['properties']['sheetId'], sh['properties']['gridProperties']['rowCount']
    raise ValueError(f'Sheet "{sheet_name}" tidak ketemu.')


def ensure_rows(service, sheet_name, needed_last_row):
    sheet_id, current_rows = get_sheet_id_and_row_count(service, sheet_name)
    if needed_last_row <= current_rows:
        return
    rows_short = needed_last_row - current_rows
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': [{
            'appendDimension': {'sheetId': sheet_id, 'dimension': 'ROWS', 'length': rows_short}
        }]}
    ).execute()
    print(f'Grid "{sheet_name}" ditambah {rows_short} baris (sekarang sampai baris {needed_last_row}).')


def extract_week_number(week_str):
    """'SEP W2 (7-13)' -> 2. None kalau formatnya gak ketemu pola 'W<angka>'."""
    m = re.search(r'W(\d+)', str(week_str or ''))
    return int(m.group(1)) if m else None


def prompt_week_choice(all_records):
    """Tampilin minggu-minggu yang ketemu di data, minta user pilih 1 (atau
    kosongin buat proses SEMUA minggu). Return angka minggu (int) atau None."""
    weeks_present = sorted(set(
        (extract_week_number(r['week']), r['week']) for r in all_records
        if extract_week_number(r['week']) is not None
    ))
    if not weeks_present:
        return None

    print('\nMinggu yang ketemu di data TAP/SC Cussons:')
    for num, label in weeks_present:
        print(f'  {num}. {label}')
    choice = input('\nMau proses WEEK berapa? (isi angka, kosongkan = SEMUA minggu): ').strip()
    if not choice:
        return None
    try:
        return int(choice)
    except ValueError:
        print(f'Input "{choice}" gak valid, dianggap SEMUA minggu.')
        return None


def write_records(service, records, next_row):
    data = []
    for idx, rec in enumerate(records):
        r = next_row + idx
        for field, col in WRITE_COLS.items():
            data.append({'range': f"'{GABUNGAN_SHEET}'!{col}{r}", 'values': [[rec[field]]]})

    chunk_rows = 500
    entries_per_row = len(WRITE_COLS)
    for start in range(0, len(records), chunk_rows):
        n = min(chunk_rows, len(records) - start)
        chunk = data[start * entries_per_row: (start + n) * entries_per_row]
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={'valueInputOption': 'USER_ENTERED', 'data': chunk}
        ).execute()
        print(f'  ... tertulis {start + n}/{len(records)} baris')


def main():
    service = get_sheets_service()

    just_created = ensure_gabungan_sheet(service)

    print('Membaca TAP Cussons September...')
    tap_records = read_tap_rows(service)
    print(f'  {len(tap_records)} baris TAP.')

    print('Membaca SC Cussons September...')
    sc_records = read_sc_rows(service)
    print(f'  {len(sc_records)} baris SC.')

    all_records = tap_records + sc_records
    if not all_records:
        print('Gak ada data di TAP/SC Cussons September, berhenti.')
        return

    target_week = prompt_week_choice(all_records)
    if target_week is not None:
        before = len(all_records)
        all_records = [r for r in all_records if extract_week_number(r['week']) == target_week]
        print(f'Filter WEEK {target_week}: {before - len(all_records)} baris dibuang, sisa {len(all_records)} baris.')
        if not all_records:
            print('Gak ada baris yang cocok WEEK itu, berhenti.')
            return

    print('Cek data yang udah ada di Gabungan TAP&SC2 (dedup count-based)...')
    existing_counts, next_row = get_existing_key_counts_and_next_row(service)
    print(f'  {sum(existing_counts.values())} baris udah ada, baris baru mulai dari {next_row}.')

    to_insert = []
    seen_now = Counter()
    skipped = 0
    for rec in all_records:
        key = f"{rec['creator']}|{rec['product_id']}|{rec['date']}|{rec['source']}"
        seen_now[key] += 1
        # baris ke-N buat kombinasi ini cuma di-skip kalau di Gabungan udah ada
        # >= N baris dengan kombinasi yang sama (dedup antar-run, BUKAN dedup
        # antar-baris yang kuncinya sama dalam 1 batch -- itu semua tetap masuk).
        if seen_now[key] <= existing_counts[key]:
            skipped += 1
            continue
        to_insert.append(rec)

    if not to_insert:
        print(f'\nSemua {len(all_records)} baris (TAP+SC) sudah ada sebelumnya di Gabungan TAP&SC2.')
        return

    needed_last_row = next_row + len(to_insert) - 1
    print(f'\n{len(to_insert)} baris baru akan ditulis ke Gabungan TAP&SC2 (baris {next_row}-{needed_last_row}).')
    print(f'{skipped} dilewati (kombinasi Creator+Product+Date+Source sudah ada).')
    ensure_rows(service, GABUNGAN_SHEET, needed_last_row)
    write_records(service, to_insert, next_row)

    print(f'\nSELESAI. {len(to_insert)} baris baru ditambahkan, {skipped} dilewati.')
    if just_created:
        print('\nINGAT: tab ini baru dibikin -- pastikan rumus Cek/CAT/Product Fokus/')
        print('Total GMV/Kuantitas Sold sudah di-drag manual sebelum dipakai produksi.')


if __name__ == '__main__':
    main()
