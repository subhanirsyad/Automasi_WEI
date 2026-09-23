"""
GABUNGIN RAW TAP + Raw Data SC -> "RAW TAP SC"  [ELLIPS]  (PER WEEK)
================================================================================
Baca tab "RAW TAP  (20 April-31 August)" (PERHATIAN: 2 spasi di nama tab) &
"Raw Data SC (20 April-31 August)" (spreadsheet master Ellips), tulis ke tab
"RAW TAP SC" yang SUDAH ADA. Beda sama pola "Gabungan" Cussons (kolom LS/SL/SV
terpisah), tab ini formatnya PANJANG (long/melted): 1 baris = 1 channel, kolom:
  Username | Payment Amount | Channel | Produk | Product ID | tipe | week

PILIH WEEK (sama kayak combine Cussons): script baca kolom Week kedua sheet
sumber, nampilin daftar minggu terbaru yang ketemu (format "SEP W3 (10-16)"),
lalu minta kamu pilih mau proses WEEK yang mana. SEMUA baris minggu itu
langsung ditulis ke bawah RAW TAP SC -- TIDAK ada cek "udah pernah masuk
atau belum", jadi jangan run week yang sama 2x. Kosongin input = minggu
paling baru.
Week Ellips: W1 = 1-2, W2 = 3-9, W3 = 10-16, W4 = 17-23, W5 = 24-akhir bulan.
Label week TAP & SC dicocokkan PERSIS (teksnya), jadi pastiin rumus Week di
kedua sheet sumber udah pakai pembagian yang sama.

DARI TAP -- 1 baris TAP = kemungkinan lebih dari 1 channel aktif dalam 1 baris:
  - Kalau CUMA 1 dari {'Affiliate video GMV','Affiliate LIVE GMV'} yang != 0
    (dan 'GMV SHARE LINK' = 0): Channel diisi ("SHORT VIDEO" / "LIVE STREAM"),
    Payment Amount = nilai channel itu.
  - Selain itu (GMV SHARE LINK doang yang isi, ATAU >=2 channel aktif
    sekaligus, ATAU semua nol): Channel DIKOSONGIN, Payment Amount = total
    'Affiliate GMV'. (Diverifikasi ke histori "RAW TAP SC" yang udah ada --
    baris Share-Link-only maupun baris multi-channel emang selalu Channel-nya
    kosong di situ.)
  Kolom 'week' diambil dari nilai yang udah kehitung di sheet sumber (kolom
  'Week'), bukan dihitung ulang.

DARI SC -- 1 baris SC = 1 order = 1 channel pasti, diambil APA ADANYA:
  Channel = kolom 'CHANNEL' (udah "SHORT VIDEO"/"LIVE STREAMING"/"SHARE LINK"),
  Payment Amount = kolom 'Payment Amount' (BUKAN Est. Commission Base -- beda
  sama konvensi Cussons, sesuai kolom target yang namanya emang "Payment
  Amount").
FILTER (baris SC yang GAK memenuhi ini DIBUANG):
  - 'cek name' != #N/A (creator gak dikenal / belum ada di roster dibuang)
  - 'Order Status' != Ineligible

CARA PAKAI:
  python combine_tap_sc_ellips.py
"""

import re

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============== KONFIGURASI ==============
CREDENTIALS_FILE = 'credentials.json'
SPREADSHEET_ID = ''  # Ellips X WEI - Community Performance 2026
TAP_SHEET = 'RAW TAP  (20 April-31 August)'  # PERHATIAN: 2 spasi
SC_SHEET = 'Raw Data SC (20 April-31 August)'
GABUNGAN_SHEET = 'RAW TAP SC'
WEEKS_SHOWN = 8  # berapa minggu terbaru yang ditampilin di pilihan
# ===========================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

GABUNGAN_DATA_START_ROW = 2  # row 1 = header

WRITE_COLS = {
    'username': 'A', 'payment_amount': 'B', 'channel': 'C',
    'produk': 'D', 'product_id': 'E', 'tipe': 'F', 'week': 'G',
}

MONTH_ORDER = {m: i for i, m in enumerate(
    ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'], start=1)}
MONTH_ORDER.update({'MEI': 5, 'AGU': 8, 'AGT': 8, 'OKT': 10, 'DES': 12})  # locale Indonesia


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def col_letter(idx0):
    """idx0: index kolom 0-based -> huruf kolom A1 notation ('A', 'B', ..., 'AA', ...)."""
    idx0 += 1
    s = ''
    while idx0 > 0:
        idx0, r = divmod(idx0 - 1, 26)
        s = chr(65 + r) + s
    return s


def get_header(service, sheet_name, last_col_letter):
    header_r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A1:{last_col_letter}1",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    header = header_r.get('values', [[]])[0] if header_r.get('values') else []
    if not header:
        raise ValueError(f'Header di "{sheet_name}" kosong/tidak ketemu.')
    return header


def normalize_week(v):
    return ' '.join(str(v or '').split()).upper()


def week_sort_key(label):
    """'SEP W3 (10-16)' -> (9, 3). Label yang gak kebaca ditaruh paling awal."""
    m = re.match(r'([A-Z]{3})\w*\s+W(\d+)', label)
    if not m:
        return (0, 0)
    return (MONTH_ORDER.get(m.group(1), 0), int(m.group(2)))


def read_week_column(service, sheet_name, week_col_letter):
    """Return {label week: [nomor baris sheet, ...]} dari seluruh kolom Week."""
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!{week_col_letter}2:{week_col_letter}",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    rows_by_week = {}
    for i, row in enumerate(r.get('values', [])):
        label = normalize_week(row[0]) if row else ''
        if label:
            rows_by_week.setdefault(label, []).append(i + 2)
    return rows_by_week


def prompt_week_choice(tap_weeks, sc_weeks):
    labels = sorted(set(tap_weeks) | set(sc_weeks), key=week_sort_key)[-WEEKS_SHOWN:]
    if not labels:
        return None

    print('\nMinggu terbaru yang ketemu di RAW TAP / Raw Data SC:')
    for n, label in enumerate(labels, start=1):
        print(f'  {n}. {label:<20} TAP: {len(tap_weeks.get(label, [])):>6} baris   '
              f'SC: {len(sc_weeks.get(label, [])):>6} baris')
    choice = input(f'\nMau proses WEEK yang mana? (isi nomor 1-{len(labels)}, kosongkan = {labels[-1]}): ').strip()
    if not choice:
        return labels[-1]
    try:
        n = int(choice)
        if 1 <= n <= len(labels):
            return labels[n - 1]
    except ValueError:
        pass
    print(f'Input "{choice}" gak valid, berhenti.')
    return None


def read_block(service, sheet_name, row_numbers, last_col_letter):
    """Baca blok baris min..max dari row_numbers, return list of (nomor baris, row)."""
    if not row_numbers:
        return []
    first, last = min(row_numbers), max(row_numbers)
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{sheet_name}'!A{first}:{last_col_letter}{last}",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    wanted = set(row_numbers)
    return [(first + i, row) for i, row in enumerate(r.get('values', [])) if first + i in wanted]


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


def read_tap_rows(service, header, row_numbers):
    def idx(name):
        return header.index(name) if name in header else -1

    i_week = idx('Week')
    i_creator = idx('Creator name')
    i_pid = idx('Product ID')
    i_pname = idx('Product name')
    i_gmv_total = idx('Affiliate GMV')
    i_video = idx('Affiliate video GMV')
    i_live = idx('Affiliate LIVE GMV')
    i_sl = idx('GMV SHARE LINK')

    required = [i_week, i_creator, i_pid, i_pname, i_gmv_total, i_video, i_live, i_sl]
    if -1 in required:
        raise ValueError(f'Header RAW TAP gak lengkap, cek lagi (idx={required}).')

    def cell(row, i):
        return row[i] if i < len(row) else None

    records = []
    for _, row in read_block(service, TAP_SHEET, row_numbers, 'BB'):
        if not row or not cell(row, i_creator):
            continue
        video = to_number(cell(row, i_video))
        live = to_number(cell(row, i_live))
        sl = to_number(cell(row, i_sl))
        total = to_number(cell(row, i_gmv_total))
        nonzero = sum(1 for v in (video, live, sl) if v)

        if nonzero == 1 and video:
            channel, amount = 'SHORT VIDEO', video
        elif nonzero == 1 and live:
            channel, amount = 'LIVE STREAM', live
        else:
            channel, amount = '', total

        records.append({
            'username': str(cell(row, i_creator)).strip(),
            'payment_amount': amount,
            'channel': channel,
            'produk': cell(row, i_pname) or '',
            'product_id': str(cell(row, i_pid) or '').strip(),
            'tipe': 'TAP',
            'week': cell(row, i_week) or '',
        })
    return records


def read_sc_rows(service, header, row_numbers):
    def idx(name):
        return header.index(name) if name in header else -1

    i_pid = idx('Product ID')
    i_pname = idx('Product Name')
    i_gmv = idx('Payment Amount')
    i_creator = idx('Creator Username')
    i_channel = idx('CHANNEL')
    i_week = idx('WEEK')
    i_cek = idx('cek name')
    i_status = idx('Order Status')

    required = [i_pid, i_pname, i_gmv, i_creator, i_channel, i_week, i_cek, i_status]
    if -1 in required:
        raise ValueError(f'Header Raw Data SC gak lengkap, cek lagi (idx={required}).')

    def cell(row, i):
        return row[i] if i < len(row) else None

    records = []
    skipped_na = 0
    skipped_ineligible = 0
    for _, row in read_block(service, SC_SHEET, row_numbers, 'AD'):
        if not row or not cell(row, i_creator):
            continue

        cek = str(cell(row, i_cek) or '').strip().upper()
        if cek.startswith('#N/A'):
            skipped_na += 1
            continue

        status = str(cell(row, i_status) or '').strip().upper()
        if status == 'INELIGIBLE':
            skipped_ineligible += 1
            continue

        records.append({
            'username': str(cell(row, i_creator)).strip(),
            'payment_amount': to_number(cell(row, i_gmv)),
            'channel': str(cell(row, i_channel) or '').strip(),
            'produk': cell(row, i_pname) or '',
            'product_id': str(cell(row, i_pid) or '').strip(),
            'tipe': 'SC',
            'week': cell(row, i_week) or '',
        })

    print(f'  SC: {skipped_na} baris dilewati (cek name = #N/A), '
          f'{skipped_ineligible} baris dilewati (Order Status = Ineligible).')
    return records


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


def get_next_row(service):
    """Baris kosong pertama setelah baris terakhir yang terisi di RAW TAP SC."""
    r = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{GABUNGAN_SHEET}'!A{GABUNGAN_DATA_START_ROW}:G",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    next_row = GABUNGAN_DATA_START_ROW
    for i, row in enumerate(r.get('values', [])):
        if any(str(v).strip() for v in row):
            next_row = GABUNGAN_DATA_START_ROW + i + 1
    return next_row


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

    tap_header = get_header(service, TAP_SHEET, 'BB')
    sc_header = get_header(service, SC_SHEET, 'AD')
    tap_week_col = col_letter(tap_header.index('Week'))
    sc_week_col = col_letter(sc_header.index('WEEK'))

    print('Membaca kolom Week RAW TAP & Raw Data SC...')
    tap_weeks = read_week_column(service, TAP_SHEET, tap_week_col)
    sc_weeks = read_week_column(service, SC_SHEET, sc_week_col)

    target_week = prompt_week_choice(tap_weeks, sc_weeks)
    if target_week is None:
        print('Gak ada week yang dipilih, berhenti.')
        return
    print(f'\nProses WEEK: {target_week}')

    tap_rows = tap_weeks.get(target_week, [])
    print(f'Membaca RAW TAP ({len(tap_rows)} baris week ini)...')
    tap_records = read_tap_rows(service, tap_header, tap_rows)
    print(f'  {len(tap_records)} baris TAP.')

    sc_rows = sc_weeks.get(target_week, [])
    print(f'Membaca Raw Data SC ({len(sc_rows)} baris week ini)...')
    sc_records = read_sc_rows(service, sc_header, sc_rows)
    print(f'  {len(sc_records)} baris SC.')

    to_insert = tap_records + sc_records
    if not to_insert:
        print('Gak ada baris TAP/SC buat week ini, berhenti.')
        return

    next_row = get_next_row(service)
    needed_last_row = next_row + len(to_insert) - 1
    print(f'\n{len(to_insert)} baris week {target_week} akan ditulis ke RAW TAP SC (baris {next_row}-{needed_last_row}).')
    ensure_rows(service, GABUNGAN_SHEET, needed_last_row)
    write_records(service, to_insert, next_row)

    print(f'\nSELESAI. {len(to_insert)} baris ditambahkan.')


if __name__ == '__main__':
    main()
