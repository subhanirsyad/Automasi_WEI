"""
IMPORT SC_GMV -> Raw Data SC (jalan lokal dari komputer, pakai Google Sheets API)
==================================================================================
Baca file SC_GMV langsung dari komputer, isi semua 26 kolom + 4 kolom formula
(cek name, CHANNEL, DATE, WEEK) ke sheet Raw Data SC.

DEDUP RINGAN: kombinasi Order ID + Product ID + SKU ID yang PERSIS sama
dengan yang udah ada di sheet akan dilewati.

FILTER BULAN: cuma order yang Time Created-nya jatuh di bulan/tahun tertentu
yang dimasukin. Set FILTER_MONTH = None kalau mau semua bulan.

TIDAK auto-nambah baris kalau kurang -- kalau baris di sheet gak cukup,
script berhenti & kasih tau berapa baris yang perlu ditambahin manual dulu.

CARA PAKAI:
1. Isi CREDENTIALS_FILE, TARGET_SPREADSHEET_ID, SOURCE_XLSX_PATH di bawah.
2. Cek nama tab data sumbernya -- BUKAN "Sheet1", tapi kayak
   "affiliate_orders_xxxxx" (angkanya beda tiap toko). Isi persis di
   SOURCE_SHEET_NAME.
3. Cek ROSTER_SHEET_NAME -- ini nama sheet roster creator bulan berjalan
   (misal "GMV Creator AUGUST"). Kalau sheet itu belum dibikin di file
   master, bikin dulu, atau ganti ke "GMV Creator JULY" sementara.
4. python import_raw_data_sc.py
"""

import csv
import datetime
from pathlib import Path

import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============== KONFIGURASI -- ISI INI ==============
CREDENTIALS_FILE = 'credentials.json'
TARGET_SPREADSHEET_ID = ''
SOURCE_XLSX_PATH = r'C:\Users\Subhan\OneDrive\Documents\Automasi\SC_GMV ellips (20-26 August) (1).csv'
SOURCE_SHEET_NAME = 'affiliate_orders_76702350440273'  # CEK & GANTI sesuai file kamu (diabaikan kalau sumbernya .csv)
TARGET_SHEET_NAME = 'Raw Data SC (20 April-31 July)'  # CEK nama tab persis di file master
ROSTER_SHEET_NAME = 'GMV Creator JULY'  # CEK -- ganti ke 'GMV Creator AUGUST' kalau sheet-nya udah ada

# Filter bulan: cuma order yang Time Created-nya jatuh di bulan/tahun ini
# yang akan dimasukin. Set FILTER_MONTH = None kalau mau semua bulan.
FILTER_YEAR = 2026
FILTER_MONTH = None  # laporan SC_GMV difilter dari "Payment time", bukan "Time Created" -- ambil semua
# ======================================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Pemetaan kolom tujuan (Raw Data SC) <- nama kolom di file sumber (SC_GMV).
COLUMN_MAP = [
    ('A', 'Order ID'),
    ('B', 'Product ID'),
    ('C', 'Product Name'),
    ('D', 'SKU ID'),
    ('E', 'Price'),
    ('F', 'Payment Amount'),
    ('G', 'Currency'),
    ('H', 'Quantity'),
    ('I', 'Fully returned or refunded'),
    ('J', 'Payment method'),
    ('K', 'Order Status'),
    ('L', 'Creator Username'),
    ('N', 'Content Type'),
    ('P', 'Content ID'),
    ('Q', 'commission model'),
    ('R', 'Standard commission rate'),
    ('S', 'Est. Commission Base'),
    ('T', 'Est. standard commission payment'),
    ('U', 'Actual Commission Base'),
    ('V', 'Actual Commission Payment'),
    ('W', 'Shop Ads commission rate'),
    ('X', 'Est. Shop Ads commission payment'),
    ('Y', 'Actual Shop Ads commission payment'),
    ('Z', 'Est. co-funded creator bonus'),
    ('AA', 'Actual co-funded creator bonus'),
    ('AB', 'Time Created'),
]


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def parse_datetime_flexible(value):
    """Parse Time Created yang bisa berupa datetime asli atau teks DD/MM/YYYY HH:MM:SS."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        s = value.strip()
        for fmt in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                return datetime.datetime.strptime(s, fmt)
            except ValueError:
                continue
    return None


def to_serial_datetime(dt):
    if dt is None:
        return ''
    base = datetime.date(1899, 12, 30)
    days = (dt.date() - base).days
    frac = (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400
    return days + frac


def read_rows_from_csv(path):
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        rows = list(csv.reader(f))
    return rows


def read_rows_from_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    if SOURCE_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Sheet "{SOURCE_SHEET_NAME}" tidak ketemu. Sheet yang ada: {wb.sheetnames}')
    ws = wb[SOURCE_SHEET_NAME]
    return list(ws.iter_rows(values_only=True))


def read_source_file(path):
    rows = read_rows_from_csv(path) if str(path).lower().endswith('.csv') else read_rows_from_xlsx(path)
    if not rows:
        raise ValueError('Sheet sumber kosong.')
    header = [str(h).strip() if h is not None else '' for h in rows[0]]

    def col_idx(name):
        return header.index(name) if name in header else -1

    src_col_index = {}
    missing = []
    for dest_col, name in COLUMN_MAP:
        idx = col_idx(name)
        if idx == -1:
            missing.append(name)
        else:
            src_col_index[dest_col] = idx
    if missing:
        raise ValueError(f'Kolom berikut tidak ketemu di header: {", ".join(missing)}')

    idx_order_id = col_idx('Order ID')
    idx_product_id = col_idx('Product ID')
    idx_sku_id = col_idx('SKU ID')
    idx_time_created = col_idx('Time Created')

    data_rows = rows[1:]
    if len(data_rows) < 4:
        print(f'PERINGATAN: sheet sumber cuma ada {len(data_rows)} baris data. '
              f'Kalau harusnya data 1 bulan penuh, kemungkinan datanya kurang lengkap.')

    valid_rows = []
    filtered_out_month = 0
    for row in data_rows:
        if not row[idx_order_id]:
            continue
        if FILTER_MONTH is not None:
            dt = parse_datetime_flexible(row[idx_time_created])
            if dt is None or (dt.year, dt.month) != (FILTER_YEAR, FILTER_MONTH):
                filtered_out_month += 1
                continue
        valid_rows.append(row)

    return valid_rows, src_col_index, idx_order_id, idx_product_id, idx_sku_id, idx_time_created, filtered_out_month


def get_sheet_id_and_row_count(service, spreadsheet_id, sheet_name):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields='sheets(properties(sheetId,title,gridProperties))'
    ).execute()
    for sh in meta['sheets']:
        if sh['properties']['title'] == sheet_name:
            return sh['properties']['sheetId'], sh['properties']['gridProperties']['rowCount']
    raise ValueError(f'Sheet "{sheet_name}" tidak ketemu di file master. Cek lagi nama tabnya persis.')


def get_sheet_row_count(service, spreadsheet_id, sheet_name):
    return get_sheet_id_and_row_count(service, spreadsheet_id, sheet_name)[1]


def ensure_rows(service, spreadsheet_id, sheet_name, rows_needed):
    """Tambah baris kosong di akhir sheet kalau grid-nya kurang."""
    sheet_id, current_rows = get_sheet_id_and_row_count(service, spreadsheet_id, sheet_name)
    if rows_needed <= 0:
        return current_rows
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [{
            'appendDimension': {'sheetId': sheet_id, 'dimension': 'ROWS', 'length': rows_needed}
        }]}
    ).execute()
    return current_rows + rows_needed


def get_existing_keys(service):
    """Baca kolom A (Order ID), B (Product ID), D (SKU ID), chunk per 30rb baris."""
    target_max_row = get_sheet_row_count(service, TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME)

    chunk_size = 30000
    col_a, col_b, col_d = [], [], []
    for start in range(1, target_max_row + 1, chunk_size):
        end = min(start + chunk_size - 1, target_max_row)
        resp = service.spreadsheets().values().batchGet(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            ranges=[f"'{TARGET_SHEET_NAME}'!A{start}:A{end}",
                    f"'{TARGET_SHEET_NAME}'!B{start}:B{end}",
                    f"'{TARGET_SHEET_NAME}'!D{start}:D{end}"]
        ).execute()
        vr = resp.get('valueRanges', [])
        a_vals = vr[0].get('values', []) if len(vr) > 0 else []
        b_vals = vr[1].get('values', []) if len(vr) > 1 else []
        d_vals = vr[2].get('values', []) if len(vr) > 2 else []
        col_a.extend(a_vals)
        col_b.extend(b_vals)
        col_d.extend(d_vals)
        if not a_vals and not b_vals and not d_vals:
            break

    header_row = None
    for i, row in enumerate(col_a):
        if row and row[0] == 'Order ID':
            header_row = i
            break
    if header_row is None:
        raise ValueError(f'Header "Order ID" tidak ketemu di kolom A sheet {TARGET_SHEET_NAME}.')

    data_start = header_row + 1
    existing_keys = set()
    last_filled = data_start - 1
    last_row = max(len(col_a), len(col_b), len(col_d))
    for i in range(data_start, last_row):
        o = col_a[i][0] if i < len(col_a) and col_a[i] else ''
        p = col_b[i][0] if i < len(col_b) and col_b[i] else ''
        s = col_d[i][0] if i < len(col_d) and col_d[i] else ''
        if o or p or s:
            last_filled = i
        if o and p:
            existing_keys.add(f'{str(o).strip()}|{str(p).strip()}|{str(s).strip()}')

    return existing_keys, data_start + 1, last_filled + 2, target_max_row


def write_new_rows(service, to_insert, src_col_index, idx_time_created, next_row):
    sheet_ref = f"'{TARGET_SHEET_NAME}'!"
    value_ranges = []

    for idx, row in enumerate(to_insert):
        r = next_row + idx
        for dest_col, _ in COLUMN_MAP:
            val = row[src_col_index[dest_col]]
            if dest_col == 'AB':
                dt = parse_datetime_flexible(val)
                val = to_serial_datetime(dt)
            value_ranges.append({'range': f'{sheet_ref}{dest_col}{r}', 'values': [[val if val is not None else '']]})

        value_ranges.append({
            'range': f'{sheet_ref}M{r}',
            'values': [[f"=VLOOKUP(L{r};'{ROSTER_SHEET_NAME}'!$B:$B;1;0)"]]
        })
        value_ranges.append({
            'range': f'{sheet_ref}O{r}',
            'values': [[f'=IF(ISNUMBER(SEARCH("livestream";N{r}));"LIVE STREAMING";'
                        f'IF(ISNUMBER(SEARCH("video";N{r}));"SHORT VIDEO";"SHARE LINK"))']]
        })
        value_ranges.append({'range': f'{sheet_ref}AC{r}', 'values': [[f'=TEXT(AB{r};"yyyy/mm/dd")']]})
        value_ranges.append({
            'range': f'{sheet_ref}AD{r}',
            'values': [[f'=UPPER(TEXT(DATEVALUE(AC{r});"mmm"))&" W"&IF(DAY(DATEVALUE(AC{r}))<=8;1;'
                        f'IF(DAY(DATEVALUE(AC{r}))<=15;2;IF(DAY(DATEVALUE(AC{r}))<=22;3;'
                        f'IF(DAY(DATEVALUE(AC{r}))<=29;4;5))))&" ("&IF(DAY(DATEVALUE(AC{r}))<=8;"1-8";'
                        f'IF(DAY(DATEVALUE(AC{r}))<=15;"9-15";IF(DAY(DATEVALUE(AC{r}))<=22;"16-22";'
                        f'IF(DAY(DATEVALUE(AC{r}))<=29;"23-29";"30-"&DAY(EOMONTH(DATEVALUE(AC{r});0))))))&")"']]
        })

    entries_per_row = len(COLUMN_MAP) + 4
    chunk_rows = 300
    for start in range(0, len(to_insert), chunk_rows):
        n = min(chunk_rows, len(to_insert) - start)
        chunk = value_ranges[start * entries_per_row: (start + n) * entries_per_row]
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            body={'valueInputOption': 'USER_ENTERED', 'data': chunk}
        ).execute()
        print(f'  ... tertulis {start + n}/{len(to_insert)} baris')


def main():
    if not Path(SOURCE_XLSX_PATH).exists():
        print(f'ERROR: file tidak ketemu di path: {SOURCE_XLSX_PATH}')
        return

    print('Membaca file sumber...')
    valid_rows, src_col_index, idx_order_id, idx_product_id, idx_sku_id, idx_time_created, filtered_out_month = \
        read_source_file(SOURCE_XLSX_PATH)
    print(f'Ketemu {len(valid_rows)} baris data valid.')
    if FILTER_MONTH is not None:
        print(f'Filter bulan aktif ({FILTER_MONTH}/{FILTER_YEAR}): {filtered_out_month} order di luar bulan itu dilewati.')
    if not valid_rows:
        print('Tidak ada baris data valid, berhenti.')
        return

    print('Connect ke Google Sheets, ambil data existing buat dedup...')
    service = get_sheets_service()
    existing_keys, data_start_sheet_row, next_row, target_max_row = get_existing_keys(service)
    print(f'Ada {len(existing_keys)} kombinasi Order ID+Product ID+SKU ID yang udah ada di sheet.')

    to_insert = []
    skipped = 0
    seen_now = set(existing_keys)
    for row in valid_rows:
        key = f'{str(row[idx_order_id]).strip()}|{str(row[idx_product_id]).strip()}|{str(row[idx_sku_id] or "").strip()}'
        if key in seen_now:
            skipped += 1
            continue
        seen_now.add(key)
        to_insert.append(row)

    if not to_insert:
        print(f'\nSemua {len(valid_rows)} baris sudah ada sebelumnya (kombinasi Order ID+Product ID+SKU ID sama).')
        return

    needed_last_row = next_row + len(to_insert) - 1
    if needed_last_row > target_max_row:
        rows_short = needed_last_row - target_max_row
        print(f'\nGrid sheet {TARGET_SHEET_NAME} kurang {rows_short} baris (butuh sampai baris {needed_last_row}, '
              f'tersedia sampai {target_max_row}). Nambah otomatis...')
        target_max_row = ensure_rows(service, TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME, rows_short)
        print(f'Baris ditambah, grid sekarang sampai baris {target_max_row}.')

    print(f'\n{len(to_insert)} baris baru akan ditambahkan ke baris {next_row}-{next_row + len(to_insert) - 1}.')
    print(f'{skipped} dilewati (kombinasi Order ID+Product ID+SKU ID sudah ada).')
    print('Menulis ke Google Sheets...')
    write_new_rows(service, to_insert, src_col_index, idx_time_created, next_row)

    print(f'\nSELESAI. {len(to_insert)} baris baru ditambahkan, {skipped} dilewati.')


if __name__ == '__main__':
    main()
