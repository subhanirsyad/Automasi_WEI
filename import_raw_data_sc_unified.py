"""
IMPORT SC_GMV -> Raw Data SC / SC Cussons  [UNIFIED: Ellips & Cussons]
================================================================================
Satu script, brand dideteksi otomatis dari NAMA FILE SOURCE_XLSX_PATH (harus
mengandung kata "ellips" atau "cussons", case-insensitive). Struktur kolom &
formula tujuan beda antar brand -- semuanya di-drive dari PROFILES di bawah,
bukan hardcode di logic.

- ELLIPS -> tab "Raw Data SC (20 April-31 July)", 26 kolom + field komisi
  lengkap (commission model, Shop Ads commission, co-funded creator bonus, dll)
- CUSSONS -> tab "SC Cussons (1-30 September)", 19 kolom (A-S), lebih sederhana, gak ada
  field komisi selengkap Ellips. (Formatnya diasumsikan sama kayak SC_GMV
  Ellips karena belum ada contoh file SC_GMV Cussons asli saat script ini
  dibuat -- validasi ulang begitu file beneran ada.)

DEDUP (kedua brand sama): kombinasi Order ID + Product ID + SKU ID yang PERSIS
sama dengan yang udah ada di sheet akan dilewati.

CARA PAKAI:
1. Isi SOURCE_XLSX_PATH -- nama filenya HARUS ada kata "ellips" atau "cussons".
2. python import_raw_data_sc_unified.py
"""

import csv
import datetime
from pathlib import Path

import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ============== KONFIGURASI -- ISI INI ==============
CREDENTIALS_FILE = 'credentials.json'
SOURCE_XLSX_PATH = r''
SOURCE_SHEET_NAME = ''  # CEK & GANTI sesuai file kamu (diabaikan kalau sumbernya .csv)

# Filter bulan: cuma order yang Time Created-nya jatuh di bulan/tahun ini
# yang akan dimasukin. Set FILTER_MONTH = None kalau mau semua bulan.
FILTER_YEAR = 2026
FILTER_MONTH = None  # laporan SC_GMV gak selalu align ke Time Created -- default ambil semua
# ======================================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def col_map_ellips():
    return [
        ('A', 'Order ID'), ('B', 'Product ID'), ('C', 'Product Name'), ('D', 'SKU ID'),
        ('E', 'Price'), ('F', 'Payment Amount'), ('G', 'Currency'), ('H', 'Quantity'),
        ('I', 'Fully returned or refunded'), ('J', 'Payment method'), ('K', 'Order Status'),
        ('L', 'Creator Username'), ('N', 'Content Type'), ('P', 'Content ID'),
        ('Q', 'commission model'), ('R', 'Standard commission rate'), ('S', 'Est. Commission Base'),
        ('T', 'Est. standard commission payment'), ('U', 'Actual Commission Base'),
        ('V', 'Actual Commission Payment'), ('W', 'Shop Ads commission rate'),
        ('X', 'Est. Shop Ads commission payment'), ('Y', 'Actual Shop Ads commission payment'),
        ('Z', 'Est. co-funded creator bonus'), ('AA', 'Actual co-funded creator bonus'),
        ('AB', 'Time Created'),
    ]


def col_map_cussons():
    return [
        ('A', 'Order ID'), ('B', 'Product ID'), ('C', 'Product Name'), ('D', 'SKU ID'),
        ('E', 'Price'), ('F', 'Payment Amount'), ('G', 'Quantity'),
        ('H', 'Fully returned or refunded'), ('I', 'Payment method'), ('J', 'Order Status'),
        ('K', 'Creator Username'), ('M', 'Content Type'), ('O', 'Content ID'),
        ('P', 'Est. Commission Base'), ('Q', 'Time Created'),
    ]


def write_formulas_ellips(sheet_ref, r, roster_sheet):
    """Ellips: M=Cek Name, O=Channel, AB=Time Created(udah ditulis via COLUMN_MAP),
    AC=Date(TEXT), AD=WEEK."""
    return [
        {'range': f'{sheet_ref}M{r}', 'values': [[f"=VLOOKUP(L{r};'{roster_sheet}'!$B:$B;1;0)"]]},
        {'range': f'{sheet_ref}O{r}', 'values': [[
            f'=IF(ISNUMBER(SEARCH("livestream";N{r}));"LIVE STREAMING";'
            f'IF(ISNUMBER(SEARCH("video";N{r}));"SHORT VIDEO";"SHARE LINK"))'
        ]]},
        {'range': f'{sheet_ref}AC{r}', 'values': [[f'=TEXT(AB{r};"yyyy/mm/dd")']]},
        {'range': f'{sheet_ref}AD{r}', 'values': [[
            f'=UPPER(TEXT(DATEVALUE(AC{r});"mmm"))&" W"&IF(DAY(DATEVALUE(AC{r}))<=8;1;'
            f'IF(DAY(DATEVALUE(AC{r}))<=15;2;IF(DAY(DATEVALUE(AC{r}))<=22;3;'
            f'IF(DAY(DATEVALUE(AC{r}))<=29;4;5))))&" ("&IF(DAY(DATEVALUE(AC{r}))<=8;"1-8";'
            f'IF(DAY(DATEVALUE(AC{r}))<=15;"9-15";IF(DAY(DATEVALUE(AC{r}))<=22;"16-22";'
            f'IF(DAY(DATEVALUE(AC{r}))<=29;"23-29";"30-"&DAY(EOMONTH(DATEVALUE(AC{r});0))))))&")"'
        ]]},
    ]


def write_formulas_cussons(sheet_ref, r, roster_sheet, row=None, src_col_index=None, idx_time_created=None):
    """Cussons: L=Cek Creator, N=Channels, R=Date(nilai polos, bukan formula), S=WEEK."""
    roster_sheet_escaped = roster_sheet.replace("'", "''")
    dt = parse_datetime_flexible(row[idx_time_created])
    entries = [
        {'range': f'{sheet_ref}R{r}', 'values': [[to_serial_date_only(dt)]]},
        {'range': f'{sheet_ref}L{r}', 'values': [[f"=VLOOKUP(K{r};'{roster_sheet_escaped}'!$B:$B;1;0)"]]},
        {'range': f'{sheet_ref}N{r}', 'values': [[
            f'=IF(ISNUMBER(SEARCH("Livestream";M{r}));"LIVE STREAMING";'
            f'IF(ISNUMBER(SEARCH("Video";M{r}));"SHORT VIDEO";"SHARELINK"))'
        ]]},
        {'range': f'{sheet_ref}S{r}', 'values': [[
            f'=UPPER(TEXT(R{r};"MMM"))&" W"&IF(DAY(R{r})<=6;1;IF(DAY(R{r})<=13;2;IF(DAY(R{r})<=20;3;'
            f'IF(DAY(R{r})<=27;4;5))))&" ("&IF(DAY(R{r})<=6;1;IF(DAY(R{r})<=13;7;IF(DAY(R{r})<=20;14;'
            f'IF(DAY(R{r})<=27;21;28))))&"-"&IF(DAY(R{r})<=6;6;IF(DAY(R{r})<=13;13;IF(DAY(R{r})<=20;20;'
            f'IF(DAY(R{r})<=27;27;DAY(EOMONTH(R{r};0))))))&")"'
        ]]},
    ]
    return entries


PROFILES = {
    'ELLIPS': {
        'spreadsheet_id': '1tIG9FhUogXwBJK6YuzpT19nFlJs5EDfXA6EYQ493paE',
        'sheet': 'Raw Data SC (20 April-31 August)',
        'roster_sheet': 'GMV Creator [SEPT]',
        'column_map': col_map_ellips(),
        'extra_fields_per_row': 4,  # M, O, AC, AD
        'needs_row_data_for_formula': False,
    },
    'CUSSONS': {
        'spreadsheet_id': '1ZBOvn5fReBECSzgXrAS6SNuq7tWDaem9Cf_QhQN4b_8',
        'sheet': 'SC Cussons (1-30 September)',
        'roster_sheet': "Creator Performance PZ Cussons September'26",
        'column_map': col_map_cussons(),
        'extra_fields_per_row': 4,  # R, L, N, S
        'needs_row_data_for_formula': True,
    },
}


def detect_brand_from_filename(path):
    name = Path(path).name.lower()
    if 'ellips' in name:
        return 'ELLIPS'
    if 'cussons' in name:
        return 'CUSSONS'
    raise ValueError(
        f'Gak bisa deteksi brand dari nama file "{Path(path).name}" -- '
        'pastikan nama filenya mengandung kata "ellips" atau "cussons".'
    )


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def parse_datetime_flexible(value):
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


def to_serial_date_only(dt):
    if dt is None:
        return ''
    base = datetime.date(1899, 12, 30)
    return (dt.date() - base).days


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


def read_source_file(path, column_map):
    rows = read_rows_from_csv(path) if str(path).lower().endswith('.csv') else read_rows_from_xlsx(path)
    if not rows:
        raise ValueError('Sheet sumber kosong.')
    header = [str(h).strip() if h is not None else '' for h in rows[0]]

    def col_idx(name):
        return header.index(name) if name in header else -1

    src_col_index = {}
    missing = []
    for dest_col, name in column_map:
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
              f'Kalau harusnya data 1 periode penuh, kemungkinan datanya kurang lengkap.')

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


def ensure_rows(service, spreadsheet_id, sheet_name, rows_needed):
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


def get_existing_keys(service, spreadsheet_id, sheet_name):
    """Baca kolom A (Order ID), B (Product ID), D (SKU ID), chunk per 30rb baris."""
    _, target_max_row = get_sheet_id_and_row_count(service, spreadsheet_id, sheet_name)

    chunk_size = 30000
    col_a, col_b, col_d = [], [], []
    for start in range(1, target_max_row + 1, chunk_size):
        end = min(start + chunk_size - 1, target_max_row)
        resp = service.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=[f"'{sheet_name}'!A{start}:A{end}",
                    f"'{sheet_name}'!B{start}:B{end}",
                    f"'{sheet_name}'!D{start}:D{end}"]
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
        raise ValueError(f'Header "Order ID" tidak ketemu di kolom A sheet {sheet_name}.')

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


def force_text_if_long_id(val):
    """ID kayak Order ID/Product ID/SKU ID/Content ID itu 16+ digit -- kalau dikirim
    apa adanya ke Sheets pakai USER_ENTERED, Sheets nganggep itu ANGKA dan digit
    belakangnya kepotong (presisi float cuma akurat ~15-16 digit). Kasih awalan
    petik satu (') biar Sheets simpen persis sebagai teks."""
    s = str(val) if val is not None else ''
    if s.isdigit() and len(s) >= 16:
        return "'" + s
    return val


def write_new_rows(service, brand, profile, to_insert, src_col_index, idx_time_created, next_row):
    spreadsheet_id = profile['spreadsheet_id']
    sheet_ref = f"'{profile['sheet']}'!"
    column_map = profile['column_map']
    value_ranges = []

    for idx, row in enumerate(to_insert):
        r = next_row + idx
        for dest_col, _ in column_map:
            val = row[src_col_index[dest_col]]
            if dest_col == ('AB' if brand == 'ELLIPS' else 'Q'):  # kolom Time Created
                dt = parse_datetime_flexible(val)
                val = to_serial_datetime(dt)
            else:
                val = force_text_if_long_id(val)
            value_ranges.append({'range': f'{sheet_ref}{dest_col}{r}', 'values': [[val if val is not None else '']]})

        if brand == 'ELLIPS':
            value_ranges.extend(write_formulas_ellips(sheet_ref, r, profile['roster_sheet']))
        else:
            value_ranges.extend(write_formulas_cussons(
                sheet_ref, r, profile['roster_sheet'],
                row=row, src_col_index=src_col_index, idx_time_created=idx_time_created
            ))

    entries_per_row = len(column_map) + profile['extra_fields_per_row']
    chunk_rows = 700
    for start in range(0, len(to_insert), chunk_rows):
        n = min(chunk_rows, len(to_insert) - start)
        chunk = value_ranges[start * entries_per_row: (start + n) * entries_per_row]
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'valueInputOption': 'USER_ENTERED', 'data': chunk}
        ).execute()
        print(f'  ... tertulis {start + n}/{len(to_insert)} baris')


def main():
    if not Path(SOURCE_XLSX_PATH).exists():
        print(f'ERROR: file tidak ketemu di path: {SOURCE_XLSX_PATH}')
        return

    brand = detect_brand_from_filename(SOURCE_XLSX_PATH)
    profile = PROFILES[brand]
    print(f'Brand terdeteksi dari nama file: {brand}  ->  target tab "{profile["sheet"]}"')

    print('Membaca file sumber...')
    valid_rows, src_col_index, idx_order_id, idx_product_id, idx_sku_id, idx_time_created, filtered_out_month = \
        read_source_file(SOURCE_XLSX_PATH, profile['column_map'])
    print(f'Ketemu {len(valid_rows)} baris data valid.')
    if FILTER_MONTH is not None:
        print(f'Filter bulan aktif ({FILTER_MONTH}/{FILTER_YEAR}): {filtered_out_month} order di luar bulan itu dilewati.')
    if not valid_rows:
        print('Tidak ada baris data valid, berhenti.')
        return

    print('Connect ke Google Sheets, ambil data existing buat dedup...')
    service = get_sheets_service()
    existing_keys, data_start_sheet_row, next_row, target_max_row = get_existing_keys(
        service, profile['spreadsheet_id'], profile['sheet']
    )
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
        print(f'\nGrid sheet {profile["sheet"]} kurang {rows_short} baris (butuh sampai baris {needed_last_row}, '
              f'tersedia sampai {target_max_row}). Nambah otomatis...')
        target_max_row = ensure_rows(service, profile['spreadsheet_id'], profile['sheet'], rows_short)
        print(f'Baris ditambah, grid sekarang sampai baris {target_max_row}.')

    print(f'\n{len(to_insert)} baris baru akan ditambahkan ke baris {next_row}-{next_row + len(to_insert) - 1}.')
    print(f'{skipped} dilewati (kombinasi Order ID+Product ID+SKU ID sudah ada).')
    print('Menulis ke Google Sheets...')
    write_new_rows(service, brand, profile, to_insert, src_col_index, idx_time_created, next_row)

    print(f'\nSELESAI. {len(to_insert)} baris baru ditambahkan, {skipped} dilewati.')


if __name__ == '__main__':
    main()
