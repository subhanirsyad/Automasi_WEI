"""
IMPORT VIDEO BANK - UNIFIED (Ellips & Cussons, TAP_SV & SC_SV)
================================================================================
Satu script buat semua kombinasi:
  - Target: ELLIPS atau CUSSONS (dideteksi dari TARGET_SPREADSHEET_ID)
  - Sumber: TAP (TAP_SV) atau SC (SC_SV) (dideteksi dari nama file / isinya)

CARA PAKAI:
1. Isi TARGET_SPREADSHEET_ID (pilih salah satu dari daftar di TARGET_PROFILES)
   dan SOURCE_XLSX_PATH di bawah.
2. python import_video_bank_unified.py
"""

import re
import csv
import datetime
from pathlib import Path

import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ============== KONFIGURASI -- ISI INI ==============
CREDENTIALS_FILE = 'credentials.json'

TARGET_SPREADSHEET_ID = '1ZBOvn5fReBECSzgXrAS6SNuq7tWDaem9Cf_QhQN4b_8'  # Cussons
# TARGET_SPREADSHEET_ID = '1tIG9FhUogXwBJK6YuzpT19nFlJs5EDfXA6EYQ493paE'  # Ellips

SOURCE_XLSX_PATH = r'C:\Users\Subhan\OneDrive\Documents\Automasi\TAP_SV Cussons (1-30 August).xlsx'
SOURCE_SHEET_NAME = 'Custom report'

FILTER_YEAR = 2026
FILTER_MONTH = 8  # 8 = Agustus. Set None kalau mau semua bulan.
# ======================================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ---------- PROFIL TARGET (Ellips vs Cussons) ----------
# Kolom VIDEO BANK beda antar client -- ELLIPS gak punya kolom "Cek"/"Product ID"
# terpisah, CUSSONS punya. Semua ini di-drive dari profil, bukan hardcode di logic.
TARGET_PROFILES = {
    '1tIG9FhUogXwBJK6YuzpT19nFlJs5EDfXA6EYQ493paE': {
        'name': 'ELLIPS',
        'video_bank_sheet': 'VIDEO BANK',
        'cols': {
            'creator': 'B', 'videoid': 'E', 'link': 'F', 'count': 'G',
            'uploaddate': 'H', 'day': 'I', 'month': 'J', 'year': 'K', 'product': 'L',
        },
        'has_cek_col': False,
        'has_product_id_col': False,
        'roster_sheet': None,   # Ellips: roster cuma dipakai utk sumber SC, diisi manual per-run
        'roster_col': None,
    },
    '1ZBOvn5fReBECSzgXrAS6SNuq7tWDaem9Cf_QhQN4b_8': {
        'name': 'CUSSONS',
        'video_bank_sheet': ' VIDEO BANK',  # PERHATIAN: ada spasi di depan nama tab
        'cols': {
            'creator': 'B', 'videoid': 'E', 'cek': 'F', 'link': 'G', 'count': 'H',
            'uploaddate': 'I', 'day': 'J', 'month': 'K', 'year': 'L',
            'product': 'M', 'productid': 'N',
        },
        'has_cek_col': True,
        'has_product_id_col': True,
        'roster_sheet': "Creator Performance PZ Cussons August'26",
        'roster_col': 'B',
    },
}

_PRODUCT_ID_RE = re.compile(r'\((\d{5,})\)\s*$')


def get_target_profile():
    profile = TARGET_PROFILES.get(TARGET_SPREADSHEET_ID)
    if profile is None:
        raise ValueError(
            f'TARGET_SPREADSHEET_ID "{TARGET_SPREADSHEET_ID}" belum ada di TARGET_PROFILES. '
            'Tambahin dulu profilnya (nama tab VIDEO BANK, kolom-kolomnya) sebelum run.'
        )
    return profile


def load_source_rows(path):
    """Baca semua baris dari file sumber, support .xlsx maupun .csv.
    Export dari TikTok kadang nyampe sebagai .csv (bukan .xlsx) walau nama
    filenya masih dianggap 'sumber xlsx' -- baca dua-duanya lewat jalur sama
    biar detect_source_type/read_tap_sv/read_sc_sv gak perlu peduli formatnya."""
    ext = Path(path).suffix.lower()
    if ext == '.csv':
        with open(path, newline='', encoding='utf-8-sig') as f:
            return [tuple(row) for row in csv.reader(f)]
    # read_only=False: banyak file export TikTok punya tag <dimension> yang
    # salah/kosong -- mode read_only openpyxl percaya tag itu apa adanya dan
    # bisa kepotong cuma jadi 1 baris. read_only=False parse penuh, aman.
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    sheet_name = SOURCE_SHEET_NAME if SOURCE_SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    return list(ws.iter_rows(values_only=True))


def detect_source_type(path):
    """TAP_SV: nama file mengandung 'TAP_SV' / header row-0 langsung punya 'Video ID'.
       SC_SV: nama file mengandung 'SC_SV' / row-0 cuma metadata (tanggal), header di row index 2."""
    fname = Path(path).name.lower()
    if 'tap_sv' in fname or 'tap sv' in fname:
        return 'TAP'
    if 'sc_sv' in fname or 'sc sv' in fname:
        return 'SC'
    # fallback: intip isi file
    rows = load_source_rows(path)
    first_row = rows[0] if rows else None
    if first_row and 'Video ID' in first_row:
        return 'TAP'
    return 'SC'


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


def extract_year_month(value):
    if value is None:
        return None
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.year, value.month
    if isinstance(value, str):
        s = value.strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d %H:%M:%S', '%Y/%m/%d',
                    '%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %H:%M', '%m/%d/%Y'):
            try:
                dt = datetime.datetime.strptime(s[:19] if ' ' in s else s, fmt)
                return dt.year, dt.month
            except ValueError:
                continue
    return None


def to_serial_date(dt):
    if dt is None:
        return ''
    if isinstance(dt, str):
        s = dt.strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d %H:%M:%S', '%Y/%m/%d',
                    '%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %H:%M', '%m/%d/%Y'):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                break
            except ValueError:
                continue
        else:
            return dt
    base = datetime.date(1899, 12, 30)
    d = dt.date() if isinstance(dt, datetime.datetime) else dt
    return (d - base).days


def extract_product_id_from_text(text):
    """Kalo nama produk udah ada '(ID)' di belakangnya (format gabungan), ambil ID-nya."""
    if not text:
        return ''
    m = _PRODUCT_ID_RE.search(str(text))
    return m.group(1) if m else ''


def get_creator_roster(service, profile):
    """Cuma dipakai kalo source_type == SC dan profile.roster_sheet keisi."""
    if not profile['roster_sheet']:
        raise ValueError('Sumber SC butuh roster, tapi profil target ini gak punya roster_sheet.')
    result = service.spreadsheets().values().get(
        spreadsheetId=TARGET_SPREADSHEET_ID,
        range=f"'{profile['roster_sheet']}'!{profile['roster_col']}1:{profile['roster_col']}5000"
    ).execute()
    values = result.get('values', [])
    roster = set()
    for row in values:
        if row and row[0] and row[0].strip():
            v = row[0].strip().lower()
            if v not in ('creatorusername', 'username'):
                roster.add(v)
    return roster


MONTH_ABBR_ID = {
    1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MEI', 6: 'JUN',
    7: 'JUL', 8: 'AGUS', 9: 'SEPT', 10: 'OKT', 11: 'NOV', 12: 'DES',
}
MONTH_NAME_EN = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June',
    7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December',
}


def roster_sheet_name(brand, year, month):
    """Nama tab roster creator per bulan -- beda pattern per brand.
    ELLIPS: 'GMV Creator [AGUS]' / CUSSONS: "Creator Performance PZ Cussons August'26"."""
    if brand == 'ELLIPS':
        return f'GMV Creator [{MONTH_ABBR_ID[month]}]'
    return f"Creator Performance PZ Cussons {MONTH_NAME_EN[month]}'{str(year)[-2:]}"


def sync_creator_roster(service, spreadsheet_id, sheet_name, creator_names):
    """TAP-only: creator di TAP itu cuma creator yang kita handle sendiri (beda dari
    SC yang gabungan sama creator luar), jadi kalau ada creator baru yang belum
    tercatat di roster bulan itu berarti roster-nya belum di-update -- tambahin
    otomatis ke kolom B baris kosong paling bawah."""
    col = 'B'
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!{col}:{col}"
        ).execute()
    except HttpError as e:
        raise ValueError(f'Tab roster "{sheet_name}" tidak ketemu / gagal dibaca: {e}')
    values = result.get('values', [])

    existing = set()
    last_filled = 0
    for i, row in enumerate(values):
        v = row[0].strip() if row and row[0] else ''
        if not v or v.lower() == 'creatorusername':
            continue
        existing.add(v.lower())
        last_filled = i + 1

    seen = set()
    missing = []
    for name in creator_names:
        name = (name or '').strip()
        if not name or name.lower() in seen:
            continue
        seen.add(name.lower())
        if name.lower() not in existing:
            missing.append(name)

    if not missing:
        print(f'  Semua creator sudah ada di roster "{sheet_name}".')
        return

    start_row = last_filled + 1
    data = [{'range': f"'{sheet_name}'!{col}{start_row + i}", 'values': [[name]]}
            for i, name in enumerate(missing)]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'valueInputOption': 'USER_ENTERED', 'data': data}
    ).execute()
    print(f'  {len(missing)} creator baru ditambahin ke roster "{sheet_name}" '
          f'(baris {start_row}-{start_row + len(missing) - 1}): {", ".join(missing)}')


def sync_roster_from_records(service, spreadsheet_id, brand, records_with_date):
    """records_with_date: list of (creator_name, date_value). Dikelompokin per bulan,
    lalu di-sync ke tab roster masing-masing bulan."""
    groups = {}
    for creator, date_val in records_with_date:
        ym = extract_year_month(date_val)
        if ym is None:
            continue
        groups.setdefault(ym, []).append(creator)

    for (year, month), creators in groups.items():
        sheet_name = roster_sheet_name(brand, year, month)
        print(f'Cek roster creator TAP di "{sheet_name}"...')
        sync_creator_roster(service, spreadsheet_id, sheet_name, creators)


def read_tap_sv(path):
    """Format TAP_SV: header di row 0, ada 'Video ID'/'Post time' langsung."""
    rows = load_source_rows(path)
    if not rows:
        raise ValueError('Sheet sumber kosong.')
    header = [str(h).strip() if h is not None else '' for h in rows[0]]

    def col_idx(name):
        return header.index(name) if name in header else -1

    i_creator = col_idx('Creator name')
    i_videoid = col_idx('Video ID')
    i_posttime = col_idx('Post time')
    i_product = col_idx('Product name')
    i_productid = col_idx('Product ID')

    missing = [n for n, i in [('Creator name', i_creator), ('Video ID', i_videoid),
                               ('Post time', i_posttime), ('Product name', i_product)] if i == -1]
    if missing:
        raise ValueError(f'Kolom berikut tidak ketemu di header: {", ".join(missing)}')

    data_rows = rows[1:]
    seen = set()
    records = []
    dup_in_source = 0
    filtered_out_month = 0
    for row in data_rows:
        video_id_raw = row[i_videoid]
        if not video_id_raw or video_id_raw == '-':
            continue
        video_id = str(video_id_raw).strip()
        if video_id in seen:
            dup_in_source += 1
            continue
        post_time = row[i_posttime]
        if FILTER_MONTH is not None:
            ym = extract_year_month(post_time)
            if ym != (FILTER_YEAR, FILTER_MONTH):
                filtered_out_month += 1
                continue
        seen.add(video_id)
        records.append({
            'creator': str(row[i_creator] or '').strip(),
            'video_id': video_id,
            'post_time': post_time,
            'product': row[i_product] or '',
            'product_id': str(row[i_productid]).strip() if i_productid != -1 and row[i_productid] else '',
        })

    return records, dup_in_source, filtered_out_month


def read_sc_sv(path):
    """Format SC_SV beda template per brand -- header row-nya dicari otomatis
    (baris pertama di antara 5 baris awal yang punya kolom 'Video ID'):
      - Ellips: row0 = metadata (tanggal range), row1 kosong, header di row2.
        Kolom 'Products' bisa gabungan banyak produk -> ambil yg pertama, ID
        produk diekstrak dari teks "(angka)" di belakang nama produk.
      - Cussons: header langsung di row0, row1 = deskripsi tiap kolom (bukan
        data, dilewat otomatis krn Video ID-nya kosong). Kolom tanggalnya
        'Post date' (bukan 'Time'), dan produk udah 'Product ID' langsung
        tanpa nama produk."""
    rows = load_source_rows(path)
    if len(rows) < 2:
        raise ValueError('Sheet sumber kosong/kurang baris.')

    header = None
    for row in rows[:5]:
        cells = [str(c).strip() if c is not None else '' for c in row]
        if 'Video ID' in cells:
            header = cells
            header_row_idx = rows.index(row)
            break
    if header is None:
        raise ValueError('Baris header (yang punya kolom "Video ID") tidak ketemu di 5 baris pertama.')

    def col_idx(name):
        return header.index(name) if name in header else -1

    i_creator = col_idx('Creator name')
    i_videoid = col_idx('Video ID')
    i_time = col_idx('Time')
    if i_time == -1:
        i_time = col_idx('Post date')
    i_products = col_idx('Products')
    i_productid = col_idx('Product ID')

    missing = [n for n, i in [('Creator name', i_creator), ('Video ID', i_videoid)] if i == -1]
    if i_time == -1:
        missing.append('Time / Post date')
    if i_products == -1 and i_productid == -1:
        missing.append('Products / Product ID')
    if missing:
        raise ValueError(f'Kolom berikut tidak ketemu di header: {", ".join(missing)}')

    data_rows = rows[header_row_idx + 1:]
    seen = set()
    records = []
    dup_in_source = 0
    filtered_out_month = 0
    for row in data_rows:
        video_id_raw = row[i_videoid]
        if not video_id_raw or video_id_raw == '-':
            continue
        video_id = str(video_id_raw).strip()
        if video_id in seen:
            dup_in_source += 1
            continue
        post_time = row[i_time]
        if FILTER_MONTH is not None:
            ym = extract_year_month(post_time)
            if ym != (FILTER_YEAR, FILTER_MONTH):
                filtered_out_month += 1
                continue
        seen.add(video_id)

        if i_products != -1:
            raw_products = row[i_products] or ''
            first = str(raw_products).split(',')[0].strip()
            m = _PRODUCT_ID_RE.search(first)
            product_id = m.group(1) if m else ''
            product_name = _PRODUCT_ID_RE.sub('', first).strip()
        else:
            raw_productid = row[i_productid] or ''
            product_id = str(raw_productid).split(',')[0].strip()
            product_name = ''

        records.append({
            'creator': str(row[i_creator] or '').strip(),
            'video_id': video_id,
            'post_time': post_time,
            'product': product_name,
            'product_id': product_id,
        })

    return records, dup_in_source, filtered_out_month


def get_existing_video_ids(service, profile):
    col = profile['cols']['videoid']
    sheet = profile['video_bank_sheet']
    result = service.spreadsheets().values().get(
        spreadsheetId=TARGET_SPREADSHEET_ID,
        range=f"'{sheet}'!{col}:{col}"
    ).execute()
    values = result.get('values', [])

    header_row = None
    for i, row in enumerate(values):
        if row and row[0] == 'Video ID':
            header_row = i
            break
    if header_row is None:
        raise ValueError(f'Header "Video ID" tidak ketemu di kolom {col} sheet "{sheet}".')

    data_start = header_row + 1
    existing_ids = set()
    last_filled = data_start - 1
    for i in range(data_start, len(values)):
        vid = values[i][0] if values[i] else None
        if vid:
            existing_ids.add(str(vid).strip())
            last_filled = i

    return existing_ids, data_start + 1, last_filled + 2


def ensure_enough_rows(service, profile, needed_last_row):
    """Grid sheet punya batas rowCount tetap -- kalau baris yang mau ditulis
    lewat batas itu, batchUpdate values bakal error 'exceeds grid limits'.
    Perbesar grid-nya dulu kalau perlu."""
    sheet_name = profile['video_bank_sheet']
    meta = service.spreadsheets().get(
        spreadsheetId=TARGET_SPREADSHEET_ID,
        fields='sheets(properties(sheetId,title,gridProperties(rowCount)))'
    ).execute()
    target = next((sh['properties'] for sh in meta.get('sheets', [])
                    if sh['properties']['title'] == sheet_name), None)
    if target is None:
        raise ValueError(f'Sheet "{sheet_name}" tidak ketemu waktu cek ukuran grid.')

    current_rows = target['gridProperties']['rowCount']
    if needed_last_row > current_rows:
        add_rows = needed_last_row - current_rows + 500
        service.spreadsheets().batchUpdate(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            body={'requests': [{
                'appendDimension': {
                    'sheetId': target['sheetId'],
                    'dimension': 'ROWS',
                    'length': add_rows,
                }
            }]}
        ).execute()
        print(f'  Grid sheet "{sheet_name}" cuma {current_rows} baris, ditambah {add_rows} baris jadi {current_rows + add_rows}.')


def write_new_records(service, profile, to_insert, next_row):
    cols = profile['cols']
    sheet_ref = f"'{profile['video_bank_sheet']}'!"
    value_ranges = []

    for idx, rec in enumerate(to_insert):
        r = next_row + idx
        value_ranges.append({'range': f"{sheet_ref}{cols['creator']}{r}", 'values': [[rec['creator']]]})
        value_ranges.append({'range': f"{sheet_ref}{cols['videoid']}{r}", 'values': [[rec['video_id']]]})
        value_ranges.append({'range': f"{sheet_ref}{cols['uploaddate']}{r}", 'values': [[to_serial_date(rec['post_time'])]]})

        if profile['has_product_id_col']:
            pid = rec['product_id'] or extract_product_id_from_text(rec['product'])
            if rec['product']:
                product_text = f"{rec['product']}({pid})" if pid else rec['product']
            else:
                product_text = ''
            value_ranges.append({'range': f"{sheet_ref}{cols['product']}{r}", 'values': [[product_text]]})
            value_ranges.append({'range': f"{sheet_ref}{cols['productid']}{r}", 'values': [[f"'{pid}" if pid else '']]})
        else:
            value_ranges.append({'range': f"{sheet_ref}{cols['product']}{r}", 'values': [[rec['product']]]})

        value_ranges.append({
            'range': f"{sheet_ref}{cols['link']}{r}",
            'values': [[f"=CONCATENATE(\"https://www.tiktok.com/\";\"@\";{cols['creator']}{r};\"/\";\"video\";\"/\";{cols['videoid']}{r})"]]
        })
        value_ranges.append({'range': f"{sheet_ref}{cols['count']}{r}", 'values': [[f"=IF(ISBLANK({cols['videoid']}{r});0;1)"]]})
        value_ranges.append({'range': f"{sheet_ref}{cols['day']}{r}", 'values': [[f"=IF(${cols['uploaddate']}{r}=\"\";\"\";DAY(${cols['uploaddate']}{r}))"]]})
        value_ranges.append({'range': f"{sheet_ref}{cols['month']}{r}", 'values': [[f"=IF(${cols['uploaddate']}{r}=\"\";\"\";MONTH(${cols['uploaddate']}{r}))"]]})
        value_ranges.append({'range': f"{sheet_ref}{cols['year']}{r}", 'values': [[f"=IF(${cols['uploaddate']}{r}=\"\";\"\";YEAR(${cols['uploaddate']}{r}))"]]})

        if profile['has_cek_col']:
            value_ranges.append({'range': f"{sheet_ref}{cols['cek']}{r}", 'values': [[f"=COUNTIF({cols['videoid']}:{cols['videoid']};{cols['videoid']}{r})"]]})

    entries_per_row = len(value_ranges) // len(to_insert) if to_insert else 0
    chunk_rows = 700
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

    profile = get_target_profile()
    source_type = detect_source_type(SOURCE_XLSX_PATH)
    print(f'Target terdeteksi: {profile["name"]}  |  Sumber terdeteksi: {source_type}')

    print('Membaca file sumber...')
    if source_type == 'TAP':
        records, dup_in_source, filtered_out_month = read_tap_sv(SOURCE_XLSX_PATH)
    else:
        records, dup_in_source, filtered_out_month = read_sc_sv(SOURCE_XLSX_PATH)
    print(f'Ketemu {len(records)} video unik ({dup_in_source} duplikat dalam file sumber dibuang).')
    if FILTER_MONTH is not None:
        print(f'Filter bulan aktif ({FILTER_MONTH}/{FILTER_YEAR}): {filtered_out_month} video di luar bulan itu dilewati.')

    service = get_sheets_service()

    roster = None
    if source_type == 'SC':
        if not profile['roster_sheet']:
            print('PERINGATAN: sumber SC tapi target ini gak punya roster_sheet, semua video diambil tanpa filter roster.')
        else:
            roster = get_creator_roster(service, profile)
            print(f'Ada {len(roster)} creator di roster {profile["roster_sheet"]}.')

    print('Ambil Video ID yang udah ada di VIDEO BANK...')
    existing_ids, data_start_sheet_row, next_row = get_existing_video_ids(service, profile)
    print(f'Ada {len(existing_ids)} Video ID yang udah ada di sheet.')

    not_in_roster = 0
    to_insert = []
    for r in records:
        if r['video_id'] in existing_ids:
            continue
        if roster is not None and r['creator'].strip().lower() not in roster:
            not_in_roster += 1
            continue
        to_insert.append(r)

    already_exists = len(records) - len(to_insert) - not_in_roster

    if not to_insert:
        msg = f'\nTidak ada video baru. ({already_exists} sudah ada di sheet'
        if roster is not None:
            msg += f', {not_in_roster} creator-nya bukan roster'
        msg += ')'
        print(msg)
        return

    print(f'\n{len(to_insert)} video baru akan ditambahkan ke baris {next_row}-{next_row + len(to_insert) - 1}.')
    extra = f', {not_in_roster} dilewati karena bukan roster' if roster is not None else ''
    print(f'{already_exists} dilewati karena sudah ada di sheet{extra}.')
    ensure_enough_rows(service, profile, next_row + len(to_insert) - 1)
    print('Menulis ke Google Sheets...')
    write_new_records(service, profile, to_insert, next_row)

    if source_type == 'TAP':
        print('\nSinkronisasi roster creator (khusus sumber TAP)...')
        sync_roster_from_records(
            service, TARGET_SPREADSHEET_ID, profile['name'],
            [(r['creator'], r['post_time']) for r in to_insert]
        )

    print(f'\nSELESAI. {len(to_insert)} video baru ditambahkan.')


if __name__ == '__main__':
    main()
