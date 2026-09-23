"""
IMPORT TAP_GMV -> RAW TAP / TAP Cussons  [UNIFIED: Ellips & Cussons]
================================================================================
Satu script, brand dideteksi otomatis dari NAMA FILE SOURCE_XLSX_PATH (harus
mengandung kata "ellips" atau "cussons", case-insensitive). Struktur kolom &
formula tujuan beda antar brand -- semuanya di-drive dari PROFILES di bawah.

- ELLIPS -> tab "RAW TAP", 35 kolom + field komisi lengkap. TIDAK auto-nambah
  baris kalau kurang (baris di sheet kalau kurang harus ditambah manual dulu).
- CUSSONS -> tab "TAP Cussons (1-30 September)", 25 kolom, formula beda
  (GMV SL, Avg item Sold, Item Sold LS/SV/SL).
  OTOMATIS nambah baris kalau grid-nya kurang.

FILTER GMV: keduanya cuma ambil baris yang Affiliate/Creator-attributed GMV-nya
!= 0 (laporan TAP_GMV itu cross-join, mayoritas barisnya GMV nol).

DEDUP (kedua brand sama pola): kombinasi Date + Creator name + Product ID yang
PERSIS sama dengan yang udah ada di sheet akan dilewati.

CARA PAKAI:
1. Isi SOURCE_XLSX_PATH -- nama filenya HARUS ada kata "ellips" atau "cussons".
2. python import_raw_tap_unified.py
"""

import datetime
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ============== KONFIGURASI -- ISI INI ==============
CREDENTIALS_FILE = 'credentials.json'
SOURCE_XLSX_PATH = r'C:\Users\Subhan\OneDrive\Documents\Automasi\TAP_GMV Cussons (24-30 August).xlsx'
SOURCE_SHEET_NAME = 'Custom report'  # cek dulu, kadang namanya "Sheet1"
# ======================================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def col_map_ellips():
    return [
        ('A', ['Date']), ('C', ['Comparison date']), ('D', ['Campaign ID']),
        ('E', ['Campaign name']), ('F', ['Campaign duration']), ('G', ['Creator name']),
        ('J', ['Product ID']), ('K', ['Product name']), ('L', ['Shop ID']), ('M', ['Shop name']),
        ('N', ['Level 1 category']), ('O', ['Level 2 category']),
        ('P', ['Creator-attributed GMV', 'Affiliate GMV']),
        ('Q', ['Affiliate video-attributed GMV', 'Affiliate video GMV']),
        ('R', ['Creator LIVE-attributed GMV', 'Affiliate LIVE GMV']),
        ('T', ['Creator-attributed orders', 'Orders']),
        ('U', ['Creator LIVE-attributed orders', 'LIVE orders']),
        ('V', ['Creator video-attributed orders', 'Video orders']),
        ('W', ['LIVE likes']), ('X', ['Video likes']), ('Y', ['Video views']),
        ('Z', ['LIVE views']), ('AA', ['LIVE streams']), ('AB', ['Videos']),
        ('AC', ['Products added to Showcase']),
        ('AD', ['Estimated affiliate partner commission']),
        ('AE', ['Actual affiliate partner commission']),
        ('AF', ['Estimated creator commission']), ('AG', ['Actual creator commission']),
        ('AH', ['GMV (refund)']), ('AI', ['Settled GMV']), ('AJ', ['Revenue (Showcase)']),
        ('AK', ['Creator-attributed items sold', 'Items sold']), ('AL', ['Link GMV']),
        ('AM', ['Link items sold']), ('AN', ['Link orders']),
        ('AO', ['Link partner est commission', 'Link partner est. commission']),
        ('AP', ['Link creator est commission', 'Link creator est. commission']),
    ]


def col_map_cussons():
    return [
        ('A', ['Date']), ('C', ['Campaign ID']), ('D', ['Campaign name']),
        ('E', ['Campaign duration']), ('F', ['Creator name']), ('H', ['Product ID']),
        ('I', ['Product name']), ('J', ['Creator-attributed GMV', 'Affiliate GMV']),
        ('K', ['Affiliate video-attributed GMV', 'Affiliate video GMV']),
        ('L', ['Creator LIVE-attributed GMV', 'Affiliate LIVE GMV']),
        ('N', ['Creator-attributed orders', 'Orders']),
        ('O', ['Creator LIVE-attributed orders', 'LIVE orders']),
        ('P', ['Creator video-attributed orders', 'Video orders']),
        ('Q', ['LIVE likes']), ('R', ['Video likes']), ('S', ['Video views']),
        ('T', ['LIVE views']), ('U', ['LIVE streams']), ('V', ['Videos']),
        ('W', ['GMV (refund)']), ('X', ['Settled GMV']), ('Y', ['Revenue (Showcase)']),
        ('Z', ['Creator-attributed items sold', 'Items sold']), ('AA', ['Link GMV']),
        ('AB', ['Link items sold']), ('AC', ['Link orders']),
    ]


def write_formulas_ellips(sheet_ref, r, roster_sheet):
    # Week Ellips: W1 = 1-2, W2 = 3-9, W3 = 10-16, W4 = 17-23, W5 = 24-akhir bulan
    d = f'DAY(DATEVALUE(A{r}))'
    roster_sheet_escaped = roster_sheet.replace("'", "''")
    return [
        {'range': f'{sheet_ref}B{r}', 'values': [[
            f'=UPPER(TEXT(DATEVALUE(A{r});"mmm"))&" W"&IF({d}<=2;1;IF({d}<=9;2;IF({d}<=16;3;'
            f'IF({d}<=23;4;5))))&" ("&IF({d}<=2;"1-2";IF({d}<=9;"3-9";IF({d}<=16;"10-16";'
            f'IF({d}<=23;"17-23";"24-"&DAY(EOMONTH(DATEVALUE(A{r});0))))))&")"'
        ]]},
        # Cek USN Ext (H) -> roster GMV Creator bulan baris itu. Kolom I (cek USN
        # GMV June) sengaja dikosongin.
        {'range': f'{sheet_ref}H{r}', 'values': [[f"=VLOOKUP(G{r};'{roster_sheet_escaped}'!$B:$B;1;0)"]]},
        {'range': f'{sheet_ref}S{r}', 'values': [[f'=AJ{r}+AL{r}']]},
    ]


def write_formulas_cussons(sheet_ref, r, roster_sheet):
    roster_sheet_escaped = roster_sheet.replace("'", "''")
    return [
        {'range': f'{sheet_ref}B{r}', 'values': [[
            f'=UPPER(TEXT(A{r};"MMM"))&" W"&IF(DAY(A{r})<=6;1;IF(DAY(A{r})<=13;2;IF(DAY(A{r})<=20;3;'
            f'IF(DAY(A{r})<=27;4;5))))&" ("&IF(DAY(A{r})<=6;1;IF(DAY(A{r})<=13;7;IF(DAY(A{r})<=20;14;'
            f'IF(DAY(A{r})<=27;21;28))))&"-"&IF(DAY(A{r})<=6;6;IF(DAY(A{r})<=13;13;IF(DAY(A{r})<=20;20;'
            f'IF(DAY(A{r})<=27;27;DAY(EOMONTH(A{r};0))))))&")"'
        ]]},
        {'range': f'{sheet_ref}G{r}', 'values': [[f"=VLOOKUP(F{r};'{roster_sheet_escaped}'!$B:$B;1;0)"]]},
        {'range': f'{sheet_ref}M{r}', 'values': [[f'=AA{r}+Y{r}']]},
        {'range': f'{sheet_ref}AD{r}', 'values': [[f'=Z{r}/N{r}']]},
        {'range': f'{sheet_ref}AE{r}', 'values': [[f'=AD{r}*O{r}']]},
        {'range': f'{sheet_ref}AF{r}', 'values': [[f'=AD{r}*P{r}']]},
        {'range': f'{sheet_ref}AG{r}', 'values': [[f'=Z{r}-SUM(AE{r};AF{r})']]},
    ]


PROFILES = {
    'ELLIPS': {
        'spreadsheet_id': '1tIG9FhUogXwBJK6YuzpT19nFlJs5EDfXA6EYQ493paE',
        'sheet': 'RAW TAP  (20 April-31 August)',
        'column_map': col_map_ellips(),
        'gmv_col': 'P',
        'dedup_cols': {'date': 'A', 'creator': 'G', 'product': 'J'},
        'auto_add_rows': False,
        'extra_fields_per_row': 3,  # B, H, S
        'write_formulas': write_formulas_ellips,
        'roster_sheet': None,  # None = ikut bulan tiap baris ('GMV Creator [SEPT]' dst)
    },
    'CUSSONS': {
        'spreadsheet_id': '1ZBOvn5fReBECSzgXrAS6SNuq7tWDaem9Cf_QhQN4b_8',
        'sheet': 'TAP Cussons (1-30 September)',
        'column_map': col_map_cussons(),
        'gmv_col': 'J',
        'dedup_cols': {'date': 'A', 'creator': 'F', 'product': 'H'},
        'auto_add_rows': True,
        'extra_fields_per_row': 7,  # B, G, M, AD, AE, AF, AG
        'write_formulas': write_formulas_cussons,
        'roster_sheet': "Creator Performance PZ Cussons September'26",
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
    ELLIPS: 'GMV Creator [SEPT]' / CUSSONS: "Creator Performance PZ Cussons September'26"."""
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


XML_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XML_RIDNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
XML_RELNS = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def find_sheet_xml_target(zf, sheet_name):
    """Cari path XML internal (misal xl/worksheets/sheet2.xml) buat sheet_name tertentu."""
    wb_xml = ET.fromstring(zf.read('xl/workbook.xml'))
    rels_xml = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
    rid_to_target = {rel.get('Id'): rel.get('Target') for rel in rels_xml.findall(f'{XML_RELNS}Relationship')}
    for sheet in wb_xml.find(f'{XML_NS}sheets'):
        if sheet.get('name') == sheet_name:
            target = rid_to_target[sheet.get(f'{XML_RIDNS}id')]
            return target[1:] if target.startswith('/') else 'xl/' + target
    raise ValueError(f'Sheet "{sheet_name}" tidak ketemu.')


def load_shared_strings(zf):
    """Load xl/sharedStrings.xml (kalau ada) -- dibutuhin buat file yang cell teksnya
    disimpan sebagai referensi index (t="s") bukan inline (t="inlineStr"). File yang
    di-resave/diedit manual di Excel/Sheets biasanya pakai shared strings, sementara
    export langsung dari TikTok biasanya inlineStr."""
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    strings = []
    for si in root:
        text = ''.join(t.text or '' for t in si.iter(f'{XML_NS}t'))
        strings.append(text)
    return strings


def iter_sheet_rows_streaming(xlsx_path, sheet_name):
    """Baca baris sheet langsung dari XML internal .xlsx pakai iterparse (hemat memori).

    File TAP_GMV bisa ratusan MB (cross-join, banyak baris) dan dimension tag-nya
    sering rusak (cuma "A1") sehingga openpyxl read_only salah baca kolom, sementara
    openpyxl full-parse (read_only=False) kelewat lambat/berat buat file segede ini.
    Baca XML langsung, baris per baris, jauh lebih cepat & ringan.
    """
    zf = zipfile.ZipFile(xlsx_path)
    target = find_sheet_xml_target(zf, sheet_name)
    shared_strings = load_shared_strings(zf)
    with zf.open(target) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag.rsplit('}', 1)[-1] != 'row':
                continue
            cells = {}
            for c in elem:
                if c.tag.rsplit('}', 1)[-1] != 'c':
                    continue
                ref = c.get('r')
                if not ref:
                    continue
                letter = re.match(r'[A-Z]+', ref).group(0)
                cell_type = c.get('t')
                if cell_type == 'inlineStr':
                    is_el = c.find(f'{XML_NS}is')
                    t_el = is_el.find(f'{XML_NS}t') if is_el is not None else None
                    val = t_el.text if t_el is not None else ''
                elif cell_type == 's':
                    v_el = c.find(f'{XML_NS}v')
                    idx = int(v_el.text) if v_el is not None and v_el.text is not None else None
                    val = shared_strings[idx] if idx is not None and idx < len(shared_strings) else ''
                else:
                    v_el = c.find(f'{XML_NS}v')
                    val = v_el.text if v_el is not None else None
                cells[letter] = val
            yield cells
            elem.clear()


def parse_rupiah(val):
    """Parse angka format 'Rp1.234.567' / 0 / None jadi float."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('Rp', '').replace('.', '').strip()
    if s in ('', '-', '--'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def date_to_text(dt):
    """Format tanggal jadi teks yyyy-mm-dd (dibutuhin formula Week yang pakai DATEVALUE)."""
    if dt is None:
        return ''
    if isinstance(dt, (datetime.datetime, datetime.date)):
        return dt.strftime('%Y-%m-%d')
    return str(dt)


def read_source_file(path, column_map, gmv_col):
    row_iter = iter_sheet_rows_streaming(path, SOURCE_SHEET_NAME)
    try:
        header_row = next(row_iter)
    except StopIteration:
        raise ValueError('Sheet sumber kosong.')

    name_to_letter = {}
    for letter, val in header_row.items():
        name = (val or '').strip()
        if name and name not in name_to_letter:
            name_to_letter[name] = letter

    def resolve(names):
        for n in names:
            if n in name_to_letter:
                return name_to_letter[n]
        return None

    src_col_index = {}
    missing = []
    for dest_col, names in column_map:
        letter = resolve(names)
        if letter is None:
            missing.append(' / '.join(names))
        else:
            src_col_index[dest_col] = letter
    if missing:
        raise ValueError(f'Kolom berikut tidak ketemu di header: {", ".join(missing)}')

    idx_date = resolve(['Date'])
    idx_creator = resolve(['Creator name'])
    idx_product = resolve(['Product ID'])
    idx_gmv = src_col_index[gmv_col]

    total_data_rows = 0
    before_gmv_filter = 0
    valid_rows = []
    for row in row_iter:
        total_data_rows += 1
        date_val = row.get(idx_date)
        if not date_val or date_val == 'Summary':
            continue
        before_gmv_filter += 1
        if parse_rupiah(row.get(idx_gmv)) == 0:
            continue
        valid_rows.append(row)

    if total_data_rows < 4:
        print(f'PERINGATAN: sheet sumber cuma ada {total_data_rows} baris data. '
              f'Kalau harusnya data 1 periode penuh, kemungkinan datanya kurang lengkap.')

    print(f'Filter GMV != 0: {before_gmv_filter - len(valid_rows)} baris GMV-nya 0 dibuang, '
          f'sisa {len(valid_rows)} baris.')

    return valid_rows, src_col_index, idx_date, idx_creator, idx_product


def get_sheet_id_and_row_count(service, spreadsheet_id, sheet_name):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields='sheets(properties(sheetId,title,gridProperties))'
    ).execute()
    for sh in meta['sheets']:
        if sh['properties']['title'] == sheet_name:
            return sh['properties']['sheetId'], sh['properties']['gridProperties']['rowCount']
    raise ValueError(f'Sheet "{sheet_name}" tidak ketemu.')


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


def get_existing_keys(service, spreadsheet_id, sheet_name, dedup_cols):
    """Baca kolom Date/Creator/Product ID (posisi beda per brand), chunk per 30rb baris."""
    _, target_max_row = get_sheet_id_and_row_count(service, spreadsheet_id, sheet_name)
    date_col, creator_col, product_col = dedup_cols['date'], dedup_cols['creator'], dedup_cols['product']

    chunk_size = 30000
    col_date, col_creator, col_product = [], [], []
    for start in range(1, target_max_row + 1, chunk_size):
        end = min(start + chunk_size - 1, target_max_row)
        resp = service.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=[f"'{sheet_name}'!{date_col}{start}:{date_col}{end}",
                    f"'{sheet_name}'!{creator_col}{start}:{creator_col}{end}",
                    f"'{sheet_name}'!{product_col}{start}:{product_col}{end}"]
        ).execute()
        vr = resp.get('valueRanges', [])
        d_vals = vr[0].get('values', []) if len(vr) > 0 else []
        c_vals = vr[1].get('values', []) if len(vr) > 1 else []
        p_vals = vr[2].get('values', []) if len(vr) > 2 else []
        col_date.extend(d_vals)
        col_creator.extend(c_vals)
        col_product.extend(p_vals)
        if not d_vals and not c_vals and not p_vals:
            break

    header_row = None
    for i, row in enumerate(col_date):
        if row and row[0] == 'Date':
            header_row = i
            break
    if header_row is None:
        raise ValueError(f'Header "Date" tidak ketemu di kolom {date_col} sheet {sheet_name}.')

    data_start = header_row + 1
    existing_keys = set()
    last_filled = data_start - 1
    last_row = max(len(col_date), len(col_creator), len(col_product))
    for i in range(data_start, last_row):
        d = col_date[i][0] if i < len(col_date) and col_date[i] else ''
        c = col_creator[i][0] if i < len(col_creator) and col_creator[i] else ''
        p = col_product[i][0] if i < len(col_product) and col_product[i] else ''
        if d or c or p:
            last_filled = i
        if d and c:
            existing_keys.add(f'{str(d).strip()}|{str(c).strip()}|{str(p).strip()}')

    return existing_keys, data_start + 1, last_filled + 2, target_max_row


def force_text_if_long_id(val):
    """ID kayak Campaign ID/Product ID/Shop ID itu 16+ digit -- kalau dikirim apa
    adanya ke Sheets pakai USER_ENTERED, Sheets nganggep itu ANGKA dan digit
    belakangnya kepotong (presisi float cuma akurat ~15-16 digit). Kasih awalan
    petik satu (') biar Sheets simpen persis sebagai teks."""
    s = str(val) if val is not None else ''
    if s.isdigit() and len(s) >= 16:
        return "'" + s
    return val


def write_new_rows(service, brand, profile, to_insert, src_col_index, next_row):
    spreadsheet_id = profile['spreadsheet_id']
    sheet_ref = f"'{profile['sheet']}'!"
    column_map = profile['column_map']
    value_ranges = []

    for idx, row in enumerate(to_insert):
        r = next_row + idx
        for dest_col, _ in column_map:
            val = row.get(src_col_index[dest_col])
            if dest_col == 'A':
                val = date_to_text(val)
            else:
                val = force_text_if_long_id(val)
            value_ranges.append({'range': f'{sheet_ref}{dest_col}{r}', 'values': [[val if val is not None else '']]})

        roster_sheet = profile['roster_sheet']
        if roster_sheet is None:
            ym = extract_year_month(row.get(src_col_index['A']))
            if ym is None:
                today = datetime.date.today()
                ym = (today.year, today.month)
            roster_sheet = roster_sheet_name(brand, *ym)
        value_ranges.extend(profile['write_formulas'](sheet_ref, r, roster_sheet))

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
    valid_rows, src_col_index, idx_date, idx_creator, idx_product = read_source_file(
        SOURCE_XLSX_PATH, profile['column_map'], profile['gmv_col']
    )
    print(f'Ketemu {len(valid_rows)} baris data valid.')
    if not valid_rows:
        print('Tidak ada baris data valid, berhenti.')
        return

    print('Connect ke Google Sheets, ambil data existing buat dedup...')
    service = get_sheets_service()
    existing_keys, data_start_sheet_row, next_row, target_max_row = get_existing_keys(
        service, profile['spreadsheet_id'], profile['sheet'], profile['dedup_cols']
    )
    print(f'Ada {len(existing_keys)} kombinasi Date+Creator+Product ID yang udah ada di sheet.')

    to_insert = []
    skipped = 0
    seen_now = set(existing_keys)
    for row in valid_rows:
        key = f'{str(row.get(idx_date, "")).strip()}|{str(row.get(idx_creator, "")).strip()}|{str(row.get(idx_product, "")).strip()}'
        if key in seen_now:
            skipped += 1
            continue
        seen_now.add(key)
        to_insert.append(row)

    if not to_insert:
        print(f'\nSemua {len(valid_rows)} baris sudah ada sebelumnya (kombinasi Date+Creator+Product ID sama).')
        return

    needed_last_row = next_row + len(to_insert) - 1
    if needed_last_row > target_max_row:
        rows_short = needed_last_row - target_max_row
        if profile['auto_add_rows']:
            print(f'\nGrid sheet {profile["sheet"]} kurang {rows_short} baris (butuh sampai baris {needed_last_row}, '
                  f'tersedia sampai {target_max_row}). Nambah otomatis...')
            target_max_row = ensure_rows(service, profile['spreadsheet_id'], profile['sheet'], rows_short)
            print(f'Baris ditambah, grid sekarang sampai baris {target_max_row}.')
        else:
            print(f'\nERROR: Baris di sheet {profile["sheet"]} gak cukup. Butuh sampai baris {needed_last_row}, '
                  f'tapi grid sheet cuma sampai baris {target_max_row} (kurang {rows_short} baris).')
            print(f'Tambahin manual dulu minimal {rows_short} baris di bagian bawah sheet '
                  f'(klik kanan nomor baris terakhir > Insert {rows_short} rows below), baru jalanin lagi script ini.')
            return

    print(f'\n{len(to_insert)} baris baru akan ditambahkan ke baris {next_row}-{next_row + len(to_insert) - 1}.')
    print(f'{skipped} dilewati (kombinasi Date+Creator+Product ID sudah ada).')
    print('Menulis ke Google Sheets...')
    write_new_rows(service, brand, profile, to_insert, src_col_index, next_row)

    print('\nSinkronisasi roster creator (khusus sumber TAP)...')
    sync_roster_from_records(
        service, profile['spreadsheet_id'], brand,
        [(row.get(idx_creator, ''), row.get(idx_date)) for row in to_insert]
    )

    print(f'\nSELESAI. {len(to_insert)} baris baru ditambahkan, {skipped} dilewati.')


if __name__ == '__main__':
    main()
