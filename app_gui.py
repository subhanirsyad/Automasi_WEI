"""
GUI - Automasi Import Data
================================================================================
Antarmuka desktop untuk import Ellips/Cussons dan sync TAP LS/SV/Content PC/BW.
Mode cloud menerima link sumber/tujuan, nama tab, tahun, bulan, dan (untuk
Content BW) week; preview bersifat baca-saja, sedangkan impor hanya berjalan
setelah konfirmasi pengguna.

CARA PAKAI:
    python app_gui.py
"""

import importlib
import io
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_PATH = BASE_DIR / 'credentials.json'

# ---------------------------------------------------------------- palette ---
COLOR_BG = ('#eef0f4', '#15161a')
COLOR_SIDEBAR = ('#e2e5eb', '#0f1013')
COLOR_CARD = ('#ffffff', '#1d1f24')
COLOR_ACCENT = ('#2b6ef2', '#4d8dff')
COLOR_ACCENT_HOVER = ('#1f57c9', '#3f78dd')
COLOR_TEXT_MUTED = ('#6b7280', '#9199a6')
COLOR_SUCCESS = ('#1f8a4c', '#4fd382')
COLOR_ERROR = ('#d64545', '#ff6b6b')
COLOR_LOG_BG = ('#12141a', '#0b0c10')
COLOR_NAV_SELECTED = ('#d7e2fb', '#22314f')

SPREADSHEET_IDS = {
    'ELLIPS': '1tIG9FhUogXwBJK6YuzpT19nFlJs5EDfXA6EYQ493paE',
    'CUSSONS': '1ZBOvn5fReBECSzgXrAS6SNuq7tWDaem9Cf_QhQN4b_8',
}

FILE_IMPORT_TYPES = {
    'Raw TAP': {
        'icon': '📊',
        'subtitle': 'dari file TAP_GMV',
        'module': 'import_raw_tap_unified',
        'filetypes': [('Excel files', '*.xlsx')],
        'auto_brand': True,
        'has_month_filter': False,
        'file_hint': ['tap_gmv'],
    },
    'Raw Data SC': {
        'icon': '🛒',
        'subtitle': 'dari file SC_GMV',
        'module': 'import_raw_data_sc_unified',
        'filetypes': [('Excel/CSV files', '*.xlsx *.csv'), ('Semua file', '*.*')],
        'auto_brand': True,
        'has_month_filter': True,
        'file_hint': ['sc_gmv'],
    },
    'Live Streaming': {
        'icon': '🔴',
        'subtitle': 'dari file TAP_LS',
        'module': 'import_live_streaming_unified',
        'filetypes': [('Excel files', '*.xlsx')],
        'auto_brand': False,
        'has_month_filter': True,
        'file_hint': ['tap_ls'],
    },
    'Video Bank': {
        'icon': '🎬',
        'subtitle': 'dari file TAP_SV / SC_SV',
        'module': 'import_video_bank_unified',
        'filetypes': [('Excel files', '*.xlsx')],
        'auto_brand': False,
        'has_month_filter': True,
        'file_hint': ['tap_sv', 'sc_sv'],
    },
}

IMPORT_TYPES = {
    'TAP LS PC (Drive)': {
        'group': 'PC',
        'label': 'TAP LS',
        'icon': '☁',
        'subtitle': 'Raw Excel di Drive → tab LS per brand (GMV / Date Live)',
        'module': 'sync_tap_ls_pc',
        'script': 'sync_tap_ls_pc.py',
        'default_source': 'https://docs.google.com/spreadsheets/d/1uM0P_7f0ybvXhW2iSVtREeqNbXtHVI-Y/edit',
        'default_target': 'https://docs.google.com/spreadsheets/d/1TjViP0sreDwohIhnsSM2heKPdlZJpyyLyTW4iLf6PTI/edit',
        'default_source_tab': 'Custom report',
        'default_target_tab': 'LS 1-30',
        'date_column': 'Date Live',
        'cloud_source': True,
    },
    'TAP SV PC (Drive)': {
        'group': 'PC',
        'label': 'TAP SV',
        'icon': '☁',
        'subtitle': 'Raw Excel di Drive → tab SV per brand (GMV / Date Post)',
        'module': 'sync_tap_sv_pc',
        'script': 'sync_tap_sv_pc.py',
        'default_source': 'https://docs.google.com/spreadsheets/d/1k8zp4wpa3xcArnwGGLa-icNKCmau84SI/edit',
        'default_target': 'https://docs.google.com/spreadsheets/d/1TjViP0sreDwohIhnsSM2heKPdlZJpyyLyTW4iLf6PTI/edit',
        'default_source_tab': 'Custom report',
        'default_target_tab': 'SV 1-30',
        'date_column': 'Date Post',
        'cloud_source': True,
    },
    'Content LS PC (Drive)': {
        'group': 'PC',
        'label': 'Content LS',
        'icon': '📺',
        'subtitle': 'LS 1-30 → tab content per brand dan PAID, room ID unik',
        'module': 'sync_pc_content',
        'script': 'sync_pc_content.py',
        'default_source': 'https://docs.google.com/spreadsheets/d/1TjViP0sreDwohIhnsSM2heKPdlZJpyyLyTW4iLf6PTI/edit',
        'default_target': 'https://docs.google.com/spreadsheets/d/1Xh7GbrH7w9EHBuBBArxvoaUA9BNBoyzOezYSh8OqbYs/edit',
        'default_source_tab': 'LS 1-30',
        'default_target_tab': 'Creator PAID',
        'date_column': 'Date Live',
        'cloud_source': True,
        'content_source': True,
        'repair_paid_routing': True,
    },
    'TAP LS BW (Sheets)': {
        'group': 'BW',
        'label': 'Raw TAP LS',
        'icon': 'BW',
        'subtitle': 'Custom report -> raw LS per brand (semua GMV / non-GMV bulan terpilih)',
        'module': 'sync_tap_ls_bw',
        'script': 'sync_tap_ls_bw.py',
        'default_source': 'https://docs.google.com/spreadsheets/d/1fsGXggZKQEozmoafhMONp_ulHd-PQbhs63e4YkiPUS8/edit',
        'default_target': 'https://docs.google.com/spreadsheets/d/17B536kbB0bXrZ1xLoLwCNyPpZQ83VEBTOjO2Inw7ipE/edit',
        'default_source_tab': 'Custom report',
        'default_target_tab': 'LS',
        'date_column': 'Date Live',
        'cloud_source': True,
    },
    'Content LS BW (Sheets)': {
        'group': 'BW',
        'label': 'Content LS',
        'icon': 'BW',
        'subtitle': 'Raw LS -> 7 tab Content brand, filter bulan + Week, tanpa split PAID',
        'module': 'sync_bw_content',
        'script': 'sync_bw_content.py',
        'default_source': 'https://docs.google.com/spreadsheets/d/17B536kbB0bXrZ1xLoLwCNyPpZQ83VEBTOjO2Inw7ipE/edit',
        'default_target': 'https://docs.google.com/spreadsheets/d/1XLtBKqCtJPWzOf0GvxGVRyoaa1xpHo6U8qTK6Ch7xMI/edit',
        'default_source_tab': 'LS',
        'default_target_tab': '7 tab brand otomatis',
        'default_week': 'Week 2',
        'date_column': 'Date Live',
        'cloud_source': True,
        'content_source': True,
        'bw_content': True,
    },
}

for brand in ('ELLIPS', 'CUSSONS'):
    for label, template in FILE_IMPORT_TYPES.items():
        IMPORT_TYPES[f'{brand} / {label}'] = {
            **template, 'group': brand, 'brand': brand, 'label': label,
        }

GROUP_LABELS = {'PC': 'PC', 'BW': 'BW', 'ELLIPS': 'Ellips', 'CUSSONS': 'Cussons'}

MONTHS = ['Semua bulan', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
          'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']


def detect_brand_from_filename(name):
    name = name.lower()
    if 'ellips' in name:
        return 'ELLIPS'
    if 'cussons' in name:
        return 'CUSSONS'
    return None


class QueueWriter(io.TextIOBase):
    """Nampung print() dari script import, nanti dibaca GUI thread lewat queue."""

    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(s)
        return len(s)

    def flush(self):
        pass


def card(parent, **kwargs):
    defaults = dict(fg_color=COLOR_CARD, corner_radius=14, border_width=1,
                     border_color=('#e1e4ea', '#2a2d34'))
    defaults.update(kwargs)
    return ctk.CTkFrame(parent, **defaults)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode('System')
        ctk.set_default_color_theme('blue')

        self.title('Automasi Import Data')
        self.geometry('1040x760')
        self.minsize(880, 620)
        self.configure(fg_color=COLOR_BG)

        self.font_title = ctk.CTkFont(family='Segoe UI', size=20, weight='bold')
        self.font_h2 = ctk.CTkFont(family='Segoe UI', size=15, weight='bold')
        self.font_body = ctk.CTkFont(family='Segoe UI', size=13)
        self.font_small = ctk.CTkFont(family='Segoe UI', size=11)
        self.font_nav = ctk.CTkFont(family='Segoe UI', size=13, weight='bold')
        self.font_mono = ctk.CTkFont(family='Consolas', size=12)

        self.log_queue = queue.Queue()
        self.running = False
        self.nav_buttons = {}
        self.nav_groups = {}
        self.group_buttons = {}
        self._active_group = None
        self._cloud_forms = {}
        self._active_cloud_type = None
        self._file_forms = {}
        self._active_file_type = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()

        self.type_var = ctk.StringVar(value=list(IMPORT_TYPES.keys())[0])
        self._select_type(self.type_var.get())
        self.after(100, self._poll_log_queue)

    # ------------------------------------------------------------ sidebar --
    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=COLOR_SIDEBAR)
        sidebar.grid(row=0, column=0, sticky='nsw')
        sidebar.grid_propagate(False)

        header = ctk.CTkFrame(sidebar, fg_color='transparent')
        header.pack(fill='x', padx=22, pady=(28, 22))
        ctk.CTkLabel(header, text='⚡ Automasi', font=self.font_title, anchor='w').pack(fill='x')
        ctk.CTkLabel(header, text='PC, Ellips & Cussons', font=self.font_small,
                     text_color=COLOR_TEXT_MUTED, anchor='w').pack(fill='x', pady=(2, 0))

        ctk.CTkFrame(sidebar, height=1, fg_color=('#d3d7de', '#2a2d34')).pack(fill='x', padx=22, pady=(0, 14))

        ctk.CTkLabel(sidebar, text='BRAND / CLIENT', font=self.font_small,
                     text_color=COLOR_TEXT_MUTED, anchor='w').pack(fill='x', padx=22, pady=(0, 6))

        nav_frame = ctk.CTkScrollableFrame(sidebar, fg_color='transparent')
        nav_frame.pack(fill='both', expand=True, padx=(10, 5))

        for group, group_label in GROUP_LABELS.items():
            group_button = ctk.CTkButton(
                nav_frame, text=f'  >  {group_label}', anchor='w', font=self.font_nav,
                height=40, corner_radius=8, fg_color='transparent',
                hover_color=('#dbe1eb', '#20232a'),
                text_color=('#1b1d22', '#e7e9ee'),
                command=lambda g=group: self._toggle_group(g),
            )
            group_button.pack(fill='x', pady=(6, 2))
            self.group_buttons[group] = group_button

            submenu = ctk.CTkFrame(nav_frame, fg_color='transparent')
            submenu.pack(fill='x')
            self.nav_groups[group] = submenu
            for name, cfg in IMPORT_TYPES.items():
                if cfg['group'] != group:
                    continue
                btn = ctk.CTkButton(
                    submenu, text=f"   {cfg['icon']}  {cfg['label']}", anchor='w',
                    font=self.font_body, height=38, corner_radius=8,
                    fg_color='transparent', hover_color=('#dbe1eb', '#20232a'),
                    text_color=('#1b1d22', '#e7e9ee'),
                    command=lambda n=name: self._select_type(n),
                )
                btn.pack(fill='x', pady=2, padx=(12, 0))
                self.nav_buttons[name] = btn
            submenu.pack_forget()

        footer = ctk.CTkFrame(sidebar, fg_color='transparent')
        footer.pack(fill='x', padx=22, pady=(0, 20))
        cred_ok = CREDENTIALS_PATH.exists()
        dot = '🟢' if cred_ok else '🔴'
        txt = 'credentials.json OK' if cred_ok else 'credentials.json tidak ketemu'
        ctk.CTkLabel(footer, text=f'{dot}  {txt}', font=self.font_small,
                     text_color=COLOR_TEXT_MUTED, anchor='w').pack(fill='x')

    # -------------------------------------------------------------- main ---
    def _build_main(self):
        outer = ctk.CTkScrollableFrame(self, fg_color='transparent')
        outer.grid(row=0, column=1, sticky='nsew', padx=28, pady=24)
        outer.grid_columnconfigure(0, weight=1)
        self.main_scroll = outer

        self.page_title = ctk.CTkLabel(outer, text='', font=self.font_title, anchor='w')
        self.page_title.grid(row=0, column=0, sticky='ew', pady=(0, 2))
        self.page_subtitle = ctk.CTkLabel(outer, text='', font=self.font_body,
                                           text_color=COLOR_TEXT_MUTED, anchor='w')
        self.page_subtitle.grid(row=1, column=0, sticky='ew', pady=(0, 18))

        # ---- card: file sumber ----
        c1 = card(outer)
        c1.grid(row=2, column=0, sticky='ew', pady=(0, 16))
        self.file_card = c1
        c1.grid_columnconfigure(0, weight=1)
        inner1 = ctk.CTkFrame(c1, fg_color='transparent')
        inner1.pack(fill='x', padx=22, pady=18)
        inner1.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(inner1, text='1.  File Sumber', font=self.font_h2, anchor='w').grid(
            row=0, column=0, columnspan=2, sticky='ew', pady=(0, 10))

        self.file_var = ctk.StringVar(value='')
        self.file_entry = ctk.CTkEntry(inner1, textvariable=self.file_var, height=38,
                                        placeholder_text='Belum pilih file...', font=self.font_body)
        self.file_entry.grid(row=1, column=0, sticky='ew', padx=(0, 8))
        self.file_var.trace_add('write', lambda *a: self._update_brand_label())

        ctk.CTkButton(inner1, text='Browse...', width=110, height=38, font=self.font_body,
                      fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER,
                      command=self._browse_file).grid(row=1, column=1)

        self.quickpick_var = ctk.StringVar(value='')
        self.quickpick_menu = ctk.CTkOptionMenu(
            inner1, values=['(cari file relevan di folder ini)'], variable=self.quickpick_var,
            command=self._on_quickpick, height=34, font=self.font_body,
            fg_color=('#eef1f6', '#26292f'), button_color=('#dde2ea', '#32363e'),
            button_hover_color=('#cdd4e0', '#3c4049'), text_color=('#1b1d22', '#e7e9ee'),
            dropdown_font=self.font_body
        )
        self.quickpick_menu.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(10, 0))

        self.brand_label = ctk.CTkLabel(inner1, text='', font=self.font_small, anchor='w',
                                         corner_radius=8, height=26)
        self.brand_label.grid(row=3, column=0, columnspan=2, sticky='w', pady=(12, 0))

        # ---- card: pengaturan ----
        self.c2 = card(outer)
        self.c2.grid(row=3, column=0, sticky='ew', pady=(0, 16))
        self.c2.grid_columnconfigure(0, weight=1)
        inner2 = ctk.CTkFrame(self.c2, fg_color='transparent')
        inner2.pack(fill='x', padx=22, pady=18)
        inner2.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(inner2, text='2.  Pengaturan', font=self.font_h2, anchor='w').grid(
            row=0, column=0, sticky='ew', pady=(0, 10))

        ctk.CTkLabel(inner2, text='Nama Sheet Sumber', font=self.font_small,
                     text_color=COLOR_TEXT_MUTED, anchor='w').grid(row=1, column=0, sticky='w')
        self.sheet_name_var = ctk.StringVar(value='Custom report')
        ctk.CTkEntry(inner2, textvariable=self.sheet_name_var, height=36,
                     font=self.font_body).grid(row=2, column=0, sticky='ew', pady=(4, 12))

        extra_row = ctk.CTkFrame(inner2, fg_color='transparent')
        extra_row.grid(row=3, column=0, sticky='ew')
        extra_row.grid_columnconfigure(0, weight=1)
        extra_row.grid_columnconfigure(1, weight=1)
        self.extra_row = extra_row

        self.month_block = ctk.CTkFrame(extra_row, fg_color='transparent')
        ctk.CTkLabel(self.month_block, text='Filter Bulan (opsional)', font=self.font_small,
                     text_color=COLOR_TEXT_MUTED, anchor='w').pack(fill='x')
        month_row = ctk.CTkFrame(self.month_block, fg_color='transparent')
        month_row.pack(fill='x', pady=(4, 0))
        self.month_var = ctk.StringVar(value=MONTHS[0])
        ctk.CTkOptionMenu(month_row, values=MONTHS, variable=self.month_var, width=150, height=36,
                           font=self.font_body, fg_color=('#eef1f6', '#26292f'),
                           button_color=('#dde2ea', '#32363e'), button_hover_color=('#cdd4e0', '#3c4049'),
                           text_color=('#1b1d22', '#e7e9ee')).pack(side='left')
        self.year_var = ctk.StringVar(value='2026')
        ctk.CTkEntry(month_row, textvariable=self.year_var, width=70, height=36,
                     font=self.font_body).pack(side='left', padx=(8, 0))

        # ---- card: TAP LS PC from Google Drive ----
        self.drive_card = card(outer)
        self.drive_card.grid(row=2, column=0, sticky='ew', pady=(0, 16))
        drive_inner = ctk.CTkFrame(self.drive_card, fg_color='transparent')
        drive_inner.pack(fill='x', padx=22, pady=18)
        drive_inner.grid_columnconfigure(0, weight=1)
        drive_inner.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(drive_inner, text='Sumber dan Tujuan Google Drive',
                     font=self.font_h2, anchor='w').grid(row=0, column=0, columnspan=2,
                                                        sticky='ew', pady=(0, 12))
        self.drive_source_var = ctk.StringVar(value=(
            'https://docs.google.com/spreadsheets/d/'
            '1uM0P_7f0ybvXhW2iSVtREeqNbXtHVI-Y/edit'
        ))
        self.drive_target_var = ctk.StringVar(value=(
            'https://docs.google.com/spreadsheets/d/'
            '1TjViP0sreDwohIhnsSM2heKPdlZJpyyLyTW4iLf6PTI/edit'
        ))
        for row_no, label, variable in (
            (1, 'Link sheet raw (file Excel di Drive)', self.drive_source_var),
            (3, 'Link sheet tujuan', self.drive_target_var),
        ):
            label_widget = ctk.CTkLabel(drive_inner, text=label, font=self.font_small,
                                        text_color=COLOR_TEXT_MUTED, anchor='w')
            label_widget.grid(row=row_no, column=0, columnspan=2, sticky='w')
            if row_no == 1:
                self.drive_source_label = label_widget
            ctk.CTkEntry(drive_inner, textvariable=variable, height=36,
                         font=self.font_body).grid(row=row_no + 1, column=0,
                                                   columnspan=2, sticky='ew', pady=(4, 12))

        self.drive_source_tab_var = ctk.StringVar(value='Custom report')
        self.drive_target_tab_var = ctk.StringVar(value='LS 1-30')
        for col_no, label, variable in (
            (0, 'Tab raw', self.drive_source_tab_var),
            (1, 'Tab tujuan', self.drive_target_tab_var),
        ):
            label_widget = ctk.CTkLabel(drive_inner, text=label, font=self.font_small,
                                        text_color=COLOR_TEXT_MUTED, anchor='w')
            label_widget.grid(row=5, column=col_no, sticky='w',
                              padx=(0, 8) if col_no == 0 else (8, 0))
            if col_no == 0:
                self.drive_source_tab_label = label_widget
            else:
                self.drive_target_tab_label = label_widget
            ctk.CTkEntry(drive_inner, textvariable=variable, height=36,
                         font=self.font_body).grid(row=6, column=col_no, sticky='ew',
                                                   padx=(0, 8) if col_no == 0 else (8, 0),
                                                   pady=(4, 12))

        self.drive_month_var = ctk.StringVar(value='September')
        self.drive_year_var = ctk.StringVar(value='2026')
        self.drive_month_label = ctk.CTkLabel(
            drive_inner, text='Bulan untuk GMV 0 (Date Live)',
            font=self.font_small, text_color=COLOR_TEXT_MUTED, anchor='w'
        )
        self.drive_month_label.grid(row=7, column=0, sticky='w')
        ctk.CTkLabel(drive_inner, text='Tahun', font=self.font_small,
                     text_color=COLOR_TEXT_MUTED, anchor='w').grid(row=7, column=1,
                                                                   sticky='w', padx=(8, 0))
        ctk.CTkOptionMenu(drive_inner, values=MONTHS[1:], variable=self.drive_month_var,
                          height=36, font=self.font_body,
                          fg_color=('#eef1f6', '#26292f'), button_color=('#dde2ea', '#32363e'),
                          button_hover_color=('#cdd4e0', '#3c4049'),
                          text_color=('#1b1d22', '#e7e9ee')).grid(
            row=8, column=0, sticky='ew', padx=(0, 8), pady=(4, 12))
        ctk.CTkEntry(drive_inner, textvariable=self.drive_year_var, height=36,
                     font=self.font_body).grid(row=8, column=1, sticky='ew',
                                               padx=(8, 0), pady=(4, 12))
        self.drive_info_label = ctk.CTkLabel(
            drive_inner,
            text='Cek Data hanya membaca. Impor Baris Baru menulis ke sheet setelah konfirmasi.',
            font=self.font_small, text_color=COLOR_TEXT_MUTED, anchor='w',
            wraplength=650,
        )
        self.drive_info_label.grid(row=9, column=0, columnspan=2, sticky='ew')
        self.drive_card.grid_remove()

        # ---- run button ----
        run_row = ctk.CTkFrame(outer, fg_color='transparent')
        run_row.grid(row=4, column=0, sticky='ew', pady=(0, 16))
        run_row.grid_columnconfigure(0, weight=1)

        self.run_button = ctk.CTkButton(
            run_row, text='▶   Jalankan Import', height=46, font=ctk.CTkFont(family='Segoe UI', size=14, weight='bold'),
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, corner_radius=10,
            command=self._on_run_clicked
        )
        self.run_button.grid(row=0, column=0, sticky='ew')

        self.apply_button = ctk.CTkButton(
            run_row, text='Impor Baris Baru', height=46,
            font=ctk.CTkFont(family='Segoe UI', size=14, weight='bold'),
            fg_color=COLOR_SUCCESS, corner_radius=10,
            command=lambda: self._on_run_clicked(apply=True)
        )
        self.apply_button.grid(row=0, column=1, sticky='ew', padx=(8, 0))
        self.apply_button.grid_remove()

        self.status_label = ctk.CTkLabel(run_row, text='Siap.', font=self.font_body,
                                          text_color=COLOR_TEXT_MUTED, anchor='w')
        self.status_label.grid(row=1, column=0, sticky='w', pady=(8, 0))

        # ---- card: log ----
        c3 = card(outer, fg_color=COLOR_LOG_BG, border_color=('#1c1e24', '#1c1e24'))
        c3.grid(row=5, column=0, sticky='nsew', pady=(0, 4))
        outer.grid_rowconfigure(5, weight=1)
        c3.grid_columnconfigure(0, weight=1)
        c3.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(c3, fg_color='transparent')
        log_header.grid(row=0, column=0, sticky='ew', padx=18, pady=(14, 6))
        log_header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(log_header, text='Log', font=self.font_h2, text_color=('#f2f2f2', '#f2f2f2'),
                     anchor='w').grid(row=0, column=0, sticky='w')
        ctk.CTkButton(log_header, text='Bersihkan', width=90, height=28, font=self.font_small,
                      fg_color='transparent', border_width=1, border_color=('#3a3d45', '#3a3d45'),
                      hover_color=('#20232a', '#20232a'), text_color=('#c7cad2', '#c7cad2'),
                      command=self._clear_log).grid(row=0, column=1, sticky='e')

        self.log_box = ctk.CTkTextbox(c3, font=self.font_mono, fg_color=COLOR_LOG_BG,
                                       text_color='#d8dbe2', wrap='word', height=260)
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=18, pady=(0, 18))
        self.log_box.configure(state='disabled')
        self.log_box.tag_config('error', foreground='#ff6b6b')
        self.log_box.tag_config('success', foreground='#4fd382')
        self.log_box.tag_config('muted', foreground='#7d8290')

    # ----------------------------------------------------------- actions ---
    def _current_config(self):
        return IMPORT_TYPES[self.type_var.get()]

    def _show_group(self, group):
        for name, submenu in self.nav_groups.items():
            if name == group:
                submenu.pack(fill='x', after=self.group_buttons[name])
            else:
                submenu.pack_forget()
            self.group_buttons[name].configure(
                text=f"  {'v' if name == group else '>'}  {GROUP_LABELS[name]}",
                fg_color=COLOR_NAV_SELECTED if name == group else 'transparent',
            )
        self._active_group = group

    def _toggle_group(self, group):
        if self._active_group == group:
            self.nav_groups[group].pack_forget()
            self.group_buttons[group].configure(
                text=f'  >  {GROUP_LABELS[group]}', fg_color='transparent'
            )
            self._active_group = None
        else:
            self._show_group(group)

    def _select_type(self, name):
        if self._active_cloud_type:
            self._cloud_forms[self._active_cloud_type] = {
                'source': self.drive_source_var.get(),
                'target': self.drive_target_var.get(),
                'source_tab': self.drive_source_tab_var.get(),
                'target_tab': self.drive_target_tab_var.get(),
                'year': self.drive_year_var.get(),
                'month': self.drive_month_var.get(),
            }
        if self._active_file_type:
            self._file_forms[self._active_file_type] = {
                'source': self.file_var.get(),
                'source_tab': self.sheet_name_var.get(),
                'year': self.year_var.get(),
                'month': self.month_var.get(),
            }
        self.type_var.set(name)
        cfg = IMPORT_TYPES[name]
        self._show_group(cfg['group'])

        for n, btn in self.nav_buttons.items():
            selected = n == name
            btn.configure(fg_color=COLOR_NAV_SELECTED if selected else 'transparent')

        self.page_title.configure(text=f"{cfg['icon']}  {cfg['label']}")
        self.page_subtitle.configure(text=f"{GROUP_LABELS[cfg['group']]} · {cfg['subtitle']}")

        if cfg.get('cloud_source'):
            self._active_file_type = None
            values = self._cloud_forms.get(name, {
                'source': cfg['default_source'],
                'target': cfg['default_target'],
                'source_tab': cfg['default_source_tab'],
                'target_tab': cfg['default_target_tab'],
                'year': '2026',
                'month': 'September',
            })
            self.drive_source_var.set(values['source'])
            self.drive_target_var.set(values['target'])
            self.drive_source_tab_var.set(values['source_tab'])
            self.drive_target_tab_var.set(values['target_tab'])
            self.drive_year_var.set(values['year'])
            self.drive_month_var.set(values['month'])
            content_mode = cfg.get('content_source', False)
            self.drive_source_label.configure(
                text='Link sheet sumber LS' if content_mode
                else 'Link sheet raw (file Excel di Drive)'
            )
            self.drive_source_tab_label.configure(
                text='Tab LS sumber' if content_mode else 'Tab raw'
            )
            self.drive_target_tab_label.configure(
                text='Tab daftar PAID' if content_mode else 'Tab tujuan'
            )
            self.drive_month_label.configure(
                text=(f"Bulan {cfg['date_column']}" if content_mode
                      else f"Bulan untuk GMV 0 ({cfg['date_column']})")
            )
            self.drive_info_label.configure(
                text=('Cek Data hanya membaca. Sinkronkan Content dapat memperbarui GMV '
                      'serta mengosongkan room ID ganda/salah tab PAID setelah konfirmasi.'
                      if content_mode else
                      'Cek Data hanya membaca. Impor Baris Baru menulis ke sheet setelah konfirmasi.')
            )
            self._active_cloud_type = name
            self.file_card.grid_remove()
            self.c2.grid_remove()
            self.drive_card.grid()
            self.apply_button.grid()
            self.apply_button.configure(
                text='Sinkronkan Content' if content_mode else 'Impor Baris Baru'
            )
            self.run_button.configure(text='Cek Data (baca saja)')
            return

        self._active_cloud_type = None
        values = self._file_forms.get(name, {
            'source': '', 'source_tab': 'Custom report',
            'year': '2026', 'month': MONTHS[0],
        })
        self.file_var.set(values['source'])
        self.sheet_name_var.set(values['source_tab'])
        self.year_var.set(values['year'])
        self.month_var.set(values['month'])
        self._active_file_type = name
        self.drive_card.grid_remove()
        self.apply_button.grid_remove()
        self.file_card.grid()
        self.run_button.configure(text='▶   Jalankan Import')

        self.c2.grid()
        if cfg['has_month_filter']:
            self.month_block.grid(row=0, column=0, sticky='ew')
        else:
            self.month_block.grid_forget()

        self._refresh_quickpick()
        self._update_brand_label()

    def _refresh_quickpick(self):
        cfg = self._current_config()
        hints = cfg['file_hint']
        matches = []
        for p in BASE_DIR.iterdir():
            if not p.is_file():
                continue
            if p.suffix.lower() not in ('.xlsx', '.csv'):
                continue
            low = p.name.lower()
            if (cfg['brand'].lower() in low
                    and any(h in low.replace(' ', '_') or h.replace('_', ' ') in low for h in hints)):
                matches.append(p.name)
        matches.sort()
        if matches:
            self.quickpick_menu.configure(values=['Pilih dari file di folder ini...'] + matches)
            self.quickpick_var.set('Pilih dari file di folder ini...')
        else:
            self.quickpick_menu.configure(values=['(gak ada file relevan ketemu di folder ini)'])
            self.quickpick_var.set('(gak ada file relevan ketemu di folder ini)')

    def _on_quickpick(self, value):
        if value and not value.startswith('Pilih') and not value.startswith('('):
            self.file_var.set(str(BASE_DIR / value))

    def _browse_file(self):
        cfg = self._current_config()
        path = filedialog.askopenfilename(
            title='Pilih file sumber', initialdir=str(BASE_DIR), filetypes=cfg['filetypes']
        )
        if path:
            self.file_var.set(path)

    def _update_brand_label(self):
        cfg = self._current_config()
        if cfg.get('cloud_source'):
            self.brand_label.configure(text='', fg_color='transparent')
            return
        name = Path(self.file_var.get()).name if self.file_var.get() else ''
        detected = detect_brand_from_filename(name)
        target = cfg['brand']
        if not name:
            self.brand_label.configure(
                text=f'  Tujuan: {target}  ', text_color='#1f8a4c',
                fg_color=('#e3f6ea', '#173323')
            )
        elif detected == target or (not detected and not cfg['auto_brand']):
            self.brand_label.configure(
                text=f'  Tujuan: {target}  ', text_color='#1f8a4c',
                fg_color=('#e3f6ea', '#173323')
            )
        else:
            self.brand_label.configure(
                text=(f'  File terdeteksi {detected}, tetapi menu tujuan {target}  '
                      if detected else f'  Nama file harus memuat "{target.lower()}"  '),
                text_color='#d64545', fg_color=('#fbe6e6', '#3a1e1e')
            )

    def _clear_log(self):
        self.log_box.configure(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.configure(state='disabled')

    def _append_log(self, text):
        tag = None
        stripped = text.strip().upper()
        if stripped.startswith('ERROR') or 'TRACEBACK' in stripped:
            tag = 'error'
        elif stripped.startswith('SELESAI'):
            tag = 'success'
        elif text.strip().startswith('...'):
            tag = 'muted'

        self.log_box.configure(state='normal')
        self.log_box.insert('end', text, tag)
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item == '__DONE__' or (isinstance(item, tuple) and item[0] == '__DRIVE_DONE__'):
                    success = item == '__DONE__' or item[1]
                    self.running = False
                    self.run_button.configure(
                        state='normal',
                        text='Cek Data (baca saja)' if self._current_config().get('cloud_source')
                        else '▶   Jalankan Import'
                    )
                    self.apply_button.configure(state='normal')
                    self.status_label.configure(
                        text='Selesai.' if success else 'Gagal. Lihat log.',
                        text_color=COLOR_SUCCESS if success else COLOR_ERROR
                    )
                else:
                    self._append_log(item)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _on_run_clicked(self, apply=False):
        if self.running:
            return

        cfg = self._current_config()
        if cfg.get('cloud_source'):
            self._on_drive_run_clicked(apply)
            return
        source_path = self.file_var.get().strip()

        if not source_path:
            self._append_log('ERROR: pilih file sumber dulu.\n')
            return
        if not Path(source_path).exists():
            self._append_log(f'ERROR: file tidak ketemu: {source_path}\n')
            return
        if not CREDENTIALS_PATH.exists():
            self._append_log(f'ERROR: credentials.json tidak ketemu di {BASE_DIR}\n')
            return
        detected = detect_brand_from_filename(Path(source_path).name)
        if detected and detected != cfg['brand']:
            self._append_log(
                f"ERROR: file terdeteksi {detected}, tetapi menu tujuan {cfg['brand']}.\n"
            )
            return
        if cfg['auto_brand'] and not detected:
            self._append_log(
                f"ERROR: nama file harus mengandung kata \"{cfg['brand'].lower()}\".\n"
            )
            return

        month_idx = MONTHS.index(self.month_var.get())
        filter_month = None if month_idx == 0 else month_idx
        try:
            filter_year = int(self.year_var.get().strip())
        except ValueError:
            filter_year = 2026

        self._clear_log()

        self.running = True
        self.run_button.configure(state='disabled', text='Sedang jalan...')
        self.status_label.configure(text='Menjalankan import...', text_color=COLOR_ACCENT)

        thread = threading.Thread(
            target=self._run_import_thread,
            args=(cfg, source_path, self.sheet_name_var.get().strip(),
                  filter_year, filter_month),
            daemon=True
        )
        thread.start()

    def _on_drive_run_clicked(self, apply):
        cfg = self._current_config()
        source_link = self.drive_source_var.get().strip()
        target_link = self.drive_target_var.get().strip()
        source_tab = self.drive_source_tab_var.get().strip()
        target_tab = self.drive_target_tab_var.get().strip()
        if not source_link or not target_link or not source_tab or not target_tab:
            self._append_log('ERROR: link raw, link tujuan, dan kedua nama tab wajib diisi.\n')
            return
        if not CREDENTIALS_PATH.exists():
            self._append_log(f'ERROR: credentials.json tidak ketemu di {BASE_DIR}\n')
            return
        try:
            module = importlib.import_module(cfg['module'])
            source_id = module.spreadsheet_id(source_link)
            target_id = module.spreadsheet_id(target_link)
            year = int(self.drive_year_var.get().strip())
            month = MONTHS.index(self.drive_month_var.get())
            if not 1900 <= year <= 2100 or not 1 <= month <= 12:
                raise ValueError('Tahun atau bulan tidak valid.')
        except (ImportError, ValueError) as exc:
            self._append_log(f'ERROR: {exc}\n')
            return

        confirmation = (
            'Aplikasi akan mengisi tab Content per brand/PAID, memperbarui total GMV, '
            'serta mengosongkan baris room ID ganda atau salah tab PAID. Lanjutkan?'
            if cfg.get('content_source') else
            f'Aplikasi akan menulis baris baru ke tab "{target_tab}". Lanjutkan?'
        )
        if apply and not messagebox.askyesno(
            'Konfirmasi impor ke Google Sheets', confirmation,
        ):
            return

        tab_argument = '--paid-tab' if cfg.get('content_source') else '--target-tab'
        command = [
            sys.executable, '-u', str(BASE_DIR / cfg['script']),
            '--source-id', source_id, '--target-id', target_id,
            '--source-tab', source_tab, tab_argument, target_tab,
            '--year', str(year), '--month', str(month),
            '--credentials', str(CREDENTIALS_PATH),
        ]
        if cfg.get('repair_paid_routing'):
            command.append('--repair-paid-routing')
        if apply:
            command.append('--apply')

        self._clear_log()
        self.running = True
        self.run_button.configure(state='disabled', text='Sedang jalan...')
        self.apply_button.configure(state='disabled')
        self.status_label.configure(
            text='Mengimpor baris baru...' if apply else 'Mengecek data (baca saja)...',
            text_color=COLOR_ACCENT,
        )
        threading.Thread(target=self._run_drive_thread, args=(command,), daemon=True).start()

    def _run_drive_thread(self, command):
        success = False
        try:
            process = subprocess.Popen(
                command, cwd=str(BASE_DIR), stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace',
            )
            for line in process.stdout:
                self.log_queue.put(line)
            success = process.wait() == 0
            if not success:
                self.log_queue.put(f'ERROR: proses berhenti dengan kode {process.returncode}.\n')
        except Exception:
            self.log_queue.put(traceback.format_exc())
        finally:
            self.log_queue.put(('__DRIVE_DONE__', success))

    def _run_import_thread(self, cfg, source_path, sheet_name, filter_year, filter_month):
        old_stdout = sys.stdout
        sys.stdout = QueueWriter(self.log_queue)
        try:
            module = importlib.import_module(cfg['module'])
            module.CREDENTIALS_FILE = str(CREDENTIALS_PATH)
            module.SOURCE_XLSX_PATH = source_path
            module.SOURCE_SHEET_NAME = sheet_name or 'Custom report'

            if cfg['module'] in ('import_live_streaming_unified', 'import_video_bank_unified'):
                module.TARGET_SPREADSHEET_ID = SPREADSHEET_IDS[cfg['brand']]

            if cfg['has_month_filter']:
                module.FILTER_YEAR = filter_year
                module.FILTER_MONTH = filter_month

            module.main()
        except Exception:
            traceback.print_exc()
        finally:
            sys.stdout = old_stdout
            self.log_queue.put('__DONE__')


if __name__ == '__main__':
    app = App()
    app.mainloop()
