"""
Cari WA Creator (Kalodata) - untuk Bang Zidane
================================================================================
Daftar creator (Excel / Google Spreadsheet) -> nomor WhatsApp & kontak lain.

  Spreadsheet : tiap creator ketemu, langsung ditambahkan ke tab "Hasil WA" di
                spreadsheet yang sama. Tidak ada file Excel.
  Excel       : hasil disimpan ke file Excel baru.

Jalankan dengan klik dua kali Jalankan.bat (atau: python app.py).
Panduan lengkap: README.md
"""

import queue
import threading
import time
import traceback
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

import inti
from kalodata_browser import DELAY, KalodataBrowser, SesiExpired

PROFIL = inti.BASE_DIR / 'profil_chrome'      # login Kalodata tersimpan di sini
XLSX = [('Excel files', '*.xlsx')]

BG = ('#eef0f4', '#15161a')
CARD = ('#ffffff', '#1d1f24')
ACCENT = ('#2b6ef2', '#4d8dff')
ACCENT_HOVER = ('#1f57c9', '#3f78dd')
MUTED = ('#6b7280', '#9199a6')
OK = ('#1f8a4c', '#4fd382')
ERR = ('#d64545', '#ff6b6b')
LOG_BG = ('#12141a', '#0b0c10')
MODE_EXCEL, MODE_SHEET = 'File Excel', 'Google Spreadsheet'


def card(parent):
    return ctk.CTkFrame(parent, fg_color=CARD, corner_radius=14, border_width=1,
                        border_color=('#e1e4ea', '#2a2d34'))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode('System')
        ctk.set_default_color_theme('blue')
        self.title('Cari WA Creator - Kalodata')
        self.geometry('820x760')
        self.minsize(700, 620)
        self.configure(fg_color=BG)

        self.f_title = ctk.CTkFont(family='Segoe UI', size=20, weight='bold')
        self.f_h2 = ctk.CTkFont(family='Segoe UI', size=15, weight='bold')
        self.f_body = ctk.CTkFont(family='Segoe UI', size=13)
        self.f_small = ctk.CTkFont(family='Segoe UI', size=11)
        self.f_mono = ctk.CTkFont(family='Consolas', size=12)

        cfg = inti.muat_pengaturan()
        self.q, self.stop, self.running = queue.Queue(), threading.Event(), False
        self.mode_var = ctk.StringVar(value=cfg.get('mode', MODE_SHEET))
        self.link_var = ctk.StringVar(value=cfg.get('link', ''))
        self.tab_var = ctk.StringVar(value=cfg.get('tab', ''))
        self.in_var = ctk.StringVar(value=cfg.get('excel', ''))
        self.out_var = ctk.StringVar()
        self.in_var.trace_add('write', lambda *a: self._auto_output())

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        outer = ctk.CTkScrollableFrame(self, fg_color='transparent')
        outer.grid(row=0, column=0, sticky='nsew', padx=28, pady=24)
        outer.grid_columnconfigure(0, weight=1)
        self._build(outer)
        self._on_mode(self.mode_var.get())
        self.after(100, self._poll)

    # ------------------------------------------------------------------ UI ---
    def _build(self, outer):
        ctk.CTkLabel(outer, text='📞  Cari WA Creator', font=self.f_title, anchor='w').grid(
            row=0, column=0, sticky='ew')
        ctk.CTkLabel(outer, font=self.f_body, text_color=MUTED, anchor='w', justify='left',
                     wraplength=700, text=(
                         'Isi daftar username TikTok creator, klik jalan. Pertama kali, jendela browser '
                         'terbuka: login Kalodata (+ OTP) di situ. Setelah itu otomatis.')
                     ).grid(row=1, column=0, sticky='ew', pady=(2, 18))

        # ---- kartu input
        c = card(outer)
        c.grid(row=2, column=0, sticky='ew', pady=(0, 16))
        inner = ctk.CTkFrame(c, fg_color='transparent')
        inner.pack(fill='x', padx=22, pady=16)
        inner.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(inner, text='1.  Daftar creator', font=self.f_h2, anchor='w').grid(
            row=0, column=0, sticky='w', pady=(0, 8))
        ctk.CTkSegmentedButton(inner, values=[MODE_SHEET, MODE_EXCEL], variable=self.mode_var,
                               command=self._on_mode).grid(row=0, column=1, sticky='e', pady=(0, 8))

        self.xl_row = ctk.CTkFrame(inner, fg_color='transparent')
        self.xl_row.grid(row=1, column=0, columnspan=2, sticky='ew')
        self.xl_row.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(self.xl_row, textvariable=self.in_var, height=38, font=self.f_body,
                     placeholder_text='Belum dipilih...').grid(row=0, column=0, sticky='ew', padx=(0, 8))
        ctk.CTkButton(self.xl_row, text='Browse...', width=110, height=38, font=self.f_body,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=self._browse_in).grid(row=0, column=1)

        self.gs_row = ctk.CTkFrame(inner, fg_color='transparent')
        self.gs_row.grid(row=1, column=0, columnspan=2, sticky='ew')
        self.gs_row.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(self.gs_row, textvariable=self.link_var, height=38, font=self.f_body,
                     placeholder_text='Link Google Spreadsheet (docs.google.com/spreadsheets/d/...)'
                     ).grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 8))
        ctk.CTkEntry(self.gs_row, textvariable=self.tab_var, height=38, font=self.f_body,
                     placeholder_text='Nama tab (kosong = tab dari link, atau tab pertama)'
                     ).grid(row=1, column=0, columnspan=2, sticky='ew')
        ctk.CTkLabel(self.gs_row, anchor='w', justify='left', wraplength=660, font=self.f_small,
                     text_color=MUTED, text=(
                         'Hasil langsung ditulis ke tab "%s" di spreadsheet ini, satu baris tiap creator '
                         'selesai. Tab input tidak diubah. Kalau dijalankan ulang, creator yang sudah '
                         'ada di tab itu dilewati.' % inti.TAB_HASIL)
                     ).grid(row=2, column=0, columnspan=2, sticky='w', pady=(8, 0))
        email = inti.email_service_account()
        self.sa_label = ctk.CTkLabel(
            self.gs_row, anchor='w', font=self.f_small, wraplength=520, justify='left',
            text_color=MUTED if email else ERR,
            text=('Share spreadsheet sebagai Editor ke: %s' % email) if email
            else 'credentials.json belum ada di folder aplikasi (lihat README).')
        self.sa_label.grid(row=3, column=0, sticky='w', pady=(6, 0))
        if email:
            ctk.CTkButton(self.gs_row, text='Salin email', width=100, height=28, font=self.f_small,
                          fg_color='transparent', border_width=1, text_color=('#1b1d22', '#e7e9ee'),
                          border_color=('#c4c9d2', '#3a3d45'), hover_color=('#dbe1eb', '#20232a'),
                          command=lambda: self._copy(email)).grid(row=3, column=1, sticky='e', pady=(6, 0))

        # ---- kartu file hasil (khusus Excel)
        self.out_card = card(outer)
        self.out_card.grid(row=3, column=0, sticky='ew', pady=(0, 16))
        oi = ctk.CTkFrame(self.out_card, fg_color='transparent')
        oi.pack(fill='x', padx=22, pady=16)
        oi.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(oi, text='2.  File hasil (Excel)', font=self.f_h2, anchor='w').grid(
            row=0, column=0, columnspan=2, sticky='ew', pady=(0, 8))
        ctk.CTkEntry(oi, textvariable=self.out_var, height=38, font=self.f_body).grid(
            row=1, column=0, sticky='ew', padx=(0, 8))
        ctk.CTkButton(oi, text='Browse...', width=110, height=38, font=self.f_body,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=self._browse_out).grid(row=1, column=1)

        # ---- jalan / stop / log
        row = ctk.CTkFrame(outer, fg_color='transparent')
        row.grid(row=4, column=0, sticky='ew', pady=(0, 6))
        row.grid_columnconfigure(0, weight=1)
        self.run_btn = ctk.CTkButton(
            row, text='▶   Cari Nomor WA', height=46, corner_radius=10,
            font=ctk.CTkFont(family='Segoe UI', size=14, weight='bold'),
            fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self._run)
        self.run_btn.grid(row=0, column=0, sticky='ew')
        self.stop_btn = ctk.CTkButton(row, text='Stop', height=46, width=90, corner_radius=10,
                                      fg_color=ERR, state='disabled', command=self.stop.set)
        self.stop_btn.grid(row=0, column=1, padx=(8, 0))
        self.bar = ctk.CTkProgressBar(outer)
        self.bar.set(0)
        self.bar.grid(row=5, column=0, sticky='ew', pady=(6, 0))
        self.status = ctk.CTkLabel(outer, text='Siap.', font=self.f_body, text_color=MUTED, anchor='w')
        self.status.grid(row=6, column=0, sticky='w', pady=(6, 12))
        self.log = ctk.CTkTextbox(outer, font=self.f_mono, fg_color=LOG_BG, text_color='#d8dbe2',
                                  wrap='word', height=300)
        self.log.grid(row=7, column=0, sticky='nsew')
        self.log.configure(state='disabled')
        self.log.tag_config('error', foreground='#ff6b6b')

    # ------------------------------------------------------------- aksi UI ---
    def _on_mode(self, mode):
        sheet = mode == MODE_SHEET
        (self.gs_row.grid if sheet else self.gs_row.grid_remove)()
        (self.xl_row.grid_remove if sheet else self.xl_row.grid)()
        (self.out_card.grid_remove if sheet else self.out_card.grid)()

    def _copy(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.configure(text='Email disalin.', text_color=OK)

    def _browse_in(self):
        p = filedialog.askopenfilename(title='Pilih Excel daftar creator', filetypes=XLSX)
        if p:
            self.in_var.set(p)

    def _browse_out(self):
        p = filedialog.asksaveasfilename(title='Simpan hasil sebagai', defaultextension='.xlsx',
                                         filetypes=XLSX, initialfile=Path(self.out_var.get()).name)
        if p:
            self.out_var.set(p)

    def _auto_output(self):
        p = Path(self.in_var.get())
        if p.suffix.lower() == '.xlsx':
            self.out_var.set(str(p.with_name(p.stem + '_WA.xlsx')))

    def _append(self, text, tag=None):
        if tag is None and text.strip().upper().startswith('ERROR'):
            tag = 'error'
        self.log.configure(state='normal')
        self.log.insert('end', text, tag)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                if isinstance(item, tuple) and item[0] == 'progress':
                    self.bar.set(item[1] / max(item[2], 1))
                    self.status.configure(text='Memproses %d / %d...' % item[1:], text_color=ACCENT)
                elif isinstance(item, tuple) and item[0] == 'done':
                    self.running = False
                    self.run_btn.configure(state='normal', text='▶   Cari Nomor WA')
                    self.stop_btn.configure(state='disabled')
                    self.status.configure(text='Selesai.' if item[1] else 'Berhenti. Lihat log.',
                                          text_color=OK if item[1] else ERR)
                else:
                    self._append(item)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    # ---------------------------------------------------------------- jalan --
    def _run(self):
        if self.running:
            return
        sheet = self.mode_var.get() == MODE_SHEET
        src = (self.link_var if sheet else self.in_var).get().strip()
        if not src or (not sheet and not Path(src).is_file()):
            self._append('ERROR: %s.\n' % ('isi link spreadsheet dulu' if sheet
                                           else 'pilih file Excel yang valid dulu'))
            return
        inti.simpan_pengaturan({'mode': self.mode_var.get(), 'link': self.link_var.get().strip(),
                                'tab': self.tab_var.get().strip(), 'excel': self.in_var.get().strip()})
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')
        self.stop.clear()
        self.running = True
        self.bar.set(0)
        self.run_btn.configure(state='disabled', text='Sedang jalan...')
        self.stop_btn.configure(state='normal')
        self.status.configure(text='Membaca daftar creator...', text_color=ACCENT)
        threading.Thread(target=self._work, args=(sheet, src, self.tab_var.get(), self.out_var.get().strip()),
                         daemon=True).start()

    def _work(self, sheet, src, tab, out):
        log, ok, sink = self.q.put, False, None
        try:
            if sheet:
                client = inti.SheetsClient()
                names, kolom, tab_in, sid = client.baca_input(src, tab)
                sudah = client.sudah_ada(sid)
                baru = [n for n in names if n not in sudah]
                log('Tab "%s", kolom "%s": %d creator unik, %d sudah ada di "%s" (dilewati), %d diproses.\n'
                    % (tab_in, kolom, len(names), len(names) - len(baru), inti.TAB_HASIL, len(baru)))
                names = baru
                if names:
                    client.siapkan_hasil(sid)
                    sink = inti.SheetSink(client, sid)
            else:
                names, kolom = inti.baca_excel(src)
                log('Kolom "%s": %d creator unik.\n' % (kolom, len(names)))
                sink = inti.ExcelSink(out or str(Path(src).with_name(Path(src).stem + '_WA.xlsx')))
            if not names:
                log('Tidak ada creator yang perlu diproses.\n')
                ok = True
                return
            log('Perkiraan %d menit. Jangan tutup jendela browser yang terbuka.\n'
                % max(1, round(len(names) * 2 * DELAY / 60)))
            self.q.put(('progress', 0, len(names)))
            with KalodataBrowser(PROFIL) as kal:
                if kal.tunggu_login(log, self.stop):
                    ok = inti.jalankan(names, kal, sink, log,
                                       lambda i, n: self.q.put(('progress', i, n)), self.stop)
        except SesiExpired as e:
            log('ERROR: %s\nLogin ulang di jendela browser lalu jalankan lagi. Creator yang sudah '
                'selesai tidak diulang (spreadsheet) / sudah tersimpan (Excel).\n' % e)
        except (PermissionError, FileNotFoundError, ValueError, RuntimeError) as e:
            log('ERROR: %s\n' % e)
        except Exception:
            log('ERROR:\n' + traceback.format_exc())
        finally:
            if isinstance(sink, inti.ExcelSink) and sink.rows:
                log('Tersimpan: %s (%d baris)\n' % (sink.simpan(), len(sink.rows)))
            self.q.put(('done', ok))


if __name__ == '__main__':
    App().mainloop()
