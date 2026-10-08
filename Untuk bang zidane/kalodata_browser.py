"""
Kalodata lewat browser asli (Chrome / Edge)
================================================================================
Browser dibuka dengan PROFIL KHUSUS (folder profil_chrome). Pengguna login Kalodata
sendiri (termasuk OTP) di jendela itu; sesinya tersimpan di profil, jadi login
berikutnya otomatis. Permintaan ke Kalodata dikirim dari DALAM halaman itu, jadi
cookie dan Cloudflare ditangani browser: tidak ada cURL / cookie yang perlu di-copy.

Endpoint yang dipakai (sama dengan yang dipakai situs Kalodata sendiri):
  POST /overview/fullText/search  nama -> creator_uid
  POST /creator/detail            creator_uid -> whatsapp, email, dll.
"""

import json
import re
import time
from datetime import date, timedelta

URL = 'https://www.kalodata.com'
START = URL + '/creator'
PROBE_ID = '6701563873919894530'   # creator uji, cuma buat cek "sudah login?"
DELAY = 1.2                        # jeda antar request; jangan <1.0 (rate-limit / Cloudflare)

_FETCH = """async ([path, body]) => {
  const r = await fetch(path, {method: 'POST', headers: {
    'content-type': 'application/json', 'accept': 'application/json, text/plain, */*',
    'country': 'ID', 'currency': 'IDR', 'language': 'en-US'}, body: JSON.stringify(body)});
  return {status: r.status, ct: r.headers.get('content-type') || '', text: await r.text()};
}"""


class SesiExpired(RuntimeError):
    """Belum login / sesi habis / ada verifikasi Cloudflare / jendela browser ditutup."""


def _body_detail(uid):
    hari = date.today()
    return {'id': uid, 'startDate': (hari - timedelta(days=31)).isoformat(),
            'endDate': (hari - timedelta(days=1)).isoformat(), 'cateIds': [], 'authority': True}


class KalodataBrowser:
    def __init__(self, profil_dir, delay=DELAY):
        self.profil, self.delay = str(profil_dir), delay
        self._pw = self._ctx = self.page = None

    def __enter__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        galat = None
        for channel in ('chrome', 'msedge'):      # Edge selalu ada di Windows 10/11
            try:
                self._ctx = self._pw.chromium.launch_persistent_context(
                    self.profil, channel=channel, headless=False, no_viewport=True,
                    chromium_sandbox=True,            # tanpa ini muncul peringatan "--no-sandbox"
                    args=['--disable-blink-features=AutomationControlled', '--start-maximized'],
                    ignore_default_args=['--enable-automation'])
                break
            except Exception as e:                # channel nggak terpasang
                galat = e
        else:
            self._pw.stop()
            raise RuntimeError('Chrome / Edge tidak bisa dibuka: %s' % galat)
        self.page = self._ctx.pages[0] if self._ctx.pages else self._ctx.new_page()
        self._maksimalkan()
        self.page.goto(START, wait_until='domcontentloaded')
        return self

    def _maksimalkan(self):
        """Jendela penuh, supaya tombol login di pojok kanan atas kelihatan (--start-maximized
        sering diabaikan kalau profil menyimpan ukuran jendela lama)."""
        try:
            cdp = self._ctx.new_cdp_session(self.page)
            wid = cdp.send('Browser.getWindowForTarget')['windowId']
            cdp.send('Browser.setWindowBounds', {'windowId': wid, 'bounds': {'windowState': 'maximized'}})
        except Exception:
            pass

    def __exit__(self, *exc):
        try:
            self._ctx.close()
        except Exception:
            pass
        self._pw.stop()

    # ------------------------------------------------------------ request ----
    def _post(self, path, body, percobaan=3):
        for ke in range(1, percobaan + 1):
            try:
                r = self.page.evaluate(_FETCH, [path, body])
            except Exception as e:
                raise SesiExpired('Jendela browser tidak bisa dipakai (%s). Jangan tutup '
                                  'jendelanya selama proses jalan.' % type(e).__name__) from e
            if r['status'] in (401, 403) or 'json' not in r['ct']:
                raise SesiExpired('Belum login, atau ada verifikasi Cloudflare di jendela browser.')
            d = json.loads(r['text'])
            time.sleep(self.delay)
            if d.get('success'):
                return d.get('data')
            if ke < percobaan:
                time.sleep(self.delay * 2 * ke)
        return None

    def _cek(self):
        """(sudah_login, alasan_kalau_belum). Dipakai berulang, jadi nggak boleh melempar error."""
        try:
            if 'kalodata.com' not in self.page.url:       # fetch relatif butuh halaman Kalodata
                self.page.goto(START, wait_until='domcontentloaded')
            r = self.page.evaluate(_FETCH, ['/creator/detail', _body_detail(PROBE_ID)])
            url = self.page.url
        except Exception as e:                            # halaman lagi berpindah (habis login)
            return False, 'halaman sedang berpindah (%s)' % type(e).__name__
        if r['status'] in (401, 403):
            return False, 'HTTP %d, halaman: %s' % (r['status'], url[:60])
        if 'json' not in r['ct']:
            return False, 'bukan JSON, kemungkinan verifikasi Cloudflare (HTTP %d)' % r['status']
        d = json.loads(r['text'])
        if d.get('success') and isinstance(d.get('data'), dict) and d['data']:
            return True, ''
        return False, 'Kalodata menjawab success=%s code=%s, halaman: %s' % (
            d.get('success'), d.get('code'), url[:60])

    def sudah_login(self):
        return self._cek()[0]

    def _klik_login(self):
        """Buka kotak login otomatis (tombol bisa terpotong di jendela sempit). Boleh gagal."""
        try:
            self.page.get_by_text('Log-in / Sign-up').first.click(timeout=8000)
            self.page.get_by_text('Log-in', exact=True).first.click(timeout=5000)   # tab "Log-in"
        except Exception:
            pass

    def tunggu_login(self, log, stop, batas=900, lanjut=None):
        """
        Tunggu pengguna selesai login di jendela browser. False kalau di-stop.
        lanjut = threading.Event: kalau di-set (tombol "Sudah login"), deteksi dilewati.
        """
        mulai, n = time.time(), 0
        while time.time() - mulai < batas:
            if stop.is_set():
                return False
            if lanjut is not None and lanjut.is_set():
                log('Dilanjutkan manual oleh pengguna.\n')
                return True
            ok, alasan = self._cek()
            if ok:
                log('Login terdeteksi, mulai memproses.\n')
                return True
            if n == 0:
                log('Jendela browser terbuka di Kalodata tapi BELUM LOGIN. Kotak login dibuka '
                    'otomatis (kalau tidak muncul, klik "Log-in / Sign-up" di pojok kanan atas '
                    'jendela itu), lalu login dan selesaikan OTP. (Tanda belum login: angka di '
                    'tabel tampil ****.) Proses lanjut otomatis setelah login.\n')
                self._klik_login()
            elif n % 5 == 0:                              # tiap ~15 detik, supaya penyebabnya kelihatan
                log('Belum terdeteksi login: %s. Kalau angka di tabel masih ****, '
                    'login dulu di jendela browser itu.\n' % alasan)
            n += 1
            time.sleep(3)
        raise SesiExpired('Waktu tunggu login habis (15 menit).')

    # ------------------------------------------------------------- Kalodata --
    def cari(self, nama):
        """Balikin (hasil_cocok | None, saran_handle | None). Cocok = handle/nickname PERSIS."""
        data = self._post('/overview/fullText/search', {
            'country_code': 'id', 'keyword': nama,
            'scope': [{'index': 'creator', 'pageNo': 1, 'pageSize': 10}]})
        hasil = (data or {}).get('creator') or []
        n = nama.lower().lstrip('@')
        for kunci in ('creator_handle', 'creator_nickname'):
            for c in hasil:
                if str(c.get(kunci, '')).lower() == n:
                    return c, None
        return None, (hasil[0].get('creator_handle') if hasil else None)

    def kontak(self, uid):
        return self._post('/creator/detail', _body_detail(uid)) or {}
