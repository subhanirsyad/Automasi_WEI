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
                    args=['--disable-blink-features=AutomationControlled'],
                    ignore_default_args=['--enable-automation'])
                break
            except Exception as e:                # channel nggak terpasang
                galat = e
        else:
            self._pw.stop()
            raise RuntimeError('Chrome / Edge tidak bisa dibuka: %s' % galat)
        self.page = self._ctx.pages[0] if self._ctx.pages else self._ctx.new_page()
        self.page.goto(START, wait_until='domcontentloaded')
        return self

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

    def sudah_login(self):
        """Dipakai berulang selama menunggu login, jadi nggak boleh melempar error."""
        try:
            d = self._post('/creator/detail', _body_detail(PROBE_ID), percobaan=1)
        except SesiExpired:
            return False
        return isinstance(d, dict) and bool(d.get('handle'))

    def tunggu_login(self, log, stop, batas=900):
        """Tunggu sampai pengguna selesai login di jendela browser. False kalau di-stop."""
        mulai, info = time.time(), False
        while time.time() - mulai < batas:
            if stop.is_set():
                return False
            if self.sudah_login():
                log('Login terdeteksi, mulai memproses.\n')
                return True
            if not info:
                log('Login Kalodata di jendela browser yang terbuka (termasuk OTP). '
                    'Proses lanjut otomatis setelah login.\n')
                info = True
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
