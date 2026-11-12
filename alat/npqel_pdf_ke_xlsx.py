#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
npqel_pdf_ke_xlsx.py - Tukar laporan PDF i-Penilaian (NPQEL) kepada SATU fail xlsx
dalam format muat naik Dashboard Penilaian NPQEL.

Borang yang dikenali (dikesan daripada KANDUNGAN PDF, bukan nama fail):
  BK15a  sesi latihan (satu PDF setiap sesi)      BK15b  khidmat sokongan
  BK15c  soalan terbuka sesi latihan              BK15d  soalan terbuka khidmat sokongan
  BK30   laporan khas pengurusan                  BK38   laporan am kursus
  BK64   senarai peserta (hanya dikira secara agregat; nama, No. KP, telefon, e-mel TIDAK disimpan)

Penggunaan baris arahan:
  python npqel_pdf_ke_xlsx.py FAIL_ATAU_FOLDER_ATAU_ZIP ... -k SRGH6 -t 2027 -o Data_SRGH6_2027.xlsx
  python npqel_pdf_ke_xlsx.py folder_pdf/ -k SRGH6 -t 2027 --fasa 2027-03-02=1 --fasa 2027-05-11=2

Antara muka web setempat (muat naik pukal melalui pelayar, tiada data keluar dari komputer):
  python npqel_pdf_ke_xlsx.py --web          (buka http://127.0.0.1:8765)

Keperluan: Python 3.9+, pdfplumber, openpyxl.
"""
import argparse, base64, collections, datetime, difflib, io, json, os, re, sys, threading, zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, unquote

SKEMA = json.loads(r'''{"skema": {"fasa": {"fasa": "t", "kod_program": "t", "tempoh": "t", "tempat": "t", "bil_ditawarkan": "n", "bil_hadir": "n", "bil_responden_sesi": "n", "bil_pensyarah_terlibat": "n", "bil_sesi_dinilai": "n", "bil_pensyarah_dinilai": "n", "ada_data_sesi": "t", "keseluruhan": "n", "keseluruhan_dilaporkan": "n", "status_keseluruhan": "t", "status_keseluruhan_dilaporkan": "t", "pengisian": "n", "pengisian_dilaporkan": "n", "status_pengisian": "t", "status_pengisian_dilaporkan": "t", "keupayaan": "n", "keupayaan_dilaporkan": "n", "status_keupayaan": "t", "status_keupayaan_dilaporkan": "t", "komposit": "n", "komposit_dilaporkan": "n", "status_komposit": "t", "status_komposit_dilaporkan": "t", "khidmat_sokongan": "n", "status_khidmat_sokongan": "t", "khidmat_sokongan_bk15b": "n", "min_khidmat_sokongan_bk15b": "n", "responden_item_kemudahan": "n", "responden_item_pengurus": "n", "respons_terbuka_sesi": "n", "respons_terbuka_sokongan": "n", "ulasan_pengurus": "t", "sumber": "t"}, "sesi": {"id": "t", "fasa": "t", "modul": "t", "topik": "t", "topik_asal": "t", "pensyarah": "t", "bil_hadir": "n", "bil_responden": "n", "L1": "n", "L2": "n", "L3": "n", "L4": "n", "L5": "n", "L6": "n", "L7": "n", "L8": "n", "purata_pengisian": "n", "purata_keupayaan": "n", "purata_keseluruhan": "n", "peratus_keseluruhan": "n", "peratus_pengisian": "n", "peratus_keupayaan": "n", "tahap": "t", "skor_bk30": "n", "sumber_skor": "t", "fail_sumber": "t", "catatan": "t"}, "sokongan": {"kunci": "t", "fasa": "t", "item": "t", "no_item": "n", "perkara": "t", "sub_aspek": "t", "no_sub_aspek": "n", "bil_menjawab": "n", "bil_berpuas_hati": "n", "peratus": "n", "min": "n", "min_dilaporkan": "n", "fail_sumber": "t"}, "terbuka_sesi": {"no": "n", "fasa": "t", "modul": "t", "topik": "t", "pensyarah": "t", "bil": "n", "v12": "t", "v13": "t", "tema_v13": "t", "sumber": "t"}, "terbuka_sokongan": {"no": "n", "fasa": "t", "bil": "n", "respons": "t", "fail_sumber": "t"}, "peserta": {"kunci": "t", "fasa": "t", "dimensi": "t", "kategori": "t", "bilangan": "n"}, "perubahan": {"no": "n", "fasa": "t", "butiran": "t", "fail_sumber": "t"}}, "opsyen": ["ada_data_sesi", "bil_hadir_sesi", "catatan", "fail_sumber", "keseluruhan_dilaporkan", "keupayaan_dilaporkan", "kod_program", "komposit_dilaporkan", "kunci", "min_dilaporkan", "min_khidmat_sokongan_bk15b", "modul", "pengisian_dilaporkan", "purata_keseluruhan", "purata_keupayaan", "purata_pengisian", "skor_bk30", "status_keseluruhan", "status_keseluruhan_dilaporkan", "status_keupayaan", "status_keupayaan_dilaporkan", "status_khidmat_sokongan", "status_komposit_dilaporkan", "status_pengisian", "status_pengisian_dilaporkan", "sumber", "sumber_skor", "tahap", "tema_v13", "tempat", "topik_asal", "ulasan_pengurus", "bil_hadir@sesi"]}''')
NAMA = ['fasa', 'sesi', 'sokongan', 'terbuka_sesi', 'terbuka_sokongan', 'peserta', 'perubahan']
OPSYEN = set(SKEMA['opsyen'])
def opsyen(n, k): return k in OPSYEN or (n == 'sesi' and k == 'bil_hadir')

# ----------------------------------------------------------------------------- utiliti
def norm(s): return re.sub(r'\s+', ' ', s or '').strip()
def ns(x): return re.sub(r'\s', '', str(x))
def natkey(s): return [int(x) if x.isdigit() else x for x in re.split(r'(\d+)', s)]
BULAN = {'JAN': 1, 'FEB': 2, 'MAC': 3, 'MAR': 3, 'APR': 4, 'MEI': 5, 'JUN': 6, 'JUL': 7, 'OGO': 8, 'AUG': 8,
         'SEP': 9, 'OKT': 10, 'OCT': 10, 'NOV': 11, 'DIS': 12, 'DEC': 12}
def iso_dmy(s):
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', s or '')
    return f'{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}' if m else None
def iso_malay(s):
    m = re.search(r'(\d{1,2})\s+([A-Za-z]{3,})\.?,?\s+(\d{4})', s or '')
    if not m or m.group(2)[:3].upper() not in BULAN: return None
    return f'{m.group(3)}-{BULAN[m.group(2)[:3].upper()]:02d}-{int(m.group(1)):02d}'
def kod_borang(t):
    m = re.search(r'\bBK\s*-?\s*(\d+[a-d]?)\b', t[:1500]); return 'BK' + m.group(1) if m else None
def teks_pdf(pdf): return '\n'.join((pg.extract_text() or '') for pg in pdf.pages)
def fld(src, lab, end):
    m = re.search(lab + r'\s*:\s*(.*?)\s*(?=' + end + ')', src, re.S); return norm(m.group(1)) if m else None
def intu(x): return int(x) if x and str(x).isdigit() else None

# ----------------------------------------------------------------------------- pengekstrak (satu PDF)
def baca_pdf(sumber, nama):
    """Baca satu PDF (laluan atau bytes). Pulangkan dict: jenis, mula (ISO), kod, hint, data, tidak_dibaca."""
    import pdfplumber
    fh = io.BytesIO(sumber) if isinstance(sumber, (bytes, bytearray)) else sumber
    fn = os.path.basename(nama); unread = []
    hm = re.search(r'Fasa\s*([1-9])\b', nama, re.I)
    R = dict(nama=nama, fail=fn, jenis=None, mula=None, kod=None, hint=int(hm.group(1)) if hm else None, data=None, tidak_dibaca=unread)
    with pdfplumber.open(fh) as pdf:
        t = teks_pdf(pdf); code = kod_borang(t); R['jenis'] = code
        if code in ('BK15a', 'BK15b', 'BK15c', 'BK15d'):
            R['mula'] = iso_dmy(fld(t, 'Tarikh Mula', 'Tarikh Tamat') or '')
            R['kod'] = fld(t, r'Kod (?:Kursus/)?Program', r'Nama (?:Kursus|Program)')
        if code == 'BK15a': R['data'] = _bk15a(t, fn, unread)
        elif code == 'BK15b': R['data'] = _bk15b(t, fn, unread)
        elif code == 'BK15c': R['data'] = _bk15c(pdf, t, fn, unread)
        elif code == 'BK15d': R['data'] = _bk15d(pdf, fn, unread)
        elif code == 'BK30':
            R['data'] = _bk30(pdf, t, fn, unread); R['mula'] = iso_dmy((R['data']['h'].get('tempoh') or ''))
        elif code == 'BK38':
            R['data'] = _bk38(t, fn, unread); R['mula'] = iso_dmy((R['data'].get('tempoh') or ''))
            hm = re.search(r'\bFasa\s*([1-9])\b', t[:900])
            R['fasa_teks'] = int(hm.group(1)) if hm else None
        elif code == 'BK64':
            R['data'] = _bk64(pdf, fn, unread)
            m = re.search(r'KOD KURSUS\s*:\s*(\S+)', t); R['kod'] = m.group(1) if m else None
            m = re.search(r'TARIKH\s*:\s*([^\n]+?)\s*-\s*\d', t); R['mula'] = iso_malay(m.group(1)) if m else None
    return R

def _bk15a(t, fn, unread):
    r = dict(fail=fn)
    r['kod_kursus'] = fld(t, r'Kod Kursus/Program', 'Nama Kursus'); r['nama_kursus'] = fld(t, r'Nama Kursus/Program', 'Topik')
    r['topik'] = fld(t, 'Topik', 'Tempat Kursus'); r['tempat'] = fld(t, r'Tempat Kursus/Program', 'Nama Pensyarah')
    r['pensyarah'] = fld(t, 'Nama Pensyarah', 'Tarikh Mula'); r['tarikh_mula'] = fld(t, 'Tarikh Mula', 'Tarikh Tamat')
    r['tarikh_tamat'] = fld(t, 'Tarikh Tamat', 'Bil Peserta'); r['bil_hadir'] = intu(fld(t, 'Bil Peserta Hadir', 'Bil Responden'))
    r['bil_responden'] = intu(fld(t, 'Bil Responden', r'\n'))
    m = re.search(r'^\s*((?:\d\.\d{2}\s+){10}\d\.\d{2})\s*$', t, re.M)
    if m:
        for k, x in zip(['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'purata_pengisian', 'purata_keupayaan', 'purata_keseluruhan'], [float(x) for x in m.group(1).split()]): r[k] = x
    else: unread.append((fn, 'baris purata L1-L8'))
    for k, lab in [('pct_keseluruhan', 'Rumusan Keseluruhan Penilaian'), ('pct_pengisian', 'Elemen Pengisian Sesi Latihan'), ('pct_keupayaan', 'Elemen Keupayaan Pensyarah')]:
        m = re.search(lab + r'\s*:\s*([\d.]+)%\s*Tahap\s*:\s*([^\n]+)', t)
        if m: r[k] = float(m.group(1)); r[k.replace('pct', 'tahap')] = norm(m.group(2))
        else: unread.append((fn, lab))
    return r

def _bk15b(t, fn, unread):
    h = dict(fail=fn, kod_program=fld(t, 'Kod Program', 'Nama Program'), tarikh_mula=fld(t, 'Tarikh Mula', 'Tarikh Tamat'))
    m = re.search(r'Prestasi\s*:\s*([\d.]+)%\s*\(Min\s*=\s*([\d.]+)\)', t)
    h['prestasi_pct_header'], h['min_header'] = (float(m.group(1)), float(m.group(2))) if m else (None, None)
    items = re.findall(r'^\s*(W\d+)\s+\((\d+)\)\s+([\d.]+)\s*%\s*\((\d+)\)\s*(.*?)\s*=\s*([\d.]+)\s*$', t, re.M)
    if len(items) != 24: unread.append((fn, 'item W: jumpa %d daripada 24' % len(items)))
    w = [dict(item=a[0], bil_menjawab=int(a[1]), peratus=float(a[2]), bil_dalam_kurungan=int(a[3]), perkara=a[4], min=float(a[5])) for a in items]
    return dict(h=h, w=w)

def _bk15d(pdf, fn, unread):
    rows = []
    for pi, pg in enumerate(pdf.pages, 1):
        for tb in pg.extract_tables():
            for row in tb:
                c = [(x or '') for x in row] + ['']
                if norm(c[0]).isdigit(): rows.append(dict(fail=fn, halaman=pi, bil=int(norm(c[0])), respons=norm(c[1])))
    if not rows: unread.append((fn, 'tiada baris respons BK15d'))
    return rows

def _bk15c(pdf, t, fn, unread):
    out = []; cur = None
    if 'Belum Dimasukkan' in t: return []
    def frag(pg, tbl, where):
        x0, top, x1, bot = tbl.bbox
        cells = [c for r in tbl.rows for c in [r.cells] if len([k for k in c if k]) == 3]
        if not cells: return None
        cc = [k for k in cells[0] if k]; b1, b2 = cc[0][2], cc[1][2]
        ws = [w_ for w_ in pg.extract_words() if x0 - 2 <= w_['x0'] <= x1 + 2 and ((where == 'tail' and w_['top'] >= bot - 1) or (where == 'head' and w_['bottom'] <= top + 1)) and 'http' not in w_['text']]
        ws = [w_ for w_ in ws if not (where == 'tail' and w_['top'] > pg.height - 25)]
        if not ws: return None
        ws.sort(key=lambda w_: (round(w_['top']), w_['x0']))
        col = lambda w_: 0 if (w_['x0'] + w_['x1']) / 2 < b1 else (1 if (w_['x0'] + w_['x1']) / 2 < b2 else 2)
        return [' '.join(w_['text'] for w_ in ws if col(w_) == i) for i in range(3)] + [' '.join(w_['text'] for w_ in ws)]
    pages = []
    for pi, pg in enumerate(pdf.pages, 1):
        tbls = pg.find_tables(); rows = [r for tb in tbls for r in tb.extract()]
        head = frag(pg, tbls[0], 'head') if (tbls and pi > 1) else None
        tail = frag(pg, tbls[-1], 'tail') if tbls else None
        pages.append((pi, rows, head, tail))
    seq = []
    for i, (pi, rows, head, tail) in enumerate(pages):
        for r in rows: seq.append((pi, r, ''))
        if tail:
            nh = pages[i + 1][2] if i + 1 < len(pages) else None
            merged = [norm(tail[k] + ' ' + (nh[k] if nh else '')) for k in range(3)]
            ro = norm(tail[3] + ' ' + (nh[3] if nh else ''))
            if ro.startswith('TAJUK KULIAH') or ro.startswith('Tiada Maklum Balas'): merged = [ro, '', '']
            seq.append((pi, merged, 'baris terpotong di sempadan halaman (dipulihkan)'))
        elif i + 1 < len(pages) and pages[i + 1][2]:
            unread.append((fn, 'serpihan atas hlm %d tanpa pasangan' % (pi + 1)))
    for pi, row, note in seq:
        c = [(x or '') for x in row] + ['', '', '']; c0 = norm(c[0])
        if note and not (c0.isdigit() or c0.startswith('TAJUK KULIAH') or c0.startswith('Tiada Maklum Balas')):
            unread.append((fn, 'baris terpotong tidak dikenali hlm %d: %s' % (pi, ' | '.join(c)[:80]))); continue
        if c0.startswith('TAJUK KULIAH'):
            m = re.match(r'TAJUK KULIAH:\s*(.*),\s*TENAGA PENGAJAR:\s*(.*)$', c0); cur = (m.group(1).strip(), m.group(2).strip()) if m else (c0, None)
        elif c0 == 'BIL': continue
        elif c0.isdigit():
            out.append(dict(fail=fn, halaman=pi, topik=cur[0] if cur else None, pensyarah=cur[1] if cur else None, bil=int(c0), V12=norm(c[1]), V13=norm(c[2])))
        elif c0 == '' and (norm(c[1]) or norm(c[2])): unread.append((fn, 'baris tanpa BIL hlm %d' % pi))
        elif c0 == '' or c0.startswith('Tiada Maklum Balas') or 'Belum Dimasukkan' in c0: continue
        else: unread.append((fn, 'baris BK15c tidak dikenali hlm %d: %s' % (pi, c0[:40])))
    return out

def _bk30(pdf, t, fn, unread):
    h = dict(fail=fn, tempoh=fld(t, r'3\. Tempoh Kursus/Program', r'4\.'), tempat=fld(t, r'2\. Tempat Kursus/Program', r'3\.'), pengurus=fld(t, r'4\. Nama Pengurus Kursus/Program', r'5\.'))
    for k, lab in [('bil_ditawarkan', 'Bilangan Peserta Yang Ditawarkan Tempat'), ('bil_hadir', 'Bilangan Peserta Yang Hadir'), ('bil_responden', r'7\.4 Bil Responden')]:
        m = re.search(lab + r'\s*:\s*(\d+)', t); h[k] = int(m.group(1)) if m else None
    for k, lab in [('pct_pengisian', 'Pengisian Sesi Latihan'), ('pct_keupayaan', 'Keupayaan Pensyarah'), ('pct_khidmat_sokongan', 'Khidmat Sokongan')]:
        m = re.search(r'Penilaian Keseluruhan ' + lab + r'\s*:\s*([\d.]+)%', t); h[k] = float(m.group(1)) if m else None
    chg = []
    m = re.search(r'8\. Laporan perubahan.*?:\s*\n(.*?)(?=\n\s*9\.\s*Skor Pensyarah)', t, re.S)
    if m:
        for ch in re.split(r'\n\s*-\s+(?=SESI PENGAJARAN)', '\n' + m.group(1)):
            ch = norm(ch).lstrip('- ').strip()
            if ch: chg.append(dict(fail=fn, perubahan=ch))
    p = []
    for pi, pg in enumerate(pdf.pages, 1):
        for tb in pg.extract_tables():
            for row in tb:
                c = [norm(x) for x in row]
                if len(c) >= 4 and c[0].isdigit(): p.append(dict(fail=fn, bil=int(c[0]), pensyarah=c[1], tajuk=c[2], skor=float(c[3]) if re.fullmatch(r'[\d.]+', c[3]) else None))
    if h['tempoh'] is None: unread.append((fn, 'tempoh kursus BK30'))
    return dict(h=h, chg=chg, p=p)

def _bk38(t, fn, unread):
    h = dict(fail=fn, tempoh=fld(t, r'1\.3\s+Tempoh Kursus/Program', r'\(Tarikh Lapor|\n\s*1\.4'), pengurus=fld(t, r'1\.4\s+Nama Pengurus/Pengurus', r'Bersama|1\.5'))
    m = re.search(r'Bilangan Pensyarah Terlibat\s*:\s*(\d+)', t); h['bil_pensyarah'] = int(m.group(1)) if m else None
    m = re.search(r'Bil Responden\s*:\s*(\d+)', t); h['bil_responden'] = int(m.group(1)) if m else None
    for k, lab in [('keseluruhan', r'1\. Rumusan Keseluruhan Penilaian'), ('pengisian', r'1\.1 Penilaian Keseluruhan Pengisian Sesi Latihan'), ('keupayaan', r'1\.2 Penilaian Keseluruhan Keupayaan Pensyarah'), ('khidmat_sokongan', r'2\. Penilaian Keseluruhan Khidmat Sokongan'), ('komposit', r'3\. Komposit Kursus \(90%\) dan Khidmat Sokongan\s*\(10%\)')]:
        m = re.search(lab + r'\s+([\d.]+)\s+([^\n]+)', t)
        h['pct_' + k] = float(m.group(1)) if m else None; h['status_' + k] = norm(m.group(2)) if m else None
        if not m: unread.append((fn, 'BK38: ' + lab.replace('\\', '')))
    m = re.search(r'ditawarkan tempat adalah (\d+) orang dan yang hadir (\d+) orang', t)
    h['bil_ditawarkan'], h['bil_hadir'] = (int(m.group(1)), int(m.group(2))) if m else (None, None)
    m = re.search(r'7\. Ulasan Pengurus[^\n]*\n(.*?)(?=\n\s*8\. Lampiran)', t, re.S); h['ulasan_pengurus'] = norm(m.group(1)) if m else None
    return h

STATES = ['JOHOR', 'KEDAH', 'KELANTAN', 'MELAKA', 'NEGERI SEMBILAN', 'PAHANG', 'PERAK', 'PERLIS', 'PULAU PINANG', 'SABAH', 'SARAWAK', 'SELANGOR', 'TERENGGANU', 'KUALA LUMPUR', 'LABUAN', 'PUTRAJAYA']
def _bk64(pdf, fn, unread):
    """Hanya jawatan, gred, negeri dan jenis tempat bertugas. Nama/No. KP/telefon/e-mel tidak dibaca ke dalam hasil."""
    out = []
    for pi, pg in enumerate(pdf.pages, 1):
        for tb in pg.extract_tables():
            for row in tb:
                c = [(x or '') for x in row] + ['', '', '', '']
                if not norm(c[0]).isdigit(): continue
                cell = c[2]
                j = re.search(r'JAWATAN:\s*(.*?)\s*GRED:', cell, re.S); g = re.search(r'GRED:\s*(.*?)\s*NO TEL:', cell, re.S)
                tb_ = norm(c[3]).upper(); st = [s for s in STATES if re.search(r'\b' + s + r'\b', tb_)]
                if not tb_.strip(' ,'): jenis = '(kosong dalam laporan)'
                elif re.match(r'(SK\b|SEKOLAH KEBANGSAAN)', tb_): jenis = 'Sekolah Kebangsaan (SK)'
                elif re.match(r'(SJK|SEKOLAH JENIS)', tb_): jenis = 'Sekolah Jenis Kebangsaan (SJK)'
                elif tb_.startswith('PEJABAT PELAJARAN'): jenis = 'Pejabat Pelajaran Daerah'
                elif tb_.startswith('JABATAN PELAJARAN'): jenis = 'Jabatan Pelajaran'
                else: jenis = 'Lain-lain'
                out.append(dict(jawatan=norm(j.group(1)) if j else None, gred=norm(g.group(1)) if g else None, negeri=st[-1] if st else 'TIADA MAKLUMAT', jenis_tempat_bertugas=jenis))
    if not out: unread.append((fn, 'tiada baris peserta BK64'))
    return out

# ----------------------------------------------------------------------------- pengumpulan masukan
def kumpul_pdf(masukan, log=print):
    """masukan: senarai laluan fail/folder/zip. Pulangkan senarai (nama_relatif, bytes)."""
    out = []
    def tambah_zip(zf, awalan):
        for zi in zf.infolist():
            if zi.is_dir() or zi.filename.startswith('__MACOSX'): continue
            if zi.filename.lower().endswith('.pdf'): out.append((awalan + zi.filename, zf.read(zi)))
            elif zi.filename.lower().endswith('.zip'): tambah_zip(zipfile.ZipFile(io.BytesIO(zf.read(zi))), awalan + zi.filename + '/')
    for m in masukan:
        if os.path.isdir(m):
            for d, _, fs in os.walk(m):
                for f in sorted(fs):
                    p = os.path.join(d, f)
                    if f.lower().endswith('.pdf'): out.append((os.path.relpath(p, os.path.dirname(os.path.abspath(m))), open(p, 'rb').read()))
                    elif f.lower().endswith('.zip'): tambah_zip(zipfile.ZipFile(p), os.path.relpath(p, m) + '/')
        elif m.lower().endswith('.zip'): tambah_zip(zipfile.ZipFile(m), '')
        elif m.lower().endswith('.pdf'): out.append((os.path.basename(m), open(m, 'rb').read()))
        else: log(f'Dilangkau (bukan PDF/ZIP/folder): {m}')
    return out

# ----------------------------------------------------------------------------- penyatuan -> dataset
def seragam(t):
    t = re.sub(r'\s+', ' ', t or '').strip()
    m = re.match(r'^(\d+(?:\.\d+)*)[.,]?\s*:?\s*(.*)$', t)
    if m: t = f'{m.group(1)} {m.group(2)}'
    return re.sub(r'\s+:', ':', t)

SUB = {**{f'W{i}': 'Bilik Latihan' for i in (1, 2, 3)}, **{f'W{i}': 'Pusat Sumber' for i in (4, 5, 6)}, **{f'W{i}': 'Kemudahan/Perkhidmatan ICT' for i in (7, 8)},
       **{f'W{i}': 'Tandas' for i in (9, 10)}, **{f'W{i}': 'Tempat Penginapan' for i in (11, 12, 13, 14)}, **{f'W{i}': 'Sajian' for i in (15, 16, 17, 18, 19)},
       **{f'W{i}': 'Kemudahan/Perkhidmatan Pengangkutan IAB' for i in (20, 21)}, **{f'W{i}': 'Pengurus Kursus' for i in (22, 23, 24)}}
SUBORD = {v: i for i, v in enumerate(dict.fromkeys(SUB.values()), 1)}

def tentukan_fasa(rekod, peta=None):
    """Kumpulkan rekod mengikut program (tarikh mula) dan tentukan nombor fasa. Pulangkan (program, amaran, ralat).
    Keutamaan: (1) --fasa/pilihan pengguna, (2) teks 'Fasa N' dalam BK38, (3) 'Fasa N' pada nama folder/fail (undi majoriti),
    (4) susunan tarikh mula - hanya sebagai CADANGAN dan ditanda 'andaian'."""
    peta = peta or {}; amaran = []; ralat = []
    prog = collections.OrderedDict()
    for r in sorted((r for r in rekod if r['mula']), key=lambda r: (r['mula'], natkey(r['nama']))):
        prog.setdefault(r['mula'], dict(mula=r['mula'], rekod=[], kod=None))['rekod'].append(r)
    for p in prog.values():
        p['kod'] = next((r['kod'] for r in p['rekod'] if r['kod']), None)
    tanpa = [r for r in rekod if not r['mula'] and r['jenis'] in ('BK15a', 'BK15b', 'BK15c', 'BK15d', 'BK30', 'BK38', 'BK64')]
    # fasa setiap program
    for p in prog.values():
        sbr = None; asas = None
        if p['mula'] in peta: sbr, asas = peta[p['mula']], 'pilihan pengguna'
        else:
            t38 = [r['fasa_teks'] for r in p['rekod'] if r['jenis'] == 'BK38' and r.get('fasa_teks')]
            hints = collections.Counter(r['hint'] for r in p['rekod'] if r['hint'])
            if t38: sbr, asas = t38[0], 'teks "Fasa N" dalam BK38'
            elif hints: sbr, asas = hints.most_common(1)[0][0], 'nama folder/fail'
        p['fasa'], p['asas'] = sbr, asas
    # cadangan mengikut susunan tarikh untuk yang tiada asas
    nyata = [p for p in prog.values() if p['fasa'] is None]
    if nyata:
        diguna = {p['fasa'] for p in prog.values() if p['fasa']}
        bebas = [n for n in (1, 2, 3) if n not in diguna]
        for p, n in zip(nyata, bebas):
            p['fasa'], p['asas'], p['andaian'] = n, 'cadangan sahaja (susunan tarikh mula) - belum disahkan', True
            ralat.append(f'Program bermula {p["mula"]}: fasa tidak dinyatakan dalam PDF atau nama folder. Cadangan mengikut susunan tarikh ialah Fasa {n}, tetapi belum disahkan. Tentukan dengan --fasa {p["mula"]}=N (antara muka web: pilih fasa pada jadual program).')
        for p in nyata[len(bebas):]: p['fasa'] = None; ralat.append(f'Program bermula {p["mula"]}: tidak dapat menentukan fasa. Gunakan --fasa {p["mula"]}=N.')
    cnt = collections.Counter(p['fasa'] for p in prog.values() if p['fasa'])
    for n, c in cnt.items():
        if c > 1: ralat.append(f'Lebih daripada satu program dikesan sebagai Fasa {n} (' + ', '.join(p['mula'] for p in prog.values() if p['fasa'] == n) + '). Satu fail muat naik hanya menerima satu kumpulan; asingkan PDF mengikut kumpulan atau betulkan dengan --fasa.')
    for p in prog.values():
        if p['fasa'] and p['fasa'] not in (1, 2, 3): ralat.append(f'Program {p["mula"]}: fasa {p["fasa"]} tidak sah (hanya Fasa 1, 2 atau 3).')
    # fail tanpa tarikh: pada hint fasa, atau program tunggal
    for r in tanpa:
        tgt = [p for p in prog.values() if r['hint'] and p['fasa'] == r['hint']]
        if not tgt and len(prog) == 1: tgt = list(prog.values()); amaran.append(f'{r["nama"]} ({r["jenis"]}): tarikh tidak dapat dibaca; dipadankan dengan satu-satunya program.')
        if tgt: tgt[0]['rekod'].append(r)
        else: ralat.append(f'{r["nama"]} ({r["jenis"]}): tidak dapat dipadankan dengan mana-mana program (tarikh dan fasa tidak dikesan).')
    return prog, amaran, ralat

def bina_dataset(prog, kumpulan, tahun):
    D = {n: [] for n in NAMA}; amaran = []; F = lambda n: f'Fasa {n}'
    aturan = sorted((p for p in prog.values() if p['fasa'] in (1, 2, 3)), key=lambda p: p['fasa'])
    no = dict(ts=0, tk=0, pr=0)
    for p in aturan:
        f = p['fasa']; Fl = F(f)
        by = collections.defaultdict(list)
        for r in sorted(p['rekod'], key=lambda r: natkey(r['nama'])): by[r['jenis']].append(r)
        for j in ('BK15b', 'BK30', 'BK38'):
            if len(by[j]) > 1: amaran.append(f'{Fl}: {len(by[j])} fail {j} dikesan; hanya {by[j][0]["nama"]} digunakan.')
            if not by[j]: amaran.append(f'{Fl}: tiada fail {j}; medan berkaitan dibiarkan kosong.')
        b15 = by['BK15b'][0]['data'] if by['BK15b'] else None
        b30 = by['BK30'][0]['data'] if by['BK30'] else None
        b38 = by['BK38'][0]['data'] if by['BK38'] else None
        s15 = [r['data'] for r in by['BK15a']]
        # ---- sesi
        for x in s15:
            x['_k'] = ns(x.get('pensyarah')) + '|' + ns(x.get('topik'))
        pool = collections.defaultdict(list)
        for x in s15: pool[x['_k']].append(x)
        sesi = []; pakai = set()
        if b30 and b30['p']:
            for q in b30['p']:
                k = ns(q['pensyarah']) + '|' + ns(q['tajuk']); x = pool[k].pop(0) if pool.get(k) else None
                if x: pakai.add(id(x))
                sesi.append((q['bil'], q, x))
            sisa = [x for x in s15 if id(x) not in pakai]
            SIG = lambda x: tuple(x.get(k) for k in ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'purata_keseluruhan', 'bil_hadir', 'bil_responden'])
            tersedia = [x for _, _, x in sesi if x]
            for x in list(sisa):
                kembar = next((y for y in tersedia if y['_k'] == x['_k'] and SIG(y) == SIG(x)), None)
                if kembar:
                    sisa.remove(x)
                    tbl = [q['pensyarah'] + ' / ' + q['tajuk'] for _, q, y in sesi if q and y is None]
                    amaran.append(f'{Fl}: {x["fail"]} mengandungi data yang SAMA dengan {kembar["fail"]} (topik, pensyarah dan semua skor) lalu dianggap pendua dan tidak dimasukkan. '
                                  + ('Jadual BK30 mempunyai baris tanpa BK15a: ' + '; '.join(tbl[:3]) + '. Semak sama ada fail ini sebenarnya milik sesi tersebut.' if tbl else ''))
            if sisa:
                amaran.append(f'{Fl}: {len(sisa)} fail BK15a tidak dipadankan dengan jadual BK30 (ditambah pada hujung): ' + '; '.join(x['fail'] for x in sisa[:5]))
                n0 = max((q['bil'] for q in b30['p']), default=0)
                for i, x in enumerate(sisa, 1): sesi.append((n0 + i, None, x))
        else:
            if s15: amaran.append(f'{Fl}: BK30 tiada/kosong; sesi dibina daripada BK15a sahaja (nombor sesi mengikut susunan nama fail).')
            for i, x in enumerate(s15, 1): sesi.append((i, None, x))
        for bil, q, x in sesi:
            t = seragam((x or {}).get('topik') if x else q['tajuk']); ph = (q or {}).get('pensyarah') or x.get('pensyarah')
            r = dict(id=f'F{f}-{bil:02d}', fasa=Fl, modul=None, topik=t, topik_asal=(q or {}).get('tajuk') or x.get('topik'), pensyarah=ph,
                     bil_hadir=x.get('bil_hadir') if x else None, bil_responden=x.get('bil_responden') if x else None)
            for k in ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'purata_pengisian', 'purata_keupayaan', 'purata_keseluruhan']: r[k] = x.get(k) if x else None
            r['peratus_keseluruhan'] = x.get('pct_keseluruhan') if x else (q or {}).get('skor')
            r['peratus_pengisian'] = x.get('pct_pengisian') if x else None; r['peratus_keupayaan'] = x.get('pct_keupayaan') if x else None
            r['tahap'] = x.get('tahap_keseluruhan') if x else None; r['skor_bk30'] = (q or {}).get('skor')
            r['sumber_skor'] = 'BK15a' if x else 'BK30 (fail BK15a tiada)'; r['fail_sumber'] = x['fail'] if x else q['fail']
            r['catatan'] = ('Bil Responden melebihi Bil Peserta Hadir' if x and x.get('bil_responden') and x.get('bil_hadir') and x['bil_responden'] > x['bil_hadir']
                            else ('Fail BK15a tiada; butiran L1-L8 tiada' if not x else ''))
            if x and q is None: r['catatan'] = (r['catatan'] + '; ' if r['catatan'] else '') + 'Tiada dalam jadual BK30'
            D['sesi'].append(r)
        ses_f = [r for r in D['sesi'] if r['fasa'] == Fl]
        # ---- fasa
        a = b38 or {}; b = b30['h'] if b30 else {}; k = b15['h'] if b15 else {}
        resp = a.get('bil_responden') if a.get('bil_responden') is not None else b.get('bil_responden')
        ada = bool(resp and resp > 0)
        def pick(key, bkey=None): return a.get('pct_' + key) if a.get('pct_' + key) is not None else (b.get('pct_' + (bkey or key)))
        tempoh = (a.get('tempoh') or b.get('tempoh') or '')
        r = dict(fasa=Fl, kod_program=k.get('kod_program') or p['kod'], tempoh=tempoh.replace(' HINGGA ', ' hingga '), tempat=b.get('tempat'),
                 bil_ditawarkan=a.get('bil_ditawarkan') if a.get('bil_ditawarkan') is not None else b.get('bil_ditawarkan'),
                 bil_hadir=a.get('bil_hadir') if a.get('bil_hadir') is not None else b.get('bil_hadir'),
                 bil_responden_sesi=resp, bil_pensyarah_terlibat=a.get('bil_pensyarah'), bil_sesi_dinilai=len(ses_f),
                 bil_pensyarah_dinilai=len({s['pensyarah'] for s in ses_f}), ada_data_sesi='Ya' if ada else 'Tidak')
        for kk in ['keseluruhan', 'pengisian', 'keupayaan', 'komposit']:
            v = a.get('pct_' + kk) if a.get('pct_' + kk) is not None else (b.get('pct_' + kk) if kk in ('pengisian', 'keupayaan') else None)
            st = a.get('status_' + kk)
            r[kk] = v if ada else None; r[kk + '_dilaporkan'] = v; r['status_' + kk] = st if ada else 'Tiada data'; r['status_' + kk + '_dilaporkan'] = st
        r['khidmat_sokongan'] = pick('khidmat_sokongan'); r['status_khidmat_sokongan'] = a.get('status_khidmat_sokongan')
        r['khidmat_sokongan_bk15b'] = k.get('prestasi_pct_header'); r['min_khidmat_sokongan_bk15b'] = k.get('min_header')
        ww = b15['w'] if b15 else []
        r['responden_item_kemudahan'] = max((w['bil_menjawab'] for w in ww if w['item'] not in ('W22', 'W23', 'W24')), default=None)
        r['responden_item_pengurus'] = max((w['bil_menjawab'] for w in ww if w['item'] in ('W22', 'W23', 'W24')), default=None)
        c_rows = [x for rr in by['BK15c'] for x in rr['data']]; d_rows = [x for rr in by['BK15d'] for x in rr['data']]
        r['respons_terbuka_sesi'] = len(c_rows); r['respons_terbuka_sokongan'] = len(d_rows)
        r['ulasan_pengurus'] = a.get('ulasan_pengurus'); r['sumber'] = ', '.join(j for j in ('BK38', 'BK30', 'BK15b') if by[j])
        D['fasa'].append(r)
        # ---- sokongan
        for w in ww:
            sa = SUB.get(w['item'], 'Lain-lain')
            D['sokongan'].append(dict(kunci=f'{Fl}|{w["item"]}', fasa=Fl, item=w['item'], no_item=int(w['item'][1:]), perkara=w['perkara'], sub_aspek=sa, no_sub_aspek=SUBORD.get(sa, 99),
                bil_menjawab=w['bil_menjawab'], bil_berpuas_hati=w['bil_dalam_kurungan'], peratus=w['peratus'] if w['bil_menjawab'] > 0 else None,
                min=w['min'] if w['bil_menjawab'] > 0 else None, min_dilaporkan=w['min'], fail_sumber=by['BK15b'][0]['fail']))
        # ---- terbuka sesi / sokongan / perubahan
        for x in c_rows:
            no['ts'] += 1; t = seragam(x['topik'])
            D['terbuka_sesi'].append(dict(no=no['ts'], fasa=Fl, modul=None, topik=t, pensyarah=x['pensyarah'], bil=x['bil'], v12=x['V12'], v13=x['V13'], tema_v13=None, sumber=f'BK15c {Fl}, hlm {x["halaman"]}'))
        for x in d_rows:
            no['tk'] += 1; D['terbuka_sokongan'].append(dict(no=no['tk'], fasa=Fl, bil=x['bil'], respons=x['respons'], fail_sumber=x['fail']))
        for rr in by['BK30'][:1]:
            for x in rr['data']['chg']:
                no['pr'] += 1; D['perubahan'].append(dict(no=no['pr'], fasa=Fl, butiran=x['perubahan'], fail_sumber=x['fail']))
        # ---- peserta (agregat)
        pe = [x for rr in by['BK64'] for x in rr['data']]
        for dim, col in [('Gred', 'gred'), ('Jawatan', 'jawatan'), ('Negeri', 'negeri'), ('Jenis tempat bertugas', 'jenis_tempat_bertugas')]:
            cnt = collections.Counter(({'gred': lambda v: v or '(kosong dalam laporan)', 'jawatan': lambda v: v or '(kosong dalam laporan)',
                                        'negeri': lambda v: '(tidak dapat ditentukan)' if v == 'TIADA MAKLUMAT' else v}.get(col, lambda v: v))(x[col]) for x in pe)
            for kat, n in cnt.items(): D['peserta'].append(dict(kunci=f'{Fl}|{dim}|{kat}', fasa=Fl, dimensi=dim, kategori=kat, bilangan=n))
        if not pe: amaran.append(f'{Fl}: tiada BK64; tab profil peserta akan kosong.')
        # topik hampir serupa dalam fasa yang sama (kemungkinan ejaan berbeza)
        tp = sorted({s['topik'] for s in ses_f} | {s['topik'] for s in D['terbuka_sesi'] if s['fasa'] == Fl})
        for i in range(len(tp)):
            for j in range(i + 1, len(tp)):
                if re.match(r'\d+(?:\.\d+)*', tp[i]) and re.match(r'\d+(?:\.\d+)*', tp[j]) and re.match(r'\d+(?:\.\d+)*', tp[i]).group() != re.match(r'\d+(?:\.\d+)*', tp[j]).group(): continue
                if difflib.SequenceMatcher(None, tp[i], tp[j]).ratio() > 0.9:
                    amaran.append(f'{Fl}: topik hampir serupa - semak ejaan: "{tp[i][:70]}" / "{tp[j][:70]}"')
    return D, amaran

def semak(D):
    """Cerminkan semakan struktur dan nilai dalam muat.js supaya fail lulus semasa dimuat naik."""
    ralat = []; amaran = []
    for n in NAMA:
        for r in D[n]:
            for k, t in SKEMA['skema'][n].items():
                v = r.get(k)
                if v is not None and t == 'n' and not isinstance(v, (int, float)): ralat.append(f'{n}.{k}: bukan nombor ({v!r})')
    fs = collections.Counter(r['fasa'] for r in D['fasa'])
    if not D['fasa']: ralat.append('Tiada baris dalam lembaran fasa.')
    if any(v > 1 for v in fs.values()): ralat.append('Fasa berulang dalam lembaran fasa.')
    ids = collections.Counter(s['id'] for s in D['sesi'])
    if any(v > 1 for v in ids.values()): ralat.append('id sesi berulang.')
    def had(l, k, lo, hi):
        for r in D[l]:
            if r.get(k) is not None and not (lo <= r[k] <= hi): ralat.append(f'{l}.{k}: {r[k]} di luar julat {lo}-{hi}'); break
    for k in ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8']: had('sesi', k, 0, 5)
    for k in ['peratus_keseluruhan', 'peratus_pengisian', 'peratus_keupayaan']: had('sesi', k, 0, 100)
    for k in ['keseluruhan', 'pengisian', 'keupayaan', 'komposit', 'khidmat_sokongan']: had('fasa', k, 0, 100)
    for r in D['fasa']:
        for k, t in SKEMA['skema']['fasa'].items():
            if not opsyen('fasa', k) and r.get(k) is None and k not in ('keseluruhan', 'pengisian', 'keupayaan', 'komposit') :
                amaran.append(f'{r["fasa"]}: medan "{k}" kosong (tidak ditemui dalam PDF).')
    return ralat, amaran

# ----------------------------------------------------------------------------- penulis xlsx
def tulis_xlsx(D, kumpulan, tahun, sasaran):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook(); ws = wb.active; ws.title = 'arahan'
    for l in ['Fail ini dijana oleh npqel_pdf_ke_xlsx.py daripada laporan PDF i-Penilaian.',
              f'Kumpulan: {kumpulan}; tahun: {tahun}. Muat naik melalui butang "Muat naik data" pada Dashboard Penilaian NPQEL.',
              'Lajur "modul" dan "tema_v13" sengaja dikosongkan; dashboard menerbitkannya daripada topik dan jawapan V13.',
              'Data peserta (lembaran peserta) ialah bilangan agregat sahaja; tiada nama, No. KP, telefon atau e-mel disimpan.',
              'Semak nilai dengan laporan PDF asal sebelum digunakan untuk keputusan rasmi.']: ws.append([l])
    ws['A1'].font = Font(bold=True, size=13); ws.column_dimensions['A'].width = 120
    m = wb.create_sheet('maklumat'); m.append(['medan', 'nilai']); m.append(['kumpulan', kumpulan]); m.append(['tahun', int(tahun)])
    for c in m[1]: c.font = Font(bold=True)
    m.column_dimensions['A'].width = 14; m.column_dimensions['B'].width = 20
    for n in NAMA:
        w = wb.create_sheet(n); cols = list(SKEMA['skema'][n].keys()); w.append(cols)
        for c in w[1]:
            c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='0B5560' if not opsyen(n, c.value) else '6B7F86'); c.alignment = Alignment(vertical='top')
        for r in D[n]: w.append([r.get(k) for k in cols])
        w.freeze_panes = 'A2'
        for i, k in enumerate(cols, 1): w.column_dimensions[w.cell(1, i).column_letter].width = min(40, max(10, len(k) + 2))
    if hasattr(sasaran, 'write'): wb.save(sasaran)
    else: wb.save(sasaran)

# ----------------------------------------------------------------------------- aliran penuh
class Sesi:
    """Menyimpan hasil bacaan PDF (bukan PDF itu sendiri) dalam ingatan."""
    def __init__(self): self.rekod = []; self.dilangkau = []; self.kunci = set(); self.hash = {}
    def tambah(self, nama, data):
        import hashlib
        h = hashlib.sha1(data).hexdigest()
        if h in self.hash:
            self.dilangkau.append((nama, f'kandungan sama dengan {self.hash[h]} (fail pendua)')); return dict(nama=nama, jenis='(pendua - dilangkau)')
        asal = nama; i = 2
        while nama in self.kunci: nama = f'{asal} #{i}'; i += 1
        self.kunci.add(nama); self.hash[h] = nama
        try: R = baca_pdf(data, nama)
        except Exception as e:
            self.dilangkau.append((nama, f'gagal dibuka: {type(e).__name__}: {e}')); return dict(nama=nama, jenis=None, ralat=str(e))
        if not R['jenis'] or R['data'] is None:
            self.dilangkau.append((nama, 'bukan borang BK15a/b/c/d, BK30, BK38 atau BK64')); R['jenis'] = R['jenis'] or '(tidak dikenali)'
        else: self.rekod.append(R)
        return dict(nama=nama, jenis=R['jenis'], mula=R['mula'])
    def ringkasan(self, peta=None):
        prog, am, ra = tentukan_fasa(self.rekod, peta)
        return [dict(mula=p['mula'], kod=p['kod'], fasa=None if p.get('andaian') else p['fasa'], cadangan=p['fasa'] if p.get('andaian') else None, asas=p['asas'], borang=dict(collections.Counter(r['jenis'] for r in p['rekod']))) for p in prog.values()], am, ra
    def jana(self, kumpulan, tahun, peta=None, sasaran=None):
        prog, am, ra = tentukan_fasa(self.rekod, peta)
        lap = dict(amaran=list(am), ralat=list(ra), dilangkau=list(self.dilangkau), tidak_dibaca=[x for r in self.rekod for x in r['tidak_dibaca']])
        if not kumpulan or not str(kumpulan).strip(): lap['ralat'].append('Nama kumpulan wajib diisi.')
        try: tahun = int(tahun)
        except Exception: lap['ralat'].append('Tahun mesti nombor (contoh 2027).'); tahun = None
        if not self.rekod: lap['ralat'].append('Tiada borang BK yang dikenali.')
        if lap['ralat']: return None, lap
        yrs = {p['mula'][:4] for p in prog.values()}
        if tahun and str(tahun) not in yrs: lap['amaran'].append(f'Tahun {tahun} tidak sepadan dengan tahun tarikh mula dalam PDF ({", ".join(sorted(yrs))}).')
        D, am2 = bina_dataset(prog, kumpulan.strip(), tahun); lap['amaran'] += am2
        r2, am3 = semak(D); lap['ralat'] += r2; lap['amaran'] += am3
        lap['ringkas'] = {n: len(D[n]) for n in NAMA}; lap['program'] = [dict(mula=p['mula'], fasa=p['fasa'], asas=p['asas'], kod=p['kod']) for p in prog.values()]
        if lap['ralat']: return None, lap
        if sasaran is None: sasaran = io.BytesIO()
        tulis_xlsx(D, kumpulan.strip(), tahun, sasaran)
        return sasaran, lap

# ----------------------------------------------------------------------------- CLI
def cetak_lap(lap, f=sys.stderr):
    if lap.get('ringkas'): print('Ringkasan baris:', ', '.join(f'{k}={v}' for k, v in lap['ringkas'].items()), file=f)
    for p in lap.get('program', []): print(f'  Program {p["mula"]} ({p["kod"]}) -> Fasa {p["fasa"]} [{p["asas"]}]', file=f)
    for t, key in (('RALAT', 'ralat'), ('AMARAN', 'amaran')):
        for x in lap.get(key, []): print(f'{t}: {x}', file=f)
    for n, x in lap.get('dilangkau', []): print(f'DILANGKAU: {n} - {x}', file=f)
    for n, x in lap.get('tidak_dibaca', []): print(f'TIDAK DIBACA: {n} - {x}', file=f)

def main(argv=None):
    ap = argparse.ArgumentParser(description='Tukar PDF i-Penilaian NPQEL kepada satu fail xlsx format muat naik dashboard.')
    ap.add_argument('masukan', nargs='*', help='fail PDF, folder, atau ZIP (boleh banyak)')
    ap.add_argument('-k', '--kumpulan', help='nama kumpulan, contoh SRGH6'); ap.add_argument('-t', '--tahun', help='tahun ambilan, contoh 2027')
    ap.add_argument('-o', '--output', help='fail xlsx hasil'); ap.add_argument('--fasa', action='append', default=[], metavar='TARIKH_MULA=N', help='tentukan fasa program, contoh 2027-03-02=1')
    ap.add_argument('--web', action='store_true', help='mulakan antara muka web setempat'); ap.add_argument('--port', type=int, default=8765)
    a = ap.parse_args(argv)
    if a.web: return web(a.port)
    if not a.masukan or not a.kumpulan or not a.tahun: ap.error('masukan, --kumpulan dan --tahun diperlukan (atau gunakan --web)')
    peta = {}
    for x in a.fasa:
        m = re.fullmatch(r'(\d{4}-\d{2}-\d{2})=([123])', x)
        if not m: ap.error('--fasa mesti dalam bentuk TARIKH_MULA=N, contoh 2027-03-02=1')
        peta[m.group(1)] = int(m.group(2))
    s = Sesi(); fail = kumpul_pdf(a.masukan)
    print(f'{len(fail)} PDF ditemui. Membaca...', file=sys.stderr)
    for i, (n, b) in enumerate(fail, 1):
        r = s.tambah(n, b); print(f'  [{i}/{len(fail)}] {n} -> {r.get("jenis")}', file=sys.stderr)
    out = a.output or f'Data_{a.kumpulan}_{a.tahun}_format_muat_naik.xlsx'
    buf, lap = s.jana(a.kumpulan, a.tahun, peta, out); cetak_lap(lap)
    if buf is None: print('Fail xlsx TIDAK dijana kerana ada ralat di atas.', file=sys.stderr); return 1
    print(f'Siap: {out}', file=sys.stderr); return 0

# ----------------------------------------------------------------------------- antara muka web setempat
HTML = r'''<!doctype html><html lang="ms"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PDF ke Excel - Data Penilaian NPQEL</title>
<style>
:root{--bg:#f4f7f7;--fg:#12292e;--mut:#5a6f74;--acc:#0b5560;--card:#fff;--line:#d3dfe0;--bad:#a4262c;--warn:#8a5a00;--ok:#16603a}
@media(prefers-color-scheme:dark){:root{--bg:#0e1b1e;--fg:#e6f0f1;--mut:#9ab0b4;--acc:#5fc3cf;--card:#15272b;--line:#2a4247;--bad:#ff8a8f;--warn:#f0c060;--ok:#6fd49b}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.5 system-ui,Segoe UI,Roboto,sans-serif}
main{max-width:820px;margin:0 auto;padding:24px 16px 64px}h1{font-size:1.5rem;margin:0 0 4px}p.s{color:var(--mut);margin:0 0 20px}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin:0 0 16px}h2{font-size:1.05rem;margin:0 0 10px}
#drop{border:2px dashed var(--line);border-radius:10px;padding:28px 12px;text-align:center;color:var(--mut)}#drop.on{border-color:var(--acc);background:color-mix(in srgb,var(--acc) 8%,transparent)}
button,label.b{font:inherit;background:var(--acc);color:#fff;border:0;border-radius:8px;padding:8px 14px;cursor:pointer;display:inline-block}
@media(prefers-color-scheme:dark){button,label.b{color:#06242a}}
button.g,label.g{background:transparent;color:var(--acc);border:1px solid var(--acc)}button:disabled{opacity:.5;cursor:not-allowed}
:focus-visible{outline:3px solid var(--acc);outline-offset:2px}
input[type=text],input[type=number],select{font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg)}
.r{display:flex;gap:12px;flex-wrap:wrap;align-items:end}.r>div{display:flex;flex-direction:column;gap:3px}
table{border-collapse:collapse;width:100%;font-size:.92rem}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
progress{width:100%;height:10px}.bad{color:var(--bad)}.warn{color:var(--warn)}.ok{color:var(--ok)}ul{margin:6px 0 0 18px;padding:0}
.k{color:var(--mut);font-size:.9rem}
</style></head><body><main>
<h1>PDF ke Excel: Data Penilaian NPQEL</h1>
<p class="s">Muat naik semua PDF i-Penilaian bagi satu kumpulan. Satu fail xlsx akan dijana untuk dimuat naik ke dashboard. Pemprosesan berlaku di komputer ini sahaja.</p>
<section><h2>1. Pilih PDF</h2>
<div id="drop" tabindex="0" role="button" aria-label="Seret fail atau folder ke sini">Seret fail PDF, folder atau ZIP ke sini<br><span class="k">atau</span><br>
<label class="b" style="margin-top:8px">Pilih fail PDF/ZIP<input id="fFail" type="file" accept=".pdf,.zip" multiple hidden></label>
<label class="b g" style="margin:8px 0 0 6px">Pilih folder<input id="fFolder" type="file" webkitdirectory multiple hidden></label></div>
<progress id="pg" max="1" value="0" hidden></progress><div id="stat" class="k" role="status" aria-live="polite"></div></section>
<section id="sP" hidden><h2>2. Semak program yang dikesan</h2><div class="k">Fasa ditentukan daripada kandungan PDF atau nama folder. Tukar jika tidak betul.</div><div id="progs"></div><div id="msgP"></div></section>
<section id="sG" hidden><h2>3. Maklumat kumpulan dan jana fail</h2>
<div class="r"><div><label for="kump">Kumpulan</label><input id="kump" type="text" placeholder="contoh: SRGH6" autocomplete="off"></div>
<div><label for="thn">Tahun ambilan</label><input id="thn" type="number" min="2000" max="2100" placeholder="2027"></div>
<button id="jana">Jana fail xlsx</button><button id="semula" class="g">Mula semula</button></div><div id="lap" role="status" aria-live="polite"></div></section>
</main><script>
const $=i=>document.getElementById(i),E=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let peta={},fasaAkhir=[];
async function hantar(files){
  const L=[...files].filter(f=>/\.(pdf|zip)$/i.test(f.name));
  if(!L.length){$('stat').textContent='Tiada fail PDF atau ZIP dipilih.';return}
  $('pg').hidden=false;$('pg').max=L.length;$('pg').value=0;
  for(let i=0;i<L.length;i++){const f=L[i],nama=f.webkitRelativePath||f.name;$('stat').textContent=`Membaca ${i+1}/${L.length}: ${nama}`;
    try{const r=await fetch('/api/fail?nama='+encodeURIComponent(nama),{method:'POST',body:await f.arrayBuffer()});paparProg(await r.json())}catch(e){$('stat').textContent='Ralat sambungan ke pelayan setempat: '+e.message;return}
    $('pg').value=i+1}
  $('stat').textContent=`Selesai membaca ${L.length} fail.`;$('pg').hidden=true}
function paparProg(d){$('sP').hidden=$('sG').hidden=false;fasaAkhir=d.program;
  $('progs').innerHTML='<table><thead><tr><th>Tarikh mula</th><th>Kod program</th><th>Borang dikesan</th><th>Fasa</th><th>Asas</th></tr></thead><tbody>'+d.program.map(p=>{
    const f=peta[p.mula]||p.fasa||'';return `<tr><td>${E(p.mula)}</td><td>${E(p.kod||'-')}</td><td>${Object.entries(p.borang).map(([k,v])=>k+'&times;'+v).join(', ')}</td>
    <td><select data-m="${E(p.mula)}" aria-label="Fasa bagi program ${E(p.mula)}"><option value="">?</option>${[1,2,3].map(n=>`<option value="${n}"${f==n?' selected':''}>Fasa ${n}</option>`).join('')}</select></td><td class="k">${E(peta[p.mula]?'pilihan pengguna':(p.asas||'tidak diketahui'))}${p.cadangan&&!peta[p.mula]?` (cadangan: Fasa ${p.cadangan})`:''}</td></tr>`}).join('')+'</tbody></table>';
  document.querySelectorAll('#progs select').forEach(s=>s.onchange=async()=>{if(s.value)peta[s.dataset.m]=+s.value;else delete peta[s.dataset.m];const r=await fetch('/api/ringkas',{method:'POST',body:JSON.stringify({fasa:peta})});const d=await r.json();const m=[...d.ralat.map(x=>`<li class="bad">${E(x)}</li>`),...d.amaran.map(x=>`<li class="warn">${E(x)}</li>`)];$('msgP').innerHTML=m.length?'<ul>'+m.join('')+'</ul>':'<p class="ok">Fasa semua program telah ditentukan.</p>'});
  const m=[...d.ralat.map(x=>`<li class="bad">${E(x)}</li>`),...d.amaran.map(x=>`<li class="warn">${E(x)}</li>`)];$('msgP').innerHTML=m.length?'<ul>'+m.join('')+'</ul>':''}
function muatDrop(e){e.preventDefault();$('drop').classList.remove('on');const it=[...(e.dataTransfer.items||[])].map(i=>i.webkitGetAsEntry&&i.webkitGetAsEntry()).filter(Boolean);
  if(!it.length){hantar(e.dataTransfer.files);return}
  const out=[];const walk=async(en,p)=>{if(en.isFile)await new Promise(ok=>en.file(f=>{Object.defineProperty(f,'webkitRelativePath',{value:p+f.name});out.push(f);ok()}));else{const rd=en.createReader();let b;do{b=await new Promise(ok=>rd.readEntries(ok));for(const c of b)await walk(c,p+en.name+'/')}while(b.length)}};
  Promise.all(it.map(en=>walk(en,''))).then(()=>hantar(out))}
$('drop').ondragover=e=>{e.preventDefault();$('drop').classList.add('on')};$('drop').ondragleave=()=>$('drop').classList.remove('on');$('drop').ondrop=muatDrop;
$('fFail').onchange=e=>hantar(e.target.files);$('fFolder').onchange=e=>hantar(e.target.files);
$('semula').onclick=async()=>{await fetch('/api/semula',{method:'POST'});peta={};$('sP').hidden=$('sG').hidden=true;$('lap').innerHTML='';$('stat').textContent='';$('progs').innerHTML=''};
$('jana').onclick=async()=>{$('jana').disabled=true;$('lap').innerHTML='<p class="k">Menjana...</p>';
  try{const r=await fetch('/api/jana',{method:'POST',body:JSON.stringify({kumpulan:$('kump').value,tahun:$('thn').value,fasa:peta})}),d=await r.json();
  let h='';if(d.ringkas)h+='<p>Baris dijana: '+Object.entries(d.ringkas).map(([k,v])=>`${k} ${v}`).join(', ')+'</p>';
  const li=(a,c)=>a&&a.length?`<ul>${a.map(x=>`<li class="${c}">${E(Array.isArray(x)?x.join(' - '):x)}</li>`).join('')}</ul>`:'';
  if(d.ralat&&d.ralat.length)h+='<p class="bad"><b>Fail tidak dijana.</b> Betulkan ralat berikut:</p>'+li(d.ralat,'bad');
  if(d.xlsx){const b=Uint8Array.from(atob(d.xlsx),c=>c.charCodeAt(0)),u=URL.createObjectURL(new Blob([b],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})),n=`Data_${$('kump').value.trim()}_${$('thn').value}_format_muat_naik.xlsx`;
    h+=`<p class="ok"><b>Fail sedia.</b> <a href="${u}" download="${E(n)}">Muat turun ${E(n)}</a></p>`}
  if(d.amaran&&d.amaran.length)h+='<p><b>Amaran (semak sebelum guna):</b></p>'+li(d.amaran,'warn');
  if(d.tidak_dibaca&&d.tidak_dibaca.length)h+='<p><b>Bahagian PDF tidak dapat dibaca:</b></p>'+li(d.tidak_dibaca,'warn');
  if(d.dilangkau&&d.dilangkau.length)h+='<p><b>Fail dilangkau:</b></p>'+li(d.dilangkau,'warn');
  $('lap').innerHTML=h}catch(e){$('lap').innerHTML='<p class="bad">Ralat: '+E(e.message)+'</p>'}$('jana').disabled=false};
</script></body></html>
'''
SES = Sesi(); KUNCI = threading.Lock()
class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _j(self, o, code=200):
        b = json.dumps(o, ensure_ascii=False).encode(); self.send_response(code); self.send_header('Content-Type', 'application/json; charset=utf-8'); self.send_header('Content-Length', str(len(b))); self.end_headers(); self.wfile.write(b)
    def do_GET(self):
        if urlparse(self.path).path == '/':
            b = HTML.encode(); self.send_response(200); self.send_header('Content-Type', 'text/html; charset=utf-8'); self.send_header('Content-Length', str(len(b))); self.end_headers(); self.wfile.write(b)
        else: self._j({'ralat': 'tiada'}, 404)
    def do_POST(self):
        global SES
        u = urlparse(self.path); q = parse_qs(u.query); n = int(self.headers.get('Content-Length') or 0); body = self.rfile.read(n)
        if self.headers.get('Host', '').split(':')[0] not in ('127.0.0.1', 'localhost'): return self._j({'ralat': 'hos ditolak'}, 403)
        with KUNCI:
            if u.path == '/api/fail':
                nama = unquote(q.get('nama', ['fail.pdf'])[0]); hasil = []
                if nama.lower().endswith('.zip'):
                    try:
                        for nm, b in kumpul_pdf_zip(body, nama): hasil.append(SES.tambah(nm, b))
                    except Exception as e: hasil.append(dict(nama=nama, jenis=None, ralat=str(e)))
                else: hasil.append(SES.tambah(nama, body))
                ring, am, ra = SES.ringkasan(); return self._j(dict(hasil=hasil, jumlah=len(SES.rekod), program=ring, amaran=am, ralat=ra))
            if u.path == '/api/ringkas':
                d = json.loads(body or b'{}'); peta = {k: int(v) for k, v in (d.get('fasa') or {}).items() if str(v) in ('1', '2', '3')}
                ring, am, ra = SES.ringkasan(peta); return self._j(dict(jumlah=len(SES.rekod), program=ring, amaran=am, ralat=ra))
            if u.path == '/api/semula': SES = Sesi(); return self._j({'ok': True})
            if u.path == '/api/jana':
                d = json.loads(body or b'{}'); peta = {k: int(v) for k, v in (d.get('fasa') or {}).items() if str(v) in ('1', '2', '3')}
                buf, lap = SES.jana(d.get('kumpulan'), d.get('tahun'), peta)
                if buf is not None: lap['xlsx'] = base64.b64encode(buf.getvalue()).decode()
                return self._j(lap)
        self._j({'ralat': 'tiada'}, 404)
def kumpul_pdf_zip(data, nama):
    tmp = []
    def rec(zf, aw):
        for zi in zf.infolist():
            if zi.is_dir() or zi.filename.startswith('__MACOSX'): continue
            if zi.filename.lower().endswith('.pdf'): tmp.append((aw + zi.filename, zf.read(zi)))
            elif zi.filename.lower().endswith('.zip'): rec(zipfile.ZipFile(io.BytesIO(zf.read(zi))), aw + zi.filename + '/')
    rec(zipfile.ZipFile(io.BytesIO(data)), os.path.splitext(nama)[0] + '/'); return tmp
def web(port):
    srv = ThreadingHTTPServer(('127.0.0.1', port), H)
    print(f'Buka http://127.0.0.1:{port} dalam pelayar. Tekan Ctrl+C untuk berhenti.', file=sys.stderr)
    try: srv.serve_forever()
    except KeyboardInterrupt: print('\nBerhenti.', file=sys.stderr)
    return 0

if __name__ == '__main__': sys.exit(main())
