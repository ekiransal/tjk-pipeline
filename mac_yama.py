#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""TARIH -> KOSU VIDEOSU yamasi v2 (Mac tarafi).

v2 yeniligi: tarihin yanina o kosudaki AT NO parantez icinde eklenir: 16.05.2026 (7) >
Hem sifirdan kurulumu hem v1 -> v2 yukseltmeyi bilir. Yedekler: .bak_video / .bak_video2

Uc dosyayi gunceller:
 1) web/parse_mockup.py   : detay satirina 43-44. kolonlar (video linki + at no)
 2) web/prototip_uret.py  : Tarih hucresi tiklanabilir + (at no) rozeti (+CSS)
 3) GUNLUK_CALISTIR.sh    : tjk_yeni_yer.py'den sonra video_link_ekle.py calistir

Calistirma: cd ~/Desktop/tjk && python3 mac_yama.py
"""
import io, os, shutil, sys, ast

KOK = os.path.dirname(os.path.abspath(__file__))
n = {"parse": 0, "proto_var": 0, "proto_tarih": 0, "proto_td": 0,
     "proto_css": 0, "baslik": 0, "kalin": 0, "yer": 0, "seyir": 0, "gunluk": 0}
hata = []


def oku(yol):
    return io.open(yol, encoding="utf-8").read()


def yaz(yol, icerik, ek=".bak_video"):
    if not os.path.exists(yol + ek):
        shutil.copy(yol, yol + ek)
    else:
        shutil.copy(yol, yol + ".bak_video2")
    io.open(yol, "w", encoding="utf-8").write(icerik)


# ---------- 1) parse_mockup.py ----------
p1 = os.path.join(KOK, "web", "parse_mockup.py")
s = oku(p1)
hedef = "for c in range(1, 45)])"
if hedef in s:
    n["parse"] = 1
else:
    d = None
    for eski in ("for c in range(1, 44)])", "for c in range(1, 43)])"):
        if eski in s:
            d = s.replace(eski, hedef, 1)
            break
    if d is None:
        hata.append("parse_mockup: range satiri bulunamadi")
    else:
        ast.parse(d)
        yaz(p1, d)
        n["parse"] = 1

# ---------- 2) prototip_uret.py ----------
p2 = os.path.join(KOK, "web", "prototip_uret.py")
s = oku(p2)
degisti = False

# 2a. hucre dongusune _vhtml degiskeni
a_eski = 'let v=String(r[c]??"").trim();\n    let cls="";'
a_yeni = 'let v=String(r[c]??"").trim();\n    let cls="";\n    let _vhtml=null;'
if "let _vhtml=null;" in s:
    n["proto_var"] = 1
elif a_eski in s:
    s = s.replace(a_eski, a_yeni, 1)
    n["proto_var"] = 1
    degisti = True
else:
    hata.append("prototip: hucre dongusu bulunamadi")

# 2b. Tarih hucresi (v2: at no rozetli)
orijinal = ('else if(c===FK.tarih){ v=tarihGoster(v); cls="num drc"; }'
            '            // tarih: Derece gibi KOYU')
v1_blok = ('else if(c===FK.tarih){ v=tarihGoster(v); cls="num drc";\n'
           '      const _vl=String(r[42]??"").trim();\n'
           '      if(/^https:\\/\\/[a-z0-9.-]*tjk\\.org\\//i.test(_vl))\n'
           '        _vhtml=\'<a class="tlink" href="\'+esc(_vl)+\'" target="_blank" '
           'rel="noopener" title="Ko\\u015fuyu izle">\'+esc(fmt(v))+\''
           '<span class="tply">\\u25b6</span></a>\';\n'
           '    }            // tarih: Derece gibi KOYU (linkliyse tiklaninca video)')
v2_blok = ('else if(c===FK.tarih){ v=tarihGoster(v); cls="num drc";\n'
           '      const _vl=String(r[42]??"").trim();\n'
           '      const _no=String(r[43]??"").trim().replace(/\\.0$/,"");\n'
           '      if(/^https:\\/\\/[a-z0-9.-]*tjk\\.org\\//i.test(_vl))\n'
           '        _vhtml=\'<a class="tlink" href="\'+esc(_vl)+\'" target="_blank" '
           'rel="noopener" title="Ko\\u015fuyu izle\'+'
           '(_no?\' \\u2014 bu ko\\u015fuda \'+esc(_no)+\' numarayd\\u0131\':\'\')+\'">\''
           '+esc(fmt(v))+(_no?\' <span class="tno">(\'+esc(_no)+\')</span>\':\'\')+'
           '\'<span class="tply">\\u25b6</span></a>\';\n'
           '    }            // tarih: Derece gibi KOYU (video + o kosudaki at no)')
if "const _no=" in s:
    n["proto_tarih"] = 1
elif v1_blok in s:
    s = s.replace(v1_blok, v2_blok, 1)
    n["proto_tarih"] = 1
    degisti = True
elif orijinal in s:
    s = s.replace(orijinal, v2_blok, 1)
    n["proto_tarih"] = 1
    degisti = True
else:
    hata.append("prototip: Tarih hucresi bulunamadi (ne orijinal ne v1)")

# 2c. td ciktisinda _vhtml kullan
c_eski = ('return `<td class="${cls}${mg2}" data-c="${c}" '
          'data-v="${esc(dv)}">${esc(fmt(v))}</td>`;')
c_yeni = ('return `<td class="${cls}${mg2}" data-c="${c}" '
          'data-v="${esc(dv)}">${_vhtml||esc(fmt(v))}</td>`;')
if "${_vhtml||esc(fmt(v))}" in s:
    n["proto_td"] = 1
elif c_eski in s:
    s = s.replace(c_eski, c_yeni, 1)
    n["proto_td"] = 1
    degisti = True
else:
    hata.append("prototip: td satiri bulunamadi")

# 2d. CSS (v2: .tno rozeti dahil)
css_anchor = ".detay tr:nth-child(even) td{background:#fafbfc}"
tply_kural = ".tlink .tply{font-size:8px;color:#1e6f5c;margin-left:2px}"
tno_kural = ".tlink .tno{font-size:9px;color:#1e6f5c;font-weight:700;margin-left:2px}"
if ".tno{" in s.replace(" ", ""):
    n["proto_css"] = 1
elif tply_kural in s:                      # v1 css var -> .tno ekle
    s = s.replace(tply_kural, tply_kural + "\n  " + tno_kural, 1)
    n["proto_css"] = 1
    degisti = True
elif css_anchor in s:                      # hic css yok -> hepsini ekle
    s = s.replace(css_anchor, css_anchor +
                  "\n  .tlink{color:inherit;text-decoration:none;"
                  "border-bottom:1px dotted #7aa08f;cursor:pointer}"
                  "\n  .tlink:hover{color:#1e6f5c;border-bottom-color:#1e6f5c}"
                  "\n  " + tply_kural + "\n  " + tno_kural, 1)
    n["proto_css"] = 1
    degisti = True
else:
    hata.append("prototip: CSS demiri bulunamadi")

# 2e. BASLIK: Tarih -> "Tarih (At No)"
e_eski = 'let ad=c===seyA?"Yarıştaki Seyri":(basliklar[c]??("K"+(c+1)));'
e_yeni = (e_eski + '\n    if(c===FK.tarih) ad="Tarih (At No)";')
if 'ad="Tarih (At No)"' in s:
    n["baslik"] = 1
elif e_eski in s:
    s = s.replace(e_eski, e_yeni, 1)
    n["baslik"] = 1
    degisti = True
else:
    hata.append("prototip: baslik satiri bulunamadi")

# 2f. KACINCI OLDU kalin (drc = font-weight 800)
f1_eski = 'c===19){ cls="orta";'
f2_eski = 'c===18){ cls="orta";'
f_say = 0
if 'c===19){ cls="orta drc";' in s:
    f_say += 1
elif f1_eski in s:
    s = s.replace(f1_eski, 'c===19){ cls="orta drc";', 1)
    f_say += 1
    degisti = True
if 'c===18){ cls="orta drc";' in s:
    f_say += 1
elif f2_eski in s:
    s = s.replace(f2_eski, 'c===18){ cls="orta drc";', 1)
    f_say += 1
    degisti = True
if f_say == 2:
    n["kalin"] = 1
else:
    hata.append("prototip: kacinci hucre satirlari eksik (%d/2)" % f_say)

# 2g. KACINCI OLDU kolonu Tarih'in hemen sagina
g2_eski = ('const zincir=[...domList.filter(c=>!domKilo.has(c)), kacinci, seyA];'
           '   // v110: Kaçıncı Oldu önce, Seyir sonra')
g2_yeni = ('tasi(kacinci, FK.drc);   // Kaçıncı -> Tarih\'in hemen sağı\n'
           '    const zincir=[...domList.filter(c=>!domKilo.has(c)), seyA];'
           '   // Derece\'den sonra: P50/P66/P75 + Seyir')
if "tasi(kacinci, FK.drc);" in s:
    n["yer"] = 1
elif g2_eski in s:
    s = s.replace(g2_eski, g2_yeni, 1)
    n["yer"] = 1
    degisti = True
else:
    hata.append("prototip: zincir satiri bulunamadi")

# 2h. SEYIR BASLIGI: "Yarıştaki Seyri" -> "O Tarihteki Seyri"
h1_eski = 'c===seyA?"Yarıştaki Seyri"'
h1_yeni = 'c===seyA?"O Tarihteki Seyri"'
h2_eski = '(c===seyA)?"Yarıştaki<br>Seyri"'
h2_yeni = '(c===seyA)?"O Tarihteki<br>Seyri"'
h_say = 0
if h1_yeni in s:
    h_say += 1
elif h1_eski in s:
    s = s.replace(h1_eski, h1_yeni, 1)
    h_say += 1
    degisti = True
if h2_yeni in s:
    h_say += 1
elif h2_eski in s:
    s = s.replace(h2_eski, h2_yeni, 1)
    h_say += 1
    degisti = True
if h_say == 2:
    n["seyir"] = 1
else:
    hata.append("prototip: seyir basligi satirlari eksik (%d/2)" % h_say)

if degisti and not hata:
    ast.parse(s)
    yaz(p2, s)

# ---------- 3) GUNLUK_CALISTIR.sh ----------
p3 = os.path.join(KOK, "GUNLUK_CALISTIR.sh")
s = oku(p3)
g_eski = "caffeinate -i python3 tjk_yeni_yer.py"
g_ek = (g_eski + "\n\n"
        'echo ""\n'
        'echo "============ 4.2/5  VIDEO LINKLERI (tarih tiklanabilir) ======="\n'
        "python3 video_link_ekle.py "
        '|| echo "UYARI: video linkleri eklenemedi (pipeline devam ediyor)"')
if "video_link_ekle.py" in s:
    n["gunluk"] = 1
elif g_eski in s:
    s = s.replace(g_eski, g_ek, 1)
    yaz(p3, s)
    n["gunluk"] = 1
else:
    hata.append("GUNLUK_CALISTIR: tjk_yeni_yer satiri bulunamadi")

# ---------- rapor ----------
print("PARSE:%d PROTO:%d%d%d%d BASLIK:%d KALIN:%d YER:%d SEYIR:%d GUNLUK:%d" %
      (n["parse"], n["proto_var"], n["proto_tarih"], n["proto_td"],
       n["proto_css"], n["baslik"], n["kalin"], n["yer"], n["seyir"], n["gunluk"]))
if hata:
    print("SORUNLAR:")
    for h in hata:
        print(" -", h)
    sys.exit(1)
vl = os.path.join(KOK, "video_link_ekle.py")
print("video_link_ekle.py:", "VAR" if os.path.exists(vl) else "YOK (zip eksik acilmis!)")
print("YAMA_TAMAM_V4")
