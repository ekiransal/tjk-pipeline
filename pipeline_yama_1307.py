#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""13.07 PİPELİNE YAMASI (Mac'te, ~/Desktop/tjk içinde çalıştır):
A) tjk_yeni_yer.py  : program sayfasından KOŞU KİMLİKLERİNİ de toplar -> web/kosu_id.json
                      (AGF çipi TJK'da o koşunun tek-koşu sayfasını açabilsin diye)
B) ERTESI_GUN...PY  : son 6 ayda ilk-4'ü olmayan atın EN SON koşusu 5.lik ise
                      o koşu dereceye (stil dahil) eklenir; 6+ ise eklenmez.
Çalıştır: cd ~/Desktop/tjk && python3 pipeline_yama_1307.py
"""
import os, sys, ast

KOK = os.path.dirname(os.path.abspath(__file__))

# ================= A) tjk_yeni_yer.py =================
Y = os.path.join(KOK, "tjk_yeni_yer.py")
s = open(Y, encoding="utf-8").read()

if "kosu_id.json" in s:
    print("A) tjk_yeni_yer zaten yamalı")
else:
    a1 = '        saatler = {}   # {sehir: {"1":"14:30", ...}}  -> web/saatler.json'
    assert a1 in s, "A1 anchor yok"
    s = s.replace(a1, a1 + '\n        kosu_idler = {}   # {sehir: {"1":"223765", ...}} -> web/kosu_id.json', 1)

    a2 = ('                for m in re.finditer(r"(\\d{1,2})\\.\\s*Ko[şs]u\\s*(\\d{1,2})[.:](\\d{2})", html):\n'
          '                    kno, sa, dk = m.group(1), m.group(2), m.group(3)\n'
          '                    saatler.setdefault(sehir, {}).setdefault(kno, f"{sa}:{dk}")\n'
          '            except Exception:\n'
          '                pass')
    assert a2 in s, "A2 anchor yok"
    ek2 = ('\n            # KOŞU KİMLİKLERİ: sekme bağlantıları href="#223765">1. Koşu ... kalıbından\n'
           '            try:\n'
           "                for m in re.finditer(r'href=\"#(\\d{4,9})\"[^>]*>\\s*(\\d{1,2})\\.\\s*Ko[şs]u', html):\n"
           '                    kid, kno = m.group(1), m.group(2)\n'
           '                    kosu_idler.setdefault(sehir, {}).setdefault(kno, kid)\n'
           '            except Exception:\n'
           '                pass')
    s = s.replace(a2, a2 + ek2, 1)

    a3 = '            print(f"      NOT: saatler.json yazılamadı: {_e}")'
    assert a3 in s, "A3 anchor yok"
    ek3 = ('\n        # WEB: koşu kimlikleri (AGF çipi -> TJK tek-koşu görünümü)\n'
           '        try:\n'
           '            import json as _json3\n'
           '            if os.path.isdir("web"):\n'
           '                _json3.dump(kosu_idler, open(os.path.join("web", "kosu_id.json"), "w",\n'
           '                                             encoding="utf-8"), ensure_ascii=False)\n'
           '                print(f"      web/kosu_id.json yazıldı ({sum(len(v) for v in kosu_idler.values())} koşu kimliği)")\n'
           '        except Exception as _e:\n'
           '            print(f"      NOT: kosu_id.json yazılamadı: {_e}")')
    s = s.replace(a3, a3 + ek3, 1)
    open(Y, "w", encoding="utf-8").write(s)
    print("A) tjk_yeni_yer: koşu kimliği toplama EKLENDİ")

# ================= B) ERTESI (derece) =================
D = os.path.join(KOK, "ERTESI_GUN_MODU_OKA_BASAN_PLUS_FAZ30_STIL_TEK_PY.py")
s = open(D, encoding="utf-8").read()

if "son_kosu_5_fallback" in s:
    print("B) derece zaten yamalı")
else:
    # Yardımcı fonksiyon: sira_ilk4_mi'nin hemen önüne
    ab = "def sira_ilk4_mi(value):"
    assert ab in s, "B anchor (sira_ilk4_mi) yok"
    yardimci = '''def son_kosu_5_fallback(df, at_adi, alti_ay_once, program_tarihi):
    """Ilk-4'u olmayan atin EN SON kosusu 5.lik ise o TEK satiri (DataFrame) dondurur.
    6+ ise ya da son kosu 6 aydan eskiyse None. (13.07.2026 istegi)"""
    try:
        d = df.copy()
        d["TarihParsed"] = pd.to_datetime(d["Tarih"], dayfirst=True, errors="coerce")
        d = d[d["TarihParsed"].notna() & (d["TarihParsed"] < pd.Timestamp(program_tarihi.date()))]
        if len(d) == 0:
            return None
        son = d.sort_values("TarihParsed").iloc[-1]
        m = re.search(r"(\\d+)", str(son.get("S", "")).strip())
        if not m or int(m.group(1)) != 5:
            return None
        if son["TarihParsed"] < alti_ay_once:
            return None
        print(f"SON KOSU 5.: {at_adi} | {son.get('Tarih')} -> dereceye ekleniyor (ilk-4 yok ama son kosu 5.)")
        return d.loc[[son.name]]
    except Exception as _e:
        print(f"SON KOSU 5. kontrol hatasi ({at_adi}): {_e}")
        return None


'''
    s = s.replace(ab, yardimci + ab, 1)

    b1 = ('            if len(df3) == 0:\n'
          '                print(f"ILK4 YOK: {at_adi} | son 6 ay filtresinden önce ilk-4 satırı bulunmadı")\n'
          '                continue')
    assert b1 in s, "B1 anchor yok"
    s = s.replace(b1,
          '            if len(df3) == 0:\n'
          '                df3 = son_kosu_5_fallback(df, at_adi, alti_ay_once, program_tarihi)\n'
          '                if df3 is None:\n'
          '                    print(f"ILK4 YOK: {at_adi} | son 6 ay filtresinden önce ilk-4 satırı bulunmadı")\n'
          '                    continue', 1)

    b2 = ('            if len(df3) == 0:\n'
          '                print(f"SON6AY ILK4 YOK: {at_adi}")\n'
          '                continue')
    assert b2 in s, "B2 anchor yok"
    s = s.replace(b2,
          '            if len(df3) == 0:\n'
          '                df3 = son_kosu_5_fallback(df, at_adi, alti_ay_once, program_tarihi)\n'
          '                if df3 is None:\n'
          '                    print(f"SON6AY ILK4 YOK: {at_adi}")\n'
          '                    continue', 1)
    open(D, "w", encoding="utf-8").write(s)
    print("B) derece: SON KOSU 5. kuralı EKLENDİ")

# ================= söz dizimi kontrolü =================
for yol in (Y, D):
    ast.parse(open(yol, encoding="utf-8").read())
    print("   syntax OK:", os.path.basename(yol))

print("PIPELINE_YAMA_1307_TAMAM — bu akşamki çekimden itibaren geçerli.")
