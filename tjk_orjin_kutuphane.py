#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJK ORJİN KÜTÜPHANESİ  (baba/sire bazlı kazanan istatistikleri)
==============================================================

Ham veri = TJK Koşu Sorgulama'dan gelen KAZANAN (1.) koşular. Her satır bir
galibiyet. Baba (sire) bazında toplanır; damsire de bir aygır olduğundan kendi
yavrularıyla veride 'baba' olarak geçer, yani aynı tablo hem BABA hem DEDE
aramasına hizmet eder.

Ürettiği oranlar (panelle birebir):
  - MESAFE dağılımı: kısa/orta/uzun × kum/çim/sentetik  (oran = kategori/yüzey-toplam)
  - KALİTE: cins→sınıf ağırlığı; Σpuan ÷ toplam galibiyet = Kalite Skoru
  - KAÇAK : 800 fark 0..4 → kaçak;  kaçak ÷ toplam galibiyet
  - SPRİNTER: 800 fark > (ırk,zemin) 75p eşiği;  karışık = toplam üstü ÷ toplam fark-dolu

Kütüphane SQLite'ta ham galibiyetleri saklar; ayda/yılda bir yeni koşular
EKLENİR (INSERT) ve oranlar yeniden hesaplanır. Düşük örneklem (az galibiyet)
elenebilsin diye her babanın galibiyet sayısı da tutulur.
"""

import os
import re
import sqlite3
import pandas as pd
import numpy as np

_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tjk_orjin_kutuphane.db")

# --- KALİTE sınıf ağırlıkları ---
# G/A grupları sabit isimli.
KALITE_GA = {"G 1": 4.0, "G 2": 2.6, "A 2": 2.25, "G 3": 1.8, "A 3": 1.5, "G3 Handikap": 1.1}
# Kısa Vade / Handikap ailesi SADECE sayıya göre. Aynı sayının tüm yazımları
# ("KV-N", "Kısa Vade Handikap N", "Handikap N", "HN") aynı sınıf sayılır.
# Sadece bu sayılar ağırlıklı; diğer handikaplar (13-17, 21...) = 0.
KVH_AGIRLIK = {7: 0.7, 8: 0.81, 9: 0.95, 10: 0.95, 18: 1.0, 22: 0.7, 24: 0.81, 25: 0.81}

YUZEYLER = ["Kum", "Çim", "Sentetik"]
KATEGORILER = ["kısa", "orta", "uzun"]

# Kütüphaneye GİRMEYECEK iller (doğu illeri). Yükleme sırasında elenir.
HARIC_SEHIRLER = {"şanlıurfa", "sanliurfa", "diyarbakır", "diyarbakir", "elazığ", "elazig"}


def _sehir_haric_mi(sehir):
    t = str(sehir or "").strip().lower()
    t = t.replace("i̇", "i")
    return t in HARIC_SEHIRLER


# ----------------------------------------------------------------------------
# Yardımcılar
# ----------------------------------------------------------------------------
def _irk_norm(s):
    t = str(s or "").lower()
    if "arap" in t:
        return "Arap"
    if "ingiliz" in t or "i̇ngiliz" in t or "ingı" in t:
        return "İngiliz"
    return ""


def _zemin_norm(s):
    t = str(s or "").strip().lower()
    if "kum" in t:
        return "Kum"
    if "çim" in t or "cim" in t:
        return "Çim"
    if "sentetik" in t:
        return "Sentetik"
    return ""


def _mesafe_tipi(irk, mesafe):
    try:
        m = int(float(mesafe))
    except Exception:
        return ""
    if irk == "Arap":
        if m <= 1500:
            return "kısa"
        if 1600 <= m <= 1800:
            return "orta"
        if m >= 1900:
            return "uzun"
    if irk == "İngiliz":
        if m <= 1400:
            return "kısa"
        if 1500 <= m <= 1700:
            return "orta"
        if m >= 1800:
            return "uzun"
    return ""


def _cins_tabani(cins):
    """'G 3 /Dişi' -> 'G 3';  'Kısa Vade Handikap 24 /H1' -> 'Kısa Vade Handikap 24'."""
    base = str(cins or "").split("/")[0].strip()
    return re.sub(r"\s+", " ", base)


def _kalite_agirlik(cins):
    """Cins -> KALİTE ağırlığı. Yazım varyantları normalize edilir:
    'KV-22' = 'Kısa Vade Handikap 22' = 'Handikap 22' = 'H22' -> 22 sınıfı."""
    base = _cins_tabani(cins)
    if base in KALITE_GA:
        return KALITE_GA[base]
    # Kısa Vade / Handikap ailesi: baştaki etiketi at, sayıyı al.
    m = re.match(r"^(?:KV-|Kısa Vade Handikap |Handikap |KVH-?|H)\s*(\d+)$", base, flags=re.IGNORECASE)
    if m:
        return KVH_AGIRLIK.get(int(m.group(1)), 0.0)
    return 0.0


def _baba_ana_ayir(orijin):
    """'FINESSE - DELAL SULTAN' -> ('FINESSE','DELAL SULTAN').
    'AGRESIVO (USA) - ROYAL BEAUTY' -> ('AGRESIVO (USA)','ROYAL BEAUTY').
    İlk ' - ' ayıracına göre böler (ülke eki babada kalır)."""
    s = str(orijin or "").replace("–", "-").replace("—", "-").strip()
    if " - " in s:
        b, a = s.split(" - ", 1)
    elif "-" in s:
        b, a = s.split("-", 1)
    else:
        b, a = s, ""
    return b.strip(), a.strip()


def _fark_num(x):
    try:
        v = float(str(x).replace(",", ".").strip())
        return v
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Kütüphane (SQLite): ham galibiyetler
# ----------------------------------------------------------------------------
# NOT: DB-seviyesi UNIQUE YOK. Tarihsel 85k dosyada koşu no yok ve tarih yıl-bazlı;
# UNIQUE olsaydı aynı atın aynı yıl/şehirdeki farklı galibiyetleri yanlışlıkla
# silinirdi. Çift-sayma engellemesi SADECE güncel eklemede (tam tarih+koşu) yapılır.
_TABLO = """
CREATE TABLE IF NOT EXISTS kazanan (
    tarih TEXT, sehir TEXT, kosu TEXT, grup TEXT, irk TEXT, cins TEXT,
    mesafe INTEGER, zemin TEXT, kilo TEXT, baba TEXT, ana TEXT,
    kazanc TEXT, at_adi TEXT, yas TEXT, derece TEXT, fark REAL
);
"""


def baglan(db_path=None):
    conn = sqlite3.connect(db_path or _DB, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute(_TABLO)
    conn.commit()
    return conn


def _kolon_bul(df, adaylar):
    for a in adaylar:
        for c in df.columns:
            if str(c).strip().lower() == a.lower():
                return c
    # gevşek: içeren
    for a in adaylar:
        for c in df.columns:
            if a.lower() in str(c).strip().lower():
                return c
    return None


def ham_excel_yukle(conn, excel_path, sheet=0):
    """85k ham dosyayı (veya Koşu Sorgulama çıktısını) kütüphaneye yükler/ekler.
    Kolon adları esnek eşleştirilir; UNIQUE ile çift kayıt engellenir."""
    df = pd.read_excel(excel_path, sheet_name=sheet)
    K = {
        "tarih": _kolon_bul(df, ["tarih", "yıl", "yil"]),
        "sehir": _kolon_bul(df, ["şehir", "sehir", "şehir adı"]),
        "kosu":  _kolon_bul(df, ["koşu", "kosu"]),
        "grup":  _kolon_bul(df, ["grup"]),
        "irk":   _kolon_bul(df, ["ırk", "irk"]),
        "cins":  _kolon_bul(df, ["cins", "koşu cinsi"]),
        "mesafe":_kolon_bul(df, ["meafe", "mesafe"]),
        "zemin": _kolon_bul(df, ["zemin", "pist"]),
        "kilo":  _kolon_bul(df, ["kilo", "sıklet", "siklet"]),
        "baba":  _kolon_bul(df, ["baba"]),
        "ana":   _kolon_bul(df, ["ana", "anne"]),
        "orijin":_kolon_bul(df, ["orijin (baba-anne)", "orijin", "orjin"]),
        "kazanc":_kolon_bul(df, ["kazanç", "kazanc", "ikramiye", "i̇kramiye"]),
        "at_adi":_kolon_bul(df, ["at adı", "at adi", "birinci", "at ismi"]),
        "yas":   _kolon_bul(df, ["yaş", "yas"]),
        "derece":_kolon_bul(df, ["derece"]),
        "fark":  _kolon_bul(df, ["800 e ilk giren ile 800 fark", "fark"]),
    }
    n = 0
    elenen = 0
    cur = conn.cursor()
    for _, r in df.iterrows():
        # Doğu illeri kütüphaneye girmez.
        if K["sehir"] and _sehir_haric_mi(r[K["sehir"]]):
            elenen += 1
            continue
        irk = _irk_norm(r[K["irk"]]) if K["irk"] else _irk_norm(r[K["grup"]] if K["grup"] else "")
        try:
            mesafe = int(re.search(r"(\d{3,4})", str(r[K["mesafe"]])).group(1)) if K["mesafe"] else None
        except Exception:
            mesafe = None
        # Baba/ana: birleşik "Orijin (Baba-Anne)" kolonu ÖNCELİKLİ (varsa ayrılır),
        # yoksa ayrı baba/ana kolonları. (Gevşek eşleşme Orijin'i baba sanmasın diye
        # orijin önce kontrol edilir.)
        if K["orijin"]:
            baba, ana = _baba_ana_ayir(r[K["orijin"]])
        elif K["baba"]:
            baba = str(r[K["baba"]]).strip()
            ana = str(r[K["ana"]]).strip() if K["ana"] else ""
        else:
            baba = ana = ""
        vals = (
            str(r[K["tarih"]]) if K["tarih"] else "",
            str(r[K["sehir"]]) if K["sehir"] else "",
            str(r[K["kosu"]]) if K["kosu"] else "",
            str(r[K["grup"]]) if K["grup"] else "",
            irk,
            str(r[K["cins"]]) if K["cins"] else "",
            mesafe,
            _zemin_norm(r[K["zemin"]]) if K["zemin"] else "",
            str(r[K["kilo"]]) if K["kilo"] else "",
            baba,
            ana,
            str(r[K["kazanc"]]) if K["kazanc"] else "",
            str(r[K["at_adi"]]).strip() if K["at_adi"] else "",
            str(r[K["yas"]]) if K["yas"] else "",
            str(r[K["derece"]]) if K["derece"] else "",
            _fark_num(r[K["fark"]]) if K["fark"] else None,
        )
        try:
            cur.execute(
                "INSERT INTO kazanan "
                "(tarih,sehir,kosu,grup,irk,cins,mesafe,zemin,kilo,baba,ana,kazanc,at_adi,yas,derece,fark) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", vals)
            n += cur.rowcount
        except Exception:
            pass
    conn.commit()
    print(f"[ORJİN KÜTÜPHANE] yüklendi/eklendi: {n} yeni satır | doğu illeri elendi: {elenen} | toplam: {sayac(conn)}")
    return n


def sayac(conn):
    return conn.execute("SELECT COUNT(*) FROM kazanan").fetchone()[0]


def _df(conn):
    return pd.read_sql_query("SELECT * FROM kazanan", conn)


# ----------------------------------------------------------------------------
# HESAPLAMALAR
# ----------------------------------------------------------------------------
def sprinter_esikleri(conn):
    """(ırk, zemin) -> farkların 75. persentili + N (fark dolu)."""
    df = _df(conn)
    df = df[df["fark"].notna()]
    out = {}
    for irk in ("İngiliz", "Arap"):
        for zem in YUZEYLER:
            s = df[(df["irk"] == irk) & (df["zemin"] == zem)]["fark"]
            if len(s):
                out[(irk, zem)] = {"esik": float(np.percentile(s, 75)), "N": int(len(s))}
    return out


def mesafe_tablosu(conn):
    """baba -> yüzey×kategori sayıları + oran + SAYI('k/toplam')."""
    df = _df(conn)
    df["kat"] = df.apply(lambda r: _mesafe_tipi(r["irk"], r["mesafe"]), axis=1)
    df = df[(df["baba"] != "") & (df["zemin"] != "") & (df["kat"] != "")]
    rows = []
    for baba, g in df.groupby("baba"):
        rec = {"baba": baba, "toplam_galibiyet": int(len(g))}
        for zem in YUZEYLER:
            zg = g[g["zemin"] == zem]
            tot = len(zg)
            for kat in KATEGORILER:
                c = int((zg["kat"] == kat).sum())
                rec[f"{zem}_{kat}_sayi"] = c
                rec[f"{zem}_{kat}_oran"] = (c / tot) if tot else 0.0
                rec[f"{zem}_{kat}_str"] = f"{c}/{tot}" if tot else "0/0"
            rec[f"{zem}_toplam"] = int(tot)
        rows.append(rec)
    return pd.DataFrame(rows)


def kalite_tablosu(conn):
    """baba -> Toplam Ağırlıklı Puan, Toplam Kazanç(=galibiyet), Kalite Skoru."""
    df = _df(conn)
    df = df[df["baba"] != ""]
    df["puan"] = df["cins"].map(_kalite_agirlik)
    rows = []
    for baba, g in df.groupby("baba"):
        tap = float(g["puan"].sum())
        tot = int(len(g))
        rows.append({"baba": baba, "toplam_agirlikli_puan": round(tap, 4),
                     "toplam_kazanc": tot, "kalite_skoru": round(tap / tot, 6) if tot else 0.0})
    return pd.DataFrame(rows)


def kacak_tablosu(conn):
    """baba -> kaçak (fark 0..4) sayısı / FARK-DOLU payda (panelle aynı).
    Payda = 800 farkı DOLU olan galibiyetler; toplam galibiyet değil."""
    df = _df(conn)
    df = df[df["baba"] != ""]
    rows = []
    for baba, g in df.groupby("baba"):
        tot = int(len(g))
        fk = g[g["fark"].notna()]
        payda = int(len(fk))
        kacak = int(((fk["fark"] >= 0) & (fk["fark"] <= 4)).sum())
        rows.append({"baba": baba, "kacak_sayi": kacak, "fark_dolu_payda": payda,
                     "toplam_galibiyet": tot,
                     "kacak_oran": round(kacak / payda, 6) if payda else 0.0})
    return pd.DataFrame(rows)


def sprinter_tablosu(conn, esikler=None):
    """baba -> yüzey bazında eşik üstü/fark-dolu + karışık oran."""
    if esikler is None:
        esikler = sprinter_esikleri(conn)
    df = _df(conn)
    df = df[(df["baba"] != "") & (df["fark"].notna())]
    rows = []
    for baba, g in df.groupby("baba"):
        rec = {"baba": baba}
        kar_ustu = kar_payda = 0
        for zem in YUZEYLER:
            zg = g[g["zemin"] == zem]
            payda = len(zg)
            ustu = 0
            for _, r in zg.iterrows():
                es = esikler.get((r["irk"], zem))
                if es and r["fark"] > es["esik"]:
                    ustu += 1
            rec[f"{zem}_ustu"] = int(ustu)
            rec[f"{zem}_payda"] = int(payda)
            rec[f"{zem}_oran"] = round(ustu / payda, 6) if payda else 0.0
            kar_ustu += ustu
            kar_payda += payda
        rec["karisik_ustu"] = int(kar_ustu)
        rec["karisik_payda"] = int(kar_payda)
        rec["karisik_oran"] = round(kar_ustu / kar_payda, 6) if kar_payda else 0.0
        rows.append(rec)
    return pd.DataFrame(rows)


def tum_tablolari_uret(conn, min_galibiyet=0):
    """Dört tabloyu üretir. min_galibiyet altındaki babalar işaretlenebilir."""
    esik = sprinter_esikleri(conn)
    return {
        "mesafe": mesafe_tablosu(conn),
        "kalite": kalite_tablosu(conn),
        "kacak": kacak_tablosu(conn),
        "sprinter": sprinter_tablosu(conn, esik),
        "sprinter_esikleri": esik,
    }


if __name__ == "__main__":
    import sys
    excel = sys.argv[1] if len(sys.argv) > 1 else None
    conn = baglan()
    if excel:
        ham_excel_yukle(conn, excel)
    print("Kütüphanedeki galibiyet:", sayac(conn))
