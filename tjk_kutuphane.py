#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJK KÜTÜPHANE  (kalıcı SQLite önbellek)
======================================

Amaç: Atlar 10-15 günde bir koşuyor. Aynı geçmiş koşunun sonuç sayfasını
her çalıştırmada tekrar tekrar TJK'dan çekmek yerine, bir kez çekip buraya
yazıyoruz. Sonraki çalıştırmalar önce kütüphaneye bakar; koşu zaten varsa
TJK'ya hiç gitmez.

Ne saklıyoruz (koşu kimliği bazında, saha seviyesi):
    - HP medyanı, Kilo medyanı
    - HP listesi, Kilo listesi   (persentil 50/66/75 tekrar hesaplanabilsin diye)
    - detay_url (referans / hata ayıklama)
    - kaynak    ('derece' | '800')

ANAHTAR: koşu kimliği = "Tarih|Şehir|Msf|Pist|Kcins"
    Bu, derece scraper'ındaki kosu_cache_anahtari(row) ile BİREBİR aynı olacak.
    URL formatına bağlı değil; böylece derece (183 gün) ve 800 (375 gün) aynı
    koşuyu aynı anahtarla bulur -> 800, derece'nin çektiğini tekrar çekmez.

TASARIM İLKESİ: Bu modül HİÇBİR koşulda hata fırlatmaz. Veritabanı bozuksa,
kilitliyse ya da yoksa; okuma None döner, yazma sessizce geçer. Yani kütüphane
çökse bile ana scraper eskisi gibi (TJK'dan çekerek) çalışmaya devam eder.
Kütüphane sadece bir HIZLANDIRICI; doğruluk kaynağı değil.
"""

import os
import json
import sqlite3
import datetime

# Varsayılan veritabanı: bu dosyanın yanında.
_VARSAYILAN_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tjk_kutuphane.db")

_TABLO_SQL = """
CREATE TABLE IF NOT EXISTS kosu_medyan (
    kosu_kimlik      TEXT PRIMARY KEY,
    hp_medyan        REAL,
    kilo_medyan      REAL,
    hp_liste         TEXT,
    kilo_liste       TEXT,
    detay_url        TEXT,
    kaynak           TEXT,
    guncelleme       TEXT,
    stil_45          TEXT,
    stil_45_kaynagi  TEXT
);
"""


def _migrate_stil(conn):
    """Eski DB'lerde (stil sütunları yok) kosu_medyan tablosuna stil_45 /
    stil_45_kaynagi kolonlarını EKLER. Var olan veriyi/medyanları bozmaz;
    eski satırlar için değer NULL olur. Asla fırlatmaz."""
    try:
        kolonlar = {r[1] for r in conn.execute("PRAGMA table_info(kosu_medyan)").fetchall()}
        if "stil_45" not in kolonlar:
            conn.execute("ALTER TABLE kosu_medyan ADD COLUMN stil_45 TEXT")
        if "stil_45_kaynagi" not in kolonlar:
            conn.execute("ALTER TABLE kosu_medyan ADD COLUMN stil_45_kaynagi TEXT")
        conn.commit()
    except Exception as e:
        print(f"[KÜTÜPHANE] stil migrasyonu (görmezden geliniyor): {e}")


def baglan(db_path=None):
    """Kütüphaneye bağlan ve tabloyu (yoksa) kur. Hata olursa None döner."""
    yol = db_path or _VARSAYILAN_DB
    try:
        conn = sqlite3.connect(yol, timeout=30)
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA busy_timeout=30000;")
        conn.execute(_TABLO_SQL)
        conn.commit()
        _migrate_stil(conn)   # eski DB'yi otomatik yükselt (stil sütunları)
        return conn
    except Exception as e:
        print(f"[KÜTÜPHANE] baglan hatası (görmezden geliniyor): {e}")
        return None


def _liste_yukle(metin):
    if not metin:
        return []
    try:
        v = json.loads(metin)
        return [float(x) for x in v]
    except Exception:
        return []


def medyan_getir(conn, kosu_kimlik):
    """Koşu kimliğine göre kayıtlı medyan/kilo + listeleri döndürür.
    Bulunamazsa ya da hata olursa None. Asla fırlatmaz."""
    if conn is None or not kosu_kimlik:
        return None
    try:
        cur = conn.execute(
            "SELECT hp_medyan, kilo_medyan, hp_liste, kilo_liste, detay_url, kaynak "
            "FROM kosu_medyan WHERE kosu_kimlik = ?",
            (str(kosu_kimlik),),
        )
        satir = cur.fetchone()
        if satir is None:
            return None
        # SADECE-STİL satırı (stil_yaz ile açılmış, medyanı henüz yok) medyan önbelleği
        # SAYILMAZ -> None dön ki scraper medyanı yine çeksin. İki alan bağımsız birikir.
        if satir[0] is None and satir[1] is None and not satir[2] and not satir[3]:
            return None
        return {
            "hp_medyan": satir[0],
            "kilo_medyan": satir[1],
            "hp_liste": _liste_yukle(satir[2]),
            "kilo_liste": _liste_yukle(satir[3]),
            "detay_url": satir[4] or "",
            "kaynak": satir[5] or "",
        }
    except Exception as e:
        print(f"[KÜTÜPHANE] medyan_getir hatası (görmezden geliniyor): {e}")
        return None


def medyan_yaz(conn, kosu_kimlik, hp_medyan, kilo_medyan,
               hp_liste=None, kilo_liste=None, detay_url="", kaynak=""):
    """Koşu medyanını kütüphaneye yazar (varsa günceller). Asla fırlatmaz.
    Döndürür: True (yazıldı) / False (yazılamadı, görmezden gelinir)."""
    if conn is None or not kosu_kimlik:
        return False
    try:
        hp_j = json.dumps([float(x) for x in (hp_liste or [])])
        kilo_j = json.dumps([float(x) for x in (kilo_liste or [])])
        zaman = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "INSERT INTO kosu_medyan "
            "(kosu_kimlik, hp_medyan, kilo_medyan, hp_liste, kilo_liste, detay_url, kaynak, guncelleme) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(kosu_kimlik) DO UPDATE SET "
            "  hp_medyan=excluded.hp_medyan, kilo_medyan=excluded.kilo_medyan, "
            "  hp_liste=excluded.hp_liste, kilo_liste=excluded.kilo_liste, "
            "  detay_url=excluded.detay_url, kaynak=excluded.kaynak, "
            "  guncelleme=excluded.guncelleme",
            (str(kosu_kimlik),
             None if hp_medyan is None else float(hp_medyan),
             None if kilo_medyan is None else float(kilo_medyan),
             hp_j, kilo_j, str(detay_url or ""), str(kaynak or ""), zaman),
        )
        conn.commit()
        return True
    except Exception as e:
        print(f"[KÜTÜPHANE] medyan_yaz hatası (görmezden geliniyor): {e}")
        return False


def stil_getir(conn, kosu_kimlik):
    """Koşunun %45 seyir sırasını (stil) döndürür -> videoyu TEKRAR AÇMADAN buradan
    okunur. Kayıt yoksa / stil boşsa / hata olursa None (o zaman video çekilir).
    Asla fırlatmaz. Anahtar medyan önbelleğiyle AYNI (derece & 800 paylaşır)."""
    if conn is None or not kosu_kimlik:
        return None
    try:
        cur = conn.execute(
            "SELECT stil_45 FROM kosu_medyan WHERE kosu_kimlik = ?",
            (str(kosu_kimlik),))
        satir = cur.fetchone()
        if not satir:
            return None
        v = satir[0]
        if v is None or str(v).strip() == "":
            return None
        return str(v).strip()
    except Exception as e:
        print(f"[KÜTÜPHANE] stil_getir hatası (görmezden geliniyor): {e}")
        return None


def stil_yaz(conn, kosu_kimlik, stil, kaynak=""):
    """Koşunun %45 stilini kütüphaneye yazar. SADECE dolu değer yazılır — boş/None
    yazılmaz ki başarısız okuma önbelleği kirletmesin (sonraki run tekrar dener).
    Var olan medyan satırını KORUR, yalnızca stil alanını günceller (medyan_yaz da
    stil'i silmez; iki alan bağımsız birikir). Asla fırlatmaz.
    Döndürür: True (yazıldı) / False (boş ya da yazılamadı)."""
    if conn is None or not kosu_kimlik:
        return False
    s = "" if stil is None else str(stil).strip()
    if s == "":
        return False
    try:
        zaman = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "INSERT INTO kosu_medyan (kosu_kimlik, stil_45, stil_45_kaynagi, guncelleme) "
            "VALUES (?, ?, ?, ?) "
            "ON CONFLICT(kosu_kimlik) DO UPDATE SET "
            "  stil_45=excluded.stil_45, stil_45_kaynagi=excluded.stil_45_kaynagi, "
            "  guncelleme=excluded.guncelleme",
            (str(kosu_kimlik), s, str(kaynak or ""), zaman))
        conn.commit()
        return True
    except Exception as e:
        print(f"[KÜTÜPHANE] stil_yaz hatası (görmezden geliniyor): {e}")
        return False


def stil_sayac(conn):
    """Kütüphanede %45 stili dolu kaç koşu var. Hata olursa 0."""
    if conn is None:
        return 0
    try:
        return conn.execute(
            "SELECT COUNT(*) FROM kosu_medyan "
            "WHERE stil_45 IS NOT NULL AND stil_45 != ''").fetchone()[0]
    except Exception:
        return 0


def sayac(conn):
    """Kütüphanede kaç koşu var. Hata olursa 0."""
    if conn is None:
        return 0
    try:
        return conn.execute("SELECT COUNT(*) FROM kosu_medyan").fetchone()[0]
    except Exception:
        return 0


def kaynaga_gore_sayac(conn):
    """Kaynağa göre dağılım (derece / 800 / boş)."""
    if conn is None:
        return {}
    try:
        cur = conn.execute("SELECT kaynak, COUNT(*) FROM kosu_medyan GROUP BY kaynak")
        return {(k or ""): n for k, n in cur.fetchall()}
    except Exception:
        return {}


# ----------------------------------------------------------------------------
# Kendi kendine test (python3 tjk_kutuphane.py)
# ----------------------------------------------------------------------------
if __name__ == "__main__":
    import tempfile

    test_db = os.path.join(tempfile.gettempdir(), "tjk_kutuphane_TEST.db")
    if os.path.exists(test_db):
        os.remove(test_db)

    print("== Kütüphane kendi kendine test ==")
    conn = baglan(test_db)
    assert conn is not None, "baglan başarısız"

    kimlik = "19.05.2026|İstanbul|1400|Çim|Handikap"
    # İlk sefer: yok
    assert medyan_getir(conn, kimlik) is None, "boş olmalıydı"

    # Yaz
    ok = medyan_yaz(conn, kimlik, hp_medyan=58.0, kilo_medyan=54.5,
                    hp_liste=[52, 55, 58, 61, 64], kilo_liste=[52, 53, 54.5, 56, 58],
                    detay_url="https://.../GunlukYarisSonuclari#224544", kaynak="derece")
    assert ok, "yazılamadı"

    # Oku
    kayit = medyan_getir(conn, kimlik)
    assert kayit is not None, "okunamadı"
    assert kayit["hp_medyan"] == 58.0, kayit
    assert kayit["kilo_medyan"] == 54.5, kayit
    assert kayit["hp_liste"] == [52.0, 55.0, 58.0, 61.0, 64.0], kayit
    assert kayit["kaynak"] == "derece", kayit
    print("  yaz/oku OK:", kayit)

    # Güncelleme (aynı kimlik -> üzerine yazar, çift kayıt olmaz)
    medyan_yaz(conn, kimlik, 59.0, 55.0, [50, 59, 68], [50, 55, 60], kaynak="800")
    kayit2 = medyan_getir(conn, kimlik)
    assert kayit2["hp_medyan"] == 59.0, kayit2
    assert sayac(conn) == 1, "çift kayıt oluştu!"
    print("  güncelleme OK, kayıt sayısı:", sayac(conn))

    # Bozuk/None girişler asla fırlatmamalı
    assert medyan_getir(None, kimlik) is None
    assert medyan_getir(conn, "") is None
    assert medyan_yaz(conn, "", 1, 2) is False
    assert medyan_yaz(None, kimlik, 1, 2) is False
    print("  güvenli-hata (fırlatmıyor) OK")

    print("  kaynak dağılımı:", kaynaga_gore_sayac(conn))

    # --- %45 STİL önbelleği ---
    # İlk sefer: stil yok -> None (video çekilmeli)
    assert stil_getir(conn, kimlik) is None, "stil boş olmalıydı"
    # Yaz (dolu) -> okunur
    assert stil_yaz(conn, kimlik, "3", kaynak="derece") is True
    assert stil_getir(conn, kimlik) == "3", stil_getir(conn, kimlik)
    # Boş stil YAZILMAZ (başarısız okuma önbelleği kirletmesin)
    assert stil_yaz(conn, kimlik, "") is False
    assert stil_yaz(conn, kimlik, None) is False
    assert stil_getir(conn, kimlik) == "3", "boş yazım eskiyi silmemeli"
    # Medyan güncellemesi stil'i SİLMEMELİ (iki alan bağımsız)
    medyan_yaz(conn, kimlik, 60.0, 56.0, [55, 60, 65], [54, 56, 58], kaynak="800")
    assert stil_getir(conn, kimlik) == "3", "medyan_yaz stil'i sildi!"
    assert medyan_getir(conn, kimlik)["hp_medyan"] == 60.0
    assert sayac(conn) == 1 and stil_sayac(conn) == 1
    # Yalnız-stil koşu (medyanı olmayan) -> yeni satır açar
    k2 = "01.01.2026|Ankara|1600|Kum|Maiden"
    assert stil_yaz(conn, k2, "1", kaynak="800") is True
    assert stil_getir(conn, k2) == "1"
    assert medyan_getir(conn, k2) is None, "medyan yokken stil olabilir"
    print("  %45 stil önbelleği OK | stilli koşu:", stil_sayac(conn))
    conn.close()
    os.remove(test_db)

    # --- ESKİ DB'DEN MİGRASYON (stil sütunları yok) ---
    eski_db = os.path.join(tempfile.gettempdir(), "tjk_kutuphane_ESKI_TEST.db")
    if os.path.exists(eski_db):
        os.remove(eski_db)
    _c = sqlite3.connect(eski_db)
    _c.execute("CREATE TABLE kosu_medyan (kosu_kimlik TEXT PRIMARY KEY, hp_medyan REAL, "
               "kilo_medyan REAL, hp_liste TEXT, kilo_liste TEXT, detay_url TEXT, "
               "kaynak TEXT, guncelleme TEXT)")
    _c.execute("INSERT INTO kosu_medyan (kosu_kimlik, hp_medyan) VALUES (?, ?)",
               ("ESKI|KOSU|1400|Çim|X", 57.0))
    _c.commit(); _c.close()
    conn2 = baglan(eski_db)   # baglan otomatik migrate etmeli
    assert conn2 is not None
    # eski veri korunmuş olmalı
    assert medyan_getir(conn2, "ESKI|KOSU|1400|Çim|X")["hp_medyan"] == 57.0, "eski veri gitti!"
    # yeni stil kolonu kullanılabilir olmalı
    assert stil_getir(conn2, "ESKI|KOSU|1400|Çim|X") is None
    assert stil_yaz(conn2, "ESKI|KOSU|1400|Çim|X", "5", kaynak="derece") is True
    assert stil_getir(conn2, "ESKI|KOSU|1400|Çim|X") == "5"
    conn2.close()
    os.remove(eski_db)
    print("  eski DB migrasyonu OK (stil sütunları eklendi, veri korundu)")

    print("== TÜM TESTLER GEÇTİ ==")
