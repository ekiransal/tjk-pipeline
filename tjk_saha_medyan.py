# -*- coding: utf-8 -*-
"""
TJK SAHA MEDYAN — BAĞIMSIZ AYRIŞTIRICI
======================================

Bir yarış SONUÇ SAYFASI HTML'inden (son800'ün zaten açtığı sayfa) o koşudaki
saha HP/Sıklet medyanını ve listelerini hesaplar. Ayrıca atın geçmiş
tablosundan eski HP'yi (referans satırın bir altındaki koşu) ve koşu kimliğini
üretir.

ÖNEMLİ: Bu modül derece_scraper.py'deki KANITLANMIŞ ayrıştırıcı fonksiyonların
BİREBİR kopyasıdır (Selenium/driver/ağ yan etkileri çıkarılmış, debug_kaydet
no-op yapılmıştır). Amaç: son800 scraper'ı derece'ye dokunmadan aynı mantıkla
saha medyanını alsın.

Genel API (en altta):
    saha_medyan_cek(html, at_adi="", detay_url="", row=None)
        -> (hp_medyan, kilo_medyan, hp_liste, kilo_liste)
    eski_hp_bul(gecmis_df, row_name)   -> "67" / "" / "0"
    kosu_kimlik(row)                   -> "Tarih|Şehir|Msf|Pist|Kcins" / ""
"""

import re
import pandas as pd
from io import StringIO

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None


# --- Ayrıştırıcının doldurduğu global listeler (derece ile aynı isimler) ---
son_medyan_hp_liste = ""
son_medyan_kilo_liste = ""
son_medyan_eslesme_liste = ""
son_medyan_hp_liste_ham = []
son_medyan_kilo_liste_ham = []


def debug_kaydet(*a, **k):
    # Bağımsız modülde debug/log yan etkisi yok.
    return


# =========================================================================
# derece_scraper.py'den BİREBİR kopyalanan fonksiyonlar (leaf-first)
# =========================================================================

def satir_kolon_getir(row, kolon_adi):
    if kolon_adi in row.index:
        return row[kolon_adi]

    hedef = str(kolon_adi).strip().upper()

    for c in row.index:
        if str(c).strip().upper() == hedef:
            return row[c]

    return ""


def metin_norm(x):
    t = str(x).strip().upper()
    t = re.sub(r"\s+", " ", t)
    return t


def kolon_adi_norm(c):
    """Pandas/MultiIndex kolon adını tek satır ve normalize metne çevirir."""
    if isinstance(c, tuple):
        c = " ".join([str(x) for x in c if str(x).strip() and str(x).strip().lower() != "nan"])
    t = str(c).strip().upper()
    t = t.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"\s+", " ", t)
    return t


def metin_norm_at(x):
    t = str(x).strip().upper()
    t = t.replace("İ", "I").replace("İ", "I")
    t = t.replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"[^A-Z0-9]+", "", t)
    return t


def at_adi_sade(at_adi):
    """Programdaki/sonuçtaki parantez numaralarını temizler.

    Örnek:
    - ALMUTANABİY (6) -> ALMUTANABİY
    - ALMUTANABİY(4) KG -> ALMUTANABİY
    Buradaki parantezler bazen koşu/program/start numarasıdır; at adı kontrolünde
    bunları kullanmak tırt, doğru tabloyu reddeder.
    """
    t = str(at_adi or "").strip()
    t = re.sub(r"\([^)]*\)", " ", t)
    t = re.sub(r"\b(KG|K|DB|SK|SGKR|GKR|YP|BB|DS|TS|AP|DILBAĞI|KULAKLIK)\b", " ", t, flags=re.IGNORECASE)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def sayi_float_temizle(x):
    if pd.isna(x):
        return None

    t = str(x).strip()
    if t == "" or t.lower() == "nan":
        return None

    t = re.sub(r"[^0-9,.-]", "", t)
    if t == "":
        return None

    # 1.234,5 gibi olursa binlik noktayı at, virgülü ondalık yap.
    if "," in t:
        t = t.replace(".", "").replace(",", ".")

    try:
        return float(t)
    except:
        return None


def siklet_float_temizle(x):
    """Yarış sonucu Sıklet değerini temizler.

    KRİTİK: TJK sonuç tablosunda 54+0,13 gibi değerlerde + sonrası
    bindirme/ek bilgidir; kilo medyanına girmeyecek.
    Örnek: 54+0,13 -> 54, 55+1.20 -> 55, 58,5 -> 58.5
    """
    if pd.isna(x):
        return None

    t = str(x).strip()
    if t == "" or t.lower() == "nan":
        return None

    if "+" in t:
        t = t.split("+", 1)[0].strip()

    t = re.sub(r"[^0-9,.-]", "", t)
    if t == "":
        return None

    if "," in t:
        t = t.replace(".", "").replace(",", ".")

    try:
        return float(t)
    except:
        return None


def liste_deger_formatla(x):
    try:
        v = float(x)
    except:
        return ""
    if v.is_integer():
        return str(int(v))
    return str(v).replace(".", ",")


def hp_deger_formatla(x):
    """HP değerini Excel çıktısı için temiz formatlar."""
    v = sayi_float_temizle(x)
    if v is None:
        return ""

    if float(v).is_integer():
        return str(int(v))

    return str(v).replace(".", ",")


def hp_gecmis_kolonu_bul(df):
    """At geçmiş tablosundaki HP kolonunu bulur.

    TJK bazen kolon adını HP gibi gösterse de pandas okurken boşluk/nokta/çoklu başlık
    karışabiliyor. Görünen başlık HP olduğu için normalize edilmiş hali HP olan kolon kabul edilir.
    """
    if df is None:
        return None

    for c in df.columns:
        cn = kolon_adi_norm(c)
        cn_sade = re.sub(r"[^A-Z0-9]", "", cn)
        if cn == "HP" or cn_sade == "HP":
            return c

    return None


def onceki_kosu_hp_getir(df, row_index):
    """İşlenen ilk-4 satırın HP'sini değil, hemen altındaki önceki koşunun HP'sini döndürür.

    TJK at geçmiş tablosunda genelde en yeni koşu üstte, daha eski koşular alttadır.
    Bu yüzden işlenen satırın bir altındaki satır, tarih olarak bir önceki koşudur.

    Kural:
    - İşlenen satırın kendi HP'si asla kullanılmaz.
    - Bir alt satır varsa onun HP'si alınır; boşsa boş yazılır.
    - Bir alt satır yoksa bu atın hayatının ilk koşusudur; HP = 0 yazılır.
    """
    if df is None or len(df) == 0:
        return "0"

    hp_col = hp_gecmis_kolonu_bul(df)
    if hp_col is None:
        return ""

    try:
        pozisyon = list(df.index).index(row_index)
    except ValueError:
        return ""

    onceki_pozisyon = pozisyon + 1
    if onceki_pozisyon >= len(df):
        return "0"

    onceki_row = df.iloc[onceki_pozisyon]
    hp_raw = onceki_row.get(hp_col, "")

    hp_formatli = hp_deger_formatla(hp_raw)
    if hp_formatli != "":
        return hp_formatli

    # Önceki koşunun HP hücresi gerçekten boş/geçersizse aynen boş bırak.
    t = str(hp_raw).strip()
    if t.lower() in ["nan", "none", "<na>"]:
        return ""
    return t


def kosu_cache_anahtari(row):
    """Aynı koşuyu URL'den değil, koşunun ortak yarış bilgilerinden yakalar.

    Güvenli anahtar: Tarih + Şehir + Mesafe + Pist + Koşu Cinsi
    """
    tarih = metin_norm(satir_kolon_getir(row, "Tarih"))
    sehir = metin_norm(satir_kolon_getir(row, "Şehir"))
    msf = metin_norm(satir_kolon_getir(row, "Msf"))
    pist = metin_norm(satir_kolon_getir(row, "Pist"))
    kcins = metin_norm(satir_kolon_getir(row, "Kcins"))

    zorunlu = [tarih, sehir, msf, pist, kcins]
    if any(x == "" or x == "NAN" for x in zorunlu):
        return ""

    return "|".join(zorunlu)


def hp_kilo_kolonlarini_bul(df):
    """Sadece başlığı TAM HP ve TAM SIKLET olan kolonları kabul eder."""
    hp_col = None
    kilo_col = None

    for c in df.columns:
        cn = kolon_adi_norm(c)
        cn_sade = re.sub(r"[^A-Z0-9]", "", cn)

        if cn_sade == "HP":
            hp_col = c
        if cn_sade == "SIKLET":
            kilo_col = c

    if hp_col is None or kilo_col is None:
        return None

    tmp = df[[hp_col, kilo_col]].copy()
    tmp.columns = ["HP", "Sıklet"]
    tmp["HP_NUM"] = tmp["HP"].apply(sayi_float_temizle)
    tmp["KILO_NUM"] = tmp["Sıklet"].apply(siklet_float_temizle)

    # HP boş olan at medyana girmez. Kilo aynı HP'si dolu atlardan alınır.
    tmp = tmp.dropna(subset=["HP_NUM", "KILO_NUM"])
    tmp = tmp[
        (tmp["HP_NUM"] >= 1) & (tmp["HP_NUM"] <= 180) &
        (tmp["KILO_NUM"] >= 35) & (tmp["KILO_NUM"] <= 75)
    ].copy()

    if len(tmp) < 2:
        return None

    return tmp


def tablo_at_adi_iceriyor_mu(df, at_adi):
    if not at_adi:
        return False
    hedef = metin_norm_at(at_adi_sade(at_adi))
    if hedef == "":
        return False
    try:
        metin = " ".join(df.astype(str).fillna("").values.ravel().tolist())
    except Exception:
        return False
    return hedef in metin_norm_at(metin)


def yaris_sonuc_sayfasi_mi(url, html):
    u = str(url or "").upper()
    h = str(html or "").upper()
    return ("GUNLUKYARISSONUCLARI" in u) or ("YARIŞ SONUÇLARI" in h) or ("YARIS SONUCLARI" in h)


def row_degeri_norm(row, ad):
    return str(satir_kolon_getir(row, ad)).strip()


def kosu_no_satirdan_al(row):
    """At geçmiş satırındaki K. No-K. Adı kolonundan koşu numarasını alır."""
    val = str(satir_kolon_getir(row, "K. No-K. Adı")).strip()
    m = re.search(r"(\d+)", val)
    if m:
        return int(m.group(1))
    return None


def html_kosu_blogu_kes(html, row=None, detay_url="", at_adi=""):
    """Yarış sonuç HTML'inden SADECE ilgili koşu bloğunu keser."""
    h = str(html or "")
    if len(h) < 500:
        return h

    kosu_no = kosu_no_satirdan_al(row) if row is not None else None
    tarih = row_degeri_norm(row, "Tarih") if row is not None else ""
    sehir = row_degeri_norm(row, "Şehir").upper() if row is not None else ""
    msf = row_degeri_norm(row, "Msf") if row is not None else ""
    kcins = row_degeri_norm(row, "Kcins").upper() if row is not None else ""

    hedef_at = metin_norm_at(at_adi_sade(at_adi))

    def blok_at_iceriyor_mu(segment):
        if not hedef_at:
            return True
        return hedef_at in metin_norm_at(segment)

    def blok_skor(segment):
        """Aynı koşu no birden fazla kez görünürse doğru sonucu seçmek için puan."""
        seg_up = segment.upper()
        puan = 0
        if blok_at_iceriyor_mu(segment):
            puan += 10000
        if sehir and sehir in seg_up:
            puan += 300
        if msf and str(msf) in segment:
            puan += 200
        if kcins and kcins.split("/")[0].strip() and kcins.split("/")[0].strip() in seg_up:
            puan += 100
        if re.search(r"At\s*İsmi|At\s+Ismi|Sıklet|Siklet|HP", segment, flags=re.IGNORECASE):
            puan += 50
        return puan

    def segment_temizle(segment):
        bitisler = []
        for bp in ["Dağıtılacak", "Dagitilacak", "KURUMSAL", "Bülten", "BULTEN"]:
            i = segment.upper().find(bp.upper())
            if i > 1500:
                bitisler.append(i)
        if bitisler:
            segment = segment[:min(bitisler)]
        return segment

    # 1) Esas yöntem: HTML'i "N. Koşu" başlıklarına böl, row'daki koşu no ile birebir eşleştir.
    pat_heading = re.compile(r"(\d+)\s*\.\s*Ko[şs]u", flags=re.IGNORECASE)
    headings = [(int(m.group(1)), m.start(), m.group(0)) for m in pat_heading.finditer(h)]

    if headings:
        adaylar = []
        for i, (no, pos, title) in enumerate(headings):
            if kosu_no is not None and no != kosu_no:
                continue
            next_pos = headings[i + 1][1] if i + 1 < len(headings) else min(len(h), pos + 120000)
            segment = segment_temizle(h[pos:next_pos])
            if len(segment) <= 500:
                continue
            if not re.search(r"At\s*İsmi|At\s+Ismi|Sıklet|Siklet|HP", segment, flags=re.IGNORECASE):
                continue
            adaylar.append((blok_skor(segment), no, pos, segment))

        if adaylar:
            adaylar = sorted(adaylar, key=lambda x: x[0], reverse=True)
            skor, no, pos, segment = adaylar[0]
            if blok_at_iceriyor_mu(segment):
                print(f"KOSU BLOK OK: {at_adi} | tarih={tarih} | kosu_no={kosu_no} | secilen={no}. Koşu | skor={skor}")
                return segment
            raise Exception(
                f"{kosu_no}. Koşu bloğu bulundu ama hedef at adı blokta yok; yanlış koşudan medyan yazmak yasak: "
                f"{at_adi} | {detay_url}"
            )

    # 2) Koşu başlığı bulunamazsa anchor fallback.
    m = re.search(r"#([A-Za-z0-9_-]+)", str(detay_url or ""))
    if m:
        anchor = m.group(1)
        for ap in [
            rf'id=["\\\']{re.escape(anchor)}["\\\']',
            rf'name=["\\\']{re.escape(anchor)}["\\\']',
            re.escape(anchor),
        ]:
            am = re.search(ap, h, flags=re.IGNORECASE)
            if not am:
                continue

            left = max(0, am.start() - 45000)
            right = min(len(h), am.start() + 90000)
            segment = segment_temizle(h[left:right])
            if len(segment) <= 500:
                continue
            if not blok_at_iceriyor_mu(segment):
                continue
            if not re.search(r"At\s*İsmi|At\s+Ismi|Sıklet|Siklet|HP", segment, flags=re.IGNORECASE):
                continue
            print(f"ANCHOR FALLBACK BLOK OK: {at_adi} | tarih={tarih} | kosu_no={kosu_no} | anchor={anchor}")
            return segment

    raise Exception(
        f"İlgili koşu bloğu kesin bulunamadı; medyan yazmak yasak: {at_adi} | "
        f"tarih={tarih} | sehir={sehir} | msf={msf} | kcins={kcins} | kosu_no={kosu_no} | {detay_url}"
    )


def soup_tablodan_hp_kilo_adaylari(html_blok, at_adi=""):
    """Sonuç tablosunu pandas yerine satır satır okur."""
    if BeautifulSoup is None:
        return []

    soup = BeautifulSoup(str(html_blok or ""), "html.parser")
    adaylar = []
    hedef = metin_norm_at(at_adi_sade(at_adi))

    for tablo_no, table in enumerate(soup.find_all("table"), start=1):
        rows = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            vals = [" ".join(c.get_text(" ", strip=True).split()) for c in cells]
            vals = [v for v in vals if v is not None]
            if vals:
                rows.append(vals)

        if not rows:
            continue

        tablo_text = " ".join(" ".join(r) for r in rows)
        if hedef and hedef not in metin_norm_at(tablo_text):
            continue

        header_idx = None
        hp_idx = None
        kilo_idx = None
        at_idx = None

        for i, vals in enumerate(rows[:8]):
            norms = [re.sub(r"[^A-Z0-9]", "", kolon_adi_norm(v)) for v in vals]
            # At İsmi / Sıklet / HP başlıklarını aynı satırda arıyoruz.
            possible_hp = [j for j, n in enumerate(norms) if n == "HP"]
            possible_kilo = [j for j, n in enumerate(norms) if n == "SIKLET"]
            possible_at = [j for j, n in enumerate(norms) if n in ["ATISMI", "ATADI", "AT"] or "ATISMI" in n]
            if possible_hp and possible_kilo:
                header_idx = i
                hp_idx = possible_hp[-1]
                kilo_idx = possible_kilo[0]
                at_idx = possible_at[0] if possible_at else 0
                break

        if header_idx is None:
            continue

        hp_liste = []
        kilo_liste = []
        eslesen_atlar = []

        for vals in rows[header_idx + 1:]:
            # Bazı nested/header satırları araya girebilir.
            if len(vals) <= max(hp_idx, kilo_idx):
                continue

            satir_text = " ".join(vals)
            satir_norm = metin_norm_at(satir_text)
            # Header tekrarıysa geç.
            if "ATISMI" in satir_norm and "SIKLET" in satir_norm:
                continue

            hp_raw = vals[hp_idx]
            kilo_raw = vals[kilo_idx]
            hp_num = sayi_float_temizle(hp_raw)
            kilo_num = siklet_float_temizle(kilo_raw)

            if hp_num is None or kilo_num is None:
                continue
            if not (1 <= float(hp_num) <= 180):
                continue
            if not (35 <= float(kilo_num) <= 75):
                continue

            hp_liste.append(float(hp_num))
            kilo_liste.append(float(kilo_num))
            if at_idx is not None and len(vals) > at_idx:
                eslesen_atlar.append(vals[at_idx])
            else:
                eslesen_atlar.append(vals[0] if vals else "")

        n = len(hp_liste)
        # KRİTİK: Küçük koşuları reddetme. 4 atlı hatta daha küçük koşuda da medyan hesaplanır.
        if n < 1 or n > 24:
            continue

        at_var = True
        if hedef:
            at_var = any(hedef in metin_norm_at(a) for a in eslesen_atlar) or hedef in metin_norm_at(tablo_text)
        if not at_var:
            continue

        hp_ref = float(pd.Series(hp_liste).median())
        kilo_ref = float(pd.Series(kilo_liste).median())

        adaylar.append({
            "kaynak": "SOUP_TR_TD_ILK_SATIR_FIX",
            "tablo_no": tablo_no,
            "n": n,
            "puan": 20000 + n * 100,
            "at_var": at_var,
            "hp_ref": hp_ref,
            "kilo_ref": round(kilo_ref, 1),
            "hp_liste": hp_liste,
            "kilo_liste": kilo_liste,
            "atlar": eslesen_atlar,
        })

    return adaylar


def html_icinden_hp_siklet_medyani(html, detay_url="", at_adi="", row=None):
    """Yarış sonucu HTML'inden HP/Sıklet medyanını alır.

    KRİTİK FIX:
    pandas read_html bazı yarışlarda ilk veri satırını düşürüyor.
    Bu yüzden önce BeautifulSoup ile tablo satırlarını elle okuyoruz.
    Böylece ilk atın HP/Kilosu, örn. PİRNUR 67 / 60,5, listeye girer.
    """
    global son_medyan_hp_liste, son_medyan_kilo_liste, son_medyan_eslesme_liste
    global son_medyan_hp_liste_ham, son_medyan_kilo_liste_ham
    son_medyan_hp_liste = ""
    son_medyan_kilo_liste = ""
    son_medyan_eslesme_liste = ""
    son_medyan_hp_liste_ham = []
    son_medyan_kilo_liste_ham = []

    if not yaris_sonuc_sayfasi_mi(detay_url, html):
        raise Exception(f"Bu HTML yarış sonuç sayfası değil; medyan alınmaz: {detay_url}")

    if "ATKOSUBILGILERI" in str(detay_url).upper():
        raise Exception(f"At geçmiş sayfasından medyan almak yasak: {detay_url}")

    html_blok = html_kosu_blogu_kes(html, row=row, detay_url=detay_url, at_adi=at_adi)

    adaylar = soup_tablodan_hp_kilo_adaylari(html_blok, at_adi=at_adi)

    # Soup hiç tablo çıkaramazsa eski pandas yolu son çare. Ama pandas yolu artık öncelik değil.
    if not adaylar:
        redler = []
        try:
            tables = pd.read_html(StringIO(html_blok))
        except Exception as e:
            debug_kaydet(at_adi, detay_url, row=row, kaynak="SOUP_AND_PANDAS_FAIL", hata=str(e))
            raise Exception(f"Sonuç sayfasında tablo okunamadı: {at_adi} | {detay_url} | {e}")

        for tablo_no, df in enumerate(tables, start=1):
            tmp = hp_kilo_kolonlarini_bul(df)
            if tmp is None:
                continue
            n = len(tmp)
            # KRİTİK: n<5 filtresi kaldırıldı. 2-3-4 atlı koşular da gerçek koşudur.
            if n < 1:
                redler.append(f"pandas tablo {tablo_no}: HP/Sıklet var ama n<1 ({n})")
                continue
            if n > 24:
                redler.append(f"pandas tablo {tablo_no}: fazla uzun tablo n={n}")
                continue
            if not tablo_at_adi_iceriyor_mu(df, at_adi):
                redler.append(f"pandas tablo {tablo_no}: hedef at adı yok")
                continue
            hp_liste = [float(x) for x in tmp["HP_NUM"].tolist()]
            kilo_liste = [float(x) for x in tmp["KILO_NUM"].tolist()]
            adaylar.append({
                "kaynak": "PANDAS_FALLBACK",
                "tablo_no": tablo_no,
                "n": n,
                "puan": 10000 + n * 100,
                "at_var": True,
                "hp_ref": float(pd.Series(hp_liste).median()),
                "kilo_ref": round(float(pd.Series(kilo_liste).median()), 1),
                "hp_liste": hp_liste,
                "kilo_liste": kilo_liste,
            })

        if not adaylar:
            detay = " | ".join(redler[:12])
            debug_kaydet(at_adi, detay_url, row=row, kaynak="NO_HP_KILO_TABLE", hata=detay)
            raise Exception(f"İlgili koşu bloğunda HP + Sıklet tablosu bulunamadı: {at_adi} | {detay_url} | {detay}")

    secilen = sorted(adaylar, key=lambda x: (x.get("at_var", False), x.get("n", 0), x.get("puan", 0)), reverse=True)[0]

    son_medyan_hp_liste = "-".join([liste_deger_formatla(x) for x in secilen["hp_liste"]])
    son_medyan_kilo_liste = "-".join([liste_deger_formatla(x) for x in secilen["kilo_liste"]])
    son_medyan_eslesme_liste = "-".join([
        f"{liste_deger_formatla(hp)}/{liste_deger_formatla(kg)}"
        for hp, kg in zip(secilen["hp_liste"], secilen["kilo_liste"])
    ])
    son_medyan_hp_liste_ham = list(secilen["hp_liste"])
    son_medyan_kilo_liste_ham = list(secilen["kilo_liste"])

    print(
        f"MEDYAN OK: {at_adi} | kaynak={secilen['kaynak']} | tablo={secilen['tablo_no']} | "
        f"n={secilen['n']} | HP_LISTE=[{son_medyan_hp_liste}] | HP={secilen['hp_ref']} | "
        f"KILO_LISTE=[{son_medyan_kilo_liste}] | KILO={secilen['kilo_ref']} | "
        f"ESLESME=[{son_medyan_eslesme_liste}] | URL={detay_url}"
    )

    return secilen["hp_ref"], secilen["kilo_ref"]


# =========================================================================
# GENEL API (son800 bunları çağırır)
# =========================================================================

def saha_medyan_cek(html, at_adi="", detay_url="", row=None):
    """Sonuç sayfası HTML'inden saha HP/kilo medyanı + ham listeleri döndürür.
    Döndürür: (hp_medyan, kilo_medyan, hp_liste, kilo_liste)
    Hata olursa fırlatır (çağıran yakalar)."""
    hp_ref, kilo_ref = html_icinden_hp_siklet_medyani(
        html, detay_url=detay_url, at_adi=at_adi, row=row)
    return (hp_ref, kilo_ref,
            list(son_medyan_hp_liste_ham), list(son_medyan_kilo_liste_ham))


def eski_hp_bul(gecmis_df, row_name):
    """Referans satırın bir altındaki (kronolojik önceki) koşunun HP'si."""
    return onceki_kosu_hp_getir(gecmis_df, row_name)


def kosu_kimlik(row):
    """Koşu kimliği: Tarih|Şehir|Msf|Pist|Kcins (derece ile birebir aynı)."""
    return kosu_cache_anahtari(row)
