#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OZELLIK KAYDEDICI (Katman 1)
web/ klasorunde, parse_mockup.py'den SONRA calisir:
    python3 ozellik_kaydet.py
mockup_parsed.json'daki her at icin bir ozellik satiri yazar:
    ../veri/ozellikler_YYYY-MM-DD.csv
Bu dosyalar birikir; birlestir.py sonuclarla esleyip gecmis_veri.csv uretir.
"""
import json, os, re, csv, sys

J = "mockup_parsed.json"
if not os.path.exists(J):
    sys.exit("mockup_parsed.json yok — once parse_mockup.py calistir.")
D = json.load(open(J, encoding="utf-8"))

hedef = (D.get("extremler") or {}).get("hedef") or ""
m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", hedef)
if not m:
    sys.exit("hedef tarih bulunamadi (extremler.hedef bos) — kayit atlanacak.")
iso = "{2}-{1}-{0}".format(m.group(1), m.group(2), m.group(3))

VERI = os.path.join("..", "veri")
os.makedirs(VERI, exist_ok=True)
CIKTI = os.path.join(VERI, "ozellikler_%s.csv" % iso)

def f(x):
    try:
        v = float(str(x).replace(",", "."))
        return v
    except Exception:
        return None

def temiz_ad(s):
    s = str(s or "").strip().upper()
    s = re.sub(r"\d+$", "", s).strip()
    return s

# gec cikis / derecesiz haritasi (at adi -> bayraklar)
GEC = {}
for at, v in (D.get("gec_cikis") or {}).items():
    if not isinstance(v, dict):
        continue
    ks = v.get("kosular") or []
    GEC[str(at).upper()] = {
        "gec": 1 if any(k.get("boy") for k in ks) else 0,
        "dsiz": 1 if any(k.get("dsiz") for k in ks) else 0,
    }

# extrem haritasi (il, kosu, atno) -> (tip, gun)
EXT = {}
for e in (D.get("extremler") or {}).get("liste") or []:
    EXT[(e.get("il"), str(e.get("kosu")), str(e.get("atno")))] = (e.get("tip"), e.get("gun"))

BASLIKLAR = ["tarih","il","kosu","atno","at","baba",
             "mesafe","zemin","irk","saat","kosu_yorum","kosu_final",
             "gecmis_n","p50_max","p50_ort","p66_max","p66_ort","p75_max","p75_ort",
             "derece_en_iyi_cs",
             "g30_en_iyi","g30_n","songalop_en_iyi","songalop_n",
             "orj_kalite","orj_mesafe","orj_sprinter","orj_kacak",
             "dede_kalite","dede_mesafe",
             "extrem_tip","extrem_gun","gec_cikis","derecesiz","gecmis_yok"]

def drc_cs(v):
    """'1:06,52' benzeri dereceyi santisaniyeye cevir (kucuk = iyi)."""
    s = str(v or "").strip().replace(",", ".")
    m2 = re.match(r"^(?:(\d+):)?(\d+)\.(\d+)$", s)
    if not m2:
        return None
    dk = int(m2.group(1) or 0); sn = int(m2.group(2)); cs = int(m2.group(3)[:2].ljust(2, "0"))
    return dk * 6000 + sn * 100 + cs

satirlar = []
babalar = D.get("babalar") or {}

for b in D.get("Sayfa1") or []:
    h = b.get("header") or {}
    il = h.get("İl", ""); kosu = str(h.get("Koşu No", "")).strip()
    baslik_norm = re.sub(r"\s+", " ", str(b.get("title") or "").strip()).upper()
    baba_map = babalar.get(baslik_norm, {})
    # --- at listesi + gecmis satir gruplari ---
    atlar = {}   # atno -> {"ad":..., "rows":[...], "gecmis_yok":0/1}
    for r in b.get("detay") or []:
        atno = str(r[0] if len(r) > 0 else "").strip()
        ad = temiz_ad(r[1] if len(r) > 1 else "")
        if not atno or not ad:
            continue
        a = atlar.setdefault(atno, {"ad": ad, "rows": [], "gecmis_yok": 0})
        if len(r) > 41 and str(r[41] or "") == "GECMIS_YOK":
            a["gecmis_yok"] = 1
        else:
            a["rows"].append(r)
    # --- galop haritalari: atno -> degerler ---
    g30, gso = {}, {}
    for g in b.get("galops") or []:
        hedef_map = gso if g.get("son") else g30
        for row in g.get("rows") or []:
            p = str(row[0] or "").split("/")
            if len(p) < 2:
                continue
            atno = p[0].strip(); deg = f(p[1])
            if atno and deg is not None:
                hedef_map.setdefault(atno, []).append(deg)
    # --- orjin haritalari ---
    orj = {}   # name -> {atno: deger}
    for t in b.get("tables") or []:
        ad = t.get("name", ""); col = t.get("col", 0)
        anah = ("Dede " if col >= 16 else "") + ad
        mp = orj.setdefault(anah, {})
        for row in t.get("rows") or []:
            mp[str(row[0]).strip()] = f(row[1])
    # --- satirlar ---
    for atno, a in sorted(atlar.items(), key=lambda x: (len(x[0]), x[0])):
        rows = a["rows"]
        def domlar(c_hp):
            vals = []
            for r in rows:
                if len(r) > c_hp + 1:
                    hp = f(r[c_hp]); kl = f(r[c_hp + 1])
                    if hp is not None:
                        vals.append(hp + 2 * (kl or 0))
            return vals
        p50, p66, p75 = domlar(24), domlar(26), domlar(28)
        drc = [drc_cs(r[14]) for r in rows if len(r) > 14]
        drc = [x for x in drc if x]
        ext = EXT.get((il, kosu, atno), ("", ""))
        gflags = GEC.get(a["ad"], {})
        satirlar.append([
            hedef, il, kosu, atno, a["ad"], baba_map.get(atno, ""),
            h.get("Mesafe",""), h.get("Zemin",""), h.get("Irk",""), h.get("Saat",""),
            h.get("Yorum",""), f(h.get("Final")),
            len(rows),
            max(p50) if p50 else "", round(sum(p50)/len(p50),2) if p50 else "",
            max(p66) if p66 else "", round(sum(p66)/len(p66),2) if p66 else "",
            max(p75) if p75 else "", round(sum(p75)/len(p75),2) if p75 else "",
            min(drc) if drc else "",
            min(g30.get(atno,[])) if g30.get(atno) else "", len(g30.get(atno,[])),
            min(gso.get(atno,[])) if gso.get(atno) else "", len(gso.get(atno,[])),
            orj.get("Kalite",{}).get(atno,""), orj.get("Mesafe",{}).get(atno,""),
            orj.get("Sprinter",{}).get(atno,""), orj.get("Kaçak",{}).get(atno,""),
            orj.get("Dede Kalite",{}).get(atno,""), orj.get("Dede Mesafe",{}).get(atno,""),
            ext[0], ext[1], gflags.get("gec",0), gflags.get("dsiz",0), a["gecmis_yok"],
        ])

with open(CIKTI, "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(BASLIKLAR)
    w.writerows(satirlar)
print("[OZELLIK] %s -> %d at, %d kosu" % (CIKTI, len(satirlar),
      len(set((s[1], s[2]) for s in satirlar))))
