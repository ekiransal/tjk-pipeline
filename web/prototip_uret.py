#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""mockup_parsed.json -> tek dosyalık TJK analiz platformu prototipi (HTML)."""
import json, datetime

DATA = json.load(open("mockup_parsed.json", encoding="utf-8"))
DATA["uretim"] = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")   # son güncelleme damgası

HTML = r"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1500">
<title>TJK Yarış Analiz</title>
<style>
  :root{
    --bg:#f6f7f9; --card:#ffffff; --line:#e7e9ee; --txt:#1c2333; --mut:#6b7385;
    --pri:#1e6f5c; --pri2:#e8f3f0; --acc:#c8542a; --gold:#f4b63f;
    --r:14px; --sh:0 1px 3px rgba(20,30,50,.06),0 4px 14px rgba(20,30,50,.05);
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
       background:var(--bg);color:var(--txt);font-size:14px;line-height:1.45}
  .wrap{max-width:none;margin:0 auto;padding:16px 24px}
  /* ÜST BAR */
  .topbar{display:flex;align-items:center;gap:12px;padding:14px 18px;background:var(--card);
          border-radius:var(--r);box-shadow:var(--sh);margin-bottom:14px}
  .logo{display:flex;align-items:center;gap:10px;font-weight:800;font-size:18px;letter-spacing:.2px}
  .logo .ic{width:34px;height:34px;border-radius:10px;background:var(--pri);color:#fff;
            display:flex;align-items:center;justify-content:center;font-size:17px}
  .topbar .sub{color:var(--mut);font-size:12.5px;margin-left:2px}
  /* SEÇİCİLER */
  .secici{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
  .secici .lbl{font-size:12px;font-weight:700;color:var(--mut);text-transform:uppercase;
               letter-spacing:.6px;margin-right:4px;min-width:64px}
  .chip{border:1.5px solid var(--line);background:var(--card);border-radius:999px;
        padding:8px 18px;font-size:14px;font-weight:600;cursor:pointer;transition:.15s;color:var(--txt)}
  .chip:hover{border-color:var(--pri)}
  .chip.on{background:var(--pri);border-color:var(--pri);color:#fff}
  .chip.num{padding:8px 0;width:42px;text-align:center}
  .chip.fchip{padding:4px 12px;font-size:12px}
  /* tablo içi filtre satırı (Excel tarzı kutucuklu) + sıralanabilir başlıklar */
  .detay th{cursor:default;user-select:none}
  .detay th.siralanir{cursor:pointer}
  .detay th.siralanir:hover{color:var(--pri)}
  
  .fsat td{padding:2px 3px !important;background:var(--pri2);border-top:1px solid var(--line)}
  .fsat select{width:100%;max-width:110px;font-size:10.5px;border:1px solid var(--line);
               border-radius:6px;padding:2px 3px;background:#fff;color:var(--txt)}
  .fcell{position:relative}
  .fbtn{width:100%;font-size:10.5px;border:1px solid var(--line);border-radius:6px;
        padding:2px 4px;background:#fff;color:var(--pri);font-weight:800;cursor:pointer}
  .fpanel{display:none;position:absolute;top:100%;left:0;z-index:50;background:#fff;
          border:1px solid var(--line);border-radius:10px;box-shadow:var(--sh);
          padding:8px 10px;max-height:260px;overflow-y:auto;min-width:140px;text-align:left;z-index:500}
  .fpanel.acik{display:block}
  .fpanel label{display:flex;align-items:center;gap:6px;font-size:12px;padding:3px 0;
                white-space:nowrap;cursor:pointer}
  .fpanel .ftumu{accent-color:var(--pri)}
  .fdrop{font-size:12.5px;border:1.5px solid var(--line);border-radius:9px;
         padding:6px 10px;background:var(--card);color:var(--txt);font-weight:600}

  /* filtre oku başlığın YANINDA */
  .detay th{position:relative}
  .fok{display:inline-block;margin-left:5px;padding:0 6px;border:1px solid var(--line);
       border-radius:5px;background:#fff;color:var(--pri);cursor:pointer;font-size:10px;line-height:16px}
  .fok:hover{border-color:var(--pri)}
  .fok.aktif{background:var(--acc);border-color:var(--acc);color:#fff}
  .fayrac{height:1px;background:var(--line);margin:6px 0}
  .ftumu-l b{color:var(--pri)}
  /* DOMİNANS GRUPLARI: 50 / 66 / 75 ayrı tonlar + grup başı kalın çizgi */
  .detay th.g50,.detay th.g66,.detay th.g75{width:46px;font-size:9.5px;letter-spacing:0;padding:4px 1px}
  .detay th.dar{width:40px;font-size:9.5px;letter-spacing:0;padding:4px 1px}

  
  /* GALOP GRUP KUTULARI: kendi içinde bütün, birbirinden keskin ayrım */
  .grup30{background:#edf6f2;border:1.5px solid #9fc4b4;border-radius:10px;padding:5px 6px}
  .grupson{background:#fdf2ea;border:1.5px solid #e4b894;border-radius:10px;padding:5px 6px}
  .grup30 .panel,.grupson .panel{background:#fff}
  /* ORJİN GRUP KUTUSU: 6 tablo tek bütün (baba+dede) */
  .grupor{background:#eef2f9;border:1.5px solid #a9bcdd;border-radius:12px;padding:7px;
          display:inline-block;max-width:100%}
  .gbaslik.or{color:#2f6fb3}

  /* GALOPLAR TEK SATIR: 30 günlük + son galop yan yana, ayrı renkli kutular */
  .galoprow{display:flex;gap:10px;align-items:stretch;overflow:visible;padding-bottom:4px;justify-content:flex-start}
  .galoprow .grup30,.galoprow .grupson{flex:0 1 auto;min-width:0}   /* SOLA YASLI: içerik kadar, gerekirse daralır */
  .gbaslik{font-size:10px;font-weight:800;letter-spacing:.5px;margin:0 0 5px;color:#1e6f5c}
  .gbaslik.son{color:var(--acc)}

  /* ====== MOBİL (telefon önceliği) ====== */
  .mtum{display:none}
  @media(max-width:820px){
    .wrap{padding:10px}
    .m-gizle{display:none}
    body.mtum-acik .m-gizle{display:table-cell}
    .mtum{display:inline-block;margin:0 0 8px;padding:7px 14px;border:1.5px solid var(--pri);
          border-radius:9px;background:#fff;color:var(--pri);font-weight:700;font-size:12.5px}
    .mtum.acik{background:var(--pri);color:#fff}
    body.mtum-acik .detay{overflow-x:auto}
    .detay table{font-size:11.5px}
    .detay th,.detay td{padding:6px 6px}
    .galoprow{flex-direction:column}
    .paneller.tekhiza{flex-wrap:nowrap;overflow-x:auto}
    .tabs{width:100%}.tab{flex:1;padding:9px 6px;text-align:center}
    .kart{padding:12px 12px}
    .ustblok{flex-direction:column}#yan{width:100%}
    .chip{padding:10px 15px;font-size:14px}.chip.num{width:44px}
    /* koşu seçici tepede sabit kalır */
    .kosusec{position:sticky;top:0;z-index:200;background:var(--bg);padding:8px 0;margin:0 -10px;
             padding-left:10px;padding-right:10px;box-shadow:0 2px 8px rgba(20,30,50,.08)}
    /* dokunma hedefleri büyüsün */
    .fok{padding:2px 10px;font-size:12px;line-height:20px}
    .fpanel label{font-size:14px;padding:7px 0}
    .detay{overflow-x:auto}
    .mtum{padding:9px 16px;font-size:13.5px}
    /* detay tablosu mobil: sıkı hücreler + at adı kısaltılır */
    .detay table{font-size:10px}
    .detay th,.detay td{padding:4px 2px}
    .detay td.atadi{max-width:74px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .detay td.ucgen{font-size:10.5px;letter-spacing:0}
    /* GALOPLAR: iki grup (30 günlük | son galop) YAN YANA, içerde paneller alt alta */
    .galoprow{flex-direction:row;gap:8px;overflow:visible}
    .grup30,.grupson{flex:1;min-width:0;padding:7px}
    .paneller.tekhiza{display:grid;grid-template-columns:1fr;gap:6px;overflow:visible}
    .paneller.tekhiza .panel{flex:none;min-width:0;padding:6px 7px}
    .panel h4{font-size:10px;margin-bottom:5px}
    .gbaslik{font-size:9.5px;margin-bottom:5px}
    .g-satir{font-size:9.5px;gap:3px;padding:3px 0}
    .atno{min-width:18px;height:18px;font-size:9.5px;border-radius:5px}
    .gdeger{min-width:32px}
    .sekil{font-size:8.5px;padding:1px 4px}
    .gtarih{font-size:8.5px}
    /* ANALİZ TABLOLARI: 3'lü yan yana, ufak punto */
    .tablolar.hepsi{grid-template-columns:repeat(3,1fr) !important;gap:6px}
    .tablo h4{font-size:10px;padding:6px 7px 4px;gap:4px}
    .tablo h4 .nk{width:7px;height:7px}
    .tablo td,.tablo th{padding:2px 4px;font-size:9px}
    .tablo td.say{font-size:8px}
    /* DETAY: daha sıkı, sağa-sola gitmeden sığmaya yakın */
    .detay table{font-size:9px}
    .detay th,.detay td{padding:3px 1px}
    .detay td.atadi{max-width:62px}
    .detay td.ucgen{font-size:9.5px}
  }
  /* SEKMELER */
  .tabs{display:flex;gap:4px;background:var(--card);padding:5px;border-radius:12px;
        box-shadow:var(--sh);margin-bottom:14px;width:max-content;max-width:100%}
  .tab{padding:9px 22px;border-radius:9px;font-weight:700;font-size:13.5px;cursor:pointer;
       color:var(--mut);border:none;background:transparent;transition:.15s}
  .tab.on{background:var(--pri2);color:var(--pri)}
  /* KOŞU KARTI */
  .kart{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:18px 20px;margin-bottom:14px}
  .kosu-baslik{font-size:16.5px;font-weight:800}
  .kosu-baslik-satir{display:flex;align-items:baseline;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px}
  .kosusaat{font-size:13px;font-weight:800;color:#1a7f4b;background:#eaf6ee;border:1px solid #bfe3cc;border-radius:20px;padding:4px 13px;white-space:nowrap}
  .meta{display:flex;flex-wrap:wrap;gap:8px}
  .meta .m{background:var(--bg);border:1px solid var(--line);border-radius:9px;
           padding:6px 12px;font-size:12.5px}
  .meta .m b{color:var(--pri);margin-right:5px;font-weight:700}
  .meta .m.vurgu{background:#fdf1ec;border-color:#f3d5c8}
  .meta .m.vurgu b{color:var(--acc)}
  /* BÖLÜM BAŞLIĞI */
  .bol-baslik{display:flex;align-items:center;gap:8px;margin:18px 2px 10px;
              font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:var(--mut)}
  .bol-baslik::after{content:"";flex:1;height:1px;background:var(--line)}
  /* Detay bölüm başlıkları koşu başlığı gibi: kalın, büyük, koyu — tablo kime ait belli olsun */
  .bol-baslik.buyuk{font-size:16.5px;font-weight:800;color:var(--txt);text-transform:none;letter-spacing:.2px}
  /* BAŞLIK ÇİPLERİ: galopa zıpla + AGF kısayolları */
  .bol-baslik .atlama{font-size:11px;font-weight:700;color:var(--mut);background:var(--card);
    border:1px solid var(--line);border-radius:14px;padding:3px 11px;cursor:pointer;
    white-space:nowrap;letter-spacing:.2px;flex:0 0 auto}
  .bol-baslik .atlama:hover{color:var(--txt);border-color:#b9c2d0}
  .bol-baslik .atlama.agfc{color:#1a7f4b;background:#eaf6ee;border-color:#bfe3cc}
  .bol-baslik .atlama.agfc:hover{color:#116338}
  .bol-baslik .ayninot{font-size:11.5px;font-weight:600;color:var(--mut);font-style:italic;letter-spacing:0;text-transform:none}
  /* GALOP PANELLERİ */
  .paneller{display:grid;grid-template-columns:repeat(auto-fill,minmax(215px,1fr));gap:10px}
  /* TEK HİZA: galop panelleri alta sarkmasın; sığmazsa yatay kaydırılır */
  .paneller.tekhiza{display:flex;flex-wrap:nowrap;overflow:visible;padding-bottom:2px;gap:4px;justify-content:flex-start}
  .paneller.tekhiza .panel{flex:0 1 auto;min-width:0;padding:3px 5px}   /* SIKI + sola yaslı */
  .panel{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
  .panel h4{font-size:10px;font-weight:800;color:var(--pri);margin-bottom:3px;letter-spacing:.2px}
  .panel.sonp h4{color:var(--acc)}
  .g-satir{display:flex;flex-wrap:wrap;align-items:center;gap:1px 2px;padding:1.5px 0;
           border-bottom:1px dashed var(--line);font-size:9.5px}  /* sığmazsa tarih alt satıra iner, komşu panele TAŞMAZ */
  .g-satir:last-child{border-bottom:none}
  .atno{min-width:16px;height:16px;border-radius:5px;background:var(--pri2);color:var(--pri);
        font-weight:800;display:flex;align-items:center;justify-content:center;font-size:9.5px}
  .gdeger{font-weight:700;min-width:28px}
  .gdeger.neg{color:#1a7f4b}.gdeger.poz{color:#c23a3a}
  .sekil{font-size:8.5px;font-weight:800;border-radius:4px;padding:1px 3px;background:#eef0f4;color:var(--mut)}
  .gtarih{color:var(--mut);font-size:9px;margin-left:auto;text-align:right;white-space:nowrap}
  .panel{overflow:hidden}   /* içerik komşu panele kesinlikle akmaz */
  .gsehir{color:var(--acc);font-size:9px;font-weight:600}
  .bos{color:var(--mut);font-size:12px;font-style:italic;padding:6px 0}
  .tablonot{color:var(--mut);font-size:11.5px;font-style:italic;margin-top:8px}
  .tablonot.ustte{margin:0 0 6px;text-align:left}
  /* ANALİZ TABLOLARI */
  .tablolar{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px}
  .tablo{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
  .tablo h4{font-size:13px;font-weight:800;padding:11px 14px 9px;color:var(--txt);
            display:flex;align-items:center;gap:7px}
  .tablo h4 .nk{width:9px;height:9px;border-radius:3px}
  .t-Kalite .nk{background:#1e6f5c}.t-Mesafe .nk{background:#2f6fb3}
  .t-Sprinter .nk{background:#c8542a}.t-Kaçak .nk{background:#8a56c9}
  .t-Dede-Kalite .nk{background:#0f4d40}.t-Dede-Mesafe .nk{background:#1d4e80}
  .tablolar.hepsi{display:flex;flex-wrap:wrap;justify-content:flex-start;gap:8px}   /* SOLA YASLI, içerik kadar */
  .tablolar.hepsi .tablo{flex:0 1 auto;width:188px}
  .tablolar.hepsi .tablo h4{font-size:11px;padding:7px 9px 5px;gap:5px}
  .tablolar.hepsi .tablo h4 .nk{width:7px;height:7px}
  .tablolar.hepsi th{font-size:8.5px;padding:3px 4px;letter-spacing:.1px}
  .tablolar.hepsi td{font-size:10px;padding:3px 4px}
  .tablolar.hepsi td.no{width:26px}
  .tablolar.hepsi td.say{font-size:9.5px}
  @media(max-width:1150px){.tablolar.hepsi{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:700px){.tablolar.hepsi{grid-template-columns:repeat(2,1fr)}}
  @media(max-width:440px){.tablolar.hepsi{grid-template-columns:1fr}}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);
     text-align:left;padding:5px 9px;background:var(--bg)}
  td{padding:5px 9px;border-top:1px solid var(--line)}
  td.no{font-weight:800;width:44px}
  td.deg{position:relative;font-weight:700}
  td.deg .bar{position:absolute;left:0;top:15%;bottom:15%;background:var(--pri2);border-radius:0 6px 6px 0;z-index:0}
  td.deg span{position:relative;z-index:1}
  td.say{color:var(--txt);font-weight:700;text-align:left;white-space:nowrap}  /* SAYI başlığının altında, orana yakın */
  /* DETAY TABLOSU — TEK SAYFA: yana kaydırma yerine sığdır (kompakt) */
  .detay{overflow-x:auto;border:1px solid var(--line);border-radius:12px;background:var(--card);
         display:inline-block;max-width:100%;vertical-align:top}  /* kutu tabloya sarılır, boş sağ kalmaz */
  .detay table{font-size:10px;white-space:normal;width:auto;margin:0}  /* genişlik: içerik kadar (DAR) */
  .detay th{position:sticky;top:0;background:var(--bg);z-index:2;padding:3px 2px;
            white-space:normal;line-height:1.1;border:1px solid #c9cfd9;border-bottom:2px solid #aab2bf;text-align:center;font-size:9.5px}
  .detay td{padding:2px 2px;border:1px solid #c9cfd9;line-height:1.2;text-align:center}
  .detay tr:nth-child(even) td{background:#fafbfc}
  .detay tr:hover td{background:#e2efe8 !important}
  .detay td.atadi{font-weight:700;text-align:center;width:64px;white-space:normal;overflow-wrap:break-word;line-height:1.15}
  .detay td.kcins{width:56px;white-space:normal;line-height:1.15}   /* Koşu Cinsi: metin kaydırılır, sütun dar */
  .detay td.yc{width:42px;white-space:normal;line-height:1.15}      /* Yaş/Cins: '4y Erkek' iki satıra iner */
  .detay td.ucgen{font-size:13px;letter-spacing:1px;color:#1c2333;white-space:nowrap;text-align:center !important}
  .detay td.orta{text-align:center}
  .detay td.num{text-align:center}
  .detay td.drc{font-weight:800;color:#111}
  .detay td.sonno{font-weight:800;background:var(--pri2)}
  .detay td.kfiyi{color:#1a7f4b;font-weight:700}
  .detay td.kfkotu{color:#c23a3a;font-weight:700}
  .detay td.dom{font-weight:700}
  .detay td.dom.poz{color:#1a7f4b}
  .detay td.dom.neg{color:#c23a3a}
  .detay td.dom.yokd{color:#b7bfcc;font-weight:400;font-size:9.5px;font-style:italic;white-space:nowrap}   /* Maiden/Şartlı 1/Şartlı 19: soluk etiket */
  /* SONUÇ İŞARETLERİ: biten koşu çipi + kazanan at (numara geçen HER YERDE) */
  .chip.bitti:not(.on){background:#eaf6ee;border-color:#bfe3cc;color:#1a7f4b}
  .detay td.kazanan{color:#1a7f4b !important;font-weight:800}
  .kzn{color:#1a7f4b !important;font-weight:800 !important;white-space:nowrap}
  td.no.kzn{width:auto !important;background:#eaf6ee}
  .ex-satir .kno.kzn{background:#eaf6ee}
  /* EXTREMLER (üstte, seçicilerin sağındaki boş alanda) */
  .ustblok{display:flex;gap:14px;align-items:flex-start}
  .ustsol{flex:1;min-width:0}
  #yan{width:300px;flex-shrink:0}
  #yan .kart{margin-bottom:0}
  .extrem h3{font-size:13px;font-weight:800;margin-bottom:8px;display:flex;align-items:center;gap:6px}
  .extrem .grup{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;
                color:var(--mut);margin:10px 0 5px}
  .ex-satir{display:flex;align-items:center;gap:7px;padding:4px 0;font-size:12.5px;
            border-bottom:1px dashed var(--line)}
  .ex-satir:last-child{border-bottom:none}
  .ex-satir .kno{font-size:10.5px;font-weight:800;background:#eef0f4;color:var(--mut);
                 border-radius:6px;padding:2px 6px;white-space:nowrap}
  .ex-satir .gun{margin-left:auto;font-weight:800;white-space:nowrap}
  .ex-satir.sik .gun{color:#c23a3a}
  .ex-satir.uzun .gun{color:#2f6fb3}
  .ex-satir.gec .gun,.ex-satir.dsiz .gun{color:#c77f00;font-size:11px}
  .ex-satir.dsiz .gun{color:#8a5ac2}
  @media(max-width:900px){.satir{flex-direction:column-reverse}#yan{width:100%}}
  /* AGF */
  .agf{padding:40px;text-align:center;color:var(--mut)}
  .agf .big{font-size:38px;margin-bottom:8px}
  footer{color:var(--mut);font-size:11.5px;text-align:center;padding:18px 0 8px}
  @media(max-width:640px){
    .wrap{padding:10px}.chip{padding:7px 14px}.tab{padding:8px 14px}
    .paneller,.tablolar{grid-template-columns:1fr 1fr}
  }
  @media(max-width:440px){.paneller,.tablolar{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <div class="logo"><div class="ic">🏇</div>TJK Yarış Analiz</div>
    <div class="sub">Günlük koşu analiz platformu — prototip<span id="uretimNot"></span></div>
  </div>

  <div class="ustblok">
    <div class="ustsol">
      <div class="secici"><span class="lbl">İl</span><span id="iller"></span></div>
      <div class="secici kosusec"><span class="lbl">Koşu</span><span id="kosular"></span></div>
      <div class="tabs" id="tabs">
        <button class="tab on" data-t="Sayfa1">Analiz</button>
        <button class="tab" data-t="AGF">AGF</button>
        <button class="tab" data-t="AGF2" style="display:none">2. AGF</button>
      </div>
    </div>
    <div id="yan"></div>
  </div>

  <div id="icerik"></div>
  <footer>Prototip — veriler örnek Excel'den alındı. © TJK Analiz</footer>
</div>

<script>
const DATA = __DATA__;
let secIl=null, secKosu=null, secTab="Sayfa1";

function bloklar(t){ return DATA[t]||[]; }
function iller(){ const s=[]; for(const b of bloklar("Sayfa1")){const il=b.header["İl"]; if(il&&!s.includes(il))s.push(il);} return s; }
function kosular(il){ return bloklar("Sayfa1").filter(b=>b.header["İl"]===il).map(b=>b.header["Koşu No"]); }
function blokBul(t,il,no){ return bloklar(t).find(b=>b.header["İl"]===il&&b.header["Koşu No"]===no); }

function ciz(){
  document.querySelectorAll("body > .fpanel").forEach(x=>x.remove());
  const ilBox=document.getElementById("iller"); ilBox.innerHTML="";
  for(const il of iller()){
    const b=document.createElement("button");
    b.className="chip"+(il===secIl?" on":""); b.textContent=il;
    b.onclick=()=>{secIl=il; secKosu=kosular(il)[0]; derF={sehir:"",msf:"",ay:""}; ciz();};
    ilBox.appendChild(b);
  }
  const koBox=document.getElementById("kosular"); koBox.innerHTML="";
  for(const no of kosular(secIl)){
    const b=document.createElement("button");
    b.className="chip num"+(no===secKosu?" on":""); b.textContent=no;
    // BİTEN KOŞU: koşu çipine ✓ (sonuclar.json'dan, tarih eşleşiyorsa)
    const _sn=(typeof sonucGecerli==="function"&&sonucGecerli())?(((SONUC.iller||{})[secIl]||{})[String(no)]):null;
    if(_sn&&_sn.kazanan!=null){ b.classList.add("bitti"); b.textContent=no+" ✓"; }
    b.onclick=()=>{secKosu=no; derF={sehir:"",msf:"",ay:""}; ciz();};
    koBox.appendChild(b);
  }
  // AGF düğmeleri: 8+ koşulu günde "1. AGF" + "2. AGF", az koşulu günde tek "AGF"
  {
    const N=kosular(secIl).length;
    const b1=document.querySelector('.tab[data-t="AGF"]');
    const b2=document.querySelector('.tab[data-t="AGF2"]');
    if(b1) b1.textContent=(N>=8)?"1. AGF":"AGF";
    if(b2) b2.style.display=(N>=8)?"":"none";
  }
  document.querySelectorAll(".tab").forEach(t=>{
    t.classList.toggle("on",t.dataset.t===secTab);
    t.onclick=()=>{
      if(t.dataset.t==="AGF"||t.dataset.t==="AGF2"){   // AGF: TJK AGF tablosu yeni sekmede
        agfAc(t.dataset.t==="AGF2");
        return;
      }
      secTab=t.dataset.t; ciz();
    };
  });
  // EXTREMLER (sağ üst) — SEÇİLİ KOŞUYA özel
  const yan=document.getElementById("yan");
  const ex=(DATA.extremler&&DATA.extremler.liste)||[];
  const exK=ex.filter(e=>e.il===secIl && String(e.kosu)===String(secKosu));
  {
    const sik=exK.filter(e=>e.tip==="sik").sort((a,b)=>a.gun-b.gun);
    const uzun=exK.filter(e=>e.tip==="uzun").sort((a,b)=>b.gun-a.gun);
    const sat=e=>`<div class="ex-satir ${e.tip}"><span class="kno">${esc(e.atno)}</span>
      <span>${esc(e.at)}</span><span class="gun">${e.gun}g</span></div>`;
    // GEÇ ÇIKIŞ / DERECESİZ: seçili koşunun atları, AT NO ile birlikte
    const gecler=(()=>{
      const gec=DATA.gec_cikis||{}; const out=[];
      const b1=blokBul("Sayfa1",secIl,secKosu);
      if(Object.keys(gec).length&&b1){
        const gor=new Set();
        for(const r of (b1.detay||[])){
          const no=String(r[0]??"").trim();
          const at=String(r[1]??"").trim().toUpperCase().replace(/\d+$/,"").trim();
          if(!at||gor.has(at)) continue; gor.add(at);
          const v=gec[at]; if(!v||!v.problem) continue;
          const g=v.kosular.filter(k=>k.boy).map(k=>`${k.tarih} (${k.boy})`).join(", ");
          const ds=v.kosular.filter(k=>k.dsiz).map(k=>k.tarih).join(", ");
          if(g)  out.push({tip:"gec",  no, at, det:g});
          if(ds) out.push({tip:"dsiz", no, at, det:ds});
        }
      }
      return out;
    })();
    const gsat=e=>`<div class="ex-satir ${e.tip}"><span class="kno">${esc(e.no)}</span>
      <span>${esc(e.at)}</span><span class="gun">${esc(e.det)}</span></div>`;
    yan.innerHTML=`<div class="kart extrem"><h3>⚠️ Dikkat — ${esc(secKosu)}. Koşu
      <span style="color:var(--mut);font-weight:600;font-size:11px">${esc(DATA.extremler.hedef||"")}</span></h3>
      ${sik.length?`<div class="grup">Sık koşan</div>${sik.map(sat).join("")}`:""}
      ${uzun.length?`<div class="grup">Uzun ara</div>${uzun.map(sat).join("")}`:""}
      ${gecler.some(e=>e.tip==="gec")?`<div class="grup">⚠ Geç çıkış</div>${gecler.filter(e=>e.tip==="gec").map(gsat).join("")}`:""}
      ${gecler.some(e=>e.tip==="dsiz")?`<div class="grup">⚠ Derecesiz</div>${gecler.filter(e=>e.tip==="dsiz").map(gsat).join("")}`:""}
      ${(!sik.length&&!uzun.length&&!gecler.length)?'<div class="bos">bu koşuda dikkat gerektiren at yok</div>':""}
    </div>`;
  }

  const ic=document.getElementById("icerik");
  if(secTab==="AGF"){
    ic.innerHTML=`<div class="kart agf"><div class="big">🎯</div>
      <h3>AGF — Altılı Ganyan Favorisi</h3>
      <p style="margin-top:6px">Bu ekran yakında: koşu bazında AGF yüzdeleri ve favori analizi.</p></div>`;
    return;
  }
  const b=blokBul(secTab,secIl,secKosu);
  if(!b){ ic.innerHTML='<div class="kart bos">Bu koşu için veri yok.</div>'; return; }
  ic.innerHTML=kartHTML(b);
  // DİKKAT PANELİ HİZASI: sağ kenarı Toplam Derece tablosunun sağ kenarına çekilir
  requestAnimationFrame(()=>{
    const ub=document.querySelector(".ustblok"), d=document.querySelector("#icerik .detay");
    if(ub&&d){
      const w=d.getBoundingClientRect().right - ub.getBoundingClientRect().left;
      ub.style.maxWidth=Math.max(620, Math.round(w))+"px";
    }
  });
  // KAZANAN İŞARETİ HER YERDE: galop satırları, orjin tabloları, dikkat paneli.
  // (Detay tablolarındaki ✓ kartHTML içinde basılır; burası numara geçen diğer yerler.)
  (()=>{
    const _sn=(typeof sonucGecerli==="function"&&sonucGecerli())
      ?(((SONUC.iller||{})[secIl]||{})[String(secKosu)]):null;
    if(!_sn||_sn.kazanan==null) return;
    const kz=String(_sn.kazanan);
    document.querySelectorAll("#icerik .g-satir .atno, #icerik .tablolar.hepsi td.no, #yan .ex-satir .kno")
      .forEach(e=>{
        if(e.textContent.replace("✓","").trim()===kz){
          e.classList.add("kzn");
          if(!e.textContent.includes("✓")) e.textContent=e.textContent.trim()+" ✓";
        }
      });
  })();
  // VARSAYILAN: detay tablosu Derece'ye göre KÜÇÜKTEN BÜYÜĞE açılır
  ic.querySelectorAll(".detay table").forEach(t=>{ if(t.dataset.drc) siralaUygula(t, t.dataset.drc); });
}

function esc(s){return String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));}

function galopSatir(r){
  const p=(r[0]||"").split("/").map(x=>x.trim());
  const at=p[0]||"", deg=p[1]||"", sek=p[2]||"";
  const neg=deg.startsWith("-");
  return `<div class="g-satir"><span class="atno">${esc(at)}</span>
    <span class="gdeger ${neg?"neg":"poz"}">${esc(deg)}</span>
    ${sek?`<span class="sekil">${esc(sek)}</span>`:""}
    <span class="gtarih">${esc(r[1]||"")}${r[2]?`<br><span class="gsehir">${esc(r[2])}</span>`:""}</span></div>`;
}

function tabloHTML(t, zemin){
  // Mesafe tablosu başlığı KOŞUNUN ZEMİNİNE göre: Kum/Çim -> "Bu Zemin ve Mesafedeki",
  // Sentetik -> "Bu Mesafedeki" (sentetikte zemin ayrımı yok)
  const sentetik=String(zemin||"").toUpperCase().includes("SENTET");
  const mesafeAd=sentetik?"Bu Mesafedeki Payı":"Bu Zemin ve Mesafedeki Payı";
  // SIRALAMA: Değer büyükten küçüğe; EŞİTSE Sayı'nın payı (267/1726 -> 267) büyük olan ÜSTTE
  const trows=t.rows.slice().sort((a,b)=>{
    const dv=s=>{const f=parseFloat(String(s).replace(",","."));return isNaN(f)?-Infinity:f;};
    const nv=s=>{const n=parseInt(String(s).split("/")[0],10);return isNaN(n)?-1:n;};
    return (dv(b[1])-dv(a[1])) || (nv(b[2])-nv(a[2]));
  });
  let maxd=0;
  for(const r of trows){const v=parseFloat(String(r[1]).replace(",","."));if(!isNaN(v))maxd=Math.max(maxd,Math.abs(v));}
  const rows=trows.map((r,i)=>{
    const v=parseFloat(String(r[1]).replace(",","."));
    const w=(maxd>0&&!isNaN(v))?Math.round(Math.abs(v)/maxd*100):0;
    return `<tr><td class="no">${esc(r[0])}</td>
      <td class="deg"><div class="bar" style="width:${w}%"></div><span>${esc(r[1])}</span></td>
      <td class="say">${esc(r[2])}</td></tr>`;
  }).join("");
  const ORJ_AD={"Kalite":"Babanın Yavruları: Elit Kazanma","Mesafe":"Babanın Yavruları: "+mesafeAd,
                "Sprinter":"Babanın Yavruları: Sprintle Kazanma","Kaçak":"Babanın Yavruları: Kaçarak Kazanma",
                "Dede Kalite":"Dedenin Yavruları: Elit Kazanma","Dede Mesafe":"Dedenin Yavruları: "+mesafeAd};
  // Sayı başlığı tabloya özel AÇIKLAMALI (mutfak eşiği ifşa edilmeden)
  const SAYI_AD={"Kalite":"Elit Kazanma / Toplam Kazanma","Mesafe":"Bu Mesafede / Toplam Kazanma",
                 "Sprinter":"Sprintle / Toplam Kazanma","Kaçak":"Kaçarak / Toplam Kazanma",
                 "Dede Kalite":"Elit Kazanma / Toplam Kazanma","Dede Mesafe":"Bu Mesafede / Toplam Kazanma"};
  return `<div class="tablo t-${t.name.replace(/\s+/g,"-")}"><h4><span class="nk"></span>${ORJ_AD[t.name]||t.name}</h4>
    <table><tr><th>At No</th><th>Değer</th><th>${SAYI_AD[t.name]||"Sayı"}</th></tr>${rows}</table></div>`;
}

// MOBİLDE GİZLENEN ikincil kolonlar (telefonda sade görünüm; 'tüm kolonlar' ile açılır)
const MGIZ = {"Sayfa1":[3,4,6,10,11,13,16,18,22,23,24,31,32], "Sayfa2":[4,5,6,9,10,12,16,17,18]};
let mTum=false;   // true = telefonda da tüm kolonlar
const BASLIK_TD = ["At No","At","Y.Kilo","E.Kilo","Kilo Farkı","Koşu Cinsi","?7","?8","Tarih","Şehir",
  "Zemin","?12","Pist","Mesafe","Derece","M.Derece","?17","Yaş/Cins","?19","Kaçıncı","",
  "İlk 3 HP'si","Güncel HP","?24","Genel HP Avantajı","Genel Kilo Avantajı","Orta HP Avantajı","Orta Kilo Avantajı","İnce HP Avantajı","İnce Kilo Avantajı",
  "Koşu Cinsi","?32","","Seyir","Üçgen"];
const BASLIK_S8 = ["At No","At","Tarih","Şehir","Zemin","Pist","?7","Mesafe","Y.Kilo","E.Kilo",
  "Kilo Farkı","Koşu Cinsi","?13","Son 800","Fark","Net Son 800","?17","?18","Kaçıncı",
  "Genel HP Avantajı","Genel Kilo Avantajı","Orta HP Avantajı","Orta Kilo Avantajı","İnce HP Avantajı","İnce Kilo Avantajı","Seyir","Üçgen","F.Üçgen"];

// SİLİNECEK KOLONLAR (1-tabanlı): at-no tekrarları + H, P, Q + mükerrer koşu cinsi (AX)
const SIL = {"Sayfa1":[7,8,12,16,17,19,24,31,32], "Sayfa2":[7,13,14,15,17,18,28]};   // Son 800 (14) + Fark (15) gizli; P/Net Son 800 (16) görünür; F.Üçgen (28) gizli
// DOMİNANS kolonları: artı YEŞİL, eksi KIRMIZI
const DOMK = {"Sayfa1":[25,26,27,28,29,30], "Sayfa2":[20,21,22,23,24,25]};
// SON HP: HP listesinin yalnız SON değeri gösterilir (AP sütunu isteği)
const SONHP = {"Sayfa1":[23], "Sayfa2":[]};
// SEYİR: stil sayısı + üçgen TEK hücrede (başlık 'Seyir')
const SEYIR = {"Sayfa1":[34,35], "Sayfa2":[26,27]};
// YAŞ/CİNS kodu çevirisi: '4y a a' -> son harf a/i/e/g=Erkek, k/d=Dişi
const YASCINS = {"Sayfa1":[18], "Sayfa2":[]};
function yasCinsCevir(v){
  const m=String(v).trim().match(/^(\d+\s*y)\b/i);
  if(!m) return v;
  const harfler=String(v).replace(/\d+\s*y/i,"").replace(/[^a-zçğıöşü]/gi,"").toLowerCase();
  const son=harfler.slice(-1);
  if("aieg".includes(son)) return m[1].replace(/\s+/g,"")+" Erkek";
  if("kd".includes(son))   return m[1].replace(/\s+/g,"")+" Dişi";
  return v;
}

function fmt(v){
  if(/^\d{2}\.\d{2}\.\d{4}$/.test(String(v).trim())) return v;   // tarihlere DOKUNMA (yıl kalsın)
  const f=parseFloat(String(v).replace(",","."));
  if(!isNaN(f) && String(v).length>8 && String(v).includes(".")) return f.toFixed(2);
  return v;
}
function detayHTML(b, tab){
  const rows=b.detay||[];
  // KAZANAN: koşu bittiyse kazanan at no (✓ işareti için)
  const _snk=(typeof sonucGecerli==="function"&&sonucGecerli())
    ?(((SONUC.iller||{})[b.header["İl"]]||{})[String(b.header["Koşu No"])]):null;
  const kazananNo=(_snk&&_snk.kazanan!=null)?String(_snk.kazanan):null;
  if(!rows.length) return '<div class="bos" style="padding:12px">Bu koşunun atlarının geçmiş koşu kaydı yok — ilk kez koşacak taylar (örn. 2 yaşlılar) olabilir.</div>';
  const basliklar = tab==="Sayfa1"? BASLIK_TD : BASLIK_S8;
  const sil=new Set((SIL[tab]||[]).map(x=>x-1));
  const domk=new Set((DOMK[tab]||[]).map(x=>x-1));
  const domList=(DOMK[tab]||[]).map(x=>x-1);
  const domKilo=new Set(domList.filter((_,i)=>i%2===1));   // kilo dominansı gizli: 2x'i HP'ye katılır
  const domGrp=c=>{const i=domList.indexOf(c);return i<0?"":(i<2?"g50":(i<4?"g66":"g75"));};
  const domBas=c=>{const i=domList.indexOf(c);return (i===0||i===2||i===4)?" gbas":"";};
  const sonhp=new Set((SONHP[tab]||[]).map(x=>x-1));
  const yascins=new Set((YASCINS[tab]||[]).map(x=>x-1));
  const [seyA,seyB]=(SEYIR[tab]||[0,0]).map(x=>x-1);
  const dolu=[];
  for(let c=0;c<42;c++){
    if(sil.has(c)||c===seyB||domKilo.has(c)||c===41) continue;   // 41 = GECMIS_YOK işaretleyici, görünmez
    if(rows.some(r=>String(r[c]??"").trim()!=="") || c===seyA || (tab==="Sayfa2"&&c===18)) dolu.push(c);   // S8: Kaçıncı sentetik
  }
  // Handikap ile İlk 3 HP yer değiştirir (Handikap önce)
  if(tab==="Sayfa1"){
    const i1=dolu.indexOf(21), i2=dolu.indexOf(22);
    if(i1>-1&&i2>-1){ dolu[i1]=22; dolu[i2]=21; }
  }
  // SON 800: kolon sırası Toplam Derece iskeletine çekilir (aynı bilgi aynı yerde)
  //   No · At · Y.Kilo · E.Kilo · Kilo Farkı · Koşu Cinsi · Tarih · Şehir · Zemin · Pist · Mesafe · Son 800 · Fark · ...
  if(tab==="Sayfa2"){
    const isk=[0,1,8,9,10,11,2,3,4,5,7,13,14];
    dolu.sort((a,b)=>{
      const ia=isk.indexOf(a), ib=isk.indexOf(b);
      return (ia<0?isk.length+a:ia)-(ib<0?isk.length+b:ib);
    });
  }
  // filtreli kolonlar + otomatik sıralama kolonu (Derece / Son 800)
  const FK = tab==="Sayfa1" ? {no:0,sehir:9,msf:13,tarih:8,drc:14,kf:4,zemin:10} : {no:0,sehir:3,msf:7,tarih:2,drc:15,kf:10,zemin:4};   // S8 sıralama: Net Son 800 (c15)
  // KOLON TAŞIMA: kol'u hedef'in hemen soluna getirir
  const tasi=(kol,hedef)=>{ const i=dolu.indexOf(kol); if(i<0) return;
    dolu.splice(i,1); const j=dolu.indexOf(hedef);
    if(j<0){ dolu.splice(i,0,kol); return; } dolu.splice(j,0,kol); };
  tasi(FK.tarih, FK.drc);                 // Tarih -> Derece'nin soluna (iki tabloda da)
  if(tab==="Sayfa1") tasi(17,22);         // Yaş/Cins -> Güncel HP'nin soluna
  // DERECE'den hemen sonra: P50 · P66 · P75 · Yarıştaki Seyri · Kaçıncı, sonra kalanlar
  {
    const kacinci = tab==="Sayfa1" ? 19 : 18;
    const zincir=[...domList.filter(c=>!domKilo.has(c)), seyA, kacinci];
    let hIdx=dolu.indexOf(FK.drc);
    for(const kol of zincir){
      const i=dolu.indexOf(kol); if(i<0) continue;
      dolu.splice(i,1);
      if(i<hIdx) hIdx--;
      dolu.splice(hIdx+1,0,kol);
      hIdx++;
    }
  }
  const kutuPanel=(c)=>{
    const uniq=[...new Set(rows.map(r=>String(r[c]??"").trim()).filter(v=>v!==""))];
    if(uniq.length<2) return "";
    uniq.sort((a,b)=>{const x=parseFloat(a.replace(",",".")),y=parseFloat(b.replace(",","."));
      return (!isNaN(x)&&!isNaN(y))?x-y:a.localeCompare(b,"tr");});
    return `<span class="fok" onclick="fAc(this,event)">▾</span><div class="fpanel" data-c="${c}" onclick="event.stopPropagation()">
      ${uniq.map(v=>`<label><input type="checkbox" class="fv" value="${esc(v)}" checked onchange="fUygula(this)"> ${esc(v)}</label>`).join("")}
      <div class="fayrac"></div>
      <label class="ftumu-l"><input type="checkbox" class="ftumu" checked onchange="fTumu(this)"> <b>Tümü</b></label>
    </div>`;
  };
  const secimPanel=(c,tur,secenekler,fn)=>`<span class="fok" onclick="fAc(this,event)">▾</span>
    <div class="fpanel" data-c="${c}" data-tur="${tur}" onclick="event.stopPropagation()">
      ${secenekler.map(([v,ad])=>`<label><input type="radio" name="f_${tur}_${tab}_${c}" class="fr" value="${v}"${v===""?" checked":""} onchange="${fn||"fUygula"}(this)"> ${ad}</label>`).join("")}
    </div>`;
  const th=dolu.map(c=>{
    let ad=c===seyA?"Yarıştaki Seyri":(basliklar[c]??("K"+(c+1)));
    if(domk.has(c)){ const g0=domGrp(c);
      ad=g0==="g50"?"P50 HP/KG Avantajı":(g0==="g66"?"P66 HP/KG Avantajı":"P75 HP/KG Avantajı"); }
    let f="";
    if(c===FK.no||c===FK.sehir||c===FK.msf||c===FK.zemin) f=kutuPanel(c);
    else if(c===FK.tarih&&tab==="Sayfa1") f=secimPanel(c,"trh",[["60","Son 2 Ay"],["","Tümü"]]);
    else if(c===seyA) f=secimPanel(c,"sey",[["KACAK","Kaçak"],["","Tümü"]]);
    else if(c===FK.kf&&tab==="Sayfa1") f=secimPanel(c,"kf",[["SIRALA","Küçükten büyüğe"],["","Normal"]],"kFarkSirala");
    const g=domGrp(c);
    const mg=(MGIZ[tab]||[]).includes(c+1)?" m-gizle":"";
    const siralanir=false;   // başlığa tıklayınca sıralama YOK (Derece otomatik, Kilo Farkı panelden)
    const dar=((tab==="Sayfa1"&&(c===19||c===21||c===22))||(tab==="Sayfa2"&&c===18)||c===FK.kf||c===FK.no||c===FK.msf||c===FK.zemin)?" dar":"";
    const adH=(c===seyA)?"Yarıştaki<br>Seyri":esc(ad);   // Seyir başlığı iki satır (dar kalsın)
    return `<th data-c="${c}" class="${g}${g?domBas(c):""}${mg}${siralanir?" siralanir":""}${dar}"${siralanir?' onclick="tabloSirala(this)" title="küçükten büyüğe sırala"':""}>${adH}${f}</th>`;
  }).join("");
  const DOM_YOK=new Set(["ŞARTLI 1","SARTLI 1","ŞARTLI 19","SARTLI 19"]);
  const KCIN = tab==="Sayfa1" ? 5 : 11;   // satırın koşu cinsi kolonu (0-tabanlı)
  let _oncekiAt="";
  const trs=rows.map(r=>{
    // KAYIP AT: geçmişi olmayan at tek satırla gösterilir (görünmez olmaz)
    if(String(r[41]||"")==="GECMIS_YOK"){
      _oncekiAt=String(r[1]??"").trim().toUpperCase();
      return `<tr class="grupbas gecmisyok"><td class="num drc" data-c="0" data-v="${esc(r[0])}">${esc(r[0])}</td>`+
        `<td class="atadi" data-c="1" data-v="${esc(r[1])}">${esc(r[1])}</td>`+
        `<td colspan="${dolu.length-2}" style="text-align:left;color:#8a93a5;font-style:italic;padding-left:10px">`+
        `son 6 ayda ilk-4 derecesi bulunmuyor</td></tr>`;
    }
    const kcinsUp=String(r[KCIN]??"").trim().toUpperCase().replace(/\s+/g," ");
    const domGizle=DOM_YOK.has(kcinsUp)||kcinsUp.startsWith("MAIDEN");   // Maiden (+/Dişi) dominans YOK
    const atBase=String(r[1]??"").trim().toUpperCase().replace(/\d+$/,"").trim();
    const grupBas=(atBase!==_oncekiAt); _oncekiAt=atBase;
    return `<tr${grupBas?' class="grupbas"':''}>`+dolu.map(c=>{
    let v=String(r[c]??"").trim();
    let cls="";
    const _kzn=(kazananNo&&String(r[0]??"").trim()===kazananNo);   // bu at koşuyu KAZANDI
    if(c===1){ cls="atadi"+(_kzn?" kazanan":""); v=v.replace(/\d+$/,"").trim(); if(_kzn&&grupBas) v="✓ "+v; }   // TESCİL1 -> TESCİL
    else if(c===FK.no){ cls="num drc"+(_kzn?" kazanan":""); }            // At No: Derece gibi KOYU
    else if(c===seyA){ v=String(r[seyB]??"").trim(); cls="ucgen"; }      // YALNIZ ÜÇGEN
    else if(c===FK.tarih){ v=tarihGoster(v); cls="num drc"; }            // tarih: Derece gibi KOYU
    else if(c===FK.drc){ cls="num drc"; }   // Derece: kalın siyah (boyasız)
    else if(c===FK.kf){                     // Kilo Farkı: EKSİ=avantaj(yeşil), ARTI=dezavantaj(kırmızı)
      const fk=parseFloat(v.replace(",","."));
      cls="num "+(fk<0?"kfiyi":(fk>0?"kfkotu":""));
    }
    else if(tab==="Sayfa1"&&c===19){ cls="orta"; if(/^\d+$/.test(v)) v=v+"."; }   // Kaçıncı: 2. 3. gibi
    else if(tab==="Sayfa2"&&c===18){ cls="orta"; v="1."; }                        // S8 Kaçıncı: hepsi 1. (kazananın son 800'ü)
    else if(basliklar[c]==="Koşu Cinsi"){ cls="kcins"; }                          // Koşu Cinsi: dar + kaydırmalı
    else if(sonhp.has(c)){
      const p=v.split("-").filter(x=>x.trim()!=="");
      v=p.length?p[p.length-1]:v; cls="num";
    }
    else if(domk.has(c)){                       // BİRLEŞİK AVANTAJ: HP + 2 x Kilo (tek sayı, RENKSİZ)
      if(domGizle){                                 // Maiden / Şartlı 1 / Şartlı 19 -> ortada soluk etiket, yanlarda tire
        const kisa=kcinsUp.startsWith("MAIDEN")?"Maiden":(kcinsUp.includes("19")?"Şartlı 19":"Şartlı 1");
        v=(c===domList[2])?kisa:"—"; cls="num dom yokd";
      }
      else{
        const hp=parseFloat(v.replace(",","."));
        const kl=parseFloat(String(r[c+1]??"").replace(",","."));
        const f2=isNaN(hp)?NaN:(hp + 2*(isNaN(kl)?0:kl));
        cls="num dom "+(f2>0?"poz":(f2<0?"neg":""));
        if(!isNaN(f2)){ const yv=Math.round(f2*100)/100; v=(yv>0?"+":"")+yv; } else v="";
      }
    }
    else if(yascins.has(c)){ v=yasCinsCevir(v); cls="yc"; }
    else if(v.includes("▶")||v.includes("▷")) cls="ucgen";
    else if(/^-?\d/.test(v)) cls="num";
    const dv=(c===seyA)?String(r[seyB]??"").trim():String(r[c]??"").trim();
    const mg2=(MGIZ[tab]||[]).includes(c+1)?" m-gizle":"";
    return `<td class="${cls}${mg2}" data-c="${c}" data-v="${esc(dv)}">${esc(fmt(v))}</td>`;
  }).join("")+"</tr>";}).join("");
  return `<div class="detay"><table data-drc="${FK.drc}"><tr>${th}</tr>${trs}</table></div>`;
}

function drcSaniye(v){
  const m=String(v).match(/^(\d+):(\d+)[,.](\d+)$/);
  return m? parseInt(m[1])*6000+parseInt(m[2])*100+parseInt(m[3]) : null;
}
function tarihGoster(v){ // her tarih dd.mm.yyyy (yıl DAHİL) görünsün
  const m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return m[3]+"."+m[2]+"."+m[1];
  return v;
}
function fAc(el,ev){
  if(ev) ev.stopPropagation();
  let p=el._panel||el.nextElementSibling;
  if(!p||!p.classList.contains("fpanel")) return;
  document.querySelectorAll(".fpanel.acik").forEach(x=>{if(x!==p)x.classList.remove("acik");});
  if(p.classList.contains("acik")){p.classList.remove("acik");return;}
  if(!el._panel){                     // ilk açılış: EN ÜST KATMANA taşı (hiçbir şey örtemez)
    const tbl=el.closest("table");
    p._tbl=tbl; p._ok=el; el._panel=p;
    (tbl._pnl=tbl._pnl||[]).push(p);
    document.body.appendChild(p);
  }
  p.classList.add("acik");
  const r=el.getBoundingClientRect();
  // SAYFAYA GÖRE konum (absolute): zoom ve kaydırmayla birlikte hareket eder,
  // okun HEMEN ÜSTÜNDE açılır (yukarı doğru), telefonda da şaşmaz.
  p.style.position="absolute";
  p.style.zIndex="9999";
  p.style.right="auto"; p.style.width=""; p.style.bottom="auto";
  const sx=window.scrollX||document.documentElement.scrollLeft;
  const sy=window.scrollY||document.documentElement.scrollTop;
  p.style.left=Math.max(8,sx+r.left-4)+"px";
  p.style.top=(sy+r.top-6)+"px";
  p.style.transform="translateY(-100%)";
  p.style.maxHeight="300px";
}
document.addEventListener("click",e=>{
  if(!e.target.closest(".fpanel")&&!e.target.closest(".fok"))
    document.querySelectorAll(".fpanel.acik").forEach(x=>x.classList.remove("acik"));
});
window.addEventListener("resize",()=>{
  document.querySelectorAll(".fpanel.acik").forEach(x=>x.classList.remove("acik"));
});
function fTumu(cb){
  cb.closest(".fpanel").querySelectorAll(".fv").forEach(x=>x.checked=cb.checked);
  fUygula(cb);
}
function trhSayi(v){
  const m=String(v).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  return m? new Date(m[3]+"-"+m[2]+"-"+m[1]) : null;
}
function sortVal(v){
  const dm=v.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if(dm) return parseInt(dm[3]+dm[2]+dm[1]);
  const dr=v.match(/^(\d+):(\d+)[,.](\d+)$/);
  if(dr) return parseInt(dr[1])*6000+parseInt(dr[2])*100+parseInt(dr[3]);
  const f=parseFloat(v.replace(",","."));
  if(!isNaN(f)&&/^[-\d]/.test(v)) return f;
  return null;
}
function siralaUygula(table,c){ // küçükten büyüğe
  const rows=[...table.querySelectorAll("tr")].filter(tr=>!tr.querySelector("th"));
  rows.sort((a,b)=>{
    const ta=a.querySelector(`td[data-c="${c}"]`), tb=b.querySelector(`td[data-c="${c}"]`);
    const x=sortVal(ta?ta.dataset.v:""), y=sortVal(tb?tb.dataset.v:"");
    if(x!==null&&y!==null) return x-y;
    if(x!==null) return -1; if(y!==null) return 1;
    return String(ta?ta.dataset.v:"").localeCompare(String(tb?tb.dataset.v:""),"tr");
  });
  rows.forEach(r=>table.appendChild(r));
}
function kFarkSirala(inp){
  const p=inp.closest(".fpanel"); const c=p.dataset.c;
  const table=p._tbl||inp.closest("table");
  const th=table.querySelector(`th[data-c="${c}"]`);
  table.querySelectorAll("th.sirali").forEach(x=>x.classList.remove("sirali"));
  const ok=p._ok||p.previousElementSibling;
  if(inp.value==="SIRALA"){
    th.classList.add("sirali");
    siralaUygula(table,c);                       // K.Fark küçükten büyüğe
    if(ok) ok.classList.add("aktif");
  }else{
    siralaUygula(table, table.dataset.drc);       // normale dön -> Derece sıralı
    if(ok) ok.classList.remove("aktif");
  }
}
function fUygula(el){
  const _pan=el.closest(".fpanel");
  const table=(_pan&&_pan._tbl)||el.closest("table");
  const kutu=[], radio=[];
  const _paneller=[...table.querySelectorAll(".fpanel"),...(table._pnl||[])];
  _paneller.forEach(p=>{
    if(p.dataset.tur==="kf") return;              // K.Fark paneli sıralama işidir, filtre değil
    if(p.dataset.tur){
      const sec=p.querySelector(".fr:checked");
      if(sec&&sec.value) radio.push([p.dataset.c,p.dataset.tur,sec.value]);
      return;
    }
    const vals=[...p.querySelectorAll(".fv")];
    const secili=vals.filter(x=>x.checked).map(x=>x.value);
    const t=p.querySelector(".ftumu"); if(t) t.checked=(secili.length===vals.length);
    // HİÇBİRİ seçili değilse filtre YOK sayılır (tablo boşalmaz)
    if(secili.length&&secili.length!==vals.length) kutu.push([p.dataset.c,new Set(secili)]);
  });
  const hedef=trhSayi((DATA.extremler&&DATA.extremler.hedef)||"");
  table.querySelectorAll("tr").forEach(tr=>{
    if(tr.querySelector("th")) return;
    let g=true;
    for(const [c,set] of kutu){
      const td=tr.querySelector(`td[data-c="${c}"]`);
      if(!td||!set.has(td.dataset.v)){g=false;break;}
    }
    if(g) for(const [c,tur,v] of radio){
      const td=tr.querySelector(`td[data-c="${c}"]`);
      if(tur==="trh"&&hedef){
        const d=td?trhSayi(td.dataset.v):null;
        const esik=new Date(hedef); esik.setDate(esik.getDate()-parseInt(v));
        if(!d||d<esik){g=false;break;}
      }
      if(tur==="sey"&&v==="KACAK"){
        if(!td||td.dataset.v!=="▷▷▷▶"){g=false;break;}
      }
    }
    tr.style.display=g?"":"none";
  });
  // AKTİF FİLTRE GÖSTERGESİ: filtre uygulanan kolonun oku turuncu yanar
  _paneller.forEach(p=>{
    if(p.dataset.tur==="kf") return;
    const ok=p._ok||p.previousElementSibling;
    let aktif=false;
    if(p.dataset.tur){ const sec=p.querySelector(".fr:checked"); aktif=!!(sec&&sec.value); }
    else{ const vals=[...p.querySelectorAll(".fv")]; const sc=vals.filter(x=>x.checked).length;
          aktif=(sc>0&&sc<vals.length); }
    if(ok&&ok.classList.contains("fok")) ok.classList.toggle("aktif",aktif);
  });
  // FİLTRE SONRASI: kullanıcı sıralaması yoksa Derece'ye göre küçükten büyüğe
  if(!table.querySelector("th.sirali")) siralaUygula(table, table.dataset.drc);
}
function tabloSirala(th){ // 1. tık: küçükten büyüğe (▲) | 2. tık: eski diziliş
  const table=th.closest("table"); const c=th.dataset.c;
  const rows=[...table.querySelectorAll("tr")].filter(tr=>!tr.querySelector("th"));
  if(!table.dataset.oi){ rows.forEach((r,i)=>r.dataset.oi=i); table.dataset.oi="1"; }
  const aktif=th.classList.contains("sirali");
  table.querySelectorAll("th.sirali").forEach(x=>x.classList.remove("sirali"));
  if(aktif){ rows.sort((a,b)=>(+a.dataset.oi)-(+b.dataset.oi)); rows.forEach(r=>table.appendChild(r)); table.classList.remove("sirali-tablo"); return; }
  th.classList.add("sirali");
  siralaUygula(table,c);
}

function drcFmt(v){
  // Derece santisaniye (7695 -> 1.16.95, 14005 -> 2.20.05); sayı değilse aynen
  const n=parseInt(String(v).trim(),10);
  if(isNaN(n)||n<=0) return v;
  const cs=n%100, tot=Math.floor(n/100), dk=Math.floor(tot/60), sn=tot%60;
  const p=x=>String(x).padStart(2,"0");
  return (dk>0? dk+"."+p(sn) : String(sn))+"."+p(cs);
}
let derF={sehir:"",msf:"",ay:""};   // Dereceler filtresi (koşu değişince sıfırlanır)
function derFsec(alan,deger){ derF[alan]=(derF[alan]===deger?"":deger); ciz(); }

function derecelerBolum(){
  const der=DATA.dereceler||{};
  if(!Object.keys(der).length) return "";
  // seçili koşunun atları (800 detayından: at no + ad, sırayla)
  const b=blokBul("Sayfa2",secIl,secKosu)||blokBul("Sayfa1",secIl,secKosu);
  const atlar=[]; const gorulen=new Set();
  for(const r of ((b&&b.detay)||[])){
    let at=String(r[1]??"").trim().toUpperCase().replace(/\d+$/,"").trim();
    const no=String(r[0]??"").trim();
    if(at&&!gorulen.has(at)){gorulen.add(at);atlar.push([no,at]);}
  }
  if(!atlar.length) return "";
  // FİLTRE: yalnız Şehir + Mesafe + tarih aralığı (Son 2 Ay / Tümü)
  const secim={sehir:new Set(),msf:new Set()};
  for(const [,at] of atlar) for(const k of (der[at]||[])){
    if(k[1])secim.sehir.add(k[1]); if(k[4])secim.msf.add(k[4]);
  }
  const fsel=(alan,etiket,degerler)=>`<select class="fdrop" onchange="derF['${alan}']=this.value;ciz()">
      <option value="">${etiket} ▾ tümü</option>
      ${degerler.map(v=>`<option${derF[alan]===v?" selected":""}>${esc(v)}</option>`).join("")}</select>`;
  const filtreler=`<div class="secici" style="margin:4px 0 10px">
    <span class="lbl">Filtre</span>
    ${fsel("sehir","Şehir",[...secim.sehir].sort((a,b)=>a.localeCompare(b,"tr")))}
    ${fsel("msf","Mesafe",[...secim.msf].sort((a,b)=>parseInt(a)-parseInt(b)))}
    <select class="fdrop" onchange="derF.ay=this.value;ciz()">
      <option value="">Tarih ▾ tümü</option>
      <option value="60"${derF.ay==="60"?" selected":""}>Son 2 Ay</option></select>
    ${(derF.sehir||derF.msf||derF.ay)?`<button class="chip fchip" style="border-color:var(--acc);color:var(--acc)" onclick="derF={sehir:'',msf:'',ay:''};ciz()">✕ temizle</button>`:""}
  </div>`;
  let esik=null;
  if(derF.ay){
    const h=trhSayi((DATA.extremler&&DATA.extremler.hedef)||"");
    if(h){esik=new Date(h); esik.setDate(esik.getDate()-parseInt(derF.ay));}
  }
  const gruplar=atlar.map(([no,at])=>{
    const kosular=(der[at]||[]).filter(k=>{
      if(derF.sehir&&k[1]!==derF.sehir) return false;
      if(derF.msf&&k[4]!==derF.msf) return false;
      if(esik){const d=trhSayi(k[0]); if(!d||d<esik) return false;}
      return true;});
    const trs=kosular.map(k=>{
      const sey=(k[9]&&k[10])?`${String(k[9]).replace(/\.0$/,"")} ${k[10]}`:(k[10]||String(k[9]||"").replace(/\.0$/,""));
      const dv=(i)=>`data-c="${i}" data-v="${esc(String(k[i]??"").trim())}"`;
      return `<tr><td ${dv(0)}>${esc(k[0])}</td><td ${dv(1)}>${esc(k[1])}</td><td ${dv(2)}>${esc(k[2])}${k[3]?" · "+esc(k[3]):""}</td>
        <td class="num" ${dv(4)}>${esc(k[4])}</td><td class="num" style="font-weight:700" ${dv(5)}>${esc(drcFmt(k[5]))}</td>
        <td ${dv(6)}>${esc(k[6])}</td><td class="num" ${dv(7)}>${esc(k[7])}</td><td class="num" ${dv(8)}>${esc(k[8])}</td>
        <td class="ucgen" data-c="9" data-v="${esc(sey)}">${esc(sey)}</td></tr>`;
    }).join("");
    return `<div style="margin-bottom:14px">
      <div style="display:flex;align-items:center;gap:8px;margin:0 0 6px">
        <span class="atno">${esc(no)}</span><b>${esc(at)}</b>
        <span style="color:var(--mut);font-size:12px">${kosular.length} koşu</span></div>
      ${kosular.length?`<div class="detay"><table>
        <tr>${["Tarih","Şehir","Pist","Msf","Derece","Koşu Cinsi","Sıra","HP","Seyir"].map((b,i)=>`<th data-c="${i}">${b}</th>`).join("")}</tr>
        ${trs}</table></div>`:'<div class="bos">derece kaydı yok</div>'}
    </div>`;
  }).join("");
  return `<div class="bol-baslik">Dereceler</div>${filtreler}${gruplar}`;
}

// AGF TABLOSU AÇ: ikinci=true -> 2. altılı. Gerçek link yoksa şehir program sayfasına düşer.
function agfAc(ikinci){
  const SL=DATA.sehir_link||{};
  const gecerli=x=>(typeof x==="string"&&/^https?:\/\//i.test(x))?x:"";   // "javascript: void(0)" gibi sahte linkler ELENİR
  const agfU=gecerli(((ikinci?SL.agf2:SL.agf1)||{})[secIl]);
  const ham=agfU||gecerli(SL.agf&&SL.agf[secIl])||gecerli(SL.program&&SL.program[secIl])||gecerli(SL[secIl])||"";
  // Info/Sehir = süssüz iç parça -> Info/Page = tam görünümlü sayfa (şehir sekmesi seçili)
  const u=ham?ham.replace(/\/Info\/Sehir\//i,"/Info/Page/")
              :"https://www.tjk.org/TR/YarisSever/Info/Page/GunlukYarisProgrami";
  window.open(u,"_blank");
}
function galopaGit(){
  // Son 800'e en yakın galop AŞAĞIDAKİ (orjinlerin üstündeki) blok — oraya kaydır
  const g=document.getElementById("galopbas2")||document.getElementById("galopbas");
  if(g) g.scrollIntoView({behavior:"smooth",block:"start"});
}

function kartHTML(b){
  const h=b.header;
  const kosuCinsi=String(b.title||"").replace(/^\d+\.\s*Koşu\s*/,"").split(",")[0].trim();
  const kosuOzet=[kosuCinsi, h["Mesafe"], h["Zemin"]].filter(Boolean).join(" ");

  const meta=["Mesafe","Zemin","Irk"].filter(k=>h[k])
    .map(k=>`<span class="m"><b>${k}</b>${esc(h[k])}</span>`).join("");
  const _fv=Math.abs(parseFloat(String(h["Final"]||"").replace(",",".")));
  const _notr=(!isNaN(_fv)&&_fv<10);                          // -10..+10 -> NÖTR
  // YORUM KURALI: çift laf yok — Sprinter baskın ("sonu kuvvetli");
  // Kaçak+Stalker birlikte -> "toplam derecesi iyi"; tekiller kendi karşılığı.
  const _tok=String(h["Yorum"]||"").toUpperCase();
  let yorumCvr="";
  if(_tok.includes("SPRINTER")) yorumCvr="sonu kuvvetli";
  else if((_tok.includes("KAÇAK")||_tok.includes("KACAK"))&&_tok.includes("STALKER")) yorumCvr="toplam derecesi iyi";
  else if(_tok.includes("KAÇAK")||_tok.includes("KACAK")) yorumCvr="önde gitmeyi seven";
  else if(_tok.includes("STALKER")) yorumCvr="toplam derecesi iyi";
  else if(_tok.includes("NÖTR")||_tok.includes("NOTR")) yorumCvr="Nötr";
  else yorumCvr=String(h["Yorum"]||"").trim();
  const yorumSon=_notr?"Nötr":yorumCvr;
  const yorum=yorumSon?`<span class="m vurgu"><b>Yorum</b>${esc(yorumSon)}</span>`:"";
  // Final: ham sayı yerine KADEME kelimesi (mutfak değeri ifşa edilmez)
  const _fk=isNaN(_fv)?"":(_fv<20?"orta":_fv<35?"belirgin":"çok belirgin");
  const final=(h["Final"]&&_fk&&!_notr)?`<span class="m vurgu"><b>Kuvvet</b>${_fk}</span>`:"";
  const gNorm=b.galops.filter(g=>!g.son), gSon=b.galops.filter(g=>g.son);
  const pnl=(g)=>`<div class="panel${g.son?" sonp":""}"><h4>${esc(String(g.name).replace(/\s*\+\s*400/g,""))}</h4>
      ${g.rows.length?g.rows.map(galopSatir).join(""):'<div class="bos">kayıt yok</div>'}</div>`;
  const setA=b.tables.filter(t=>t.col<16);
  // DEDE SADELEŞTİRME: Dede Kalite gösterilmez; Dede Mesafe yalnız UZUN koşuda
  // (İngiliz >=1800m, Arap >=1900m). Veri boru hattında duruyor, sadece gizli.
  const _msfN=parseInt(String(h["Mesafe"]||"").replace(/\D/g,""),10)||0;
  const _arapMi=/ARAP/i.test(String(h["Irk"]||"")+" "+String(b.title||""));
  const _uzunKosu=_arapMi?(_msfN>=1900):(_msfN>=1800);
  const setDede=b.tables.filter(t=>t.col>=16)
                 .map(t=>({...t,name:"Dede "+t.name}))
                 .filter(t=>t.name!=="Dede Kalite")
                 .filter(t=>t.name!=="Dede Mesafe"||_uzunKosu);
  // BAŞLIK ÇİPLERİ: galopa zıpla + AGF aç (8+ koşulu günde iki altılı)
  const _N=kosular(secIl).length;
  const agfCip=`<span class="atlama agfc" onclick="agfAc(false)">${_N>=8?"1. AGF":"AGF"}</span>`+
    (_N>=8?`<span class="atlama agfc" onclick="agfAc(true)">2. AGF</span>`:"");
  const cipler=`<span class="atlama" onclick="galopaGit()">⬇ Galoplar</span>`+agfCip;
  // GALOP BLOĞU: hem sayfa başında hem ORJİNLERİN ÜSTÜNDE aynen gösterilir
  const galopBlok=`<div class="galoprow">
      <div class="grup30"><div class="gbaslik">🟢 30 GÜNLÜK GALOP</div>
        <div class="paneller tekhiza">${gNorm.map(pnl).join("")||'<div class="bos">galop verisi yok</div>'}</div></div>
      <div class="grupson"><div class="gbaslik son">🟠 SON GALOP</div>
        <div class="paneller tekhiza">${gSon.map(pnl).join("")||'<div class="bos">son galop verisi yok</div>'}</div></div>
    </div>`;
  return `
  <div class="kart">
    <div class="kosu-baslik-satir">
      <div class="kosu-baslik">${esc(b.title||(h["Koşu No"]+". Koşu"))}</div>
      ${h["Saat"]?`<span class="kosusaat">🕒 Koşu Saati: ${esc(h["Saat"])}</span>`:""}
    </div>
    <div class="meta">
      <span class="m"><b>İl</b>${esc(h["İl"]||"")}</span>
      <span class="m"><b>Koşu</b>${esc(h["Koşu No"]||"")}</span>
      ${meta}${yorum}${final}
    </div>

    <div class="bol-baslik buyuk" id="galopbas">Galoplar</div>
    ${galopBlok}

    <div class="bol-baslik buyuk">Toplam Derece — Detay · ${kosuOzet}</div>
    <div class="tablonot ustte">Not: Tablo genelinde yeşil yazılar avantajı, kırmızı yazılar dezavantajı gösterir. P50, P66 ve P75 HP/KG Avantajı sütunları, atın o tarihte koştuğu koşu ile şimdi koşacağı koşuyu karşılaştırır ve ata yeni koşuda avantaj mı dezavantaj mı doğduğunu gösterir.</div>
    ${detayHTML(b, "Sayfa1")}

    ${(()=>{const b8=blokBul("Sayfa2", h["İl"], h["Koşu No"]);
      return b8?`<div class="bol-baslik buyuk">Son 800 — Detay · ${kosuOzet}${cipler}</div>
    <div class="tablonot ustte">Not: Tablo genelinde yeşil yazılar avantajı, kırmızı yazılar dezavantajı gösterir. P50, P66 ve P75 HP/KG Avantajı sütunları, atın o tarihte koştuğu koşu ile şimdi koşacağı koşuyu karşılaştırır ve ata yeni koşuda avantaj mı dezavantaj mı doğduğunu gösterir.</div>
    ${detayHTML(b8,"Sayfa2")}`:"";})()}

    <div class="bol-baslik buyuk" id="galopbas2">Galoplar <span class="ayninot">— yukarıdaki galopların aynısı</span></div>
    ${galopBlok}

    <div class="bol-baslik buyuk">Orijin Analizi · ${kosuOzet}${agfCip}</div>
    <div class="grupor"><div class="gbaslik or">🔵 ORİJİN — Babanın ve Dedenin (annenin babası) yavruları</div>
    <div class="tablolar hepsi">${setA.concat(setDede).map(t=>tabloHTML(t, h["Zemin"])).join("")}</div></div>
  </div>`;
}

if(DATA.uretim){const _u=document.getElementById("uretimNot");
  if(_u) _u.textContent="  ·  veri: "+DATA.uretim;}
secIl=iller()[0]; secKosu=kosular(secIl)[0];
// SONUÇLAR: sunucu yarım saatte bir sonuclar.json günceller; sayfa 10 dk'da bir okur.
// Tarih eşleşmezse (dünün sonucu / yarının programı) işaret BASILMAZ.
let SONUC={}, _sonucStr="";
function sonucGecerli(){ return SONUC && SONUC.tarih && DATA.extremler && SONUC.tarih===DATA.extremler.hedef; }
function sonucYukle(){
  try{
    fetch("sonuclar.json",{cache:"no-store"}).then(r=>r.ok?r.json():null).then(d=>{
      if(!d) return;
      const s=JSON.stringify(d);
      if(s!==_sonucStr){ _sonucStr=s; SONUC=d; ciz(); }
    }).catch(()=>{});
  }catch(e){}
}
sonucYukle(); setInterval(sonucYukle, 10*60*1000);
ciz();
</script>
</body>
</html>
"""

html = HTML.replace("__DATA__", json.dumps(DATA, ensure_ascii=False))
open("tjk_analiz_prototip.html", "w", encoding="utf-8").write(html)
print("yazıldı: tjk_analiz_prototip.html", len(html), "byte")
