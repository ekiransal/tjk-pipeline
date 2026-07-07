#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""mockup_parsed.json -> tek dosyalık TJK analiz platformu prototipi (HTML)."""
import json

DATA = json.load(open("mockup_parsed.json", encoding="utf-8"))

HTML = r"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
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
  th.sirali{color:var(--pri) !important}
  th.sirali::after{content:" ▲"}
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
  .detay th.g50,.detay td.g50{background:#e7f3ec !important}
  .detay th.g66,.detay td.g66{background:#e9effa !important}
  .detay th.g75,.detay td.g75{background:#f2ecfa !important}
  
  /* GALOP GRUP KUTULARI: kendi içinde bütün, birbirinden keskin ayrım */
  .grup30{background:#edf6f2;border:1.5px solid #cfe5db;border-radius:14px;padding:10px}
  .grupson{background:#fdf2ea;border:1.5px solid #f2dbc7;border-radius:14px;padding:10px}
  .grup30 .panel,.grupson .panel{background:#fff}

  /* GALOPLAR TEK SATIR: 30 günlük + son galop yan yana, ayrı renkli kutular */
  .galoprow{display:flex;gap:14px;align-items:stretch;overflow-x:auto;padding-bottom:4px}
  .gbaslik{font-size:11px;font-weight:800;letter-spacing:.7px;margin:0 0 8px;color:#1e6f5c}
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
    .chip{padding:9px 14px}.chip.num{width:40px}
  }
  /* SEKMELER */
  .tabs{display:flex;gap:4px;background:var(--card);padding:5px;border-radius:12px;
        box-shadow:var(--sh);margin-bottom:14px;width:max-content;max-width:100%}
  .tab{padding:9px 22px;border-radius:9px;font-weight:700;font-size:13.5px;cursor:pointer;
       color:var(--mut);border:none;background:transparent;transition:.15s}
  .tab.on{background:var(--pri2);color:var(--pri)}
  /* KOŞU KARTI */
  .kart{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:18px 20px;margin-bottom:14px}
  .kosu-baslik{font-size:16.5px;font-weight:800;margin-bottom:10px}
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
  /* GALOP PANELLERİ */
  .paneller{display:grid;grid-template-columns:repeat(auto-fill,minmax(215px,1fr));gap:10px}
  /* TEK HİZA: galop panelleri alta sarkmasın; sığmazsa yatay kaydırılır */
  .paneller.tekhiza{display:flex;flex-wrap:nowrap;overflow-x:auto;padding-bottom:4px}
  .paneller.tekhiza .panel{flex:0 0 215px;min-width:215px}
  .panel{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
  .panel h4{font-size:12.5px;font-weight:800;color:var(--pri);margin-bottom:8px;letter-spacing:.3px}
  .panel.sonp h4{color:var(--acc)}
  .g-satir{display:flex;align-items:center;gap:7px;padding:4px 0;border-bottom:1px dashed var(--line);font-size:12.5px}
  .g-satir:last-child{border-bottom:none}
  .atno{min-width:26px;height:26px;border-radius:8px;background:var(--pri2);color:var(--pri);
        font-weight:800;display:flex;align-items:center;justify-content:center;font-size:12.5px}
  .gdeger{font-weight:700;min-width:48px}
  .gdeger.neg{color:#1a7f4b}.gdeger.poz{color:#c23a3a}
  .sekil{font-size:10.5px;font-weight:800;border-radius:6px;padding:2px 6px;background:#eef0f4;color:var(--mut)}
  .gtarih{color:var(--mut);font-size:11.5px;margin-left:auto;text-align:right}
  .gsehir{color:var(--acc);font-size:11px;font-weight:600}
  .bos{color:var(--mut);font-size:12px;font-style:italic;padding:6px 0}
  .tablonot{color:var(--mut);font-size:11.5px;font-style:italic;margin-top:8px}
  .tablonot.ustte{margin:0 0 6px;text-align:right}
  /* ANALİZ TABLOLARI */
  .tablolar{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px}
  .tablo{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
  .tablo h4{font-size:13px;font-weight:800;padding:11px 14px 9px;color:var(--txt);
            display:flex;align-items:center;gap:7px}
  .tablo h4 .nk{width:9px;height:9px;border-radius:3px}
  .t-Kalite .nk{background:#1e6f5c}.t-Mesafe .nk{background:#2f6fb3}
  .t-Sprinter .nk{background:#c8542a}.t-Kaçak .nk{background:#8a56c9}
  .t-Dede-Kalite .nk{background:#0f4d40}.t-Dede-Mesafe .nk{background:#1d4e80}
  .tablolar.hepsi{grid-template-columns:repeat(6,1fr)}
  @media(max-width:1150px){.tablolar.hepsi{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:700px){.tablolar.hepsi{grid-template-columns:repeat(2,1fr)}}
  @media(max-width:440px){.tablolar.hepsi{grid-template-columns:1fr}}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);
     text-align:left;padding:5px 9px;background:var(--bg)}
  td{padding:5px 9px;border-top:1px solid var(--line)}
  td.no{font-weight:800;width:44px}
  tr.r1 td.no span,tr.r2 td.no span,tr.r3 td.no span{display:inline-flex;width:24px;height:24px;
     border-radius:7px;align-items:center;justify-content:center;color:#fff;font-size:12px}
  tr.r1 td.no span{background:var(--gold)} tr.r2 td.no span{background:#9aa4b5} tr.r3 td.no span{background:#c98d5c}
  td.deg{position:relative;font-weight:700}
  td.deg .bar{position:absolute;left:0;top:15%;bottom:15%;background:var(--pri2);border-radius:0 6px 6px 0;z-index:0}
  td.deg span{position:relative;z-index:1}
  td.say{color:var(--mut);font-size:11.5px;text-align:right;white-space:nowrap}
  /* DETAY TABLOSU — TEK SAYFA: yana kaydırma yerine sığdır (kompakt) */
  .detay{overflow-x:auto;border:1px solid var(--line);border-radius:12px;background:var(--card)}
  .detay table{font-size:11px;white-space:normal;width:100%}
  .detay th{position:sticky;top:0;background:var(--bg);z-index:2;padding:5px 5px;
            white-space:normal;line-height:1.15;border:1px solid #c9cfd9;border-bottom:2px solid #aab2bf;text-align:center}
  .detay td{padding:4px 5px;border:1px solid #c9cfd9;line-height:1.25;text-align:center}
  .detay tr:nth-child(even) td{background:#fafbfc}
  .detay tr:hover td{background:#e2efe8 !important}
  .detay td.atadi{font-weight:700;text-align:center}
  .detay td.ucgen{font-size:13px;letter-spacing:1px;color:#1c2333;white-space:nowrap;text-align:center !important}
  .detay td.orta{text-align:center}
  .detay td.num{text-align:center}
  .detay td.drc{font-weight:800;color:#111}
  .detay td.dom.poz{color:#1a7f4b;font-weight:700}
  .detay td.dom.neg{color:#c23a3a;font-weight:700}
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
    <div class="sub">Günlük koşu analiz platformu — prototip</div>
  </div>

  <div class="ustblok">
    <div class="ustsol">
      <div class="secici"><span class="lbl">İl</span><span id="iller"></span></div>
      <div class="secici"><span class="lbl">Koşu</span><span id="kosular"></span></div>
      <div class="tabs" id="tabs">
        <button class="tab on" data-t="Sayfa1">Toplam Derece</button>
        <button class="tab" data-t="Sayfa2">Son 800 Analizi</button>
        <button class="tab" data-t="AGF">AGF</button>
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
    b.onclick=()=>{secKosu=no; derF={sehir:"",msf:"",ay:""}; ciz();};
    koBox.appendChild(b);
  }
  document.querySelectorAll(".tab").forEach(t=>{
    t.classList.toggle("on",t.dataset.t===secTab);
    t.onclick=()=>{secTab=t.dataset.t; ciz();};
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
    yan.innerHTML=`<div class="kart extrem"><h3>⚠️ Dikkat — ${esc(secKosu)}. Koşu
      <span style="color:var(--mut);font-weight:600;font-size:11px">${esc(DATA.extremler.hedef||"")}</span></h3>
      ${sik.length?`<div class="grup">Sık koşan (≤${DATA.extremler.sik_gun} gün)</div>${sik.map(sat).join("")}`:""}
      ${uzun.length?`<div class="grup">Uzun ara (≥${DATA.extremler.uzun_gun} gün)</div>${uzun.map(sat).join("")}`:""}
      ${(!sik.length&&!uzun.length)?'<div class="bos">bu koşuda extrem at yok</div>':""}
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

function tabloHTML(t){
  let maxd=0;
  for(const r of t.rows){const v=parseFloat(String(r[1]).replace(",","."));if(!isNaN(v))maxd=Math.max(maxd,Math.abs(v));}
  const rows=t.rows.map((r,i)=>{
    const v=parseFloat(String(r[1]).replace(",","."));
    const w=(maxd>0&&!isNaN(v))?Math.round(Math.abs(v)/maxd*100):0;
    const rk=i<3?` class="r${i+1}"`:"";
    const no=i<3?`<span>${esc(r[0])}</span>`:esc(r[0]);
    return `<tr${rk}><td class="no">${no}</td>
      <td class="deg"><div class="bar" style="width:${w}%"></div><span>${esc(r[1])}</span></td>
      <td class="say">${esc(r[2])}</td></tr>`;
  }).join("");
  return `<div class="tablo t-${t.name.replace(/\s+/g,"-")}"><h4><span class="nk"></span>${t.name}</h4>
    <table><tr><th>No</th><th>Değer</th><th>Sayı</th></tr>${rows}</table></div>`;
}

// MOBİLDE GİZLENEN ikincil kolonlar (telefonda sade görünüm; 'tüm kolonlar' ile açılır)
const MGIZ = {"Sayfa1":[3,4,6,11,13,16,18,22,23,27,28,29,30], "Sayfa2":[5,6,9,10,12,14,22,23,24,25,28]};
let mTum=false;   // true = telefonda da tüm kolonlar
const BASLIK_TD = ["No","At","Kilo","E.Kilo","K.Fark","Koşu Cinsi","?7","?8","Tarih","Şehir",
  "Zemin","?12","Pist","Mesafe","Derece","M.Derece","?17","Yaş/Cins","?19","Kaçıncı","",
  "İlk 3 HP","Handikap","?24","Genel HP Avantajı","Genel Kilo Avantajı","Orta HP Avantajı","Orta Kilo Avantajı","İnce HP Avantajı","İnce Kilo Avantajı",
  "Koşu Cinsi","?32","","Seyir","Üçgen"];
const BASLIK_S8 = ["No","At","Tarih","Şehir","Zemin","Pist","?7","Mesafe","Kilo","E.Kilo",
  "K.Fark","Koşu Cinsi","?13","Son 800","Fark","?16","?17","?18","",
  "Genel HP Avantajı","Genel Kilo Avantajı","Orta HP Avantajı","Orta Kilo Avantajı","İnce HP Avantajı","İnce Kilo Avantajı","Seyir","Üçgen","F.Üçgen"];

// SİLİNECEK KOLONLAR (1-tabanlı): at-no tekrarları + H, P, Q + mükerrer koşu cinsi (AX)
const SIL = {"Sayfa1":[7,8,12,16,17,19,24,31,32], "Sayfa2":[7,13,16,17,18]};
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
  if(!rows.length) return '<div class="bos" style="padding:12px">Bu koşunun atlarının geçmiş koşu kaydı yok — ilk kez koşacak taylar (örn. 2 yaşlılar) olabilir.</div>';
  const basliklar = tab==="Sayfa1"? BASLIK_TD : BASLIK_S8;
  const sil=new Set((SIL[tab]||[]).map(x=>x-1));
  const domk=new Set((DOMK[tab]||[]).map(x=>x-1));
  const domList=(DOMK[tab]||[]).map(x=>x-1);
  const domGrp=c=>{const i=domList.indexOf(c);return i<0?"":(i<2?"g50":(i<4?"g66":"g75"));};
  const domBas=c=>{const i=domList.indexOf(c);return (i===0||i===2||i===4)?" gbas":"";};
  const sonhp=new Set((SONHP[tab]||[]).map(x=>x-1));
  const yascins=new Set((YASCINS[tab]||[]).map(x=>x-1));
  const [seyA,seyB]=(SEYIR[tab]||[0,0]).map(x=>x-1);
  const dolu=[];
  for(let c=0;c<42;c++){
    if(sil.has(c)||c===seyB) continue;
    if(rows.some(r=>String(r[c]??"").trim()!=="") || c===seyA) dolu.push(c);
  }
  // Handikap ile İlk 3 HP yer değiştirir (Handikap önce)
  if(tab==="Sayfa1"){
    const i1=dolu.indexOf(21), i2=dolu.indexOf(22);
    if(i1>-1&&i2>-1){ dolu[i1]=22; dolu[i2]=21; }
  }
  // filtreli kolonlar + otomatik sıralama kolonu (Derece / Son 800)
  const FK = tab==="Sayfa1" ? {no:0,sehir:9,msf:13,tarih:8,drc:14,kf:4} : {no:0,sehir:3,msf:7,tarih:2,drc:13,kf:10};
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
    const ad=c===seyA?"Seyir":(basliklar[c]??("K"+(c+1)));
    let f="";
    if(c===FK.no||c===FK.sehir||c===FK.msf) f=kutuPanel(c);
    else if(c===FK.tarih) f=secimPanel(c,"trh",[["30","Son 1 Ay"],["90","Son 3 Ay"],["","Tümü"]]);
    else if(c===seyA) f=secimPanel(c,"sey",[["KACAK","Kaçak"],["","Tümü"]]);
    else if(c===FK.kf) f=secimPanel(c,"kf",[["SIRALA","Küçükten büyüğe"],["","Normal"]],"kFarkSirala");
    const g=domGrp(c);
    const mg=(MGIZ[tab]||[]).includes(c+1)?" m-gizle":"";
    const siralanir=(c===FK.drc||c===FK.kf);   // SADECE Derece ve K.Fark tıklanınca sıralanır
    return `<th data-c="${c}" class="${g}${g?domBas(c):""}${mg}${siralanir?" siralanir":""}"${siralanir?' onclick="tabloSirala(this)" title="küçükten büyüğe sırala"':""}>${esc(ad)}${f}</th>`;
  }).join("");
  const DOM_YOK=new Set(["ŞARTLI 1","SARTLI 1","ŞARTLI 19","SARTLI 19"]);
  const KCIN = tab==="Sayfa1" ? 5 : 11;   // satırın koşu cinsi kolonu (0-tabanlı)
  let _oncekiAt="";
  const trs=rows.map(r=>{
    const kcinsUp=String(r[KCIN]??"").trim().toUpperCase().replace(/\s+/g," ");
    const domGizle=DOM_YOK.has(kcinsUp);
    const atBase=String(r[1]??"").trim().toUpperCase().replace(/\d+$/,"").trim();
    const grupBas=(atBase!==_oncekiAt); _oncekiAt=atBase;
    return `<tr${grupBas?' class="grupbas"':''}>`+dolu.map(c=>{
    let v=String(r[c]??"").trim();
    let cls="";
    if(c===1){ cls="atadi"; v=v.replace(/\d+$/,"").trim(); }   // TESCİL1 -> TESCİL
    else if(c===seyA){ v=String(r[seyB]??"").trim(); cls="ucgen"; }      // YALNIZ ÜÇGEN
    else if(c===FK.tarih){ v=tarihGoster(v); cls="num"; }                // tarih dd.mm.yyyy (yıl dahil)
    else if(c===FK.drc){ cls="num drc"; }   // Derece: kalın siyah (boyasız)
    else if(tab==="Sayfa1"&&c===19){ cls="orta"; }                       // Kaçıncı: ortalı
    else if(sonhp.has(c)){
      const p=v.split("-").filter(x=>x.trim()!=="");
      v=p.length?p[p.length-1]:v; cls="num";
    }
    else if(domk.has(c)){                                                // dominans grupları renkli
      if(domGizle){ v=""; cls="num dom "+domGrp(c); }                     // Şartlı 1 / Şartlı 19 -> dominans YAZILMAZ
      else{
        const f2=parseFloat(v.replace(",","."));
        cls="num dom "+domGrp(c)+domBas(c)+" "+(f2>0?"poz":(f2<0?"neg":""));
        if(!isNaN(f2)) v=(Math.round(f2*100)/100).toString();
      }
    }
    else if(yascins.has(c)) v=yasCinsCevir(v);
    else if(v.includes("▶")||v.includes("▷")) cls="ucgen";
    else if(/^-?\d/.test(v)) cls="num";
    const dv=(c===seyA)?String(r[seyB]??"").trim():String(r[c]??"").trim();
    const mg2=(MGIZ[tab]||[]).includes(c+1)?" m-gizle":"";
    return `<td class="${cls}${mg2}" data-c="${c}" data-v="${esc(dv)}">${esc(fmt(v))}</td>`;
  }).join("")+"</tr>";}).join("");
  return `<button class="mtum${mTum?" acik":""}" onclick="mTum=!mTum;document.body.classList.toggle('mtum-acik',mTum);this.classList.toggle('acik',mTum)">${mTum?"◀ Sade görünüm":"Tüm kolonlar ▶"}</button>
  <div class="detay"><table data-drc="${FK.drc}"><tr>${th}</tr>${trs}</table></div>`;
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
  p.style.position="fixed";
  p.style.zIndex="9999";
  p.style.left=Math.max(8,Math.min(r.left,window.innerWidth-190))+"px";
  // HER ZAMAN YUKARI doğru açılır; üstte yer azsa panel kendi içinde kaydırılır
  p.style.top="auto"; p.style.bottom=(window.innerHeight-r.top+4)+"px";
  p.style.maxHeight=Math.max(140,Math.min(280,r.top-10))+"px";
}
document.addEventListener("click",e=>{
  if(!e.target.closest(".fpanel")&&!e.target.closest(".fok"))
    document.querySelectorAll(".fpanel.acik").forEach(x=>x.classList.remove("acik"));
});
window.addEventListener("scroll",e=>{
  if(e.target instanceof Element && e.target.closest && e.target.closest(".fpanel")) return;
  document.querySelectorAll(".fpanel.acik").forEach(x=>x.classList.remove("acik"));
},true);
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

function gecUyarilar(b, tab){
  const gec=DATA.gec_cikis||{};
  if(!Object.keys(gec).length) return "";
  const atlar=new Set();
  for(const r of (b.detay||[])){
    let at=String(r[1]??"").trim().toUpperCase();
    if(tab==="Sayfa1") at=at.replace(/\d+$/,"").trim();
    if(at) atlar.add(at);
  }
  const uyari=[];
  for(const at of atlar){
    const v=gec[at];
    if(v && v.problem){
      const d=v.kosular.filter(k=>k.boy).map(k=>`${k.tarih} (${k.boy})`).join(", ");
      uyari.push(`<span class="m vurgu"><b>⚠ Geç Çıkış</b>${esc(at)}: ${esc(d)}</span>`);
    }
  }
  return uyari.length?`<div class="meta" style="margin-top:8px">${uyari.join("")}</div>`:"";
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
  // FİLTRE: yalnız Şehir + Mesafe + tarih aralığı (Son 1 Ay / 3 Ay / Tümü)
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
      <option value="30"${derF.ay==="30"?" selected":""}>Son 1 Ay</option>
      <option value="90"${derF.ay==="90"?" selected":""}>Son 3 Ay</option></select>
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
        <tr>${["Tarih","Şehir","Pist","Msf","Derece","Koşu Cinsi","Sıra","HP","Seyir"].map((b,i)=>`<th data-c="${i}" onclick="tabloSirala(this)" title="küçükten büyüğe sırala">${b}</th>`).join("")}</tr>
        ${trs}</table></div>`:'<div class="bos">derece kaydı yok</div>'}
    </div>`;
  }).join("");
  return `<div class="bol-baslik">Dereceler</div>${filtreler}${gruplar}`;
}

function kartHTML(b){
  const h=b.header;
  const meta=["Mesafe","Zemin","Irk"].filter(k=>h[k])
    .map(k=>`<span class="m"><b>${k}</b>${esc(h[k])}</span>`).join("");
  const yorum=h["Yorum"]?`<span class="m vurgu"><b>Yorum</b>${esc(h["Yorum"])}</span>`:"";
  const final=h["Final"]?`<span class="m vurgu"><b>Final</b>${esc(Number(h["Final"])?Number(h["Final"]).toFixed(2):h["Final"])}</span>`:"";
  const gNorm=b.galops.filter(g=>!g.son), gSon=b.galops.filter(g=>g.son);
  const pnl=(g)=>`<div class="panel${g.son?" sonp":""}"><h4>${esc(String(g.name).replace(/\s*\+\s*400/g,""))}</h4>
      ${g.rows.length?g.rows.map(galopSatir).join(""):'<div class="bos">kayıt yok</div>'}</div>`;
  const setA=b.tables.filter(t=>t.col<16);
  const setDede=b.tables.filter(t=>t.col>=16)
                 .map(t=>({...t,name:"Dede "+t.name}));
  return `
  <div class="kart">
    <div class="kosu-baslik">${esc(b.title||(h["Koşu No"]+". Koşu"))}</div>
    <div class="meta">
      <span class="m"><b>İl</b>${esc(h["İl"]||"")}</span>
      <span class="m"><b>Koşu</b>${esc(h["Koşu No"]||"")}</span>
      ${meta}${yorum}${final}
    </div>
    ${gecUyarilar(b, secTab)}

    <div class="bol-baslik">Analiz — Kalite · Mesafe · Sprinter · Kaçak · Dede</div>
    <div class="tablolar hepsi">${setA.concat(setDede).map(tabloHTML).join("")}</div>

    <div class="bol-baslik">Galoplar</div>
    <div class="galoprow">
      <div class="grup30"><div class="gbaslik">🟢 30 GÜNLÜK GALOP</div>
        <div class="paneller tekhiza">${gNorm.map(pnl).join("")||'<div class="bos">galop verisi yok</div>'}</div></div>
      <div class="grupson"><div class="gbaslik son">🟠 SON GALOP</div>
        <div class="paneller tekhiza">${gSon.map(pnl).join("")||'<div class="bos">son galop verisi yok</div>'}</div></div>
    </div>

    <div class="bol-baslik">${secTab==="Sayfa1"?"Toplam Derece — Detay":"Son 800 — Detay"}</div>
    <div class="tablonot ustte">Not: At yarışı genel hükümlerine göre her handikap puanı, sıklette ½ kg’a tekabül eder. Tablodaki kilo değerini 2 ile çarparak yanındaki HP değeriyle kıyaslayabilir; atın derece sıralamasında ne kadar öne çıkabileceğini ya da geride kalabileceğini kolayca hesaplayabilirsiniz. Yeşil değerler avantajı, kırmızılar dezavantajı gösterir.</div>
    ${detayHTML(b, secTab)}
  </div>`;
}

secIl=iller()[0]; secKosu=kosular(secIl)[0];
ciz();
</script>
</body>
</html>
"""

html = HTML.replace("__DATA__", json.dumps(DATA, ensure_ascii=False))
open("tjk_analiz_prototip.html", "w", encoding="utf-8").write(html)
print("yazıldı: tjk_analiz_prototip.html", len(html), "byte")
