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
  .wrap{max-width:1240px;margin:0 auto;padding:16px}
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
     text-align:left;padding:5px 14px;background:var(--bg)}
  td{padding:6px 14px;border-top:1px solid var(--line)}
  td.no{font-weight:800;width:44px}
  tr.r1 td.no span,tr.r2 td.no span,tr.r3 td.no span{display:inline-flex;width:24px;height:24px;
     border-radius:7px;align-items:center;justify-content:center;color:#fff;font-size:12px}
  tr.r1 td.no span{background:var(--gold)} tr.r2 td.no span{background:#9aa4b5} tr.r3 td.no span{background:#c98d5c}
  td.deg{position:relative;font-weight:700}
  td.deg .bar{position:absolute;left:0;top:15%;bottom:15%;background:var(--pri2);border-radius:0 6px 6px 0;z-index:0}
  td.deg span{position:relative;z-index:1}
  td.say{color:var(--mut);font-size:12px;text-align:right}
  /* DETAY TABLOSU */
  .detay{overflow-x:auto;border:1px solid var(--line);border-radius:12px;background:var(--card)}
  .detay table{font-size:12px;white-space:nowrap}
  .detay th{position:sticky;top:0;background:var(--bg);z-index:2;padding:7px 9px}
  .detay td{padding:5px 9px;border-top:1px solid var(--line)}
  .detay tr:nth-child(even) td{background:#fafbfc}
  .detay td.atadi{font-weight:700}
  .detay td.ucgen{font-size:13px;letter-spacing:1px;color:var(--acc)}
  .detay td.num{text-align:right}
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

  <div class="secici"><span class="lbl">İl</span><span id="iller"></span></div>
  <div class="secici"><span class="lbl">Koşu</span><span id="kosular"></span></div>

  <div class="tabs" id="tabs">
    <button class="tab on" data-t="Sayfa1">Toplam Derece</button>
    <button class="tab" data-t="Sayfa2">Son 800</button>
    <button class="tab" data-t="AGF">AGF</button>
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
  const ilBox=document.getElementById("iller"); ilBox.innerHTML="";
  for(const il of iller()){
    const b=document.createElement("button");
    b.className="chip"+(il===secIl?" on":""); b.textContent=il;
    b.onclick=()=>{secIl=il; secKosu=kosular(il)[0]; ciz();};
    ilBox.appendChild(b);
  }
  const koBox=document.getElementById("kosular"); koBox.innerHTML="";
  for(const no of kosular(secIl)){
    const b=document.createElement("button");
    b.className="chip num"+(no===secKosu?" on":""); b.textContent=no;
    b.onclick=()=>{secKosu=no; ciz();};
    koBox.appendChild(b);
  }
  document.querySelectorAll(".tab").forEach(t=>{
    t.classList.toggle("on",t.dataset.t===secTab);
    t.onclick=()=>{secTab=t.dataset.t; ciz();};
  });
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

const BASLIK_TD = ["Koşu No","At","Kilo","E.Kilo","K.Fark","Koşu Cinsi","?7","?8","Tarih","Şehir",
  "Zemin","?12","Pist","Mesafe","Derece","M.Derece","?17","Yaş/Cins","Sıra","?20","",
  "Galop","HP Listesi","?24","D50 HP","D50 Kilo","D66 HP","D66 Kilo","D75 HP","D75 Kilo",
  "Koşu Cinsi","?32","%45 Stil","Üçgen"];
const BASLIK_S8 = ["Koşu No","At","Tarih","Şehir","Zemin","Pist","?7","Mesafe","Kilo","E.Kilo",
  "K.Fark","Koşu Cinsi","?13","Son 800","Fark","?16","?17","?18","",
  "D50 HP","D50 Kilo","D66 HP","D66 Kilo","D75 HP","D75 Kilo","%45 Stil","Üçgen","F.Üçgen"];

function fmt(v){
  const f=parseFloat(String(v).replace(",","."));
  if(!isNaN(f) && String(v).length>8 && String(v).includes(".")) return f.toFixed(2);
  return v;
}
function detayHTML(b, tab){
  const rows=b.detay||[];
  if(!rows.length) return '<div class="bos" style="padding:12px">detay verisi yok</div>';
  const basliklar = tab==="Sayfa1"? BASLIK_TD : BASLIK_S8;
  // dolu kolonları bul (boş sütunlar otomatik gizlenir)
  const dolu=[];
  for(let c=0;c<42;c++){ if(rows.some(r=>String(r[c]??"").trim()!=="")) dolu.push(c); }
  const th=dolu.map(c=>`<th>${esc(basliklar[c]??("K"+(c+1)))}</th>`).join("");
  const trs=rows.map(r=>"<tr>"+dolu.map(c=>{
    const v=String(r[c]??"");
    const cls = c===1?"atadi" : v.includes("▶")||v.includes("▷")?"ucgen"
              : /^-?\d/.test(v)?"num":"";
    return `<td${cls?` class="${cls}"`:""}>${esc(fmt(v))}</td>`;
  }).join("")+"</tr>").join("");
  return `<div class="detay"><table><tr>${th}</tr>${trs}</table></div>`;
}

function kartHTML(b){
  const h=b.header;
  const meta=["Mesafe","Zemin","Irk"].filter(k=>h[k])
    .map(k=>`<span class="m"><b>${k}</b>${esc(h[k])}</span>`).join("");
  const yorum=h["Yorum"]?`<span class="m vurgu"><b>Yorum</b>${esc(h["Yorum"])}</span>`:"";
  const final=h["Final"]?`<span class="m vurgu"><b>Final</b>${esc(Number(h["Final"])?Number(h["Final"]).toFixed(2):h["Final"])}</span>`:"";
  const gNorm=b.galops.filter(g=>!g.son), gSon=b.galops.filter(g=>g.son);
  const pnl=(g)=>`<div class="panel${g.son?" sonp":""}"><h4>${esc(g.name)}</h4>
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

    <div class="bol-baslik">Analiz — Kalite · Mesafe · Sprinter · Kaçak · Dede</div>
    <div class="tablolar hepsi">${setA.concat(setDede).map(tabloHTML).join("")}</div>

    <div class="bol-baslik">Galop</div>
    <div class="paneller">${gNorm.map(pnl).join("")||'<div class="bos">galop verisi yok</div>'}</div>

    <div class="bol-baslik">Son Galop</div>
    <div class="paneller">${gSon.map(pnl).join("")||'<div class="bos">son galop verisi yok</div>'}</div>

    <div class="bol-baslik">${secTab==="Sayfa1"?"Toplam Derece — Detay":"Son 800 — Detay"}</div>
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
