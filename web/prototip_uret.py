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
  .chip.fchip{padding:4px 12px;font-size:12px}
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
  .paneller.tekhiza .panel{flex:1 0 200px;min-width:200px}
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
  /* DETAY TABLOSU — TEK SAYFA: yana kaydırma yerine sığdır (kompakt) */
  .detay{overflow-x:auto;border:1px solid var(--line);border-radius:12px;background:var(--card)}
  .detay table{font-size:11px;white-space:normal;width:100%}
  .detay th{position:sticky;top:0;background:var(--bg);z-index:2;padding:5px 5px;
            white-space:normal;line-height:1.15}
  .detay td{padding:4px 5px;border-top:1px solid var(--line);line-height:1.25}
  .detay tr:nth-child(even) td{background:#fafbfc}
  .detay td.atadi{font-weight:700}
  .detay td.ucgen{font-size:13px;letter-spacing:1px;color:var(--acc);white-space:nowrap}
  .detay td.num{text-align:right}
  .detay td.dom.poz{color:#1a7f4b;font-weight:700}
  .detay td.dom.neg{color:#c23a3a;font-weight:700}
  /* EXTREMLER (sağ üst kutu) */
  .satir{display:flex;gap:14px;align-items:flex-start}
  .satir>#icerik{flex:1;min-width:0}
  #yan{width:280px;flex-shrink:0}
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

  <div class="secici"><span class="lbl">İl</span><span id="iller"></span></div>
  <div class="secici"><span class="lbl">Koşu</span><span id="kosular"></span></div>

  <div class="tabs" id="tabs">
    <button class="tab on" data-t="Sayfa1">Toplam Derece</button>
    <button class="tab" data-t="Sayfa2">Son 800 Analizi</button>
    <button class="tab" data-t="AGF">AGF</button>
  </div>

  <div class="satir">
    <div id="icerik"></div>
    <div id="yan"></div>
  </div>
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
    b.onclick=()=>{secIl=il; secKosu=kosular(il)[0]; derF={sehir:"",pist:"",msf:""}; ciz();};
    ilBox.appendChild(b);
  }
  const koBox=document.getElementById("kosular"); koBox.innerHTML="";
  for(const no of kosular(secIl)){
    const b=document.createElement("button");
    b.className="chip num"+(no===secKosu?" on":""); b.textContent=no;
    b.onclick=()=>{secKosu=no; derF={sehir:"",pist:"",msf:""}; ciz();};
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
    yan.innerHTML=`<div class="kart extrem"><h3>⚡ Extremler — ${esc(secKosu)}. Koşu
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

const BASLIK_TD = ["No","At","Kilo","E.Kilo","K.Fark","Koşu Cinsi","?7","?8","Tarih","Şehir",
  "Zemin","?12","Pist","Mesafe","Derece","M.Derece","?17","Yaş/Cins","?19","?20","",
  "Galop","Son HP","?24","D50 HP","D50 Kilo","D66 HP","D66 Kilo","D75 HP","D75 Kilo",
  "Koşu Cinsi","?32","Seyir","Üçgen"];
const BASLIK_S8 = ["No","At","Tarih","Şehir","Zemin","Pist","?7","Mesafe","Kilo","E.Kilo",
  "K.Fark","Koşu Cinsi","?13","Son 800","Fark","?16","?17","?18","",
  "D50 HP","D50 Kilo","D66 HP","D66 Kilo","D75 HP","D75 Kilo","Seyir","Üçgen","F.Üçgen"];

// SİLİNECEK KOLONLAR (1-tabanlı): at-no tekrarları + H, P, Q + mükerrer koşu cinsi (AX)
const SIL = {"Sayfa1":[7,8,12,16,17,19,24,31,32], "Sayfa2":[7,13,18]};
// DOMİNANS kolonları: artı YEŞİL, eksi KIRMIZI
const DOMK = {"Sayfa1":[25,26,27,28,29,30], "Sayfa2":[20,21,22,23,24,25]};
// SON HP: HP listesinin yalnız SON değeri gösterilir (AP sütunu isteği)
const SONHP = {"Sayfa1":[23], "Sayfa2":[]};
// SEYİR: stil sayısı + üçgen TEK hücrede (başlık 'Seyir')
const SEYIR = {"Sayfa1":[33,34], "Sayfa2":[26,27]};
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
  const sonhp=new Set((SONHP[tab]||[]).map(x=>x-1));
  const yascins=new Set((YASCINS[tab]||[]).map(x=>x-1));
  const [seyA,seyB]=(SEYIR[tab]||[0,0]).map(x=>x-1);   // stil no + üçgen -> tek hücre
  // dolu kolonlar (boşlar gizli) - silinenler ve üçgen(SEYİR ikincisi) hariç
  const dolu=[];
  for(let c=0;c<42;c++){
    if(sil.has(c)||c===seyB) continue;
    if(rows.some(r=>String(r[c]??"").trim()!=="") || c===seyA) dolu.push(c);
  }
  const th=dolu.map(c=>`<th>${esc(c===seyA?"Seyir":(basliklar[c]??("K"+(c+1))))}</th>`).join("");
  const trs=rows.map(r=>"<tr>"+dolu.map(c=>{
    let v=String(r[c]??"").trim();
    let cls="";
    if(c===1) cls="atadi";
    else if(c===seyA){                       // Seyir: sayı + üçgen birlikte
      const u=String(r[seyB]??"").trim();
      const s=v.replace(/\.0$/,"");
      v=(s&&u)?`${s} ${u}`:(s||u);
      cls="ucgen";
    }
    else if(sonhp.has(c)){                   // HP listesi -> yalnız SON HP
      const p=v.split("-").filter(x=>x.trim()!=="");
      v=p.length?p[p.length-1]:v;
      cls="num";
    }
    else if(domk.has(c)){                    // dominans: +yeşil -kırmızı
      const f=parseFloat(v.replace(",","."));
      cls="num dom "+(f>0?"poz":(f<0?"neg":""));
      if(!isNaN(f)) v=(Math.round(f*100)/100).toString();
    }
    else if(yascins.has(c)) v=yasCinsCevir(v);   // '4y a a' -> '4y Erkek/Dişi'
    else if(v.includes("▶")||v.includes("▷")) cls="ucgen";
    else if(/^-?\d/.test(v)) cls="num";
    return `<td${cls?` class="${cls}"`:""}>${esc(fmt(v))}</td>`;
  }).join("")+"</tr>").join("");
  return `<div class="detay"><table><tr>${th}</tr>${trs}</table></div>`;
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
let derF={sehir:"",pist:"",msf:""};   // Dereceler filtresi (koşu değişince sıfırlanır)
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
  // FİLTRE seçenekleri: bu koşunun atlarının tüm geçmişinden
  const secim={sehir:new Set(),pist:new Set(),msf:new Set()};
  for(const [,at] of atlar) for(const k of (der[at]||[])){
    if(k[1])secim.sehir.add(k[1]); if(k[2])secim.pist.add(k[2]); if(k[4])secim.msf.add(k[4]);
  }
  const fchip=(alan,v)=>`<button class="chip fchip${derF[alan]===v?" on":""}"
      onclick="derFsec('${alan}','${String(v).replace(/'/g,"")}')">${esc(v)}</button>`;
  const filtreler=`<div class="secici" style="margin:4px 0 10px">
    <span class="lbl">Filtre</span>
    ${[...secim.sehir].sort().map(v=>fchip("sehir",v)).join("")}
    ${[...secim.pist].sort().map(v=>fchip("pist",v)).join("")}
    ${[...secim.msf].sort((a,b)=>parseInt(a)-parseInt(b)).map(v=>fchip("msf",v)).join("")}
    ${(derF.sehir||derF.pist||derF.msf)?`<button class="chip fchip" style="border-color:var(--acc);color:var(--acc)" onclick="derF={sehir:'',pist:'',msf:''};ciz()">✕ temizle</button>`:""}
  </div>`;
  const gruplar=atlar.map(([no,at])=>{
    const kosular=(der[at]||[]).filter(k=>
      (!derF.sehir||k[1]===derF.sehir)&&(!derF.pist||k[2]===derF.pist)&&(!derF.msf||k[4]===derF.msf));
    const trs=kosular.map(k=>{
      const sey=(k[9]&&k[10])?`${String(k[9]).replace(/\.0$/,"")} ${k[10]}`:(k[10]||String(k[9]||"").replace(/\.0$/,""));
      return `<tr><td>${esc(k[0])}</td><td>${esc(k[1])}</td><td>${esc(k[2])}${k[3]?" · "+esc(k[3]):""}</td>
        <td class="num">${esc(k[4])}</td><td class="num" style="font-weight:700">${esc(drcFmt(k[5]))}</td>
        <td>${esc(k[6])}</td><td class="num">${esc(k[7])}</td><td class="num">${esc(k[8])}</td>
        <td class="ucgen">${esc(sey)}</td></tr>`;
    }).join("");
    return `<div style="margin-bottom:14px">
      <div style="display:flex;align-items:center;gap:8px;margin:0 0 6px">
        <span class="atno">${esc(no)}</span><b>${esc(at)}</b>
        <span style="color:var(--mut);font-size:12px">${kosular.length} koşu</span></div>
      ${kosular.length?`<div class="detay"><table>
        <tr><th>Tarih</th><th>Şehir</th><th>Pist</th><th>Msf</th><th>Derece</th><th>Koşu Cinsi</th><th>Sıra</th><th>HP</th><th>Seyir</th></tr>
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

    <div class="bol-baslik">Galop (30 gün)</div>
    <div class="paneller tekhiza">${gNorm.map(pnl).join("")||'<div class="bos">galop verisi yok</div>'}</div>

    <div class="bol-baslik">Son Galop</div>
    <div class="paneller tekhiza">${gSon.map(pnl).join("")||'<div class="bos">son galop verisi yok</div>'}</div>

    <div class="bol-baslik">${secTab==="Sayfa1"?"Toplam Derece — Detay":"Son 800 — Detay"}</div>
    ${detayHTML(b, secTab)}
    ${secTab==="Sayfa1"?derecelerBolum():""}
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
