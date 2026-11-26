const AWAL=window.AWAL||JSON.parse(document.getElementById("data-awal").textContent), K="mutabaah-excel-v2", RT=typeof claude!=="undefined";
let D=AWAL, tab="ringkasan", fk="", q="";
if(!RT){try{const r=localStorage.getItem(K);if(r)D=JSON.parse(r)}catch(e){}}
const $=s=>document.querySelector(s);
function el(t,a={},...c){const e=document.createElement(t);for(const k in a){if(k==="class")e.className=a[k];else if(k.startsWith("on"))e.addEventListener(k.slice(2),a[k]);else e.setAttribute(k,a[k])}c.flat().forEach(x=>{if(x!=null)e.append(x)});return e}
const key=n=>String(n).trim().toLowerCase().replace(/\s+/g," ");
const num=v=>typeof v==="number"&&isFinite(v)?v:null;
const avg=a=>{a=a.filter(x=>x!=null);return a.length?a.reduce((s,x)=>s+x,0)/a.length:null};
const f1=v=>v==null?"–":(Math.round(v*10)/10).toLocaleString("id-ID");

/* ---- Baca Excel: sheet mutabaah, murojaah, klasmen (dicocokkan dari nama sheet dan judul kolom) ---- */
function tabel(wb,nama){
  const sn=wb.SheetNames.find(n=>n.trim().toLowerCase()===nama); if(!sn)return null;
  const a=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null}); if(!a.length)return null;
  const hs=[];a[0].forEach(c=>{const h=c==null?"":String(c).trim().toLowerCase();let k=h,i=2;while(k&&hs.includes(k))k=h+" ("+i++ +")";hs.push(k)});
  const rows=a.slice(1).map(r=>{const d={};hs.forEach((h,i)=>{if(h&&r[i]!=null&&r[i]!=="")d[h]=r[i]});return d}).filter(d=>d["nama anak"]);
  return {hs,rows};
}
function telp(v){if(v==null)return null;const t=String(v).replace(/\D/g,"");return t?(/^(0|62)/.test(t)?t:"0"+t):null}
function bacaExcel(wb){
  const M=tabel(wb,"mutabaah"); if(!M)throw new Error("Sheet “mutabaah” tidak ditemukan.");
  const J=tabel(wb,"murojaah")||{rows:[]}, Kl=tabel(wb,"klasmen")||{rows:[]};
  const i=M.hs.indexOf("no telp"), nh=M.hs.slice(i+1).filter(Boolean);
  const recs={};
  M.rows.filter(d=>d.kelas!=null).forEach(d=>{recs[key(d["nama anak"])]={nama:String(d["nama anak"]).trim(),kelas:d.kelas,telp:telp(d["no telp"]),surah:d.surah??null,baris:d.baris??null,tasmi:d.tasmi??null,sots:d.sots??null,nilai:nh.map(h=>[h,d[h]??null]),juz:null,hal:null,awal:null,akhir:null,barisK:null}});
  const cari=n=>{const k=key(n);if(recs[k])return recs[k];const c=Object.keys(recs).filter(x=>x.startsWith(k)||k.startsWith(x));return c.length===1?recs[c[0]]:null};
  J.rows.forEach(d=>{const r=cari(d["nama anak"]);if(r){r.juz=d.juz??null;r.hal=d.hal??null}});
  Kl.rows.forEach(d=>{const r=cari(d["nama anak"]);if(r){r.awal=d["hafalan awal"]??null;r.akhir=d["hafalan akhir"]??null;r.barisK=d.baris??null;if(!r.telp)r.telp=telp(d["no telp"])}});
  const rows=Object.values(recs); if(!rows.length)throw new Error("Tidak ada baris santri yang terbaca dari sheet “mutabaah”.");
  return rows;
}
$("#btn-impor").addEventListener("click",()=>$("#berkas").click());
function htmlBaru(){
  const c=document.documentElement.cloneNode(true);
  c.querySelector("#view").innerHTML="";c.querySelector("#pesan").textContent="";c.querySelector("#sumber").textContent="";
  c.querySelector("#data-awal").textContent=JSON.stringify(D).replace(/</g,"\\u003c");
  return "<!DOCTYPE html>\n"+c.outerHTML;
}
$("#berkas").addEventListener("change",async e=>{
  const f=e.target.files[0]; if(!f)return; const p=$("#pesan"); p.className="msg";
  try{
    const wb=XLSX.read(await f.arrayBuffer(),{type:"array"});
    D={file:f.name,waktu:new Date().toISOString(),rows:bacaExcel(wb)};
    const a=RT?await claude.use("artifact"):null;
    if(a){p.className="msg ok";p.textContent="Menyimpan agar orang tua ikut melihat data terbaru…";await a.publish(htmlBaru());return}
    try{localStorage.setItem(K,JSON.stringify(D))}catch(x){}
    p.className="msg ok";p.textContent="Berhasil membaca "+D.rows.length+" santri dari "+f.name+".";render();
  }catch(x){D=AWAL;p.className="msg";p.textContent="Gagal: "+(x&&x.message||x)}
  e.target.value="";
});

/* ---- Tampilan ---- */
function ranked(){const r=D.rows.filter(x=>num(x.barisK)!=null).sort((a,b)=>b.barisK-a.barisK);r.forEach((x,i)=>{x._r=i&&r[i-1].barisK===x.barisK?r[i-1]._r:i+1});return r}
const kelasList=()=>[...new Set(D.rows.map(r=>r.kelas))].sort((a,b)=>a-b);
function saring(list){return list.filter(r=>(!fk||String(r.kelas)===fk)&&(!q||key(r.nama).includes(key(q))))}
function filterBar(onchg,withSearch){
  const sel=el("select",{id:"fk","aria-label":"Filter kelas",onchange:e=>{fk=e.target.value;onchg()}},el("option",{value:""},"Semua kelas"),kelasList().map(k=>{const o=el("option",{value:String(k)},"Kelas "+k);if(String(k)===fk)o.selected=true;return o}));
  const items=[el("div",{},el("label",{for:"fk"},"Kelas"),sel)];
  if(withSearch)items.push(el("div",{},el("label",{for:"q"},"Cari nama"),el("input",{id:"q",type:"search",value:q,autocomplete:"off",oninput:e=>{q=e.target.value;onchg()}})));
  return el("div",{class:"row"},items);
}
function barNilai(v){return el("div",{class:"bar"},el("i",{class:v<75?"low":"",style:"width:"+Math.max(0,Math.min(100,v))+"%"}))}

function vRingkasan(v){
  const t=ranked(), nl=D.rows[0]?D.rows[0].nilai.map(x=>x[0]):[];
  v.append(el("section",{class:"box"},el("h2",{},"Ringkasan"),el("div",{class:"stats"},
    el("div",{class:"stat"},el("b",{},String(D.rows.length)),el("small",{},"santri")),
    el("div",{class:"stat"},el("b",{},String(kelasList().length)),el("small",{},"kelas")),
    el("div",{class:"stat"},el("b",{},f1(avg(D.rows.map(r=>num(r.barisK))))),el("small",{},"rata-rata baris")))));
  const kb=el("section",{class:"box"},el("h2",{},"Per kelas"),el("div",{class:"scroll"}));
  const tb=el("table",{},el("tr",{},["Kelas","Santri","Rata-rata baris","Rata-rata nilai"].map(h=>el("th",{},h))));
  kelasList().forEach(k=>{const r=D.rows.filter(x=>x.kelas===k);
    tb.append(el("tr",{},el("td",{},"Kelas "+k),el("td",{class:"n"},String(r.length)),el("td",{class:"n"},f1(avg(r.map(x=>num(x.barisK))))),el("td",{class:"n"},f1(avg(r.flatMap(x=>x.nilai.map(n=>num(n[1]))))))))});
  kb.lastChild.append(tb);v.append(kb);
  const nb=el("section",{class:"box"},el("h2",{},"Rata-rata nilai per penilaian"));
  nl.forEach((h,i)=>{const a=avg(D.rows.map(r=>num(r.nilai[i][1])));
    nb.append(el("div",{class:"nl"},el("span",{},h),a==null?el("span",{}):barNilai(a),el("span",{},f1(a))))});
  v.append(nb);
  const jz={};D.rows.forEach(r=>{if(r.juz)(jz[r.juz]=jz[r.juz]||[]).push(r.nama)});
  const jb=el("section",{class:"box"},el("h2",{},"Murajaah per juz"));
  Object.keys(jz).sort((a,b)=>parseInt(a.replace(/\D/g,""))-parseInt(b.replace(/\D/g,""))).forEach(k=>jb.append(el("p",{style:"margin:.3rem 0"},el("b",{},k+": "),jz[k].length+" santri")));
  if(!Object.keys(jz).length)jb.append(el("p",{class:"kosong"},"Tidak ada data murajaah."));
  v.append(jb);
}
function vKlasemen(v){
  const box=el("section",{class:"box"},el("h2",{},"Klasemen hafalan (berdasarkan baris)")), out=el("div",{class:"scroll"});
  function isi(){out.innerHTML="";const r=saring(ranked());
    if(!r.length){out.append(el("p",{class:"kosong"},"Tidak ada santri yang cocok."));return}
    const tb=el("table",{},el("tr",{},["#","Nama","Kelas","Hafalan awal","Hafalan akhir","Baris"].map(h=>el("th",{},h))));
    r.forEach(x=>tb.append(el("tr",{},el("td",{class:"rk"},String(x._r)),el("td",{},x.nama),el("td",{},String(x.kelas)),el("td",{},x.awal??"–"),el("td",{},x.akhir??"–"),el("td",{class:"n"},String(x.barisK)))));
    out.append(tb)}
  box.append(filterBar(isi,true),out);v.append(box);isi();
}
function vSantri(v){
  const box=el("section",{class:"box"},el("h2",{},"Data santri")), out=el("div");
  const rk=Object.fromEntries(ranked().map(r=>[r.nama,r._r]));
  function isi(){out.innerHTML="";const r=saring([...D.rows].sort((a,b)=>a.kelas-b.kelas||a.nama.localeCompare(b.nama)));
    if(!r.length){out.append(el("p",{class:"kosong"},"Tidak ada santri yang cocok."));return}
    r.forEach(x=>{
      const d=el("dl",{});
      const add=(k,val)=>{if(val!=null&&val!=="")d.append(el("dt",{},k),el("dd",{},String(val)))};
      add("Setoran surah",x.surah);add("Baris",x.baris);add("Hafalan awal",x.awal);add("Hafalan akhir",x.akhir);
      add("Baris klasemen",x.barisK);add("Peringkat",rk[x.nama]);add("Tasmi’",x.tasmi);add("Sots",x.sots);
      add("Murajaah",x.juz?(x.juz+(x.hal?", hal "+x.hal:"")):null);
      if(x.telp){const w="62"+x.telp.replace(/^0|^62/,"");d.append(el("dt",{},"No. telepon"),el("dd",{},el("a",{href:"https://wa.me/"+w,target:"_blank",rel:"noopener"},x.telp)))}
      const nl=el("div",{});x.nilai.forEach(([h,val])=>{if(num(val)!=null)nl.append(el("div",{class:"nl"},el("span",{},h),barNilai(val),el("span",{},String(val))))});
      const rata=avg(x.nilai.map(n=>num(n[1])));
      out.append(el("details",{},el("summary",{},el("span",{},x.nama),el("small",{},"Kelas "+x.kelas+(rata!=null?" · rata-rata "+f1(rata):""))),d,nl))});
  }
  box.append(filterBar(isi,true),out);v.append(box);isi();
}
function render(){
  document.querySelectorAll("nav button").forEach(b=>b.setAttribute("aria-selected",b.dataset.t===tab));
  const s=D.waktu?"Dibaca dari “"+D.file+"” pada "+new Date(D.waktu).toLocaleString("id-ID"):"Data awal dari “"+D.file+"”";
  $("#sumber").textContent=s+" · "+D.rows.length+" santri";
  const v=$("#view");v.innerHTML="";({ringkasan:vRingkasan,klasemen:vKlasemen,santri:vSantri})[tab](v);
}
document.querySelectorAll("nav button").forEach(b=>b.addEventListener("click",()=>{tab=b.dataset.t;render();scrollTo(0,0)}));
render();
$("#kontrol").hidden=RT;
if(RT)(async()=>{try{const u=await claude.use("user");$("#kontrol").hidden=!(u&&await u.canEdit())}catch(e){}})();
