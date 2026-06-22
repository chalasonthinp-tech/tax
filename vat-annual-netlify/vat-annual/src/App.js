import { useState, useCallback } from "react";

// ─── Tokens ───────────────────────────────────────────────────────────────────
const C = {
  bg:"#F0F4FF", surface:"#FFFFFF", border:"#DDE3F0",
  accent:"#2563EB", accentDark:"#1D4ED8", accentL:"#EFF6FF", accentL2:"#DBEAFE",
  purple:"#7C3AED", purpleL:"#F5F3FF",
  green:"#059669", greenL:"#ECFDF5",
  yellow:"#D97706", yellowL:"#FFFBEB",
  red:"#DC2626", redL:"#FEF2F2",
  text:"#111827", muted:"#6B7280", light:"#9CA3AF",
};

const MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const toBase64 = (file) => new Promise((res,rej) => {
  const r = new FileReader();
  r.onload = () => res(r.result.split(",")[1]);
  r.onerror = rej;
  r.readAsDataURL(file);
});

const fmtNum = (n) => {
  if (n === null || n === undefined || n === "") return "—";
  const v = parseFloat(String(n).replace(/,/g,""));
  if (isNaN(v)) return "—";
  return v.toLocaleString("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2});
};

const num = (n) => { const v=parseFloat(String(n||0).replace(/,/g,"")); return isNaN(v)?0:v; };

// ─── Claude API ───────────────────────────────────────────────────────────────
const SYS_SALES = `คุณเป็นผู้เชี่ยวชาญภาษีมูลค่าเพิ่มไทย ส่งคืน JSON เท่านั้น ห้ามมีข้อความอื่น
{
  "company":"","taxId":"","period":"MM/YYYY","reportNo":"",
  "items":[{"no":1,"date":"","invoiceNo":"","customer":"","customerTaxId":"","branch":"","amount":0,"vat":0}],
  "totalAmount":0,"totalVat":0,"itemCount":0
}
ใบลดหนี้ให้ใส่ตัวเลขติดลบ`;

const SYS_PP30 = `คุณเป็นผู้เชี่ยวชาญภาษีมูลค่าเพิ่มไทย ส่งคืน JSON เท่านั้น ห้ามมีข้อความอื่น
{
  "company":"","taxId":"","period":"MM/YYYY","refNo":"","filedDate":"",
  "row1_salesTotal":0,"row2_zeroRate":0,"row3_exempt":0,"row4_taxableSales":0,
  "row5_outputTax":0,"row6_purchaseTotal":0,"row7_inputTax":0,
  "row8_taxPayable":0,"row9_taxCredit":0,"row10_bfCredit":0,
  "row11_netPayable":0,"row12_netCredit":0
}`;

async function callClaude(file, systemPrompt, userText) {
  const b64  = await toBase64(file);
  const mime = file.type || "application/pdf";
  const block = mime === "application/pdf"
    ? { type:"document", source:{type:"base64",media_type:mime,data:b64} }
    : { type:"image",    source:{type:"base64",media_type:mime,data:b64} };

  const res = await fetch("/.netlify/functions/anthropic-proxy", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({
      model:"claude-sonnet-4-6", max_tokens:4000,
      system: systemPrompt,
      messages:[{role:"user", content:[block,{type:"text",text:userText}]}],
    }),
  });
  if (!res.ok) { const e=await res.json().catch(()=>{}); throw new Error(e?.error?.message||`API ${res.status}`); }
  const data = await res.json();
  const raw  = data.content.find(b=>b.type==="text")?.text||"";
  return JSON.parse(raw.replace(/```json|```/g,"").trim());
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
function exportExcel(months) {
  const XLSX = window.XLSX;
  const wb   = XLSX.utils.book_new();

  // ── Sheet 1: ภาพรวมทั้งปี ──
  const annualRows = [
    ["สรุปภาพรวมทั้งปี — เปรียบเทียบรายงานภาษีขาย vs ภ.พ.30"],
    [],
    ["เดือน","มูลค่าขาย (รายงาน)","มูลค่าขาย (ภพ30)","ผลต่างมูลค่า","ภาษีขาย (รายงาน)","ภาษีขาย (ภพ30)","ผลต่างภาษีขาย","ภาษีซื้อ","ภาษีต้องชำระ","สถานะ"],
  ];
  let totSalesAmt=0,totSalesVat=0,totPP30Amt=0,totPP30Vat=0,totInputTax=0,totPayable=0;
  months.forEach(m => {
    if (!m.salesData && !m.pp30Data) return;
    const sAmt  = num(m.salesData?.totalAmount);
    const sVat  = num(m.salesData?.totalVat);
    const pAmt  = num(m.pp30Data?.row1_salesTotal);
    const pVat  = num(m.pp30Data?.row5_outputTax);
    const iTax  = num(m.pp30Data?.row7_inputTax);
    const pay   = num(m.pp30Data?.row11_netPayable);
    const dAmt  = m.salesData && m.pp30Data ? sAmt-pAmt : null;
    const dVat  = m.salesData && m.pp30Data ? sVat-pVat : null;
    const ok    = dVat===null?"—":Math.abs(dVat)<0.02?"✅ ตรงกัน":Math.abs(dVat)<=1000?"⚠️ ต่างเล็กน้อย":"❌ ต่างมาก";
    annualRows.push([m.label, m.salesData?sAmt:"—", m.pp30Data?pAmt:"—", dAmt, m.salesData?sVat:"—", m.pp30Data?pVat:"—", dVat, m.pp30Data?iTax:"—", m.pp30Data?pay:"—", ok]);
    totSalesAmt+=sAmt; totSalesVat+=sVat; totPP30Amt+=pAmt; totPP30Vat+=pVat; totInputTax+=iTax; totPayable+=pay;
  });
  annualRows.push([]);
  annualRows.push(["รวมทั้งปี", totSalesAmt, totPP30Amt, totSalesAmt-totPP30Amt, totSalesVat, totPP30Vat, totSalesVat-totPP30Vat, totInputTax, totPayable, ""]);
  const ws0 = XLSX.utils.aoa_to_sheet(annualRows);
  ws0["!cols"] = [{wch:14},{wch:20},{wch:20},{wch:16},{wch:20},{wch:20},{wch:16},{wch:16},{wch:16},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws0, "ภาพรวมทั้งปี");

  // ── Sheet per month ──
  months.forEach(m => {
    if (!m.salesData && !m.pp30Data) return;

    // Comparison sheet
    const compRows = [
      [`เปรียบเทียบ ${m.label}`],[],
      ["รายการ","รายงานภาษีขาย","ภ.พ.30","ผลต่าง","สถานะ"],
      ["1. ยอดขายรวม", m.salesData?.totalAmount, m.pp30Data?.row1_salesTotal, m.salesData&&m.pp30Data?num(m.salesData.totalAmount)-num(m.pp30Data.row1_salesTotal):null, ""],
      ["5. ภาษีขาย", m.salesData?.totalVat, m.pp30Data?.row5_outputTax, m.salesData&&m.pp30Data?num(m.salesData.totalVat)-num(m.pp30Data.row5_outputTax):null, ""],
      ["6. ยอดซื้อ","—",m.pp30Data?.row6_purchaseTotal,null,""],
      ["7. ภาษีซื้อ","—",m.pp30Data?.row7_inputTax,null,""],
      ["8. ภาษีต้องชำระ","—",m.pp30Data?.row8_taxPayable,null,""],
      ["10. ภาษียกมา","—",m.pp30Data?.row10_bfCredit,null,""],
      ["11. สุทธิต้องชำระ","—",m.pp30Data?.row11_netPayable,null,""],
    ];
    // fill status
    compRows.slice(3).forEach(r => {
      if (r[3]!==null && r[3]!==undefined && r[3]!=="") {
        r[4] = Math.abs(r[3])<0.02?"✅ ตรงกัน":Math.abs(r[3])<=1000?"⚠️ ต่างเล็กน้อย":"❌ ต่างมาก";
      }
    });

    const wsC = XLSX.utils.aoa_to_sheet(compRows);
    wsC["!cols"] = [{wch:28},{wch:20},{wch:20},{wch:16},{wch:14}];
    XLSX.utils.book_append_sheet(wb, wsC, `${m.short} เปรียบ`.slice(0,31));

    // Sales detail sheet
    if (m.salesData?.items?.length) {
      const salesRows = [
        [`รายงานภาษีขาย ${m.label}`],
        [`บริษัท: ${m.salesData.company}`, `เลขที่แบบ: ${m.salesData.reportNo}`, `เลขผู้เสียภาษี: ${m.salesData.taxId}`],[],
        ["#","วันที่","เลขที่ใบกำกับ","ผู้ซื้อ","เลขผู้เสียภาษี","สาขา","มูลค่า","VAT"],
        ...m.salesData.items.map(r=>[r.no,r.date,r.invoiceNo,r.customer,r.customerTaxId,r.branch,r.amount,r.vat]),
        [],["","","","","","รวม",m.salesData.totalAmount,m.salesData.totalVat],
      ];
      const wsS = XLSX.utils.aoa_to_sheet(salesRows);
      wsS["!cols"] = [{wch:4},{wch:13},{wch:18},{wch:32},{wch:16},{wch:7},{wch:18},{wch:14}];
      XLSX.utils.book_append_sheet(wb, wsS, `${m.short} ภาษีขาย`.slice(0,31));
    }

    // PP30 sheet
    if (m.pp30Data) {
      const pd = m.pp30Data;
      const pp30Rows = [
        [`ภ.พ.30 ${m.label}`],
        [`บริษัท: ${pd.company}`, `ยื่นวันที่: ${pd.filedDate}`, `อ้างอิง: ${pd.refNo}`],[],
        ["ข้อ","รายการ","จำนวนเงิน (บาท)"],
        ["1","ยอดขายในเดือนนี้",pd.row1_salesTotal],
        ["2","ลบ ยอดขายอัตรา 0%",pd.row2_zeroRate],
        ["3","ลบ ยอดขายที่ยกเว้น",pd.row3_exempt],
        ["4","ยอดขายที่ต้องเสียภาษี",pd.row4_taxableSales],
        ["5","ภาษีขายเดือนนี้",pd.row5_outputTax],
        ["6","ยอดซื้อที่มีสิทธิหักภาษีซื้อ",pd.row6_purchaseTotal],
        ["7","ภาษีซื้อเดือนนี้",pd.row7_inputTax],
        ["8","ภาษีที่ต้องชำระ",pd.row8_taxPayable],
        ["9","ภาษีชำระเกิน",pd.row9_taxCredit],
        ["10","ภาษียกมา",pd.row10_bfCredit],
        ["11","สุทธิต้องชำระ",pd.row11_netPayable],
        ["12","สุทธิชำระเกิน",pd.row12_netCredit],
      ];
      const wsP = XLSX.utils.aoa_to_sheet(pp30Rows);
      wsP["!cols"] = [{wch:5},{wch:34},{wch:20}];
      XLSX.utils.book_append_sheet(wb, wsP, `${m.short} ภพ30`.slice(0,31));
    }
  });

  const buf  = XLSX.write(wb,{bookType:"xlsx",type:"array"});
  const blob = new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=`VAT_Annual_${Date.now()}.xlsx`; a.click();
  URL.revokeObjectURL(url);
}

// ─── Sub-components ───────────────────────────────────────────────────────────
function Spinner({ color }) {
  return <span style={{ display:"inline-block",width:14,height:14,
    border:`2px solid ${color||C.accentL2}55`,borderTopColor:color||C.accent,
    borderRadius:"50%",animation:"spin .7s linear infinite",verticalAlign:"middle" }} />;
}

function FileBtn({ label, color, file, onFile, loading, done, error }) {
  return (
    <label style={{ display:"flex",flexDirection:"column",gap:4,cursor:"pointer" }}>
      <input type="file" accept="image/*,.pdf" style={{display:"none"}}
        onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{
        padding:"8px 10px", border:`1.5px dashed ${file?color:C.border}`,
        borderRadius:8, background:file?`${color}11`:C.bg,
        fontSize:11, textAlign:"center", transition:"all .15s",
        minHeight:52, display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:3,
      }}>
        <div style={{fontSize:18}}>{loading?<Spinner color={color}/>:done?"✅":error?"❌":file?"📄":"⬆️"}</div>
        <div style={{fontWeight:600,color:file?color:C.muted,lineHeight:1.2,wordBreak:"break-all"}}>
          {loading?"กำลังอ่าน...":error?"อ่านไม่สำเร็จ":done?"อ่านแล้ว":file?file.name.slice(0,18)+(file.name.length>18?"…":""):label}
        </div>
      </div>
    </label>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
const initMonths = () => MONTHS.map((name,i) => ({
  idx:    i,
  label:  name,
  short:  `${String(i+1).padStart(2,"0")}`,
  salesFile:  null,
  pp30File:   null,
  salesData:  null,
  pp30Data:   null,
  salesLoading: false,
  pp30Loading:  false,
  salesError:   null,
  pp30Error:    null,
}));

export default function App() {
  const [year,    setYear]    = useState(new Date().getFullYear()+543);
  const [months,  setMonths]  = useState(initMonths);
  const [running, setRunning] = useState(false);
  const [tab,     setTab]     = useState("upload"); // upload | result

  const updMonth = useCallback((idx, patch) =>
    setMonths(prev => prev.map((m,i) => i===idx ? {...m,...patch} : m)), []);

  // ── Scan one file ──────────────────────────────────────────────────────────
  const scanFile = useCallback(async (mIdx, type) => {
    const m = months[mIdx];
    const file = type==="sales" ? m.salesFile : m.pp30File;
    if (!file) return;

    updMonth(mIdx, type==="sales" ? {salesLoading:true,salesError:null} : {pp30Loading:true,pp30Error:null});
    try {
      if (type==="sales") {
        const d = await callClaude(file, SYS_SALES, "นี่คือรายงานภาษีขาย ดึงข้อมูลทุกรายการและยอดรวม");
        updMonth(mIdx, {salesData:d, salesLoading:false});
      } else {
        const d = await callClaude(file, SYS_PP30, "นี่คือ ภ.พ.30 ดึงตัวเลขทุกช่อง ข้อ 1-12");
        updMonth(mIdx, {pp30Data:d, pp30Loading:false});
      }
    } catch(e) {
      updMonth(mIdx, type==="sales"
        ? {salesLoading:false, salesError:e.message}
        : {pp30Loading:false,  pp30Error:e.message});
    }
  }, [months, updMonth]);

  // ── Scan all uploaded files ────────────────────────────────────────────────
  const scanAll = useCallback(async () => {
    setRunning(true);
    for (let i=0; i<12; i++) {
      const m = months[i];
      if (m.salesFile && !m.salesData) await scanFile(i,"sales");
      if (m.pp30File  && !m.pp30Data)  await scanFile(i,"pp30");
    }
    setRunning(false);
    setTab("result");
  }, [months, scanFile]);

  // ── Stats ──────────────────────────────────────────────────────────────────
  const uploadedCount = months.filter(m=>m.salesFile||m.pp30File).length;
  const doneCount     = months.filter(m=>m.salesData||m.pp30Data).length;
  const hasResults    = doneCount > 0;

  // ── Annual totals ──────────────────────────────────────────────────────────
  const annual = months.reduce((acc,m) => {
    acc.salesAmt  += num(m.salesData?.totalAmount);
    acc.salesVat  += num(m.salesData?.totalVat);
    acc.pp30Amt   += num(m.pp30Data?.row1_salesTotal);
    acc.pp30Vat   += num(m.pp30Data?.row5_outputTax);
    acc.inputTax  += num(m.pp30Data?.row7_inputTax);
    acc.netPayable+= num(m.pp30Data?.row11_netPayable);
    return acc;
  }, {salesAmt:0,salesVat:0,pp30Amt:0,pp30Vat:0,inputTax:0,netPayable:0});

  // ─── Render ──────────────────────────────────────────────────────────────
  return (
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Sarabun','Noto Sans Thai',sans-serif"}}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} *{box-sizing:border-box}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-thumb{background:#CBD5E1;border-radius:3px}`}
      </style>

      {/* ── Header ── */}
      <div style={{background:`linear-gradient(135deg,${C.accentDark},${C.purple})`,color:"#fff",padding:"0 24px"}}>
        <div style={{maxWidth:1200,margin:"0 auto",padding:"18px 0",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div style={{fontSize:32}}>🧾</div>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:20}}>VAT Annual Compare</div>
            <div style={{fontSize:12,opacity:.85}}>เปรียบเทียบภาษีขาย vs ภ.พ.30 ครบ 12 เดือน + สรุปภาพรวมทั้งปี</div>
          </div>
          {/* Year selector */}
          <div style={{display:"flex",alignItems:"center",gap:8,background:"rgba(255,255,255,.15)",
            borderRadius:10,padding:"6px 14px"}}>
            <span style={{fontSize:13,fontWeight:600}}>ปี พ.ศ.</span>
            <input type="number" value={year} onChange={e=>setYear(e.target.value)}
              style={{width:72,background:"none",border:"none",color:"#fff",fontSize:15,
                fontWeight:800,outline:"none",fontFamily:"inherit"}} />
          </div>
          {/* Tabs */}
          {["upload","result"].map(t => (
            <button key={t} onClick={()=>setTab(t)} style={{
              background:tab===t?"rgba(255,255,255,.25)":"none",
              border:"1px solid rgba(255,255,255,.3)", borderRadius:8,
              color:"#fff",padding:"7px 16px",cursor:"pointer",fontSize:13,fontWeight:700,
            }}>{t==="upload"?"📂 อัปโหลด":"📊 ผลเปรียบเทียบ"}</button>
          ))}
        </div>
      </div>

      <div style={{maxWidth:1200,margin:"0 auto",padding:"22px 16px"}}>

        {/* ══ TAB: UPLOAD ══════════════════════════════════════════════════════ */}
        {tab==="upload" && (
          <>
            {/* Progress bar */}
            <div style={{background:C.surface,borderRadius:12,padding:"16px 20px",
              marginBottom:16,border:`1px solid ${C.border}`,display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:6}}>
                  อัปโหลดแล้ว {uploadedCount} เดือน · อ่านแล้ว {doneCount} เดือน
                </div>
                <div style={{height:8,background:C.border,borderRadius:4,overflow:"hidden"}}>
                  <div style={{height:"100%",borderRadius:4,transition:"width .4s",
                    background:`linear-gradient(90deg,${C.accent},${C.purple})`,
                    width:`${Math.round(uploadedCount/12*100)}%`}} />
                </div>
              </div>
              <button onClick={scanAll}
                disabled={running||uploadedCount===0}
                style={{
                  background:running||uploadedCount===0?C.border:`linear-gradient(135deg,${C.accent},${C.purple})`,
                  color:running||uploadedCount===0?C.muted:"#fff",
                  border:"none",borderRadius:10,padding:"11px 24px",
                  fontWeight:800,fontSize:14,cursor:running||uploadedCount===0?"not-allowed":"pointer",
                  display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap",
                  boxShadow:running||uploadedCount===0?"none":"0 4px 14px rgba(37,99,235,.3)",
                }}>
                {running?<><Spinner/>กำลังวิเคราะห์...</>:"🤖 ให้ AI อ่านทั้งหมด"}
              </button>
              {hasResults && (
                <button onClick={()=>setTab("result")} style={{
                  background:C.greenL,color:C.green,border:`1px solid ${C.green}`,
                  borderRadius:10,padding:"11px 20px",fontWeight:700,fontSize:14,cursor:"pointer",whiteSpace:"nowrap",
                }}>📊 ดูผล</button>
              )}
            </div>

            {/* Month grid */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:12}}>
              {months.map((m,i) => (
                <div key={i} style={{background:C.surface,borderRadius:12,padding:14,
                  border:`1px solid ${(m.salesData||m.pp30Data)?C.accentL2:C.border}`,transition:"border .2s"}}>

                  {/* Month label */}
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                    <div style={{width:30,height:30,borderRadius:8,
                      background:m.salesData&&m.pp30Data?`linear-gradient(135deg,${C.accent},${C.purple})`:C.accentL,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      fontWeight:800,fontSize:13,color:m.salesData&&m.pp30Data?"#fff":C.accent}}>
                      {String(i+1).padStart(2,"0")}
                    </div>
                    <div style={{fontWeight:700,fontSize:13,color:C.text}}>{m.label}</div>
                    {m.salesData&&m.pp30Data && <span style={{fontSize:14,marginLeft:"auto"}}>✅</span>}
                  </div>

                  {/* File buttons */}
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                    <FileBtn label="ภาษีขาย" color={C.accent}
                      file={m.salesFile} loading={m.salesLoading}
                      done={!!m.salesData} error={!!m.salesError}
                      onFile={f => { updMonth(i,{salesFile:f,salesData:null,salesError:null}); }} />
                    <FileBtn label="ภ.พ.30" color={C.purple}
                      file={m.pp30File} loading={m.pp30Loading}
                      done={!!m.pp30Data} error={!!m.pp30Error}
                      onFile={f => { updMonth(i,{pp30File:f,pp30Data:null,pp30Error:null}); }} />
                  </div>

                  {/* Quick stats */}
                  {m.salesData && (
                    <div style={{marginTop:8,fontSize:10,color:C.accent,background:C.accentL,
                      borderRadius:6,padding:"4px 8px"}}>
                      📊 {m.salesData.itemCount} รายการ · {fmtNum(m.salesData.totalVat)} VAT
                    </div>
                  )}
                  {m.pp30Error && (
                    <div style={{marginTop:6,fontSize:10,color:C.red,background:C.redL,borderRadius:6,padding:"4px 8px"}}>
                      ❌ {m.pp30Error.slice(0,40)}
                    </div>
                  )}
                  {m.salesError && (
                    <div style={{marginTop:6,fontSize:10,color:C.red,background:C.redL,borderRadius:6,padding:"4px 8px"}}>
                      ❌ {m.salesError.slice(0,40)}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </>
        )}

        {/* ══ TAB: RESULTS ═════════════════════════════════════════════════════ */}
        {tab==="result" && (
          <>
            {!hasResults ? (
              <div style={{textAlign:"center",padding:"60px 0",color:C.muted}}>
                <div style={{fontSize:48,marginBottom:12}}>📭</div>
                <div style={{fontSize:15}}>ยังไม่มีข้อมูล — กลับไปอัปโหลดและให้ AI อ่านก่อนครับ</div>
                <button onClick={()=>setTab("upload")} style={{marginTop:16,background:C.accent,color:"#fff",
                  border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontWeight:700}}>
                  📂 ไปหน้าอัปโหลด
                </button>
              </div>
            ) : (
              <>
                {/* ── Annual KPI cards ── */}
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:12,marginBottom:20}}>
                  {[
                    {label:"มูลค่าขายรวม (รายงาน)",val:annual.salesAmt,color:C.accent,icon:"💰"},
                    {label:"ภาษีขายรวม (รายงาน)",  val:annual.salesVat, color:C.purple,icon:"📤"},
                    {label:"มูลค่าขาย (ภพ30)",      val:annual.pp30Amt,  color:C.accent,icon:"📋"},
                    {label:"ภาษีขาย (ภพ30)",        val:annual.pp30Vat,  color:C.purple,icon:"📋"},
                    {label:"ภาษีซื้อรวม",           val:annual.inputTax, color:C.yellow,icon:"📥"},
                    {label:"VAT สุทธิที่ชำระ",       val:annual.netPayable,color:C.green,icon:"✅"},
                  ].map((k,i) => (
                    <div key={i} style={{background:C.surface,borderRadius:12,padding:"14px 16px",
                      border:`1px solid ${C.border}`}}>
                      <div style={{fontSize:20,marginBottom:6}}>{k.icon}</div>
                      <div style={{fontSize:11,color:C.muted,marginBottom:4}}>{k.label}</div>
                      <div style={{fontWeight:800,fontSize:16,color:k.color,fontVariantNumeric:"tabular-nums"}}>
                        {fmtNum(k.val)}
                      </div>
                      <div style={{fontSize:10,color:C.light}}>บาท</div>
                    </div>
                  ))}
                </div>

                {/* ── Diff summary banner ── */}
                {(() => {
                  const dAmt = annual.salesAmt - annual.pp30Amt;
                  const dVat = annual.salesVat - annual.pp30Vat;
                  const ok   = Math.abs(dVat)<0.02;
                  return (
                    <div style={{background:ok?C.greenL:C.yellowL,border:`1px solid ${ok?"#6EE7B7":"#FCD34D"}`,
                      borderRadius:12,padding:"14px 20px",marginBottom:20,
                      display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
                      <div style={{fontSize:28}}>{ok?"✅":"⚠️"}</div>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:800,fontSize:14,color:ok?C.green:C.yellow}}>
                          {ok?"ข้อมูลทั้งปีตรงกัน":"พบความต่างในข้อมูลทั้งปี"}
                        </div>
                        <div style={{fontSize:12,color:C.muted,marginTop:2}}>
                          ผลต่างมูลค่าขาย {fmtNum(dAmt)} บาท · ผลต่างภาษีขาย {fmtNum(dVat)} บาท
                        </div>
                      </div>
                      <button onClick={()=>exportExcel(months)} style={{
                        background:`linear-gradient(135deg,${C.green},#0D9488)`,color:"#fff",
                        border:"none",borderRadius:10,padding:"11px 24px",
                        fontWeight:800,fontSize:14,cursor:"pointer",whiteSpace:"nowrap",
                        boxShadow:"0 4px 14px rgba(5,150,105,.3)",
                      }}>⬇️ Export Excel ทั้งปี</button>
                    </div>
                  );
                })()}

                {/* ── Month-by-month table ── */}
                <div style={{background:C.surface,borderRadius:14,border:`1px solid ${C.border}`,overflow:"hidden",marginBottom:20}}>
                  <div style={{padding:"14px 20px",borderBottom:`1px solid ${C.border}`,
                    fontWeight:800,fontSize:14,color:C.text}}>
                    📅 สรุปรายเดือน — ปี พ.ศ. {year}
                  </div>
                  <div style={{overflowX:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead>
                        <tr style={{background:"#F8FAFF"}}>
                          {["เดือน","มูลค่าขาย\n(รายงาน)","มูลค่าขาย\n(ภพ30)","ผลต่าง\nมูลค่า","ภาษีขาย\n(รายงาน)","ภาษีขาย\n(ภพ30)","ผลต่าง\nภาษี","ภาษีซื้อ","ต้องชำระ","สถานะ"]
                            .map((h,i)=>(
                              <th key={i} style={{padding:"10px 12px",textAlign:i===0?"left":"right",
                                fontWeight:700,color:C.muted,borderBottom:`1px solid ${C.border}`,
                                whiteSpace:"pre-line",lineHeight:1.3,fontSize:11}}>{h}</th>
                            ))}
                        </tr>
                      </thead>
                      <tbody>
                        {months.map((m,i) => {
                          const sAmt = num(m.salesData?.totalAmount);
                          const sVat = num(m.salesData?.totalVat);
                          const pAmt = num(m.pp30Data?.row1_salesTotal);
                          const pVat = num(m.pp30Data?.row5_outputTax);
                          const dAmt = m.salesData&&m.pp30Data ? sAmt-pAmt : null;
                          const dVat = m.salesData&&m.pp30Data ? sVat-pVat : null;
                          const hasData = m.salesData||m.pp30Data;
                          if (!hasData) return null;
                          const statusColor = dVat===null?C.muted:Math.abs(dVat)<0.02?C.green:Math.abs(dVat)<=1000?C.yellow:C.red;
                          const statusBg    = dVat===null?"#F3F4F6":Math.abs(dVat)<0.02?C.greenL:Math.abs(dVat)<=1000?C.yellowL:C.redL;
                          const statusText  = dVat===null?"—":Math.abs(dVat)<0.02?"✅ ตรง":Math.abs(dVat)<=1000?"⚠️ เล็กน้อย":"❌ ต่างมาก";
                          return (
                            <tr key={i} style={{background:i%2===0?C.surface:"#FAFBFF",
                              borderBottom:`1px solid ${C.border}`}}>
                              <td style={{padding:"9px 12px",fontWeight:700,color:C.text}}>
                                {String(i+1).padStart(2,"0")} {m.label}
                              </td>
                              {[
                                m.salesData?sAmt:null,
                                m.pp30Data?pAmt:null,
                                dAmt,
                                m.salesData?sVat:null,
                                m.pp30Data?pVat:null,
                                dVat,
                                m.pp30Data?num(m.pp30Data.row7_inputTax):null,
                                m.pp30Data?num(m.pp30Data.row11_netPayable):null,
                              ].map((v,ci) => (
                                <td key={ci} style={{padding:"9px 12px",textAlign:"right",
                                  fontFamily:"monospace",
                                  color: (ci===2||ci===5)&&v!==null ? (Math.abs(v)<0.02?C.green:Math.abs(v)<=1000?C.yellow:C.red) : C.muted,
                                  fontWeight: (ci===2||ci===5)&&v!==null&&Math.abs(v)>0.02?700:400,
                                }}>
                                  {v===null?"—":(ci===2||ci===5)&&v!==null?(v>=0?"+":"")+fmtNum(v):fmtNum(v)}
                                </td>
                              ))}
                              <td style={{padding:"9px 12px",textAlign:"right"}}>
                                <span style={{background:statusBg,color:statusColor,
                                  borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                                  {statusText}
                                </span>
                              </td>
                            </tr>
                          );
                        })}
                        {/* Total row */}
                        <tr style={{background:C.accentL,borderTop:`2px solid ${C.accentL2}`}}>
                          <td style={{padding:"10px 12px",fontWeight:800,color:C.accent}}>รวมทั้งปี</td>
                          {[annual.salesAmt,annual.pp30Amt,annual.salesAmt-annual.pp30Amt,
                            annual.salesVat,annual.pp30Vat,annual.salesVat-annual.pp30Vat,
                            annual.inputTax,annual.netPayable].map((v,ci)=>(
                            <td key={ci} style={{padding:"10px 12px",textAlign:"right",
                              fontFamily:"monospace",fontWeight:800,
                              color:(ci===2||ci===5)?Math.abs(v)<0.02?C.green:C.red:C.accent}}>
                              {(ci===2||ci===5)?(v>=0?"+":"")+fmtNum(v):fmtNum(v)}
                            </td>
                          ))}
                          <td/>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* ── Per-month detail cards ── */}
                {months.filter(m=>m.salesData||m.pp30Data).map((m,i) => {
                  const dVat = m.salesData&&m.pp30Data
                    ? num(m.salesData.totalVat)-num(m.pp30Data.row5_outputTax) : null;
                  const hasIssue = dVat!==null&&Math.abs(dVat)>0.02;
                  return (
                    <details key={i} style={{background:C.surface,borderRadius:12,
                      border:`1px solid ${hasIssue?"#FCA5A5":C.border}`,marginBottom:10,overflow:"hidden"}}>
                      <summary style={{padding:"13px 20px",cursor:"pointer",
                        background:hasIssue?C.redL:C.accentL,
                        display:"flex",alignItems:"center",gap:10,listStyle:"none",userSelect:"none"}}>
                        <span style={{fontSize:18}}>{hasIssue?"⚠️":"✅"}</span>
                        <span style={{fontWeight:800,fontSize:14,flex:1}}>
                          {String(m.idx+1).padStart(2,"0")} {m.label}
                          {m.salesData&&` · ${m.salesData.itemCount} รายการ`}
                          {m.pp30Data&&` · ต้องชำระ ${fmtNum(m.pp30Data.row11_netPayable)} บาท`}
                        </span>
                        <span style={{fontSize:12,color:C.muted}}>คลิกเพื่อขยาย ▼</span>
                      </summary>

                      <div style={{padding:"16px 20px"}}>
                        {/* Comparison mini-table */}
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginBottom:12}}>
                          <thead>
                            <tr style={{background:"#F8FAFF"}}>
                              {["รายการ","รายงานภาษีขาย","ภ.พ.30","ผลต่าง","สถานะ"].map((h,j)=>(
                                <th key={j} style={{padding:"8px 12px",textAlign:j===0?"left":"right",
                                  fontWeight:700,color:C.muted,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {[
                              ["1. ยอดขายรวม",      m.salesData?.totalAmount, m.pp30Data?.row1_salesTotal],
                              ["5. ภาษีขาย",        m.salesData?.totalVat,    m.pp30Data?.row5_outputTax],
                              ["6. ยอดซื้อ",        null,                     m.pp30Data?.row6_purchaseTotal],
                              ["7. ภาษีซื้อ",       null,                     m.pp30Data?.row7_inputTax],
                              ["11. สุทธิต้องชำระ", null,                     m.pp30Data?.row11_netPayable],
                            ].map(([label,s,p],ri)=>{
                              const diff = s!==null&&p!==null&&s!==undefined&&p!==undefined ? num(s)-num(p) : null;
                              return (
                                <tr key={ri} style={{background:ri%2===0?C.surface:"#FAFBFF",borderBottom:`1px solid ${C.border}`}}>
                                  <td style={{padding:"8px 12px",fontWeight:600}}>{label}</td>
                                  <td style={{padding:"8px 12px",textAlign:"right",fontFamily:"monospace",color:C.muted}}>{s!=null?fmtNum(s):"—"}</td>
                                  <td style={{padding:"8px 12px",textAlign:"right",fontFamily:"monospace",color:C.muted}}>{p!=null?fmtNum(p):"—"}</td>
                                  <td style={{padding:"8px 12px",textAlign:"right",fontFamily:"monospace",fontWeight:diff!==null?700:400,
                                    color:diff===null?C.light:Math.abs(diff)<0.02?C.green:Math.abs(diff)<=1000?C.yellow:C.red}}>
                                    {diff===null?"—":(diff>=0?"+":"")+fmtNum(diff)}
                                  </td>
                                  <td style={{padding:"8px 12px",textAlign:"right"}}>
                                    {diff!==null&&<span style={{
                                      background:Math.abs(diff)<0.02?C.greenL:Math.abs(diff)<=1000?C.yellowL:C.redL,
                                      color:Math.abs(diff)<0.02?C.green:Math.abs(diff)<=1000?C.yellow:C.red,
                                      borderRadius:5,padding:"2px 7px",fontSize:11,fontWeight:700}}>
                                      {Math.abs(diff)<0.02?"✅":Math.abs(diff)<=1000?"⚠️":"❌"}
                                    </span>}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>

                        {/* Invoice list */}
                        {m.salesData?.items?.length>0 && (
                          <details style={{marginTop:8}}>
                            <summary style={{cursor:"pointer",fontSize:12,color:C.accent,
                              fontWeight:700,padding:"6px 0",listStyle:"none"}}>
                              📋 รายการใบกำกับภาษี ({m.salesData.items.length} รายการ) ▼
                            </summary>
                            <div style={{overflowX:"auto",marginTop:8}}>
                              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                                <thead>
                                  <tr style={{background:"#F8FAFF"}}>
                                    {["#","วันที่","เลขที่ใบกำกับ","ชื่อผู้ซื้อ","เลขผู้เสียภาษี","สาขา","มูลค่า","VAT"]
                                      .map((h,j)=><th key={j} style={{padding:"6px 10px",textAlign:j>5?"right":"left",
                                        fontWeight:700,color:C.muted,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>)}
                                  </tr>
                                </thead>
                                <tbody>
                                  {m.salesData.items.map((item,j)=>(
                                    <tr key={j} style={{background:j%2===0?C.surface:"#FAFBFF",borderBottom:`1px solid ${C.border}`}}>
                                      <td style={{padding:"6px 10px",color:C.light}}>{item.no}</td>
                                      <td style={{padding:"6px 10px",whiteSpace:"nowrap"}}>{item.date}</td>
                                      <td style={{padding:"6px 10px",fontFamily:"monospace",whiteSpace:"nowrap"}}>{item.invoiceNo}</td>
                                      <td style={{padding:"6px 10px",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.customer}</td>
                                      <td style={{padding:"6px 10px",fontFamily:"monospace"}}>{item.customerTaxId}</td>
                                      <td style={{padding:"6px 10px",textAlign:"center"}}>{item.branch}</td>
                                      <td style={{padding:"6px 10px",textAlign:"right",fontFamily:"monospace"}}>{fmtNum(item.amount)}</td>
                                      <td style={{padding:"6px 10px",textAlign:"right",fontFamily:"monospace"}}>{fmtNum(item.vat)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </details>
                        )}
                      </div>
                    </details>
                  );
                })}

                {/* Export button bottom */}
                <div style={{textAlign:"center",paddingTop:12,paddingBottom:24}}>
                  <button onClick={()=>exportExcel(months)} style={{
                    background:`linear-gradient(135deg,${C.green},#0D9488)`,color:"#fff",
                    border:"none",borderRadius:12,padding:"15px 40px",
                    fontWeight:800,fontSize:16,cursor:"pointer",
                    boxShadow:"0 6px 20px rgba(5,150,105,.3)",
                    display:"inline-flex",alignItems:"center",gap:10,
                  }}>
                    ⬇️ Export Excel ทั้งปี ({doneCount} เดือน)
                  </button>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}
