import { useState, useEffect, useRef, useCallback } from "react";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Legend } from "recharts";
import * as XLSX from "xlsx";

const P = {
  bg:"#0a0a0f", surface:"#12121a", card:"#18181f", border:"#252535",
  accent:"#7c6fff", green:"#22c55e", red:"#ef4444", yellow:"#f59e0b",
  text:"#eeeef5", muted:"#6b6b8a",
};
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const TABS = ["Budget","Actuals","Variance","Import"];
const PIE_COLORS = ["#7c6fff","#22c55e","#ef4444","#f59e0b","#06b6d4","#ec4899","#84cc16","#f97316","#8b5cf6","#14b8a6"];
const fmt = (n) => `R ${Number(n||0).toLocaleString("en-ZA",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtS = (n) => `R ${Number(n||0).toLocaleString("en-ZA",{minimumFractionDigits:0,maximumFractionDigits:0})}`;

const DEFAULT_INCOME = [{id:1,name:"Salary"},{id:2,name:"Freelance"},{id:3,name:"Business"}];
const DEFAULT_EXPENSE = [{id:4,name:"Rent"},{id:5,name:"Groceries"},{id:6,name:"Transport"},{id:7,name:"Utilities"},{id:8,name:"Entertainment"},{id:9,name:"Medical"},{id:10,name:"Other"}];
const DEFAULT_BUDGET = {Salary:25000,Freelance:5000,Business:3000,Rent:8000,Groceries:3500,Transport:1500,Utilities:900,Entertainment:1200,Medical:500,Other:800};
let nextId = 100;
const STORAGE_KEY = "budget_tracker_v3";
const save = (d) => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)); } catch(e) {} };
const load = () => { try { const d = localStorage.getItem(STORAGE_KEY); return d ? JSON.parse(d) : null; } catch(e) { return null; } };

// AmountInput is defined OUTSIDE the main component so it never remounts
// This is the key fix — defining it inside causes remount on every render
// which resets the cursor/value while typing
const AmountInput = ({ value, onSave, width, fontSize, readOnly }) => {
  const [local, setLocal] = useState(value === 0 ? "" : String(value || ""));
  const prevValue = useRef(value);
  useEffect(() => {
    // Only sync from parent if value actually changed externally (e.g. month switch)
    if (prevValue.current !== value) {
      prevValue.current = value;
      setLocal(value === 0 ? "" : String(value || ""));
    }
  }, [value]);
  return (
    <input
      type="text"
      inputMode="decimal"
      readOnly={!!readOnly}
      value={local}
      onChange={e => {
        if (readOnly) return;
        const v = e.target.value.replace(/[^0-9.]/g, "");
        setLocal(v);
      }}
      onBlur={() => { if (!readOnly && onSave) onSave(parseFloat(local) || 0); }}
      onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }}
      style={{
        background:"none", border:"none", outline:"none",
        color: readOnly ? P.muted : P.text,
        fontFamily:"inherit", textAlign:"right",
        width: width || 90, fontSize: fontSize || 13,
      }}
    />
  );
};

// Bank detection
const detectBank = (rows) => {
  const flat = rows.slice(0,15).map(r=>String(r||"").toLowerCase()).join(" ");
  if (flat.includes("first national")||flat.includes("fnb")) return "fnb";
  if (flat.includes("standard bank")) return "standardbank";
  if (flat.includes("absa")) return "absa";
  if (flat.includes("nedbank")) return "nedbank";
  if (flat.includes("capitec")) return "capitec";
  if (flat.includes("investec")) return "investec";
  if (flat.includes("african bank")) return "africanbank";
  if (flat.includes("tyme")) return "tymebank";
  return "generic";
};
const parseAmt = (v) => { if (!v) return 0; return parseFloat(String(v).replace(/[R,\s]/g,"").replace(/[()]/g,"-")) || 0; };
const parseDt = (v) => { if (!v) return ""; if (v instanceof Date) return v.toLocaleDateString("en-ZA"); return String(v); };

const bankParsers = {
  fnb: (rows) => rows.map(r=>r&&r.length>=4?{date:parseDt(r[0]),description:String(r[1]||r[2]||""),amount:parseAmt(r[3]||r[4])}:null).filter(t=>t&&t.date&&t.description&&t.amount!==0),
  standardbank: (rows) => rows.map(r=>r&&r.length>=3?{date:parseDt(r[0]),description:String(r[1]||""),amount:parseAmt(r[3])||(-Math.abs(parseAmt(r[2])))}:null).filter(t=>t&&t.date&&t.description&&t.amount!==0),
  absa: (rows) => rows.map(r=>r&&r.length>=3?{date:parseDt(r[0]),description:String(r[2]||r[1]||""),amount:parseAmt(r[3]||r[4])}:null).filter(t=>t&&t.date&&t.description&&t.amount!==0),
  nedbank: (rows) => rows.map(r=>r&&r.length>=3?{date:parseDt(r[0]),description:String(r[1]||""),amount:parseAmt(r[3]||r[2])}:null).filter(t=>t&&t.date&&t.description&&t.amount!==0),
  capitec: (rows) => rows.map(r=>r&&r.length>=3?{date:parseDt(r[0]),description:String(r[1]||""),amount:parseAmt(r[2]||r[3])}:null).filter(t=>t&&t.date&&t.description&&t.amount!==0),
  generic: (rows) => rows.map(r=>{
    if (!r||r.length<3) return null;
    let amount=0;
    for (let i=2;i<Math.min(r.length,6);i++){const v=parseAmt(r[i]);if(v!==0){amount=v;break;}}
    return {date:parseDt(r[0]),description:String(r[1]||r[2]||""),amount};
  }).filter(t=>t&&t.date&&t.description&&t.amount!==0),
};
bankParsers.investec=bankParsers.fnb;
bankParsers.africanbank=bankParsers.generic;
bankParsers.tymebank=bankParsers.generic;

const guessCategory = (desc, income, expense) => {
  const d = desc.toLowerCase();
  const all = [...income,...expense];
  for (const c of all) { if (d.includes(c.toLowerCase())) return c; }
  if (d.includes("salary")||d.includes("payroll")||d.includes("wages")) return income[0]||"Other";
  if (d.includes("freelance")||d.includes("consulting")) return income[1]||income[0]||"Other";
  if (d.includes("rent")||d.includes("lease")) return expense.find(c=>c.toLowerCase().includes("rent"))||expense[0];
  if (d.includes("grocery")||d.includes("checkers")||d.includes("pick n pay")||d.includes("spar")||d.includes("woolworths")||d.includes("shoprite")) return expense.find(c=>c.toLowerCase().includes("grocer"))||expense[0];
  if (d.includes("uber")||d.includes("bolt")||d.includes("petrol")||d.includes("fuel")||d.includes("engen")||d.includes("sasol")||d.includes("shell")||d.includes("caltex")) return expense.find(c=>c.toLowerCase().includes("transport"))||expense[0];
  if (d.includes("eskom")||d.includes("electricity")||d.includes("telkom")||d.includes("internet")||d.includes("mtn")||d.includes("vodacom")||d.includes("water")) return expense.find(c=>c.toLowerCase().includes("util"))||expense[0];
  if (d.includes("netflix")||d.includes("dstv")||d.includes("showmax")||d.includes("spotify")||d.includes("restaurant")||d.includes("nando")||d.includes("kfc")) return expense.find(c=>c.toLowerCase().includes("entertain"))||expense[0];
  if (d.includes("pharmacy")||d.includes("clicks")||d.includes("dischem")||d.includes("medical")||d.includes("doctor")||d.includes("hospital")) return expense.find(c=>c.toLowerCase().includes("medical"))||expense[0];
  return expense[expense.length-1]||"Other";
};

const parsePDFText = (text, income, expense) => {
  const lines = text.split("\n").map(l=>l.trim()).filter(Boolean);
  const txns = [];
  const dateRe = /(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}[/\-.]\d{2}[/\-.]\d{2})/;
  const amtRe = /(-?R?\s?[\d,]+[.]?\d{0,2})/g;
  lines.forEach(line=>{
    const dm = line.match(dateRe);
    if (!dm) return;
    const amounts = [...line.matchAll(amtRe)].map(m=>parseAmt(m[1])).filter(v=>v!==0);
    if (!amounts.length) return;
    const desc = line.replace(dateRe,"").replace(amtRe,"").trim().replace(/\s+/g," ");
    if (desc.length>2) txns.push({date:dm[1],description:desc,amount:amounts[amounts.length-1],category:guessCategory(desc,income,expense)});
  });
  return txns;
};

export default function BudgetTracker() {
  const curMonth = new Date().getMonth();
  const [activeTab, setActiveTab] = useState("Budget");
  const [selectedMonth, setSelectedMonth] = useState(curMonth);
  const [showYTD, setShowYTD] = useState(false);
  const [incomeItems, setIncomeItems] = useState(DEFAULT_INCOME);
  const [expenseItems, setExpenseItems] = useState(DEFAULT_EXPENSE);
  const [monthlyBudget, setMonthlyBudget] = useState(() => { const m={}; for(let i=0;i<12;i++) m[i]={...DEFAULT_BUDGET}; return m; });
  const [monthlyActuals, setMonthlyActuals] = useState(() => { const m={}; for(let i=0;i<12;i++) m[i]={}; return m; });
  const [transactions, setTransactions] = useState([]);
  const [uploadStatus, setUploadStatus] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editingVal, setEditingVal] = useState("");
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);
  const fileRef = useRef();

  useEffect(() => {
    const h = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);

  useEffect(() => {
    const d = load();
    if (!d) return;
    if (d.incomeItems) setIncomeItems(d.incomeItems);
    if (d.expenseItems) setExpenseItems(d.expenseItems);
    if (d.monthlyBudget) setMonthlyBudget(d.monthlyBudget);
    if (d.monthlyActuals) setMonthlyActuals(d.monthlyActuals);
    if (d.transactions) setTransactions(d.transactions);
  }, []);

  useEffect(() => {
    save({incomeItems,expenseItems,monthlyBudget,monthlyActuals,transactions});
  }, [incomeItems,expenseItems,monthlyBudget,monthlyActuals,transactions]);

  const allIncome = incomeItems.map(i=>i.name);
  const allExpense = expenseItems.map(i=>i.name);
  const allCats = [...allIncome,...allExpense];

  const getBudget = useCallback((m) => monthlyBudget[m]||{}, [monthlyBudget]);
  const getActuals = useCallback((m) => monthlyActuals[m]||{}, [monthlyActuals]);
  const ytdMonths = Array.from({length:selectedMonth+1},(_,i)=>i);
  const viewMonths = showYTD ? ytdMonths : [selectedMonth];

  const computeTotals = (months) => {
    const b={}, a={};
    allCats.forEach(c=>{
      b[c]=months.reduce((s,m)=>(s+(getBudget(m)[c]||0)),0);
      a[c]=months.reduce((s,m)=>(s+(getActuals(m)[c]||0)),0);
    });
    return {
      b, a,
      budgetIncome:allIncome.reduce((s,c)=>s+(b[c]||0),0),
      budgetExpense:allExpense.reduce((s,c)=>s+(b[c]||0),0),
      actualIncome:allIncome.reduce((s,c)=>s+(a[c]||0),0),
      actualExpense:allExpense.reduce((s,c)=>s+(a[c]||0),0),
    };
  };

  const {b:vB, a:vA, budgetIncome, budgetExpense, actualIncome, actualExpense} = computeTotals(viewMonths);

  const addItem = (type) => {
    const name = `New ${type==="income"?"Income":"Expense"} ${nextId}`;
    const item = {id:nextId++, name};
    if (type==="income") setIncomeItems(p=>[...p,item]); else setExpenseItems(p=>[...p,item]);
    setEditingId(item.id); setEditingVal(name);
  };

  const deleteItem = (type, id, name) => {
    if (type==="income") setIncomeItems(p=>p.filter(i=>i.id!==id)); else setExpenseItems(p=>p.filter(i=>i.id!==id));
    setMonthlyBudget(mb=>{const n={...mb};for(let m=0;m<12;m++){n[m]={...n[m]};delete n[m][name];}return n;});
    setMonthlyActuals(ma=>{const n={...ma};for(let m=0;m<12;m++){n[m]={...n[m]};delete n[m][name];}return n;});
  };

  const commitRename = (type, id, oldName) => {
    const newName = editingVal.trim();
    if (!newName||newName===oldName){setEditingId(null);return;}
    const upd = items=>items.map(i=>i.id===id?{...i,name:newName}:i);
    if (type==="income") setIncomeItems(upd); else setExpenseItems(upd);
    setMonthlyBudget(mb=>{const n={...mb};for(let m=0;m<12;m++){n[m]={...n[m]};n[m][newName]=n[m][oldName]||0;delete n[m][oldName];}return n;});
    setMonthlyActuals(ma=>{const n={...ma};for(let m=0;m<12;m++){n[m]={...n[m]};n[m][newName]=n[m][oldName]||0;delete n[m][oldName];}return n;});
    setEditingId(null);
  };

  const setBV = (month,cat,val) => setMonthlyBudget(mb=>({...mb,[month]:{...mb[month],[cat]:val}}));
  const setAV = (month,cat,val) => setMonthlyActuals(ma=>({...ma,[month]:{...ma[month],[cat]:val}}));

  const applyTransactions = (txns, month) => {
    setTransactions(prev=>[...prev.filter(t=>t.month!==month),...txns]);
    const na={};
    txns.forEach(({category,amount})=>{na[category]=(na[category]||0)+Math.abs(amount);});
    setMonthlyActuals(ma=>({...ma,[month]:{...ma[month],...na}}));
  };

  const processExcel = (file, month) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result,{type:"array",cellDates:true});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws,{header:1,raw:false});
        const bank = detectBank(rows.flat().map(String));
        const parser = bankParsers[bank]||bankParsers.generic;
        const raw = parser(rows.slice(1).filter(r=>r&&r.length>1));
        const txns = raw.map(t=>({...t,category:guessCategory(t.description,allIncome,allExpense),month}));
        applyTransactions(txns, month);
        setUploadStatus(`Loaded ${txns.length} transactions (${bank.toUpperCase()} detected)`);
      } catch(err) { setUploadStatus("Could not parse file. Please use a valid bank statement export."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const processPDF = async (file, month) => {
    setUploadStatus("Reading PDF...");
    try {
      const ab = await file.arrayBuffer();
      const str = new TextDecoder("latin1").decode(new Uint8Array(ab));
      let text = "";
      const matches = str.matchAll(/stream([\s\S]*?)endstream/g);
      for (const m of matches) { text += m[1].replace(/[^\x20-\x7E\n]/g," ").replace(/\s+/g," ") + "\n"; }
      const txns = parsePDFText(text,allIncome,allExpense).map(t=>({...t,month}));
      if (!txns.length) { setUploadStatus("Could not extract transactions. Try exporting as Excel/CSV instead."); return; }
      applyTransactions(txns, month);
      setUploadStatus(`Extracted ${txns.length} transactions from PDF`);
    } catch(e) { setUploadStatus("PDF failed. Please use Excel/CSV export."); }
  };

  const handleFile = (file, month) => {
    if (!file) return;
    const n = file.name.toLowerCase();
    if (n.endsWith(".xlsx")||n.endsWith(".xls")||n.endsWith(".csv")) processExcel(file,month);
    else if (n.endsWith(".pdf")) processPDF(file,month);
    else setUploadStatus("Please upload PDF, Excel (.xlsx) or CSV.");
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const md=[["Month","Category","Type","Budgeted","Actual","Variance","Variance %"]];
    MONTHS.forEach((mn,mi)=>{
      allCats.forEach(cat=>{
        const bv=getBudget(mi)[cat]||0, av=getActuals(mi)[cat]||0, vr=av-bv;
        md.push([mn,cat,allIncome.includes(cat)?"Income":"Expense",bv,av,vr,bv?((vr/bv)*100).toFixed(1)+"%":"N/A"]);
      });
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(md),"Monthly Detail");
    const {b:yb,a:ya}=computeTotals(ytdMonths);
    const yd=[["Category","Type","YTD Budgeted","YTD Actual","YTD Variance","Variance %"]];
    allCats.forEach(cat=>{
      const bv=yb[cat]||0,av=ya[cat]||0,vr=av-bv;
      yd.push([cat,allIncome.includes(cat)?"Income":"Expense",bv,av,vr,bv?((vr/bv)*100).toFixed(1)+"%":"N/A"]);
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(yd),"YTD Summary");
    if (transactions.length) {
      const td=[["Month","Date","Description","Amount","Category"]];
      transactions.forEach(t=>td.push([MONTHS[t.month]||"",t.date,t.description,t.amount,t.category]));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(td),"Transactions");
    }
    XLSX.writeFile(wb,`Budget_${new Date().getFullYear()}.xlsx`);
  };

  const varData = allCats.map(cat=>({
    name:cat, Budgeted:vB[cat]||0, Actual:vA[cat]||0,
    variance:(vA[cat]||0)-(vB[cat]||0), isIncome:allIncome.includes(cat),
  }));
  const expPie = allExpense.map(cat=>({name:cat,budgeted:vB[cat]||0,actual:vA[cat]||0})).filter(d=>d.budgeted>0||d.actual>0);
  const incPie = allIncome.map(cat=>({name:cat,budgeted:vB[cat]||0,actual:vA[cat]||0})).filter(d=>d.budgeted>0||d.actual>0);

  const card = {background:P.card,border:`1px solid ${P.border}`,borderRadius:12,padding:isMobile?14:20};
  const sTitle = {fontFamily:"'Syne',sans-serif",fontWeight:700,fontSize:13,letterSpacing:1,textTransform:"uppercase",marginBottom:14};
  const amtBox = {display:"flex",alignItems:"center",gap:4,background:P.bg,border:`1px solid ${P.border}`,borderRadius:8,padding:"5px 10px"};

  // CatRow also defined outside render in spirit — it uses AmountInput which is stable
  const CatRow = ({item, type, budgetVal, onBudget}) => {
    const isEd = editingId===item.id;
    return (
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:9}}>
        {isEd ? (
          <input autoFocus value={editingVal} onChange={e=>setEditingVal(e.target.value)}
            onBlur={()=>commitRename(type,item.id,item.name)}
            onKeyDown={e=>{if(e.key==="Enter")commitRename(type,item.id,item.name);if(e.key==="Escape")setEditingId(null);}}
            style={{flex:1,background:P.bg,border:`1px solid ${P.accent}`,borderRadius:6,padding:"4px 8px",color:P.text,fontSize:12,fontFamily:"inherit",outline:"none"}}/>
        ):(
          <span onClick={()=>{setEditingId(item.id);setEditingVal(item.name);}} style={{flex:1,fontSize:12,color:P.muted,cursor:"pointer"}}>{item.name}</span>
        )}
        <div style={amtBox}>
          <span style={{color:P.muted,fontSize:11}}>R</span>
          <AmountInput
            value={budgetVal||0}
            onSave={(v)=>onBudget(item.name,v)}
            width={isMobile?70:90}
            fontSize={12}
          />
        </div>
        <button onClick={()=>{setEditingId(item.id);setEditingVal(item.name);}} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,padding:"2px 3px"}}>✏️</button>
        <button onClick={()=>deleteItem(type,item.id,item.name)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,padding:"2px 3px"}}>🗑</button>
      </div>
    );
  };

  return (
    <div style={{fontFamily:"'DM Mono','Courier New',monospace",background:P.bg,minHeight:"100vh",color:P.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:4px;} ::-webkit-scrollbar-track{background:#0a0a0f;} ::-webkit-scrollbar-thumb{background:#252535;border-radius:2px;}
        select option{background:#18181f;}
        .tab-btn{background:none;border:none;cursor:pointer;font-family:inherit;transition:all 0.2s;white-space:nowrap;}
        .add-btn{background:rgba(124,111,255,0.1);border:1px dashed #7c6fff;border-radius:8px;color:#9d95ff;cursor:pointer;font-family:inherit;font-size:11px;padding:7px;width:100%;margin-top:8px;letter-spacing:1px;}
        .add-btn:hover{background:rgba(124,111,255,0.2);}
        .upload-zone{border:2px dashed #252535;border-radius:12px;padding:28px;text-align:center;cursor:pointer;transition:all 0.2s;}
        .upload-zone:hover{border-color:#7c6fff;background:rgba(124,111,255,0.04);}
      `}</style>

      {/* Header */}
      <div style={{background:P.surface,borderBottom:`1px solid ${P.border}`,padding:isMobile?"14px 16px":"16px 28px",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:100}}>
        <div style={{width:32,height:32,background:P.accent,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>B</div>
        <div>
          <div style={{fontFamily:"'Syne',sans-serif",fontSize:isMobile?14:17,fontWeight:800,letterSpacing:1}}>BUDGET TRACKER</div>
          <div style={{fontSize:10,color:P.muted}}>Actual vs Forecasted — ZA</div>
        </div>
        <div style={{marginLeft:"auto",display:"flex",gap:6,overflowX:"auto"}}>
          {TABS.map(t=>(
            <button key={t} className="tab-btn" onClick={()=>setActiveTab(t)}
              style={{padding:isMobile?"6px 10px":"7px 16px",borderRadius:8,fontSize:isMobile?11:12,fontWeight:500,color:activeTab===t?"#fff":P.muted,background:activeTab===t?P.accent:"transparent",border:activeTab===t?"none":`1px solid ${P.border}`}}>
              {t}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:isMobile?"14px 12px":"22px 28px",maxWidth:1100,margin:"0 auto"}}>

        {/* Month bar */}
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:20}}>
          <select value={selectedMonth} onChange={e=>setSelectedMonth(Number(e.target.value))}
            style={{background:P.card,border:`1px solid ${P.border}`,borderRadius:8,color:P.text,padding:"7px 12px",fontSize:13,fontFamily:"inherit",cursor:"pointer",flex:isMobile?"1":"unset"}}>
            {MONTHS.map((m,i)=><option key={i} value={i}>{m}</option>)}
          </select>
          <button onClick={()=>setShowYTD(v=>!v)}
            style={{padding:"7px 14px",borderRadius:8,fontSize:12,cursor:"pointer",fontFamily:"inherit",background:showYTD?P.accent:"transparent",color:showYTD?"#fff":P.muted,border:showYTD?"none":`1px solid ${P.border}`}}>
            {showYTD?`YTD Jan–${MONTHS[selectedMonth].slice(0,3)}`:"YTD"}
          </button>
          <button onClick={exportExcel}
            style={{padding:"7px 14px",borderRadius:8,fontSize:12,cursor:"pointer",fontFamily:"inherit",background:"rgba(34,197,94,0.15)",color:P.green,border:"1px solid rgba(34,197,94,0.3)",marginLeft:"auto"}}>
            Export Excel
          </button>
        </div>

        {/* Summary cards */}
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:12,marginBottom:12}}>
          {[{l:"Budget Income",v:fmt(budgetIncome),c:P.green},{l:"Budget Expenses",v:fmt(budgetExpense),c:P.red},{l:"Actual Income",v:fmt(actualIncome),c:P.green},{l:"Actual Expenses",v:fmt(actualExpense),c:P.red}].map(({l,v,c})=>(
            <div key={l} style={card}>
              <div style={{fontSize:10,color:P.muted,textTransform:"uppercase",letterSpacing:1.2,marginBottom:6}}>{l}</div>
              <div style={{fontSize:isMobile?13:17,fontWeight:500,color:c}}>{v}</div>
            </div>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:12,marginBottom:22}}>
          {[{l:"Budgeted Net",v:budgetIncome-budgetExpense},{l:"Actual Net",v:actualIncome-actualExpense}].map(({l,v})=>(
            <div key={l} style={{...card,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontSize:10,color:P.muted,textTransform:"uppercase",letterSpacing:1.2}}>{l}</span>
              <span style={{fontSize:isMobile?16:20,fontWeight:500,color:v>=0?P.green:P.red}}>{fmt(v)}</span>
            </div>
          ))}
        </div>

        {/* BUDGET TAB */}
        {activeTab==="Budget" && (
          <div>
            <div style={{...card,marginBottom:14,fontSize:11,color:P.muted}}>
              Click any category name to rename it. Use the pencil to rename, bin to delete, and the + button to add new lines.
            </div>
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:20}}>
              {[{title:"Income Forecast",items:incomeItems,type:"income",color:P.green},{title:"Expense Forecast",items:expenseItems,type:"expense",color:P.red}].map(({title,items,type,color})=>(
                <div key={title} style={card}>
                  <div style={{...sTitle,color}}>{title}</div>
                  {items.map(item=>(
                    <CatRow key={item.id} item={item} type={type}
                      budgetVal={showYTD ? ytdMonths.reduce((s,m)=>s+(getBudget(m)[item.name]||0),0) : getBudget(selectedMonth)[item.name]||0}
                      onBudget={(name,val)=>{
                        if (showYTD) { ytdMonths.forEach(m=>setBV(m,name,val/ytdMonths.length)); }
                        else { setBV(selectedMonth,name,val); }
                      }}
                    />
                  ))}
                  <button className="add-btn" onClick={()=>addItem(type)}>+ ADD {type.toUpperCase()} LINE</button>
                  <div style={{borderTop:`1px solid ${P.border}`,marginTop:10,paddingTop:10,display:"flex",justifyContent:"space-between"}}>
                    <span style={{fontSize:11,color:P.muted}}>TOTAL</span>
                    <span style={{fontWeight:500,fontSize:14,color}}>{fmt(items.reduce((s,i)=>s+(vB[i.name]||0),0))}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ACTUALS TAB */}
        {activeTab==="Actuals" && (
          <div style={{display:"flex",flexDirection:"column",gap:16}}>
            <div style={card}>
              <div style={{...sTitle,color:P.accent}}>Manual Entry — {showYTD?`YTD Jan–${MONTHS[selectedMonth]}`:MONTHS[selectedMonth]}</div>
              <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10}}>
                {allCats.map(cat=>(
                  <div key={cat} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
                    <span style={{fontSize:12,color:allIncome.includes(cat)?P.green:P.muted,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cat}</span>
                    <div style={amtBox}>
                      <span style={{color:P.muted,fontSize:11}}>R</span>
                      <AmountInput
                        value={showYTD ? ytdMonths.reduce((s,m)=>s+(getActuals(m)[cat]||0),0) : getActuals(selectedMonth)[cat]||0}
                        onSave={(v)=>{ if (!showYTD) setAV(selectedMonth,cat,v); }}
                        readOnly={showYTD}
                        width={isMobile?80:100}
                        fontSize={13}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </div>
            {transactions.filter(t=>showYTD?t.month<=selectedMonth:t.month===selectedMonth).length>0 && (
              <div style={card}>
                <div style={{...sTitle,color:P.accent}}>
                  Imported Transactions ({transactions.filter(t=>showYTD?t.month<=selectedMonth:t.month===selectedMonth).length})
                </div>
                <div style={{maxHeight:280,overflowY:"auto"}}>
                  {transactions.filter(t=>showYTD?t.month<=selectedMonth:t.month===selectedMonth).slice(0,100).map((t,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${P.border}`}}>
                      <div style={{flex:1,minWidth:0,paddingRight:8}}>
                        <div style={{fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.description||"—"}</div>
                        <div style={{fontSize:10,color:P.muted}}>{t.date} · {MONTHS[t.month]} · {t.category}</div>
                      </div>
                      <span style={{fontSize:13,fontWeight:500,color:t.amount>=0?P.green:P.red,flexShrink:0}}>{fmt(Math.abs(t.amount))}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* VARIANCE TAB */}
        {activeTab==="Variance" && (
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            <div style={card}>
              <div style={{...sTitle,color:P.accent}}>Budget vs Actual</div>
              <ResponsiveContainer width="100%" height={isMobile?200:280}>
                <BarChart data={varData} margin={{top:0,right:0,left:isMobile?0:10,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={P.border}/>
                  <XAxis dataKey="name" tick={{fill:P.muted,fontSize:isMobile?8:10}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:P.muted,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R${(v/1000).toFixed(0)}k`} width={isMobile?38:50}/>
                  <Tooltip contentStyle={{background:P.card,border:`1px solid ${P.border}`,borderRadius:8,fontSize:12}} formatter={v=>fmt(v)}/>
                  <Legend wrapperStyle={{fontSize:11}}/>
                  <Bar dataKey="Budgeted" fill={P.accent} radius={[4,4,0,0]}/>
                  <Bar dataKey="Actual" radius={[4,4,0,0]}>
                    {varData.map((e,i)=>{
                      const over=e.variance>0;
                      return <Cell key={i} fill={e.isIncome?(over?P.green:P.yellow):(over?P.red:P.green)}/>;
                    })}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:20}}>
              {[{title:"Income",data:incPie},{title:"Expenses",data:expPie}].map(({title,data})=>(
                <div key={title} style={card}>
                  <div style={{...sTitle,color:P.accent}}>{title} — Budgeted vs Actual</div>
                  {["budgeted","actual"].map(key=>(
                    <div key={key}>
                      <div style={{fontSize:10,color:P.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:4,marginTop:key==="actual"?16:0}}>{key}</div>
                      <ResponsiveContainer width="100%" height={180}>
                        <PieChart>
                          <Pie data={data} dataKey={key} nameKey="name" cx="50%" cy="50%" outerRadius={65}
                            label={({name,percent})=>percent>0.05?`${name} ${(percent*100).toFixed(0)}%`:""} labelLine={false} fontSize={9}>
                            {data.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                          </Pie>
                          <Tooltip formatter={v=>fmt(v)}/>
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  ))}
                </div>
              ))}
            </div>
            <div style={card}>
              <div style={{...sTitle,color:P.accent}}>Variance Detail</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:isMobile?11:12,minWidth:480}}>
                  <thead>
                    <tr style={{color:P.muted,fontSize:10,textTransform:"uppercase",letterSpacing:1}}>
                      {["Category","Type","Budgeted","Actual","Variance","Status"].map(h=>(
                        <th key={h} style={{textAlign:["Category","Type","Status"].includes(h)?"left":"right",padding:"8px",borderBottom:`1px solid ${P.border}`}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {varData.map(row=>{
                      const good=row.isIncome?row.variance>=0:row.variance<=0;
                      const pct=row.Budgeted?((row.variance/row.Budgeted)*100).toFixed(1):"—";
                      return (
                        <tr key={row.name} style={{borderBottom:`1px solid ${P.border}`}}>
                          <td style={{padding:"9px 8px"}}>{row.name}</td>
                          <td style={{padding:"9px 8px",color:row.isIncome?P.green:P.red}}>{row.isIncome?"Income":"Expense"}</td>
                          <td style={{padding:"9px 8px",textAlign:"right",color:P.muted}}>{fmtS(row.Budgeted)}</td>
                          <td style={{padding:"9px 8px",textAlign:"right"}}>{fmtS(row.Actual)}</td>
                          <td style={{padding:"9px 8px",textAlign:"right",color:good?P.green:P.red}}>{row.variance>0?"+":""}{fmtS(row.variance)}</td>
                          <td style={{padding:"9px 8px"}}>
                            <span style={{background:good?"rgba(34,197,94,0.15)":"rgba(239,68,68,0.15)",color:good?P.green:P.red,borderRadius:4,padding:"2px 8px",fontSize:10}}>
                              {row.Actual===0?"No Data":good?`+${Math.abs(pct)}%`:`-${Math.abs(pct)}%`}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* IMPORT TAB */}
        {activeTab==="Import" && (
          <div style={{display:"flex",flexDirection:"column",gap:16}}>
            <div style={card}>
              <div style={{...sTitle,color:P.accent}}>Import Bank Statement</div>
              <div style={{marginBottom:14}}>
                <div style={{fontSize:10,color:P.muted,textTransform:"uppercase",letterSpacing:1.2,marginBottom:6}}>Select month for this statement</div>
                <select value={selectedMonth} onChange={e=>setSelectedMonth(Number(e.target.value))}
                  style={{background:P.bg,border:`1px solid ${P.border}`,borderRadius:8,color:P.text,padding:"8px 14px",fontSize:13,fontFamily:"inherit",width:"100%"}}>
                  {MONTHS.map((m,i)=><option key={i} value={i}>{m} {new Date().getFullYear()}</option>)}
                </select>
              </div>
              <div className="upload-zone" onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()}
                onDrop={e=>{e.preventDefault();handleFile(e.dataTransfer.files[0],selectedMonth);}}>
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf" style={{display:"none"}} onChange={e=>handleFile(e.target.files[0],selectedMonth)}/>
                <div style={{fontSize:32,marginBottom:10}}>📂</div>
                <div style={{fontSize:13,marginBottom:6}}>Drop your bank statement here or click to browse</div>
                <div style={{fontSize:11,color:P.muted,lineHeight:1.8}}>
                  Supports: PDF, Excel (.xlsx), CSV<br/>
                  FNB · Standard Bank · ABSA · Nedbank · Capitec · Investec · TymeBank · African Bank
                </div>
              </div>
              {uploadStatus && (
                <div style={{marginTop:12,padding:"10px 14px",background:P.bg,borderRadius:8,fontSize:12,color:P.muted,border:`1px solid ${P.border}`}}>{uploadStatus}</div>
              )}
            </div>
            <div style={card}>
              <div style={{...sTitle,color:P.accent}}>Import Tips</div>
              <div style={{fontSize:12,color:P.muted,lineHeight:2}}>
                <div>Always export directly from your bank online portal — not scanned copies</div>
                <div>FNB: Login → Accounts → Download Statement → Excel/CSV</div>
                <div>Standard Bank: Online Banking → Statement → Download → Excel</div>
                <div>ABSA: Online Banking → Statement → Export → CSV</div>
                <div>Nedbank: Money app → Statements → Download</div>
                <div>Capitec: Internet Banking → Transactions → Export</div>
                <div>Transactions are auto-categorised based on description keywords</div>
                <div>Re-importing the same month replaces previous data for that month</div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}