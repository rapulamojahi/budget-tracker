import { useState, useEffect, useRef, useCallback } from "react";
import { PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid } from "recharts";
import * as XLSX from "xlsx";

// ── Palette ──────────────────────────────────────────────────────────────────
const P = {
  bg: "#0a0a0f", surface: "#12121a", card: "#18181f", border: "#252535",
  accent: "#7c6fff", accentHover: "#9d95ff", green: "#22c55e", red: "#ef4444",
  yellow: "#f59e0b", text: "#eeeef5", muted: "#6b6b8a", faint: "#1e1e2e",
};

const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const TABS = ["Budget","Actuals","Variance","Import"];
const PIE_COLORS = ["#7c6fff","#22c55e","#ef4444","#f59e0b","#06b6d4","#ec4899","#84cc16","#f97316","#8b5cf6","#14b8a6"];

const fmt = (n) => `R ${Number(n||0).toLocaleString("en-ZA",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtShort = (n) => `R ${Number(n||0).toLocaleString("en-ZA",{minimumFractionDigits:0,maximumFractionDigits:0})}`;

// ── Default Data ──────────────────────────────────────────────────────────────
const DEFAULT_INCOME = [{id:1,name:"Salary"},{id:2,name:"Freelance"},{id:3,name:"Business"}];
const DEFAULT_EXPENSE = [{id:4,name:"Rent"},{id:5,name:"Groceries"},{id:6,name:"Transport"},{id:7,name:"Utilities"},{id:8,name:"Entertainment"},{id:9,name:"Medical"},{id:10,name:"Other"}];
const DEFAULT_BUDGET = {Salary:25000,Freelance:5000,Business:3000,Rent:8000,Groceries:3500,Transport:1500,Utilities:900,Entertainment:1200,Medical:500,Other:800};

let nextId = 100;

// ── Bank Parsers ──────────────────────────────────────────────────────────────
const detectBank = (rows) => {
  const flat = rows.slice(0,15).map(r=>String(r||"").toLowerCase()).join(" ");
  if (flat.includes("first national") || flat.includes("fnb")) return "fnb";
  if (flat.includes("standard bank")) return "standardbank";
  if (flat.includes("absa")) return "absa";
  if (flat.includes("nedbank")) return "nedbank";
  if (flat.includes("capitec")) return "capitec";
  if (flat.includes("investec")) return "investec";
  if (flat.includes("african bank")) return "africanbank";
  if (flat.includes("tyme") || flat.includes("tymebank")) return "tymebank";
  return "generic";
};

const parseAmount = (val) => {
  if (!val) return 0;
  const s = String(val).replace(/[R,\s]/g,"").replace(/[()]/g,"-");
  return parseFloat(s) || 0;
};

const parseDate = (val) => {
  if (!val) return "";
  if (val instanceof Date) return val.toLocaleDateString("en-ZA");
  return String(val);
};

const bankParsers = {
  fnb: (rows) => {
    const txns = [];
    rows.forEach(row => {
      if (!row || row.length < 4) return;
      const date = parseDate(row[0]);
      const desc = String(row[1]||row[2]||"");
      const amount = parseAmount(row[3]||row[4]);
      if (date && desc && amount !== 0) txns.push({date,description:desc,amount});
    });
    return txns;
  },
  standardbank: (rows) => {
    const txns = [];
    rows.forEach(row => {
      if (!row || row.length < 3) return;
      const date = parseDate(row[0]);
      const desc = String(row[1]||"");
      const debit = parseAmount(row[2]);
      const credit = parseAmount(row[3]);
      const amount = credit !== 0 ? credit : -Math.abs(debit);
      if (date && desc && amount !== 0) txns.push({date,description:desc,amount});
    });
    return txns;
  },
  absa: (rows) => {
    const txns = [];
    rows.forEach(row => {
      if (!row || row.length < 3) return;
      const date = parseDate(row[0]);
      const desc = String(row[2]||row[1]||"");
      const amount = parseAmount(row[3]||row[4]);
      if (date && desc && amount !== 0) txns.push({date,description:desc,amount});
    });
    return txns;
  },
  nedbank: (rows) => {
    const txns = [];
    rows.forEach(row => {
      if (!row || row.length < 3) return;
      const date = parseDate(row[0]);
      const desc = String(row[1]||"");
      const amount = parseAmount(row[3]||row[2]);
      if (date && desc && amount !== 0) txns.push({date,description:desc,amount});
    });
    return txns;
  },
  capitec: (rows) => {
    const txns = [];
    rows.forEach(row => {
      if (!row || row.length < 3) return;
      const date = parseDate(row[0]);
      const desc = String(row[1]||"");
      const amount = parseAmount(row[2]||row[3]);
      if (date && desc && amount !== 0) txns.push({date,description:desc,amount});
    });
    return txns;
  },
  generic: (rows) => {
    const txns = [];
    rows.forEach(row => {
      if (!row || row.length < 3) return;
      const date = parseDate(row[0]);
      const desc = String(row[1]||row[2]||"");
      let amount = 0;
      for (let i=2; i<Math.min(row.length,6); i++) {
        const v = parseAmount(row[i]);
        if (v !== 0) { amount = v; break; }
      }
      if (date && desc && amount !== 0) txns.push({date,description:desc,amount});
    });
    return txns;
  },
};
bankParsers.investec = bankParsers.fnb;
bankParsers.africanbank = bankParsers.generic;
bankParsers.tymebank = bankParsers.generic;

// ── Category Guesser ──────────────────────────────────────────────────────────
const guessCategory = (desc, allIncome, allExpense) => {
  const d = desc.toLowerCase();
  const all = [...allIncome,...allExpense];
  for (const cat of all) { if (d.includes(cat.toLowerCase())) return cat; }
  if (d.includes("salary")||d.includes("payroll")||d.includes("wages")||d.includes("remuneration")) return allIncome[0]||"Other";
  if (d.includes("freelance")||d.includes("consulting")||d.includes("invoice")) return allIncome[1]||allIncome[0]||"Other";
  if (d.includes("rent")||d.includes("lease")||d.includes("accommodation")) return allExpense.find(c=>c.toLowerCase().includes("rent"))||allExpense[0];
  if (d.includes("grocery")||d.includes("checkers")||d.includes("pick n pay")||d.includes("pnp")||d.includes("spar")||d.includes("woolworths")||d.includes("shoprite")||d.includes("usave")||d.includes("food")) return allExpense.find(c=>c.toLowerCase().includes("grocer"))||allExpense[0];
  if (d.includes("uber")||d.includes("bolt")||d.includes("petrol")||d.includes("fuel")||d.includes("engen")||d.includes("sasol")||d.includes("shell")||d.includes("caltex")||d.includes("bp ")||d.includes("taxi")||d.includes("bus")) return allExpense.find(c=>c.toLowerCase().includes("transport"))||allExpense[0];
  if (d.includes("eskom")||d.includes("electricity")||d.includes("telkom")||d.includes("internet")||d.includes("mtn")||d.includes("vodacom")||d.includes("cell c")||d.includes("rain")||d.includes("water")||d.includes("rates")) return allExpense.find(c=>c.toLowerCase().includes("util"))||allExpense[0];
  if (d.includes("netflix")||d.includes("dstv")||d.includes("showmax")||d.includes("amazon prime")||d.includes("spotify")||d.includes("gaming")||d.includes("cinema")||d.includes("restaurant")||d.includes("nando")||d.includes("kfc")||d.includes("steers")||d.includes("mcdo")) return allExpense.find(c=>c.toLowerCase().includes("entertain"))||allExpense[0];
  if (d.includes("pharmacy")||d.includes("clicks")||d.includes("dischem")||d.includes("medical")||d.includes("doctor")||d.includes("hospital")||d.includes("clinic")||d.includes("medihelp")||d.includes("discovery health")) return allExpense.find(c=>c.toLowerCase().includes("medical"))||allExpense[0];
  return allExpense[allExpense.length-1]||"Other";
};

// ── PDF Text Parser ───────────────────────────────────────────────────────────
const parsePDFText = (text, allIncome, allExpense) => {
  const lines = text.split("\n").map(l=>l.trim()).filter(Boolean);
  const txns = [];
  const dateRe = /(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}[/\-.]\d{2}[/\-.]\d{2})/;
  const amtRe = /(-?R?\s?[\d,]+[.]?\d{0,2})/g;
  lines.forEach(line => {
    const dateMatch = line.match(dateRe);
    if (!dateMatch) return;
    const date = dateMatch[1];
    const amounts = [...line.matchAll(amtRe)].map(m=>parseAmount(m[1])).filter(v=>v!==0);
    if (amounts.length === 0) return;
    const desc = line.replace(dateRe,"").replace(amtRe,"").trim().replace(/\s+/g," ");
    const amount = amounts[amounts.length-1];
    if (desc.length > 2) txns.push({date, description:desc, amount, category:guessCategory(desc,allIncome,allExpense)});
  });
  return txns;
};

// ── Storage Helpers ───────────────────────────────────────────────────────────
const STORAGE_KEY = "budget_tracker_v2";
const saveToStorage = (data) => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch(e) {} };
const loadFromStorage = () => { try { const d = localStorage.getItem(STORAGE_KEY); return d ? JSON.parse(d) : null; } catch(e) { return null; } };

// ── Main App ──────────────────────────────────────────────────────────────────
export default function BudgetTracker() {
  const currentMonth = new Date().getMonth();
  const [activeTab, setActiveTab] = useState("Budget");
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);
  const [showYTD, setShowYTD] = useState(false);
  const [incomeItems, setIncomeItems] = useState(DEFAULT_INCOME);
  const [expenseItems, setExpenseItems] = useState(DEFAULT_EXPENSE);
  // monthlyBudget[month][category] = amount
  const [monthlyBudget, setMonthlyBudget] = useState(() => {
    const mb = {};
    for (let i=0;i<12;i++) mb[i] = {...DEFAULT_BUDGET};
    return mb;
  });
  // monthlyActuals[month][category] = amount
  const [monthlyActuals, setMonthlyActuals] = useState(() => {
    const ma = {};
    for (let i=0;i<12;i++) ma[i] = {};
    return ma;
  });
  const [transactions, setTransactions] = useState([]);
  const [uploadStatus, setUploadStatus] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editingVal, setEditingVal] = useState("");
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);
  const fileRef = useRef();

  // Responsive
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);

  // Load from storage
  useEffect(() => {
    const saved = loadFromStorage();
    if (!saved) return;
    if (saved.incomeItems) setIncomeItems(saved.incomeItems);
    if (saved.expenseItems) setExpenseItems(saved.expenseItems);
    if (saved.monthlyBudget) setMonthlyBudget(saved.monthlyBudget);
    if (saved.monthlyActuals) setMonthlyActuals(saved.monthlyActuals);
    if (saved.transactions) setTransactions(saved.transactions);
  }, []);

  // Save to storage on every change
  useEffect(() => {
    saveToStorage({incomeItems, expenseItems, monthlyBudget, monthlyActuals, transactions});
  }, [incomeItems, expenseItems, monthlyBudget, monthlyActuals, transactions]);

  const allIncomeCats = incomeItems.map(i=>i.name);
  const allExpenseCats = expenseItems.map(i=>i.name);
  const allCats = [...allIncomeCats,...allExpenseCats];

  // ── Computed values ──────────────────────────────────────────────────────────
  const getBudget = useCallback((month) => monthlyBudget[month]||{}, [monthlyBudget]);
  const getActuals = useCallback((month) => monthlyActuals[month]||{}, [monthlyActuals]);

  const computeTotals = (months) => {
    const budget = {}, actuals = {};
    allCats.forEach(c => {
      budget[c] = months.reduce((s,m)=>(s+(getBudget(m)[c]||0)),0);
      actuals[c] = months.reduce((s,m)=>(s+(getActuals(m)[c]||0)),0);
    });
    const budgetIncome = allIncomeCats.reduce((s,c)=>s+(budget[c]||0),0);
    const budgetExpense = allExpenseCats.reduce((s,c)=>s+(budget[c]||0),0);
    const actualIncome = allIncomeCats.reduce((s,c)=>s+(actuals[c]||0),0);
    const actualExpense = allExpenseCats.reduce((s,c)=>s+(actuals[c]||0),0);
    return {budget, actuals, budgetIncome, budgetExpense, actualIncome, actualExpense};
  };

  const ytdMonths = Array.from({length:selectedMonth+1},(_,i)=>i);
  const viewMonths = showYTD ? ytdMonths : [selectedMonth];
  const {budget:viewBudget, actuals:viewActuals, budgetIncome, budgetExpense, actualIncome, actualExpense} = computeTotals(viewMonths);

  // ── Category Management ──────────────────────────────────────────────────────
  const addItem = (type) => {
    const name = `New ${type==="income"?"Income":"Expense"} ${nextId}`;
    const item = {id:nextId++, name};
    if (type==="income") setIncomeItems(p=>[...p,item]);
    else setExpenseItems(p=>[...p,item]);
    setEditingId(item.id); setEditingVal(name);
  };

  const deleteItem = (type, id, name) => {
    if (type==="income") setIncomeItems(p=>p.filter(i=>i.id!==id));
    else setExpenseItems(p=>p.filter(i=>i.id!==id));
    setMonthlyBudget(mb=>{ const n={...mb}; for(let m=0;m<12;m++){n[m]={...n[m]};delete n[m][name];} return n; });
    setMonthlyActuals(ma=>{ const n={...ma}; for(let m=0;m<12;m++){n[m]={...n[m]};delete n[m][name];} return n; });
  };

  const commitRename = (type, id, oldName) => {
    const newName = editingVal.trim();
    if (!newName||newName===oldName){setEditingId(null);return;}
    const upd = items=>items.map(i=>i.id===id?{...i,name:newName}:i);
    if (type==="income") setIncomeItems(upd); else setExpenseItems(upd);
    setMonthlyBudget(mb=>{ const n={...mb}; for(let m=0;m<12;m++){n[m]={...n[m]};n[m][newName]=n[m][oldName]||0;delete n[m][oldName];} return n; });
    setMonthlyActuals(ma=>{ const n={...ma}; for(let m=0;m<12;m++){n[m]={...n[m]};n[m][newName]=n[m][oldName]||0;delete n[m][oldName];} return n; });
    setEditingId(null);
  };

  const setBudgetVal = (month, cat, val) => setMonthlyBudget(mb=>({...mb,[month]:{...mb[month],[cat]:parseFloat(val)||0}}));
  const setActualVal = (month, cat, val) => setMonthlyActuals(ma=>({...ma,[month]:{...ma[month],[cat]:parseFloat(val)||0}}));

  // ── File Import ──────────────────────────────────────────────────────────────
  const processExcel = (file, targetMonth) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {type:"array", cellDates:true});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
        const bank = detectBank(rows.flat().map(String));
        const parser = bankParsers[bank]||bankParsers.generic;
        const dataRows = rows.filter(r=>r&&r.length>1).slice(1);
        const raw = parser(dataRows);
        const txns = raw.map(t=>({...t, category:guessCategory(t.description,allIncomeCats,allExpenseCats), month:targetMonth}));
        applyTransactions(txns, targetMonth);
        setUploadStatus(`✓ Loaded ${txns.length} transactions (${bank.toUpperCase()} format detected)`);
      } catch(err) { setUploadStatus("⚠ Could not parse file. Ensure it's a valid bank statement export."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const processPDF = async (file, targetMonth) => {
    setUploadStatus("⏳ Reading PDF...");
    try {
      const arrayBuffer = await file.arrayBuffer();
      const uint8 = new Uint8Array(arrayBuffer);
      let text = "";
      // Extract readable text from PDF bytes
      const str = new TextDecoder("latin1").decode(uint8);
      const streamMatches = str.matchAll(/stream([\s\S]*?)endstream/g);
      for (const match of streamMatches) {
        const chunk = match[1].replace(/[^\x20-\x7E\n]/g," ").replace(/\s+/g," ");
        text += chunk + "\n";
      }
      const txns = parsePDFText(text, allIncomeCats, allExpenseCats).map(t=>({...t, month:targetMonth}));
      if (txns.length === 0) {
        setUploadStatus("⚠ Could not extract transactions from PDF. Try exporting as Excel/CSV from your bank's online portal instead.");
        return;
      }
      applyTransactions(txns, targetMonth);
      setUploadStatus(`✓ Extracted ${txns.length} transactions from PDF`);
    } catch(err) { setUploadStatus("⚠ PDF reading failed. Please use Excel/CSV export instead."); }
  };

  const applyTransactions = (txns, month) => {
    setTransactions(prev=>[...prev.filter(t=>t.month!==month),...txns]);
    const newActuals = {};
    txns.forEach(({category,amount})=>{
      newActuals[category]=(newActuals[category]||0)+Math.abs(amount);
    });
    setMonthlyActuals(ma=>({...ma,[month]:{...ma[month],...newActuals}}));
  };

  const handleFile = (file, targetMonth) => {
    if (!file) return;
    const name = file.name.toLowerCase();
    if (name.endsWith(".xlsx")||name.endsWith(".xls")||name.endsWith(".csv")) processExcel(file, targetMonth);
    else if (name.endsWith(".pdf")) processPDF(file, targetMonth);
    else setUploadStatus("⚠ Please upload PDF, Excel (.xlsx) or CSV file.");
  };

  // ── Export to Excel ──────────────────────────────────────────────────────────
  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    // Monthly sheet
    const monthData = [["Month","Category","Type","Budgeted","Actual","Variance","Variance %"]];
    MONTHS.forEach((mName,mi)=>{
      const mb=getBudget(mi); const ma=getActuals(mi);
      allCats.forEach(cat=>{
        const bVal=mb[cat]||0; const aVal=ma[cat]||0; const vari=aVal-bVal;
        const pct=bVal?((vari/bVal)*100).toFixed(1)+"%":"N/A";
        monthData.push([mName,cat,allIncomeCats.includes(cat)?"Income":"Expense",bVal,aVal,vari,pct]);
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(monthData), "Monthly Detail");
    // YTD sheet
    const ytdData = [["Category","Type","YTD Budgeted","YTD Actual","YTD Variance","Variance %"]];
    const {budget:yb,actuals:ya} = computeTotals(ytdMonths);
    allCats.forEach(cat=>{
      const bVal=yb[cat]||0; const aVal=ya[cat]||0; const vari=aVal-bVal;
      const pct=bVal?((vari/bVal)*100).toFixed(1)+"%":"N/A";
      ytdData.push([cat,allIncomeCats.includes(cat)?"Income":"Expense",bVal,aVal,vari,pct]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ytdData), "YTD Summary");
    // Transactions sheet
    if (transactions.length>0){
      const txData=[["Month","Date","Description","Amount","Category"]];
      transactions.forEach(t=>txData.push([MONTHS[t.month]||"",t.date,t.description,t.amount,t.category]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(txData), "Transactions");
    }
    XLSX.writeFile(wb, `Budget_Tracker_${new Date().getFullYear()}.xlsx`);
  };

  // ── Variance Data ────────────────────────────────────────────────────────────
  const varData = allCats.map(cat=>({
    name:cat, Budgeted:viewBudget[cat]||0, Actual:viewActuals[cat]||0,
    variance:(viewActuals[cat]||0)-(viewBudget[cat]||0),
    isIncome:allIncomeCats.includes(cat),
  }));

  const expensePieData = allExpenseCats.map(cat=>({name:cat,budgeted:viewBudget[cat]||0,actual:viewActuals[cat]||0})).filter(d=>d.budgeted>0||d.actual>0);
  const incomePieData = allIncomeCats.map(cat=>({name:cat,budgeted:viewBudget[cat]||0,actual:viewActuals[cat]||0})).filter(d=>d.budgeted>0||d.actual>0);

  // ── Styles ───────────────────────────────────────────────────────────────────
  const s = {
    card: { background:P.card, border:`1px solid ${P.border}`, borderRadius:12, padding:isMobile?14:20 },
    label: { fontSize:10, color:P.muted, textTransform:"uppercase", letterSpacing:1.2, marginBottom:6 },
    sectionTitle: { fontFamily:"'Syne',sans-serif", fontWeight:700, fontSize:13, letterSpacing:1, textTransform:"uppercase", marginBottom:14 },
    input: { background:"none", border:"none", outline:"none", color:P.text, fontFamily:"inherit", textAlign:"right" },
    amtBox: { display:"flex", alignItems:"center", gap:4, background:P.bg, border:`1px solid ${P.border}`, borderRadius:8, padding:"5px 10px" },
    badge: (good) => ({ background:good?"rgba(34,197,94,0.15)":"rgba(239,68,68,0.15)", color:good?P.green:P.red, borderRadius:4, padding:"2px 8px", fontSize:10 }),
  };

  // ── Category Row ─────────────────────────────────────────────────────────────
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
        <div style={s.amtBox}>
          <span style={{color:P.muted,fontSize:11}}>R</span>
          <input type="text" inputMode="decimal" value={budgetVal||""} onChange={e=>{const v=e.target.value.replace(/[^0-9.]/g,"");onBudget(item.name,v);}} style={{...s.input,width:isMobile?70:90,fontSize:12}}/>
        </div>
        <button onClick={()=>{setEditingId(item.id);setEditingVal(item.name);}} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,padding:"2px 3px"}}>✏️</button>
        <button onClick={()=>deleteItem(type,item.id,item.name)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,padding:"2px 3px"}}>🗑</button>
      </div>
    );
  };

  // ── Month/YTD Selector Bar ───────────────────────────────────────────────────
  const MonthBar = () => (
    <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:20}}>
      <select value={selectedMonth} onChange={e=>setSelectedMonth(Number(e.target.value))}
        style={{background:P.card,border:`1px solid ${P.border}`,borderRadius:8,color:P.text,padding:"7px 12px",fontSize:13,fontFamily:"inherit",cursor:"pointer",flex:isMobile?"1":"unset"}}>
        {MONTHS.map((m,i)=><option key={i} value={i}>{m}</option>)}
      </select>
      <button onClick={()=>setShowYTD(v=>!v)}
        style={{padding:"7px 14px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"inherit",background:showYTD?P.accent:"transparent",color:showYTD?"#fff":P.muted,border:showYTD?"none":`1px solid ${P.border}`,transition:"all 0.2s"}}>
        {showYTD?`YTD (Jan–${MONTHS[selectedMonth].slice(0,3)})`:"YTD"}
      </button>
      <button onClick={exportExcel}
        style={{padding:"7px 14px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"inherit",background:"rgba(34,197,94,0.15)",color:P.green,border:`1px solid rgba(34,197,94,0.3)`,marginLeft:"auto"}}>
        ⬇ Export Excel
      </button>
    </div>
  );

  // ── Summary Cards ─────────────────────────────────────────────────────────────
  const SummaryCards = () => (
    <>
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:12,marginBottom:12}}>
        {[{label:"Budget Income",val:fmt(budgetIncome),color:P.green},{label:"Budget Expenses",val:fmt(budgetExpense),color:P.red},{label:"Actual Income",val:fmt(actualIncome),color:P.green},{label:"Actual Expenses",val:fmt(actualExpense),color:P.red}].map(({label,val,color})=>(
          <div key={label} style={s.card}>
            <div style={s.label}>{label}</div>
            <div style={{fontSize:isMobile?14:17,fontWeight:500,color}}>{val}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:12,marginBottom:22}}>
        {[{label:"Budgeted Net",val:budgetIncome-budgetExpense},{label:"Actual Net",val:actualIncome-actualExpense}].map(({label,val})=>(
          <div key={label} style={{...s.card,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={s.label}>{label}</span>
            <span style={{fontSize:isMobile?16:20,fontWeight:500,color:val>=0?P.green:P.red}}>{fmt(val)}</span>
          </div>
        ))}
      </div>
    </>
  );

  // ── Render ───────────────────────────────────────────────────────────────────
  return (
    <div style={{fontFamily:"'DM Mono','Courier New',monospace",background:P.bg,minHeight:"100vh",color:P.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:4px;} ::-webkit-scrollbar-track{background:#0a0a0f;} ::-webkit-scrollbar-thumb{background:#252535;border-radius:2px;}
        
        select option{background:#18181f;}
        .tab-btn{background:none;border:none;cursor:pointer;font-family:inherit;transition:all 0.2s;white-space:nowrap;}
        .tab-btn:hover{color:#fff;}
        .add-btn{background:rgba(124,111,255,0.1);border:1px dashed #7c6fff;border-radius:8px;color:#9d95ff;cursor:pointer;font-family:inherit;font-size:11px;padding:7px;width:100%;margin-top:8px;letter-spacing:1px;}
        .add-btn:hover{background:rgba(124,111,255,0.2);}
        .upload-zone{border:2px dashed #252535;border-radius:12px;padding:28px;text-align:center;cursor:pointer;transition:all 0.2s;}
        .upload-zone:hover{border-color:#7c6fff;background:rgba(124,111,255,0.04);}
      `}</style>

      {/* Header */}
      <div style={{background:P.surface,borderBottom:`1px solid ${P.border}`,padding:isMobile?"14px 16px":"16px 28px",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:100}}>
        <div style={{width:32,height:32,background:P.accent,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>₿</div>
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
        <MonthBar/>
        <SummaryCards/>

        {/* ── BUDGET TAB ── */}
        {activeTab==="Budget" && (
          <div>
            <div style={{...s.card,marginBottom:14,fontSize:11,color:P.muted}}>
              💡 <strong style={{color:P.accent}}>Tip:</strong> Click any name to rename. Changes apply to the selected month only. Toggle YTD to see cumulative totals.
            </div>
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:20}}>
              {[{title:"Income Forecast",items:incomeItems,type:"income",color:P.green},{title:"Expense Forecast",items:expenseItems,type:"expense",color:P.red}].map(({title,items,type,color})=>(
                <div key={title} style={s.card}>
                  <div style={{...s.sectionTitle,color}}>{title}</div>
                  {items.map(item=>(
                    <CatRow key={item.id} item={item} type={type}
                      budgetVal={showYTD?ytdMonths.reduce((s,m)=>(s+(getBudget(m)[item.name]||0)),0):getBudget(selectedMonth)[item.name]}
                      onBudget={showYTD?(name,val)=>{ytdMonths.forEach(m=>setBudgetVal(m,name,(parseFloat(val)||0)/ytdMonths.length));}:(name,val)=>setBudgetVal(selectedMonth,name,val)}/>
                  ))}
                  <button className="add-btn" onClick={()=>addItem(type)}>+ ADD {type.toUpperCase()} LINE</button>
                  <div style={{borderTop:`1px solid ${P.border}`,marginTop:10,paddingTop:10,display:"flex",justifyContent:"space-between"}}>
                    <span style={{fontSize:11,color:P.muted}}>TOTAL</span>
                    <span style={{fontWeight:500,fontSize:14,color}}>{fmt(items.reduce((s,i)=>(s+(viewBudget[i.name]||0)),0))}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── ACTUALS TAB ── */}
        {activeTab==="Actuals" && (
          <div>
            <div style={s.card}>
              <div style={{...s.sectionTitle,color:P.accent}}>Manual Actual Entry — {showYTD?`YTD (Jan–${MONTHS[selectedMonth]})`:MONTHS[selectedMonth]}</div>
              <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10}}>
                {allCats.map(cat=>(
                  <div key={cat} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
                    <span style={{fontSize:12,color:allIncomeCats.includes(cat)?P.green:P.muted,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cat}</span>
                    <div style={s.amtBox}>
                      <span style={{color:P.muted,fontSize:11}}>R</span>
                      <input type="text" inputMode="decimal" value={showYTD?ytdMonths.reduce((s,m)=>(s+(getActuals(m)[cat]||0)),0):getActuals(selectedMonth)[cat]||""}
                        onChange={e=>{ if(!showYTD){ const v=e.target.value.replace(/[^0-9.]/g,""); setActualVal(selectedMonth,cat,v); }}}
                        readOnly={showYTD}
                        style={{...s.input,width:isMobile?80:100,fontSize:13,opacity:showYTD?0.5:1}}/>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Transactions list */}
            {transactions.filter(t=>showYTD?t.month<=selectedMonth:t.month===selectedMonth).length>0 && (
              <div style={{...s.card,marginTop:16}}>
                <div style={{...s.sectionTitle,color:P.accent}}>
                  Imported Transactions ({transactions.filter(t=>showYTD?t.month<=selectedMonth:t.month===selectedMonth).length})
                </div>
                <div style={{maxHeight:280,overflowY:"auto"}}>
                  {transactions.filter(t=>showYTD?t.month<=selectedMonth:t.month===selectedMonth).slice(0,100).map((t,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${P.border}`}}>
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

        {/* ── VARIANCE TAB ── */}
        {activeTab==="Variance" && (
          <div style={{display:"flex",flexDirection:"column",gap:20}}>
            {/* Bar Chart */}
            <div style={s.card}>
              <div style={{...s.sectionTitle,color:P.accent}}>Budget vs Actual</div>
              <ResponsiveContainer width="100%" height={isMobile?220:280}>
                <BarChart data={varData} margin={{top:0,right:0,left:isMobile?0:10,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={P.border}/>
                  <XAxis dataKey="name" tick={{fill:P.muted,fontSize:isMobile?8:10}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:P.muted,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R${(v/1000).toFixed(0)}k`} width={isMobile?40:50}/>
                  <Tooltip contentStyle={{background:P.card,border:`1px solid ${P.border}`,borderRadius:8,fontSize:12}} formatter={v=>fmt(v)}/>
                  <Legend wrapperStyle={{fontSize:11}}/>
                  <Bar dataKey="Budgeted" fill={P.accent} radius={[4,4,0,0]}/>
                  <Bar dataKey="Actual" radius={[4,4,0,0]}>
                    {varData.map((entry,i)=>{
                      const over=entry.variance>0;
                      const color=entry.isIncome?(over?P.green:P.yellow):(over?P.red:P.green);
                      return <Cell key={i} fill={color}/>;
                    })}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Pie Charts */}
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:20}}>
              {[{title:"Income Breakdown",data:incomePieData},{title:"Expense Breakdown",data:expensePieData}].map(({title,data})=>(
                <div key={title} style={s.card}>
                  <div style={{...s.sectionTitle,color:P.accent}}>{title} — Budgeted</div>
                  <ResponsiveContainer width="100%" height={200}>
                    <PieChart>
                      <Pie data={data} dataKey="budgeted" nameKey="name" cx="50%" cy="50%" outerRadius={70} label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={false} fontSize={10}>
                        {data.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                      </Pie>
                      <Tooltip formatter={v=>fmt(v)}/>
                    </PieChart>
                  </ResponsiveContainer>
                  <div style={{...s.sectionTitle,color:P.accent,marginTop:16}}>{title} — Actual</div>
                  <ResponsiveContainer width="100%" height={200}>
                    <PieChart>
                      <Pie data={data} dataKey="actual" nameKey="name" cx="50%" cy="50%" outerRadius={70} label={({name,percent})=>percent>0.05?`${name} ${(percent*100).toFixed(0)}%`:""} labelLine={false} fontSize={10}>
                        {data.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                      </Pie>
                      <Tooltip formatter={v=>fmt(v)}/>
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              ))}
            </div>

            {/* Variance Table */}
            <div style={s.card}>
              <div style={{...s.sectionTitle,color:P.accent}}>Variance Detail</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:isMobile?11:12,minWidth:isMobile?500:"unset"}}>
                  <thead>
                    <tr style={{color:P.muted,fontSize:10,textTransform:"uppercase",letterSpacing:1}}>
                      {["Category","Type","Budgeted","Actual","Variance","Status"].map(h=>(
                        <th key={h} style={{textAlign:["Category","Type","Status"].includes(h)?"left":"right",padding:"8px 8px",borderBottom:`1px solid ${P.border}`}}>{h}</th>
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
                          <td style={{padding:"9px 8px",textAlign:"right",color:P.muted}}>{fmtShort(row.Budgeted)}</td>
                          <td style={{padding:"9px 8px",textAlign:"right"}}>{fmtShort(row.Actual)}</td>
                          <td style={{padding:"9px 8px",textAlign:"right",color:good?P.green:P.red}}>{row.variance>0?"+":""}{fmtShort(row.variance)}</td>
                          <td style={{padding:"9px 8px"}}>
                            <span style={s.badge(good)}>{row.Actual===0?"No Data":good?`+${Math.abs(pct)}%`:`-${Math.abs(pct)}%`}</span>
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

        {/* ── IMPORT TAB ── */}
        {activeTab==="Import" && (
          <div style={{display:"flex",flexDirection:"column",gap:16}}>
            <div style={s.card}>
              <div style={{...s.sectionTitle,color:P.accent}}>Import Bank Statement</div>
              <div style={{marginBottom:14}}>
                <div style={s.label}>Select month for this statement</div>
                <select value={selectedMonth} onChange={e=>{ setSelectedMonth(Number(e.target.value)); }}
                  style={{background:P.bg,border:`1px solid ${P.border}`,borderRadius:8,color:P.text,padding:"8px 14px",fontSize:13,fontFamily:"inherit",width:"100%"}}>
                  {MONTHS.map((m,i)=><option key={i} value={i}>{m} {new Date().getFullYear()}</option>)}
                </select>
              </div>
              <div className="upload-zone" onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()}
                onDrop={e=>{e.preventDefault();handleFile(e.dataTransfer.files[0],selectedMonth);}}>
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf" style={{display:"none"}} onChange={e=>handleFile(e.target.files[0],selectedMonth)}/>
                <div style={{fontSize:32,marginBottom:10}}>📂</div>
                <div style={{fontSize:13,marginBottom:6}}>Drop your bank statement here or click to browse</div>
                <div style={{fontSize:11,color:P.muted,lineHeight:1.6}}>
                  Supports: PDF, Excel (.xlsx), CSV<br/>
                  Banks: FNB · Standard Bank · ABSA · Nedbank · Capitec · Investec · TymeBank · African Bank
                </div>
              </div>
              {uploadStatus && (
                <div style={{marginTop:12,padding:"10px 14px",background:P.bg,borderRadius:8,fontSize:12,color:P.muted,border:`1px solid ${P.border}`}}>{uploadStatus}</div>
              )}
            </div>

            <div style={s.card}>
              <div style={{...s.sectionTitle,color:P.accent}}>Import Tips</div>
              <div style={{fontSize:12,color:P.muted,lineHeight:1.9}}>
                <div>📌 Always export directly from your bank's online portal — not scanned copies</div>
                <div>📌 FNB: Login → Accounts → Download Statement → Excel/CSV</div>
                <div>📌 Standard Bank: Online Banking → Statement → Download → Excel</div>
                <div>📌 ABSA: Online Banking → Statement → Export → CSV</div>
                <div>📌 Nedbank: Money app → Statements → Download</div>
                <div>📌 Capitec: Internet Banking → Transactions → Export</div>
                <div>📌 Transactions are auto-categorised — you can adjust in the Actuals tab</div>
                <div>📌 Re-importing the same month replaces previous import data for that month</div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}