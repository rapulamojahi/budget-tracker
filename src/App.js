import { useState, useRef } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell } from "recharts";
import * as XLSX from "xlsx";

const PALETTE = {
  bg: "#0f0f13", surface: "#16161d", card: "#1c1c27", border: "#2a2a3d",
  accent: "#6c63ff", accentLight: "#8b85ff", green: "#22c55e",
  red: "#ef4444", yellow: "#f59e0b", text: "#e8e8f0", muted: "#7b7b9a",
};

const fmt = (n) => `R ${Number(n || 0).toLocaleString("en-ZA", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const tabs = ["Budget", "Actuals", "Variance"];

const DEFAULT_INCOME = [
  { id: 1, name: "Salary" },
  { id: 2, name: "Freelance" },
  { id: 3, name: "Business" },
];
const DEFAULT_EXPENSE = [
  { id: 4, name: "Rent" },
  { id: 5, name: "Groceries" },
  { id: 6, name: "Transport" },
  { id: 7, name: "Utilities" },
  { id: 8, name: "Entertainment" },
  { id: 9, name: "Medical" },
  { id: 10, name: "Other" },
];

let nextId = 11;

export default function BudgetTracker() {
  const [activeTab, setActiveTab] = useState("Budget");
  const [incomeItems, setIncomeItems] = useState(DEFAULT_INCOME);
  const [expenseItems, setExpenseItems] = useState(DEFAULT_EXPENSE);
  const [budget, setBudget] = useState({ Salary: 25000, Freelance: 5000, Business: 3000, Rent: 8000, Groceries: 3500, Transport: 1500, Utilities: 900, Entertainment: 1200, Medical: 500, Other: 800 });
  const [actuals, setActuals] = useState({});
  const [transactions, setTransactions] = useState([]);
  const [uploadStatus, setUploadStatus] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editingVal, setEditingVal] = useState("");
  const fileRef = useRef();

  const allIncomeCats = incomeItems.map((i) => i.name);
  const allExpenseCats = expenseItems.map((i) => i.name);
  const allCats = [...allIncomeCats, ...allExpenseCats];

  const budgetIncome = allIncomeCats.reduce((s, c) => s + (budget[c] || 0), 0);
  const budgetExpense = allExpenseCats.reduce((s, c) => s + (budget[c] || 0), 0);
  const actualIncome = allIncomeCats.reduce((s, c) => s + (actuals[c] || 0), 0);
  const actualExpense = allExpenseCats.reduce((s, c) => s + (actuals[c] || 0), 0);

  const addItem = (type) => {
    const name = `New ${type === "income" ? "Income" : "Expense"} ${nextId}`;
    const item = { id: nextId++, name };
    if (type === "income") setIncomeItems((p) => [...p, item]);
    else setExpenseItems((p) => [...p, item]);
    setEditingId(item.id);
    setEditingVal(name);
  };

  const deleteItem = (type, id, name) => {
    if (type === "income") setIncomeItems((p) => p.filter((i) => i.id !== id));
    else setExpenseItems((p) => p.filter((i) => i.id !== id));
    setBudget((b) => { const nb = { ...b }; delete nb[name]; return nb; });
    setActuals((a) => { const na = { ...a }; delete na[name]; return na; });
  };

  const startRename = (id, name) => { setEditingId(id); setEditingVal(name); };

  const commitRename = (type, id, oldName) => {
    const newName = editingVal.trim();
    if (!newName || newName === oldName) { setEditingId(null); return; }
    const update = (items) => items.map((i) => i.id === id ? { ...i, name: newName } : i);
    if (type === "income") setIncomeItems(update);
    else setExpenseItems(update);
    setBudget((b) => { const nb = { ...b }; nb[newName] = nb[oldName] || 0; delete nb[oldName]; return nb; });
    setActuals((a) => { const na = { ...a }; na[newName] = na[oldName] || 0; delete na[oldName]; return na; });
    setEditingId(null);
  };

  const guessCategory = (desc = "") => {
    const d = desc.toLowerCase();
    for (const cat of allCats) { if (d.includes(cat.toLowerCase())) return cat; }
    if (d.includes("salary") || d.includes("payroll")) return allIncomeCats[0] || "Other";
    if (d.includes("rent") || d.includes("lease")) return allExpenseCats.find(c => c.toLowerCase().includes("rent")) || allExpenseCats[0];
    if (d.includes("grocery") || d.includes("checkers") || d.includes("pick n pay") || d.includes("spar") || d.includes("woolworths")) return allExpenseCats.find(c => c.toLowerCase().includes("grocer")) || allExpenseCats[0];
    if (d.includes("uber") || d.includes("bolt") || d.includes("petrol") || d.includes("fuel")) return allExpenseCats.find(c => c.toLowerCase().includes("transport")) || allExpenseCats[0];
    if (d.includes("eskom") || d.includes("electricity") || d.includes("telkom") || d.includes("internet")) return allExpenseCats.find(c => c.toLowerCase().includes("util")) || allExpenseCats[0];
    return allExpenseCats[allExpenseCats.length - 1] || "Other";
  };

  const processExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const txns = [];
        rows.slice(1).forEach((row) => {
          const desc = String(row[1] || row[2] || "");
          const amount = parseFloat(row[3] || row[4] || 0);
          if (amount !== 0) txns.push({ date: row[0] || "", description: desc, amount, category: guessCategory(desc) });
        });
        setTransactions(txns);
        const newActuals = {};
        txns.forEach(({ category, amount }) => { newActuals[category] = (newActuals[category] || 0) + Math.abs(amount); });
        setActuals(newActuals);
        setUploadStatus(`✓ Loaded ${txns.length} transactions from Excel`);
      } catch { setUploadStatus("⚠ Could not parse file. Ensure columns: Date | Description | Amount"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFile = (file) => {
    if (!file) return;
    const name = file.name.toLowerCase();
    if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv")) processExcel(file);
    else if (name.endsWith(".pdf")) setUploadStatus("ℹ PDF: Please use Excel/CSV exports from FNB Online Banking instead.");
    else setUploadStatus("⚠ Please upload an Excel (.xlsx) or CSV file.");
  };

  const varData = allCats.map((cat) => ({
    name: cat, Budgeted: budget[cat] || 0, Actual: actuals[cat] || 0,
    variance: (actuals[cat] || 0) - (budget[cat] || 0),
    isIncome: allIncomeCats.includes(cat),
  }));

  const CategoryRow = ({ item, type, budgetVal, onBudgetChange }) => {
    const isEditing = editingId === item.id;
    return (
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        {isEditing ? (
          <input autoFocus value={editingVal} onChange={(e) => setEditingVal(e.target.value)}
            onBlur={() => commitRename(type, item.id, item.name)}
            onKeyDown={(e) => { if (e.key === "Enter") commitRename(type, item.id, item.name); if (e.key === "Escape") setEditingId(null); }}
            style={{ flex: 1, background: PALETTE.bg, border: `1px solid ${PALETTE.accent}`, borderRadius: 6, padding: "4px 8px", color: PALETTE.text, fontSize: 12, fontFamily: "inherit", outline: "none" }} />
        ) : (
          <span onClick={() => startRename(item.id, item.name)} title="Click to rename"
            style={{ flex: 1, fontSize: 12, color: PALETTE.muted, cursor: "pointer" }}>
            {item.name}
          </span>
        )}
        <div style={{ display: "flex", alignItems: "center", gap: 4, background: PALETTE.bg, border: `1px solid ${PALETTE.border}`, borderRadius: 8, padding: "5px 10px" }}>
          <span style={{ color: PALETTE.muted, fontSize: 11 }}>R</span>
          <input type="number" value={budgetVal || ""} onChange={(e) => onBudgetChange(item.name, e.target.value)}
            style={{ background: "none", border: "none", outline: "none", color: PALETTE.text, width: 90, fontSize: 12, textAlign: "right", fontFamily: "inherit" }} />
        </div>
        <button onClick={() => startRename(item.id, item.name)} title="Rename"
          style={{ background: "none", border: "none", cursor: "pointer", color: PALETTE.muted, fontSize: 13, padding: "2px 4px" }}>✏️</button>
        <button onClick={() => deleteItem(type, item.id, item.name)} title="Delete"
          style={{ background: "none", border: "none", cursor: "pointer", color: PALETTE.red, fontSize: 13, padding: "2px 4px" }}>🗑</button>
      </div>
    );
  };

  return (
    <div style={{ fontFamily: "'DM Mono','Courier New',monospace", background: PALETTE.bg, minHeight: "100vh", color: PALETTE.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@700;800&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 4px; } ::-webkit-scrollbar-track { background: #0f0f13; } ::-webkit-scrollbar-thumb { background: #2a2a3d; border-radius: 2px; }
        input[type=number] { -moz-appearance: textfield; } input[type=number]::-webkit-inner-spin-button { -webkit-appearance: none; }
        .tab-btn { background: none; border: none; cursor: pointer; font-family: inherit; transition: all 0.2s; }
        .tab-btn:hover { color: #fff; }
        .upload-zone { border: 2px dashed #2a2a3d; border-radius: 12px; padding: 32px; text-align: center; cursor: pointer; transition: all 0.2s; }
        .upload-zone:hover { border-color: #6c63ff; background: rgba(108,99,255,0.05); }
        .card { background: #1c1c27; border: 1px solid #2a2a3d; border-radius: 12px; padding: 20px; }
        .add-btn { background: rgba(108,99,255,0.1); border: 1px dashed #6c63ff; border-radius: 8px; color: #8b85ff; cursor: pointer; font-family: inherit; font-size: 11px; padding: 7px; width: 100%; margin-top: 10px; transition: all 0.2s; letter-spacing: 1px; }
        .add-btn:hover { background: rgba(108,99,255,0.2); }
        .badge-green { background: rgba(34,197,94,0.15); color: #22c55e; border-radius: 4px; padding: 2px 8px; font-size: 11px; }
        .badge-red { background: rgba(239,68,68,0.15); color: #ef4444; border-radius: 4px; padding: 2px 8px; font-size: 11px; }
        .badge-yellow { background: rgba(245,158,11,0.15); color: #f59e0b; border-radius: 4px; padding: 2px 8px; font-size: 11px; }
      `}</style>

      {/* Header */}
      <div style={{ background: PALETTE.surface, borderBottom: `1px solid ${PALETTE.border}`, padding: "20px 28px", display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ width: 36, height: 36, background: PALETTE.accent, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>₿</div>
        <div>
          <div style={{ fontFamily: "'Syne',sans-serif", fontSize: 18, fontWeight: 800, letterSpacing: 1 }}>BUDGET TRACKER</div>
          <div style={{ fontSize: 11, color: PALETTE.muted }}>Actual vs Forecasted — South Africa</div>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          {tabs.map((t) => (
            <button key={t} className="tab-btn" onClick={() => setActiveTab(t)}
              style={{ padding: "8px 18px", borderRadius: 8, fontSize: 12, fontWeight: 500, color: activeTab === t ? "#fff" : PALETTE.muted, background: activeTab === t ? PALETTE.accent : "transparent", border: activeTab === t ? "none" : `1px solid ${PALETTE.border}` }}>
              {t}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: "24px 28px", maxWidth: 1100, margin: "0 auto" }}>

        {/* Summary Cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 16, marginBottom: 16 }}>
          {[
            { label: "Budget Income", val: fmt(budgetIncome), color: PALETTE.green },
            { label: "Budget Expenses", val: fmt(budgetExpense), color: PALETTE.red },
            { label: "Actual Income", val: fmt(actualIncome), color: PALETTE.green },
            { label: "Actual Expenses", val: fmt(actualExpense), color: PALETTE.red },
          ].map(({ label, val, color }) => (
            <div key={label} className="card">
              <div style={{ fontSize: 10, color: PALETTE.muted, marginBottom: 8, textTransform: "uppercase", letterSpacing: 1 }}>{label}</div>
              <div style={{ fontSize: 18, fontWeight: 500, color }}>{val}</div>
            </div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
          {[{ label: "Budgeted Net", val: budgetIncome - budgetExpense }, { label: "Actual Net", val: actualIncome - actualExpense }].map(({ label, val }) => (
            <div key={label} className="card" style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 11, color: PALETTE.muted, textTransform: "uppercase", letterSpacing: 1 }}>{label}</span>
              <span style={{ fontSize: 20, fontWeight: 500, color: val >= 0 ? PALETTE.green : PALETTE.red }}>{fmt(val)}</span>
            </div>
          ))}
        </div>

        {/* BUDGET TAB */}
        {activeTab === "Budget" && (
          <div>
            <div style={{ background: PALETTE.card, border: `1px solid ${PALETTE.border}`, borderRadius: 8, padding: "10px 16px", marginBottom: 16, fontSize: 11, color: PALETTE.muted }}>
              💡 <strong style={{ color: PALETTE.accentLight }}>Tip:</strong> Click any category name or ✏️ to rename it. Use 🗑 to delete, and the button below to add new lines.
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
              <div className="card">
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 16, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.green }}>Income Forecast</div>
                {incomeItems.map((item) => (
                  <CategoryRow key={item.id} item={item} type="income" budgetVal={budget[item.name]}
                    onBudgetChange={(name, val) => setBudget((b) => ({ ...b, [name]: parseFloat(val) || 0 }))} />
                ))}
                <button className="add-btn" onClick={() => addItem("income")}>+ ADD INCOME LINE</button>
                <div style={{ borderTop: `1px solid ${PALETTE.border}`, marginTop: 12, paddingTop: 12, display: "flex", justifyContent: "space-between" }}>
                  <span style={{ fontSize: 11, color: PALETTE.muted }}>TOTAL</span>
                  <span style={{ fontWeight: 500, fontSize: 14, color: PALETTE.green }}>{fmt(budgetIncome)}</span>
                </div>
              </div>
              <div className="card">
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 16, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.red }}>Expense Forecast</div>
                {expenseItems.map((item) => (
                  <CategoryRow key={item.id} item={item} type="expense" budgetVal={budget[item.name]}
                    onBudgetChange={(name, val) => setBudget((b) => ({ ...b, [name]: parseFloat(val) || 0 }))} />
                ))}
                <button className="add-btn" onClick={() => addItem("expense")}>+ ADD EXPENSE LINE</button>
                <div style={{ borderTop: `1px solid ${PALETTE.border}`, marginTop: 12, paddingTop: 12, display: "flex", justifyContent: "space-between" }}>
                  <span style={{ fontSize: 11, color: PALETTE.muted }}>TOTAL</span>
                  <span style={{ fontWeight: 500, fontSize: 14, color: PALETTE.red }}>{fmt(budgetExpense)}</span>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ACTUALS TAB */}
        {activeTab === "Actuals" && (
          <div>
            <div className="card" style={{ marginBottom: 20 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 16, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.accentLight }}>Upload Bank Statement</div>
              <div className="upload-zone" onClick={() => fileRef.current.click()} onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }}>
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
                <div style={{ fontSize: 28, marginBottom: 8 }}>📂</div>
                <div style={{ fontSize: 13, marginBottom: 4 }}>Drop your FNB statement here or click to browse</div>
                <div style={{ fontSize: 11, color: PALETTE.muted }}>Supports Excel (.xlsx), CSV — export from FNB Online Banking</div>
              </div>
              {uploadStatus && (
                <div style={{ marginTop: 12, padding: "10px 14px", background: PALETTE.bg, borderRadius: 8, fontSize: 12, color: PALETTE.muted, border: `1px solid ${PALETTE.border}` }}>{uploadStatus}</div>
              )}
            </div>

            <div className="card" style={{ marginBottom: 20 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 16, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.accentLight }}>Manual Actual Entry</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                {allCats.map((cat) => (
                  <div key={cat} style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <span style={{ fontSize: 12, color: allIncomeCats.includes(cat) ? PALETTE.green : PALETTE.muted }}>{cat}</span>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, background: PALETTE.bg, border: `1px solid ${PALETTE.border}`, borderRadius: 8, padding: "6px 12px" }}>
                      <span style={{ color: PALETTE.muted, fontSize: 12 }}>R</span>
                      <input type="number" value={actuals[cat] || ""} onChange={(e) => setActuals((a) => ({ ...a, [cat]: parseFloat(e.target.value) || 0 }))}
                        style={{ background: "none", border: "none", outline: "none", color: PALETTE.text, width: 90, fontSize: 13, textAlign: "right", fontFamily: "inherit" }} />
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {transactions.length > 0 && (
              <div className="card">
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 16, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.accentLight }}>
                  Imported Transactions ({transactions.length})
                </div>
                <div style={{ maxHeight: 260, overflowY: "auto" }}>
                  {transactions.slice(0, 50).map((t, i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${PALETTE.border}` }}>
                      <div>
                        <div style={{ fontSize: 12 }}>{t.description || "—"}</div>
                        <div style={{ fontSize: 10, color: PALETTE.muted }}>{t.date} · {t.category}</div>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 500, color: t.amount >= 0 ? PALETTE.green : PALETTE.red }}>{fmt(Math.abs(t.amount))}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* VARIANCE TAB */}
        {activeTab === "Variance" && (
          <div>
            <div className="card" style={{ marginBottom: 24 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 20, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.accentLight }}>Budget vs Actual by Category</div>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={varData} margin={{ top: 0, right: 0, left: 10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={PALETTE.border} />
                  <XAxis dataKey="name" tick={{ fill: PALETTE.muted, fontSize: 9 }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fill: PALETTE.muted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={(v) => `R${(v / 1000).toFixed(0)}k`} />
                  <Tooltip contentStyle={{ background: PALETTE.card, border: `1px solid ${PALETTE.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v) => fmt(v)} />
                  <Legend wrapperStyle={{ fontSize: 11, color: PALETTE.muted }} />
                  <Bar dataKey="Budgeted" fill={PALETTE.accent} radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Actual" radius={[4, 4, 0, 0]}>
                    {varData.map((entry, i) => {
                      const over = entry.variance > 0;
                      const color = entry.isIncome ? (over ? PALETTE.green : PALETTE.yellow) : (over ? PALETTE.red : PALETTE.green);
                      return <Cell key={i} fill={color} />;
                    })}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="card">
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13, marginBottom: 16, letterSpacing: 1, textTransform: "uppercase", color: PALETTE.accentLight }}>Variance Detail</div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ color: PALETTE.muted, fontSize: 10, textTransform: "uppercase", letterSpacing: 1 }}>
                    {["Category", "Type", "Budgeted", "Actual", "Variance", "Status"].map((h) => (
                      <th key={h} style={{ textAlign: ["Category", "Type", "Status"].includes(h) ? "left" : "right", padding: "8px 10px", borderBottom: `1px solid ${PALETTE.border}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {varData.map((row) => {
                    const good = row.isIncome ? row.variance >= 0 : row.variance <= 0;
                    const pct = row.Budgeted ? ((row.variance / row.Budgeted) * 100).toFixed(1) : "—";
                    return (
                      <tr key={row.name} style={{ borderBottom: `1px solid ${PALETTE.border}` }}>
                        <td style={{ padding: "10px 10px" }}>{row.name}</td>
                        <td style={{ padding: "10px 10px", color: row.isIncome ? PALETTE.green : PALETTE.red }}>{row.isIncome ? "Income" : "Expense"}</td>
                        <td style={{ padding: "10px 10px", textAlign: "right", color: PALETTE.muted }}>{fmt(row.Budgeted)}</td>
                        <td style={{ padding: "10px 10px", textAlign: "right" }}>{fmt(row.Actual)}</td>
                        <td style={{ padding: "10px 10px", textAlign: "right", color: good ? PALETTE.green : PALETTE.red }}>{row.variance > 0 ? "+" : ""}{fmt(row.variance)}</td>
                        <td style={{ padding: "10px 10px" }}>
                          <span className={good ? "badge-green" : Math.abs(parseFloat(pct)) < 10 ? "badge-yellow" : "badge-red"}>
                            {row.Actual === 0 ? "No Data" : good ? `+${Math.abs(pct)}%` : `-${Math.abs(pct)}%`}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
