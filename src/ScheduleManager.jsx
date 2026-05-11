import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { Upload, X, Download, User, Calendar, BarChart2, ChevronLeft, ChevronRight, Check, AlertTriangle, FileSpreadsheet, MapPin } from "lucide-react";

const DAYS_ID = { Minggu:0,Senin:1,Selasa:2,Rabu:3,Kamis:4,Jumat:5,Sabtu:6 };
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const CLASH_CFG = {
  hard:   { label:"Hard Clash",     shortLabel:"Hard",    bg:"var(--color-background-danger)",  text:"var(--color-text-danger)",  border:"1px solid var(--color-border-danger)" },
  city:   { label:"City Clash",     shortLabel:"City",    bg:"var(--color-background-warning)", text:"var(--color-text-warning)", border:"1px solid var(--color-border-warning)" },
  travel: { label:"Travel Warning", shortLabel:"Travel",  bg:"var(--color-background-info)",    text:"var(--color-text-info)",    border:"1px solid var(--color-border-info)" },
};
const SESSION_STYLE = {
  "Session":    { bg:"var(--color-background-secondary)", text:"var(--color-text-secondary)" },
  "Mid Exam":   { bg:"var(--color-background-warning)",   text:"var(--color-text-warning)" },
  "Final Exam": { bg:"var(--color-background-danger)",    text:"var(--color-text-danger)" },
};
const LOC_STYLE = {
  Jakarta: { bg:"#dbeafe", text:"#1e40af" },
  Bandung: { bg:"#dcfce7", text:"#166534" },
};

function extractLecturers(raw) {
  if (!raw) return [];
  const s = String(raw).trim();
  // Returns [{name, sks}] — sks is the weight in parentheses e.g. "(1.5)"
  const matches = [...s.matchAll(/([A-Za-z][^(]*?)\s*\(\s*([\d.]+)\s*\)/g)];
  if (matches.length > 0) {
    return matches.map(m => ({ name: m[1].replace(/\s*[-–]\s*$/,"").trim(), sks: parseFloat(m[2])||1 })).filter(l=>l.name);
  }
  const cleaned = s.replace(/\s*\([^)]*\)\s*/g,"").replace(/\s*[-–]\s*$/,"").trim();
  return cleaned ? [{ name: cleaned, sks: 1 }] : [];
}

function getTimeForDate(jam, hari, date, sheetType) {
  if (sheetType === "ENMARK") return "18.10 - 22.20";
  const j = jam ? String(jam).trim() : "";
  if (sheetType === "Executive") {
    if (j.includes(",")) { const p = j.split(",").map(x=>x.trim()); return date.getDay()===6?p[0]:(p[1]||p[0]); }
    return j;
  }
  if (/\bdan\b/i.test(j)) {
    const times = j.split(/\s+dan\s+/i).map(t=>t.trim());
    const days = (hari||"").replace(/\s*\([^)]*\)/g,"").split(/[,\s\/]+/).map(d=>d.trim()).filter(d=>d in DAYS_ID);
    const idx = days.findIndex(d=>DAYS_ID[d]===date.getDay());
    if (idx>=0 && times[idx]) return times[idx];
  }
  return j.replace(/\s*\([^)]*\)/g,"").trim();
}

function parseTimeRange(s) {
  if (!s) return null;
  const nums = [...String(s).matchAll(/(\d{1,2})[.:](\d{2})/g)];
  if (nums.length < 2) return null;
  const toMin = m => parseInt(m[1])*60+parseInt(m[2]);
  return { start: toMin(nums[0]), end: toMin(nums[nums.length-1]) };
}
function timesOverlap(a,b) { const r1=parseTimeRange(a),r2=parseTimeRange(b); return r1&&r2&&r1.start<r2.end&&r2.start<r1.end; }
function getSessionType(h) { const s=String(h).toLowerCase(); return s.includes("mid")?"Mid Exam":s.includes("final")?"Final Exam":"Session"; }
function fmtDate(d) { return d instanceof Date?d.toLocaleDateString("id-ID",{weekday:"short",day:"numeric",month:"short",year:"numeric"}):""; }
function dk(d) { return d instanceof Date?d.toISOString().split("T")[0]:""; }
function isDate(v) { return v instanceof Date&&!isNaN(v); }

function processSheet(rawRows, sheetType) {
  if (!rawRows||rawRows.length<2) return [];
  const hdrs = rawRows[0]||[];
  const find = (...names) => { for (const n of names) { const i=hdrs.findIndex(h=>h&&String(h).toLowerCase().trim()===n.toLowerCase()); if (i>=0) return i; } return -1; };
  const ci = { prodi:find("prodi"), loc:find("location"), kode:find("kode"), nama:find("nama"), kelas:find("kelas"), team:find("team teaching"), jam:find("jam"), hari:find("hari"), ruang:find("ruang","ruangan") };
  const result=[], seen=new Set();
  for (let i=1;i<rawRows.length;i++) {
    const row=rawRows[i];
    if (!row||row.every(v=>v==null||v==="")) continue;
    const kode = row[ci.kode]?String(row[ci.kode]).trim():"";
    if (!kode||kode==="MMXXXX") continue;
    const lecturers = extractLecturers(row[ci.team]); // [{name, sks}]
    const list = lecturers.length>0?lecturers:[{name:"",sks:1}];
    const shared = {
      program: row[ci.prodi]?String(row[ci.prodi]).trim():(sheetType==="ENMARK"?"ENMARK":""),
      location: row[ci.loc]?String(row[ci.loc]).trim():"",
      class: row[ci.kelas]?String(row[ci.kelas]).trim():"",
      course: `${kode} ${row[ci.nama]?String(row[ci.nama]).trim():""}`.trim(),
      jam: row[ci.jam]?String(row[ci.jam]).trim():"",
      hari: row[ci.hari]?String(row[ci.hari]).trim():"",
      room: row[ci.ruang]?String(row[ci.ruang]).trim():"",
      sourceSheet: sheetType,
    };
    for (const {name: lecturer, sks: lecturerSKS} of list) {
      for (let j=0;j<hdrs.length;j++) {
        const val=row[j];
        if (!isDate(val)) continue;
        const time = getTimeForDate(shared.jam,shared.hari,val,sheetType);
        const dedup = `${lecturer}|${shared.course}|${shared.class}|${dk(val)}|${time}`;
        if (seen.has(dedup)) continue;
        seen.add(dedup);
        result.push({ id:`${sheetType}-${i}-${j}-${encodeURIComponent(lecturer)}`, lecturer, lecturerSKS, _rowIndex:i, hasLecturer:!!lecturer, ...shared, date:val, time, sessionType:getSessionType(hdrs[j]), hasRoom:!!shared.room });
      }
    }
  }
  return result;
}

// Divide team-teaching session dates chronologically by SKS weight.
// Rules:
//   - Regular sessions: proportional by SKS, first lecturer gets earliest dates
//   - 2 lecturers: all mid exams go to first lecturer, all finals to second
//   - 3+ lecturers: mid AND final exams distributed proportionally
function redistributeTeamTeachingDates(allRows) {
  const groups = {};
  for (const row of allRows) {
    const key = `${row.course}||${row.class}`;
    (groups[key] = groups[key]||[]).push(row);
  }
  const result = [];
  for (const rows of Object.values(groups)) {
    const namedLecturers = [...new Set(rows.filter(r=>r.lecturer).map(r=>r.lecturer))];
    if (namedLecturers.length <= 1) { result.push(...rows); continue; }

    // Unique dates per session type, sorted ascending
    const uniqueSorted = (type) => {
      const seen=new Set(), out=[];
      for (const r of rows) {
        if (r.sessionType!==type) continue;
        const k=dk(r.date);
        if (!seen.has(k)) { seen.add(k); out.push(r.date); }
      }
      return out.sort((a,b)=>a-b);
    };
    const regularDates = uniqueSorted("Session");
    const midDates     = uniqueSorted("Mid Exam");
    const finalDates   = uniqueSorted("Final Exam");

    // Build lecturer list: ordered by first spreadsheet row appearance, SKS from their rows
    const lecMap = {};
    for (const r of rows) {
      if (!r.lecturer) continue;
      if (!lecMap[r.lecturer]) lecMap[r.lecturer] = { name:r.lecturer, sks:r.lecturerSKS||1, order:r._rowIndex };
      else if (r._rowIndex < lecMap[r.lecturer].order) lecMap[r.lecturer].order = r._rowIndex;
    }
    const lecInfo = Object.values(lecMap).sort((a,b)=>a.order-b.order);
    const totalSKS = lecInfo.reduce((s,l)=>s+l.sks, 0)||1;

    // Proportional splitter: returns {dateKey → lecturerName}
    const splitDates = (dates) => {
      const assign={};
      let offset=0;
      for (let i=0;i<lecInfo.length;i++) {
        const count = i<lecInfo.length-1 ? Math.round(dates.length*(lecInfo[i].sks/totalSKS)) : dates.length-offset;
        for (const d of dates.slice(offset,offset+count)) assign[dk(d)]=lecInfo[i].name;
        offset+=count;
      }
      return assign;
    };

    const regularAssign = splitDates(regularDates);
    const midAssign={}, finalAssign={};
    if (lecInfo.length===2) {
      midDates.forEach(d   => midAssign[dk(d)]   = lecInfo[0].name);
      finalDates.forEach(d => finalAssign[dk(d)] = lecInfo[1].name);
    } else {
      Object.assign(midAssign,   splitDates(midDates));
      Object.assign(finalAssign, splitDates(finalDates));
    }

    for (const row of rows) {
      if (!row.lecturer) { result.push(row); continue; }
      const d=dk(row.date);
      const owner = row.sessionType==="Session"    ? regularAssign[d]
                  : row.sessionType==="Mid Exam"   ? midAssign[d]
                  : row.sessionType==="Final Exam" ? finalAssign[d] : null;
      if (owner===row.lecturer) result.push(row);
    }
  }
  return result;
}

function detectClashes(rows) {
  const clashes=[], seen=new Set();
  const byLD={};
  for (const r of rows) { if (!r.lecturer) continue; const k=`${r.lecturer}||${dk(r.date)}`; (byLD[k]=byLD[k]||[]).push(r); }
  for (const group of Object.values(byLD)) {
    for (let i=0;i<group.length;i++) for (let j=i+1;j<group.length;j++) {
      const a=group[i],b=group[j];
      if (a.course===b.course&&a.class===b.class) continue;
      const pk=[a.id,b.id].sort().join("||");
      if (seen.has(pk)) continue; seen.add(pk);
      if (timesOverlap(a.time,b.time)) clashes.push({id:`hard-${pk}`,type:"hard",lecturer:a.lecturer,date:a.date,rows:[a,b]});
      else if (a.location&&b.location&&a.location!==b.location) clashes.push({id:`city-${pk}`,type:"city",lecturer:a.lecturer,date:a.date,rows:[a,b]});
    }
  }
  const byL={};
  for (const r of rows) { if (!r.lecturer||!r.location) continue; if (!byL[r.lecturer]) byL[r.lecturer]={}; if (!byL[r.lecturer][dk(r.date)]) byL[r.lecturer][dk(r.date)]={date:r.date,locs:new Set()}; byL[r.lecturer][dk(r.date)].locs.add(r.location); }
  for (const [lec,dayMap] of Object.entries(byL)) {
    const days=Object.values(dayMap).sort((a,b)=>a.date-b.date);
    for (let i=0;i<days.length-1;i++) {
      const d1=days[i],d2=days[i+1];
      if (Math.round((d2.date-d1.date)/86400000)!==1) continue;
      const l1=[...d1.locs],l2=[...d2.locs];
      if (l1.some(l=>!l2.includes(l))||l2.some(l=>!l1.includes(l))) {
        const tk=`travel-${lec}-${dk(d1.date)}-${dk(d2.date)}`;
        if (!seen.has(tk)) { seen.add(tk); const r1=rows.find(r=>r.lecturer===lec&&dk(r.date)===dk(d1.date)),r2=rows.find(r=>r.lecturer===lec&&dk(r.date)===dk(d2.date)); clashes.push({id:tk,type:"travel",lecturer:lec,date:d1.date,rows:[r1,r2]}); }
      }
    }
  }
  return clashes;
}

function getWarnings(rows) {
  const noLec=rows.filter(r=>!r.hasLecturer&&r.sessionType==="Session").length;
  const noRoom=rows.filter(r=>!r.hasRoom).length;
  const noLoc=rows.filter(r=>!r.location).length;
  return [noLec&&{id:"no_lec",msg:`${noLec} sessions have no lecturer assigned`},noRoom&&{id:"no_room",msg:`${noRoom} sessions have no room assigned`},noLoc&&{id:"no_loc",msg:`${noLoc} sessions have no location assigned`}].filter(Boolean);
}

const S = {
  card: { background:"var(--color-background-primary)", borderRadius:"var(--border-radius-lg)", border:"0.5px solid var(--color-border-tertiary)", overflow:"hidden" },
  th: { padding:"10px 14px", textAlign:"left", fontWeight:500, fontSize:11, color:"var(--color-text-secondary)", background:"var(--color-background-secondary)", borderBottom:"0.5px solid var(--color-border-tertiary)", whiteSpace:"nowrap", textTransform:"uppercase", letterSpacing:"0.06em" },
  td: { padding:"9px 14px", fontSize:13, borderBottom:"0.5px solid var(--color-border-tertiary)", color:"var(--color-text-primary)", verticalAlign:"middle" },
  input: { padding:"7px 12px", borderRadius:"var(--border-radius-md)", border:"0.5px solid var(--color-border-secondary)", fontSize:13, background:"var(--color-background-primary)", color:"var(--color-text-primary)", outline:"none", width:"100%", boxSizing:"border-box" },
  select: { padding:"7px 10px", borderRadius:"var(--border-radius-md)", border:"0.5px solid var(--color-border-secondary)", fontSize:13, background:"var(--color-background-primary)", color:"var(--color-text-primary)", cursor:"pointer" },
  btn: { padding:"7px 14px", borderRadius:"var(--border-radius-md)", border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)", cursor:"pointer", fontSize:13, fontWeight:500, color:"var(--color-text-primary)", display:"inline-flex", alignItems:"center", gap:6, whiteSpace:"nowrap" },
  btnPrimary: { padding:"7px 14px", borderRadius:"var(--border-radius-md)", border:"none", background:"#1d4ed8", cursor:"pointer", fontSize:13, fontWeight:500, color:"white", display:"inline-flex", alignItems:"center", gap:6 },
  link: { background:"none", border:"none", cursor:"pointer", color:"#2563eb", fontWeight:500, fontSize:13, padding:0, textDecoration:"none" },
};

function Badge({ text, bg, color }) {
  return <span style={{ padding:"2px 8px", borderRadius:20, fontSize:11, fontWeight:500, background:bg, color, whiteSpace:"nowrap" }}>{text}</span>;
}

function LocBadge({ loc }) {
  const s = LOC_STYLE[loc];
  return s ? <Badge text={loc} bg={s.bg} color={s.text} /> : <span style={{color:"var(--color-text-secondary)",fontSize:12}}>—</span>;
}

const FILTER_LABELS = { lecturer:"Lecturer", class:"Class", program:"Program", course:"Course", sheet:"Sheet" };

export default function ScheduleManager() {
  const [rows, setRows] = useState([]);
  const [clashes, setClashes] = useState([]);
  const [warnings, setWarnings] = useState([]);
  const [acked, setAcked] = useState({});
  const [notes, setNotes] = useState({});
  const [dismissed, setDismissed] = useState(new Set());
  const [view, setView] = useState("mcp");
  const [filters, setFilters] = useState({});       // {lecturer, class, program, course, sheet}
  const [sortCfg, setSortCfg] = useState({ col:"date", dir:"asc" });
  const [statsTab, setStatsTab] = useState("lecturer");
  const [search, setSearch] = useState("");
  const [monthF, setMonthF] = useState("all");
  const [locF, setLocF] = useState("all");
  const [clashF, setClashF] = useState("all");
  const [calDate, setCalDate] = useState(new Date(2026,5,1));
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");

  // ── Shared XLSX processing ─────────────────────────────────────────────────
  // Call this with an ArrayBuffer from either a file upload or a fetch().
  const processBuffer = useCallback((buffer, name) => {
    const wb = XLSX.read(new Uint8Array(buffer), { type:"array", cellDates:true });
    const all = [];
    for (const [sName,sType] of [["Regular","Regular"],["Executive","Executive"],["ENMARK","ENMARK"]]) {
      if (wb.SheetNames.includes(sName)) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets[sName], { header:1, raw:true, cellDates:true, defval:null });
        all.push(...processSheet(data, sType));
      }
    }
    const distributed = redistributeTeamTeachingDates(all);
    setRows(distributed); setClashes(detectClashes(distributed)); setWarnings(getWarnings(distributed));
    setAcked({}); setNotes({}); setDismissed(new Set()); setFilters({}); setSearch(""); setMonthF("all"); setLocF("all"); setView("mcp");
    setFileName(name);
  }, []);

  // ── Auto-load: fetch the bundled schedule on first mount ───────────────────
  // Put your Excel file at:  public/schedule.xlsx  (CRA / Vite / Next.js)
  // Change the path below if you name it differently.
  const AUTO_LOAD_PATH = "/schedule.xlsx";

  useEffect(() => {
    setLoading(true);
    fetch(AUTO_LOAD_PATH)
      .then(res => {
        if (!res.ok) throw new Error(`${res.status} – file not found at ${AUTO_LOAD_PATH}`);
        return res.arrayBuffer();
      })
      .then(buf => {
        processBuffer(buf, AUTO_LOAD_PATH.split("/").pop());
      })
      .catch(err => {
        console.warn("Auto-load skipped:", err.message);
      })
      .finally(() => setLoading(false));
  }, [processBuffer]);

  // ── Manual upload (overrides the auto-loaded file) ─────────────────────────
  const handleFile = useCallback(e => {
    const file = e.target.files[0]; if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = evt => {
      try { processBuffer(evt.target.result, file.name); }
      catch(err) { console.error(err); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  }, [processBuffer]);

  const months = useMemo(() => [...new Set(rows.map(r=>r.date.getMonth()))].sort(), [rows]);

  // Filter helpers
  const toggleFilter = (dim, val) => setFilters(f => ({ ...f, [dim]: f[dim]===val ? null : val }));
  const clearFilter  = (dim)      => setFilters(f => { const n={...f}; delete n[dim]; return n; });
  const clearAll     = ()         => { setFilters({}); setSearch(""); setMonthF("all"); setLocF("all"); };
  const activeFilterEntries = Object.entries(filters).filter(([,v])=>v);
  const hasAnyFilter = activeFilterEntries.length>0 || search || monthF!=="all" || locF!=="all";
  const toggleSort   = (col) => setSortCfg(s => ({ col, dir: s.col===col&&s.dir==="asc" ? "desc" : "asc" }));

  const filtered = useMemo(() => {
    const arr = rows.filter(r => {
      if (filters.lecturer && r.lecturer !== filters.lecturer) return false;
      if (filters.class    && r.class    !== filters.class)    return false;
      if (filters.program  && r.program  !== filters.program)  return false;
      if (filters.course   && r.course   !== filters.course)   return false;
      if (filters.sheet    && r.sourceSheet !== filters.sheet) return false;
      if (monthF !== "all" && r.date.getMonth() !== +monthF)   return false;
      if (locF   !== "all" && r.location !== locF)             return false;
      if (search) {
        const q = search.toLowerCase();
        if (![r.lecturer, r.class, r.program, r.course, r.room].some(v=>v?.toLowerCase().includes(q))) return false;
      }
      return true;
    });
    const { col, dir } = sortCfg;
    const m = dir==="asc" ? 1 : -1;
    return arr.sort((a,b) => {
      const va = col==="date" ? (a.date?.getTime()||0) : String(a[col]||"").toLowerCase();
      const vb = col==="date" ? (b.date?.getTime()||0) : String(b[col]||"").toLowerCase();
      return va<vb ? -m : va>vb ? m : 0;
    });
  }, [rows, filters, search, monthF, locF, sortCfg]);

  const filtClashes = useMemo(() => clashes.filter(c => {
    if (clashF !== "all" && c.type !== clashF) return false;
    if (filters.lecturer && c.lecturer !== filters.lecturer) return false;
    if (filters.class   && !c.rows.some(r=>r?.class   === filters.class))   return false;
    if (filters.program && !c.rows.some(r=>r?.program === filters.program)) return false;
    return true;
  }), [clashes, clashF, filters]);

  const counts = useMemo(() => ({
    hard:   clashes.filter(c=>c.type==="hard"  &&!acked[c.id]).length,
    city:   clashes.filter(c=>c.type==="city"  &&!acked[c.id]).length,
    travel: clashes.filter(c=>c.type==="travel"&&!acked[c.id]).length,
  }), [clashes, acked]);

  const lecStats = useMemo(() => {
    const map = {};
    for (const r of rows) {
      if (!r.lecturer) continue;
      const s = map[r.lecturer] || (map[r.lecturer]={ name:r.lecturer, total:0, sessions:0, mid:0, final:0, jakarta:0, bandung:0, programs:new Set(), classes:new Set() });
      s.total++; if(r.sessionType==="Session")s.sessions++; if(r.sessionType==="Mid Exam")s.mid++; if(r.sessionType==="Final Exam")s.final++;
      if(r.location==="Jakarta")s.jakarta++; if(r.location==="Bandung")s.bandung++;
      r.program&&s.programs.add(r.program); r.class&&s.classes.add(r.class);
    }
    return Object.values(map).sort((a,b)=>b.total-a.total);
  }, [rows]);

  const classStats = useMemo(() => {
    const map = {};
    for (const r of rows) {
      if (!r.class) continue;
      const s = map[r.class] || (map[r.class]={ name:r.class, total:0, lecturers:new Set(), programs:new Set(), jakarta:0, bandung:0 });
      s.total++; r.lecturer&&s.lecturers.add(r.lecturer); r.program&&s.programs.add(r.program);
      if(r.location==="Jakarta")s.jakarta++; if(r.location==="Bandung")s.bandung++;
    }
    return Object.values(map).sort((a,b)=>b.total-a.total);
  }, [rows]);

  const programStats = useMemo(() => {
    const map = {};
    for (const r of rows) {
      if (!r.program) continue;
      const s = map[r.program] || (map[r.program]={ name:r.program, total:0, lecturers:new Set(), classes:new Set(), jakarta:0, bandung:0 });
      s.total++; r.lecturer&&s.lecturers.add(r.lecturer); r.class&&s.classes.add(r.class);
      if(r.location==="Jakarta")s.jakarta++; if(r.location==="Bandung")s.bandung++;
    }
    return Object.values(map).sort((a,b)=>b.total-a.total);
  }, [rows]);

  const calData = useMemo(() => {
    const y=calDate.getFullYear(), m=calDate.getMonth(), byDay={}, clashDays={};
    for (const r of rows) {
      if (r.date.getFullYear()!==y||r.date.getMonth()!==m) continue;
      if (filters.lecturer&&r.lecturer!==filters.lecturer) continue;
      if (filters.class   &&r.class   !==filters.class)    continue;
      if (filters.program &&r.program !==filters.program)  continue;
      (byDay[r.date.getDate()]=byDay[r.date.getDate()]||[]).push(r);
    }
    for (const c of clashes) {
      if (c.date.getFullYear()!==y||c.date.getMonth()!==m) continue;
      const d=c.date.getDate();
      if (!clashDays[d]) clashDays[d]={hard:0,city:0,travel:0};
      if (!acked[c.id]) clashDays[d][c.type]++;
    }
    return {byDay, clashDays};
  }, [rows, clashes, calDate, filters, acked]);

  // Sortable header cell
  const SortTh = ({ col, label }) => {
    const active = sortCfg.col === col;
    return (
      <th onClick={()=>toggleSort(col)} style={{ ...S.th, cursor:"pointer", userSelect:"none" }}>
        {label} <span style={{ opacity:active?1:0.25, fontSize:9 }}>{active?(sortCfg.dir==="asc"?"↑":"↓"):"↕"}</span>
      </th>
    );
  };

  const exportMCP = () => {
    const data = filtered.map((r,i) => ({ No:i+1, Lecturer:r.lecturer||"(Unassigned)", Class:r.class, Program:r.program, Course:r.course, Date:fmtDate(r.date), Time:r.time, Room:r.room, Location:r.location, "Session Type":r.sessionType, Source:r.sourceSheet }));
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "MCP Output"); XLSX.writeFile(wb,"MCP_Output.xlsx");
  };

  const exportClashes = () => {
    const data = filtClashes.map((c,i) => ({ No:i+1, Type:CLASH_CFG[c.type].label, Lecturer:c.lecturer, "Date 1":fmtDate(c.rows[0]?.date), "Entry 1":`${c.rows[0]?.course} | ${c.rows[0]?.time} | ${c.rows[0]?.room} | ${c.rows[0]?.location}`, "Date 2":fmtDate(c.rows[1]?.date), "Entry 2":`${c.rows[1]?.course} | ${c.rows[1]?.time} | ${c.rows[1]?.room} | ${c.rows[1]?.location}`, Acknowledged:acked[c.id]?"Yes":"No", Note:notes[c.id]||"" }));
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Clash Report"); XLSX.writeFile(wb,"Clash_Report.xlsx");
  };

  // ── EMPTY STATE ──
  if (!rows.length) return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:"60vh", padding:32 }}>
      <div style={{ textAlign:"center", maxWidth:440 }}>
        <div style={{ width:64, height:64, borderRadius:"var(--border-radius-lg)", background:"var(--color-background-secondary)", border:"0.5px solid var(--color-border-tertiary)", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 20px" }}>
          <FileSpreadsheet size={28} color="var(--color-text-secondary)" />
        </div>
        {loading ? (
          <>
            <h2 style={{ fontSize:20, fontWeight:500, margin:"0 0 8px", color:"var(--color-text-primary)" }}>Loading schedule…</h2>
            <p style={{ fontSize:13, color:"var(--color-text-secondary)", margin:0 }}>Fetching <code style={{fontSize:12}}>{AUTO_LOAD_PATH}</code></p>
          </>
        ) : (
          <>
            <h2 style={{ fontSize:22, fontWeight:500, margin:"0 0 8px", color:"var(--color-text-primary)" }}>Schedule Manager</h2>
            <p style={{ fontSize:14, color:"var(--color-text-secondary)", lineHeight:1.6, margin:"0 0 6px" }}>
              No schedule file found at <code style={{fontSize:12}}>{AUTO_LOAD_PATH}</code>.
            </p>
            <p style={{ fontSize:13, color:"var(--color-text-secondary)", lineHeight:1.6, margin:"0 0 28px" }}>
              Place your Excel file there, or upload one manually below.
            </p>
            <label style={{ ...S.btnPrimary, padding:"10px 22px", fontSize:14, cursor:"pointer", borderRadius:"var(--border-radius-md)" }}>
              <Upload size={16} /> Upload Excel File
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />
            </label>
          </>
        )}
      </div>
    </div>
  );

  const allClear = counts.hard+counts.city+counts.travel === 0;

  // ── MAIN APP ──
  return (
    <div style={{ fontFamily:"var(--font-sans)", color:"var(--color-text-primary)" }}>

      {/* Header */}
      <div style={{ padding:"12px 20px", borderBottom:"0.5px solid var(--color-border-tertiary)", display:"flex", alignItems:"center", gap:12, flexWrap:"wrap", background:"var(--color-background-primary)" }}>
        <div style={{ flex:1, minWidth:0 }}>
          <div style={{ fontWeight:500, fontSize:15, display:"flex", alignItems:"center", gap:8 }}>
            <FileSpreadsheet size={16} color="var(--color-text-secondary)" />
            Schedule Manager
          </div>
          <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginTop:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{fileName} · {rows.length.toLocaleString()} sessions</div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          {allClear
            ? <Badge text="✓ No active clashes" bg="var(--color-background-success)" color="var(--color-text-success)" />
            : <>
              {counts.hard>0 && <button onClick={()=>{setView("clashes");setClashF("hard")}} style={{...S.btn, background:"var(--color-background-danger)", color:"var(--color-text-danger)", border:"0.5px solid var(--color-border-danger)"}}>{counts.hard} Hard</button>}
              {counts.city>0 && <button onClick={()=>{setView("clashes");setClashF("city")}} style={{...S.btn, background:"var(--color-background-warning)", color:"var(--color-text-warning)", border:"0.5px solid var(--color-border-warning)"}}>{counts.city} City</button>}
              {counts.travel>0 && <button onClick={()=>{setView("clashes");setClashF("travel")}} style={{...S.btn, background:"var(--color-background-info)", color:"var(--color-text-info)", border:"0.5px solid var(--color-border-info)"}}>{counts.travel} Travel</button>}
            </>
          }
          <label style={{ ...S.btn, cursor:"pointer" }}>
            <Upload size={14} /> New File
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />
          </label>
        </div>
      </div>

      {/* Warnings */}
      {warnings.filter(w=>!dismissed.has(w.id)).map(w => (
        <div key={w.id} style={{ padding:"8px 20px", background:"var(--color-background-warning)", borderBottom:"0.5px solid var(--color-border-warning)", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span style={{ fontSize:13, color:"var(--color-text-warning)", display:"flex", alignItems:"center", gap:6 }}><AlertTriangle size={14} /> {w.msg}</span>
          <button onClick={()=>setDismissed(p=>new Set([...p,w.id]))} style={{ background:"none", border:"none", cursor:"pointer", color:"var(--color-text-warning)", padding:0 }}><X size={14} /></button>
        </div>
      ))}

      {/* Active filter badge strip */}
      {hasAnyFilter && (
        <div style={{ padding:"7px 20px", background:"var(--color-background-info)", borderBottom:"0.5px solid var(--color-border-info)", display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
          <span style={{ fontSize:11, color:"var(--color-text-secondary)", textTransform:"uppercase", letterSpacing:"0.05em", whiteSpace:"nowrap" }}>Filtered by</span>
          {activeFilterEntries.map(([dim, val]) => (
            <span key={dim} style={{ display:"inline-flex", alignItems:"center", gap:5, background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-secondary)", borderRadius:20, padding:"2px 8px 2px 10px", fontSize:12 }}>
              <span style={{ color:"var(--color-text-secondary)", fontSize:10, textTransform:"uppercase", letterSpacing:"0.05em" }}>{FILTER_LABELS[dim]}</span>
              <span style={{ color:"var(--color-text-primary)", fontWeight:500, maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={val}>{val}</span>
              <button onClick={()=>clearFilter(dim)} style={{ background:"none", border:"none", cursor:"pointer", color:"var(--color-text-secondary)", padding:"0 2px", lineHeight:1, fontSize:14 }}>×</button>
            </span>
          ))}
          {(search||monthF!=="all"||locF!=="all") && (
            <span style={{ fontSize:12, color:"var(--color-text-secondary)" }}>+ search/date/location filters active</span>
          )}
          <button onClick={clearAll} style={{ marginLeft:"auto", ...S.btn, fontSize:11, padding:"3px 10px" }}><X size={11}/> Clear all</button>
        </div>
      )}

      {/* Tabs */}
      <div style={{ background:"var(--color-background-primary)", borderBottom:"0.5px solid var(--color-border-tertiary)", padding:"0 20px", display:"flex", gap:0 }}>
        {[
          { id:"mcp",     icon:<FileSpreadsheet size={14}/>, label:"MCP Output",  count:filtered.length },
          { id:"clashes", icon:<AlertTriangle size={14}/>,  label:"Clashes",      count:filtClashes.length, alert:!allClear },
          { id:"calendar",icon:<Calendar size={14}/>,       label:"Calendar" },
          { id:"stats",   icon:<BarChart2 size={14}/>,      label:"Stats" },
        ].map(tab => (
          <button key={tab.id} onClick={()=>setView(tab.id)} style={{ padding:"11px 16px", border:"none", borderBottom:view===tab.id?"2px solid #1d4ed8":"2px solid transparent", background:"transparent", cursor:"pointer", fontSize:13, fontWeight:view===tab.id?500:400, color:view===tab.id?"#1d4ed8":"var(--color-text-secondary)", display:"flex", alignItems:"center", gap:6, whiteSpace:"nowrap" }}>
            {tab.icon} {tab.label}
            {tab.count!==undefined && <span style={{ fontSize:11, background:view===tab.id?"#dbeafe":"var(--color-background-secondary)", color:view===tab.id?"#1e40af":"var(--color-text-secondary)", padding:"1px 7px", borderRadius:20, fontWeight:500 }}>{tab.count}</span>}
            {tab.alert && counts.hard+counts.city+counts.travel>0 && <span style={{ width:6, height:6, borderRadius:"50%", background:"#ef4444", display:"inline-block" }} />}
          </button>
        ))}
      </div>

      <div style={{ padding:"20px", maxWidth:1400, margin:"0 auto" }}>

        {/* ── MCP OUTPUT ── */}
        {view==="mcp" && (
          <div>
            <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
              <div style={{ width:240 }}><input style={S.input} placeholder="Search lecturer, class, course…" value={search} onChange={e=>setSearch(e.target.value)} /></div>
              <select style={S.select} value={monthF} onChange={e=>setMonthF(e.target.value)}>
                <option value="all">All months</option>
                {months.map(m=><option key={m} value={m}>{MONTHS[m]}</option>)}
              </select>
              <select style={S.select} value={locF} onChange={e=>setLocF(e.target.value)}>
                <option value="all">All locations</option>
                <option value="Jakarta">Jakarta</option>
                <option value="Bandung">Bandung</option>
              </select>
              {hasAnyFilter && <button style={S.btn} onClick={clearAll}><X size={13}/> Clear all</button>}
              <div style={{ marginLeft:"auto" }}><button style={S.btnPrimary} onClick={exportMCP}><Download size={14}/> Export MCP</button></div>
            </div>
            <div style={S.card}>
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr>
                    <SortTh col="lecturer"    label="Lecturer" />
                    <SortTh col="class"       label="Class" />
                    <SortTh col="program"     label="Program" />
                    <SortTh col="course"      label="Course" />
                    <SortTh col="date"        label="Date" />
                    <SortTh col="time"        label="Time" />
                    <th style={S.th}>Room</th>
                    <SortTh col="location"    label="Location" />
                    <SortTh col="sessionType" label="Type" />
                    <SortTh col="sourceSheet" label="Source" />
                  </tr></thead>
                  <tbody>
                    {filtered.slice(0,300).map((r,i) => {
                      const hasC = clashes.some(c=>!acked[c.id]&&c.rows.some(cr=>cr?.id===r.id));
                      return (
                        <tr key={r.id} style={{ background: hasC?"#fff5f5": i%2===0?"var(--color-background-primary)":"var(--color-background-secondary)" }}>
                          <td style={{...S.td, fontWeight:500}}>
                            {r.lecturer
                              ? <button style={{...S.link}} onClick={()=>toggleFilter("lecturer",r.lecturer)}>{r.lecturer}{hasC&&" ⚠"}</button>
                              : <span style={{color:"var(--color-text-danger)",fontSize:12}}>Unassigned</span>}
                          </td>
                          <td style={S.td}><button style={{...S.link,color:filters.class===r.class?"#1d4ed8":"var(--color-text-primary)"}} onClick={()=>toggleFilter("class",r.class)}>{r.class}</button></td>
                          <td style={S.td}><button style={{...S.link,color:filters.program===r.program?"#1d4ed8":"var(--color-text-primary)"}} onClick={()=>toggleFilter("program",r.program)}>{r.program||"—"}</button></td>
                          <td style={{...S.td, maxWidth:190, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}} title={r.course}>
                            <button style={{...S.link,color:filters.course===r.course?"#1d4ed8":"var(--color-text-primary)",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",display:"block"}} onClick={()=>toggleFilter("course",r.course)} title={r.course}>{r.course}</button>
                          </td>
                          <td style={{...S.td, whiteSpace:"nowrap"}}>{fmtDate(r.date)}</td>
                          <td style={{...S.td, whiteSpace:"nowrap", fontSize:12}}>{r.time||<span style={{color:"var(--color-text-secondary)"}}>—</span>}</td>
                          <td style={{...S.td, fontSize:12}}>{r.room||<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                          <td style={S.td}><LocBadge loc={r.location}/></td>
                          <td style={S.td}><Badge text={r.sessionType} bg={SESSION_STYLE[r.sessionType]?.bg} color={SESSION_STYLE[r.sessionType]?.text}/></td>
                          <td style={{...S.td, fontSize:11}}>
                            <button style={{...S.link,fontSize:11,color:filters.sheet===r.sourceSheet?"#1d4ed8":"var(--color-text-secondary)"}} onClick={()=>toggleFilter("sheet",r.sourceSheet)}>{r.sourceSheet}</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {filtered.length > 300 && <div style={{ padding:"10px 14px", fontSize:12, color:"var(--color-text-secondary)", borderTop:"0.5px solid var(--color-border-tertiary)" }}>Showing 300 of {filtered.length} rows — use filters to narrow down.</div>}
              {filtered.length === 0 && <div style={{ padding:40, textAlign:"center", color:"var(--color-text-secondary)", fontSize:13 }}>No sessions match the current filters.</div>}
            </div>
          </div>
        )}

        {/* ── CLASHES ── */}
        {view==="clashes" && (
          <div>
            <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
              {[["all","All types"],["hard","Hard"],["city","City"],["travel","Travel"]].map(([id,label]) => (
                <button key={id} onClick={()=>setClashF(id)} style={clashF===id?{...S.btnPrimary}:S.btn}>{label}</button>
              ))}
              <div style={{ marginLeft:"auto" }}><button style={{ ...S.btn, color:"var(--color-text-warning)", borderColor:"var(--color-border-warning)" }} onClick={exportClashes}><Download size={14}/> Export report</button></div>
            </div>

            {filtClashes.length===0
              ? <div style={{ ...S.card, padding:48, textAlign:"center" }}><Check size={32} color="var(--color-text-success)" style={{ margin:"0 auto 8px", display:"block" }}/><div style={{ color:"var(--color-text-secondary)", fontSize:13 }}>No clashes found for this filter.</div></div>
              : <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                {filtClashes.map(c => {
                  const cfg=CLASH_CFG[c.type], isAcked=acked[c.id];
                  return (
                    <div key={c.id} style={{ background:isAcked?"var(--color-background-secondary)":cfg.bg, border:isAcked?"0.5px solid var(--color-border-tertiary)":cfg.border, borderRadius:"var(--border-radius-lg)", padding:16, opacity:isAcked?0.8:1 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:12, flexWrap:"wrap" }}>
                        <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
                          <Badge text={cfg.label} bg={cfg.bg} color={cfg.text}/>
                          <button style={{...S.link, fontWeight:500, fontSize:14}} onClick={()=>toggleFilter("lecturer",c.lecturer)}>{c.lecturer}</button>
                          <span style={{ fontSize:12, color:"var(--color-text-secondary)" }}>
                            {c.type==="travel" ? `${fmtDate(c.rows[0]?.date)} → ${fmtDate(c.rows[1]?.date)}` : fmtDate(c.date)}
                          </span>
                        </div>
                        <button onClick={()=>setAcked(p=>({...p,[c.id]:!p[c.id]}))} style={isAcked?{...S.btnPrimary,fontSize:12}:{...S.btn,fontSize:12}}>
                          {isAcked?<><Check size={12}/> Acknowledged</>:"Acknowledge"}
                        </button>
                      </div>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginTop:12 }}>
                        {c.rows.map((r,ri) => r && (
                          <div key={ri} style={{ background:"var(--color-background-primary)", borderRadius:"var(--border-radius-md)", padding:12, border:"0.5px solid var(--color-border-tertiary)", fontSize:12, lineHeight:1.8 }}>
                            <div style={{ fontWeight:500, color:"var(--color-text-primary)", marginBottom:4, fontSize:13 }}>{r.course}</div>
                            <div style={{ color:"var(--color-text-secondary)", display:"flex", gap:6, alignItems:"center" }}><MapPin size={11}/> <LocBadge loc={r.location}/></div>
                            <div style={{ color:"var(--color-text-secondary)" }}>Time: {r.time||"—"}</div>
                            <div style={{ color:"var(--color-text-secondary)" }}>Room: {r.room||"—"}</div>
                            <div style={{ color:"var(--color-text-secondary)" }}>
                              Class: <button style={{...S.link,fontSize:12}} onClick={()=>toggleFilter("class",r.class)}>{r.class}</button>
                            </div>
                            {c.type==="travel" && <div style={{ color:"var(--color-text-secondary)" }}>Date: {fmtDate(r.date)}</div>}
                          </div>
                        ))}
                      </div>
                      {isAcked && (
                        <div style={{ marginTop:10 }}>
                          <input style={{...S.input, fontSize:12}} placeholder="Add a resolution note…" value={notes[c.id]||""} onChange={e=>setNotes(p=>({...p,[c.id]:e.target.value}))} />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            }
          </div>
        )}

        {/* ── CALENDAR ── */}
        {view==="calendar" && (() => {
          const y=calDate.getFullYear(), m=calDate.getMonth();
          const firstDOW=new Date(y,m,1).getDay(), dim=new Date(y,m+1,0).getDate();
          const cells=[...Array(firstDOW).fill(null),...Array.from({length:dim},(_,i)=>i+1)];
          return (
            <div>
              <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16 }}>
                <button style={S.btn} onClick={()=>setCalDate(new Date(y,m-1,1))}><ChevronLeft size={15}/></button>
                <span style={{ fontSize:17, fontWeight:500, flex:1, textAlign:"center" }}>{MONTHS[m]} {y}</span>
                <button style={S.btn} onClick={()=>setCalDate(new Date(y,m+1,1))}><ChevronRight size={15}/></button>
              </div>
              <div style={{ display:"flex", gap:16, marginBottom:12, fontSize:12, color:"var(--color-text-secondary)", flexWrap:"wrap" }}>
                <span style={{ display:"flex", alignItems:"center", gap:4 }}><span style={{ width:10, height:10, borderRadius:2, background:"#bfdbfe", display:"inline-block" }}/> Jakarta</span>
                <span style={{ display:"flex", alignItems:"center", gap:4 }}><span style={{ width:10, height:10, borderRadius:2, background:"#bbf7d0", display:"inline-block" }}/> Bandung</span>
                <span style={{ display:"flex", alignItems:"center", gap:4 }}><span style={{ width:10, height:10, borderRadius:2, background:"#fde68a", display:"inline-block" }}/> Both cities</span>
                <span style={{ display:"flex", alignItems:"center", gap:4 }}><span style={{ width:8, height:8, borderRadius:"50%", background:"#ef4444", display:"inline-block" }}/> Has clash</span>
                <span style={{ fontSize:11, color:"var(--color-text-secondary)", marginLeft:4 }}>Click a name to filter · click date chip to zoom</span>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(7, 1fr)", gap:6 }}>
                {["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].map(d=><div key={d} style={{ textAlign:"center", fontWeight:500, fontSize:11, color:"var(--color-text-secondary)", padding:"4px 0", textTransform:"uppercase", letterSpacing:"0.08em" }}>{d}</div>)}
                {cells.map((day,idx) => {
                  if (!day) return <div key={`e${idx}`}/>;
                  const dr=calData.byDay[day]||[], dc=calData.clashDays[day];
                  const hasJ=dr.some(r=>r.location==="Jakarta"), hasB=dr.some(r=>r.location==="Bandung");
                  const hasClash=dc&&(dc.hard+dc.city+dc.travel)>0;
                  const bg=dr.length===0?"var(--color-background-secondary)":hasJ&&hasB?"#fffbeb":hasJ?"#eff6ff":"#f0fdf4";
                  const lecs=[...new Set(dr.map(r=>r.lecturer).filter(Boolean))];
                  return (
                    <div key={day} style={{ background:bg, borderRadius:"var(--border-radius-md)", padding:"7px 6px", minHeight:76, border:hasClash?"1.5px solid #ef4444":"0.5px solid var(--color-border-tertiary)", position:"relative" }}>
                      <div style={{ fontWeight:500, fontSize:13, marginBottom:4 }}>{day}</div>
                      {lecs.slice(0,3).map((l,li) => {
                        const loc=dr.find(r=>r.lecturer===l)?.location;
                        return <div key={li} onClick={()=>toggleFilter("lecturer",l)} style={{ fontSize:9, background:loc==="Jakarta"?"#bfdbfe":"#bbf7d0", borderRadius:3, padding:"1px 4px", marginBottom:2, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", cursor:"pointer" }}>{l}</div>;
                      })}
                      {lecs.length>3 && <div style={{ fontSize:9, color:"var(--color-text-secondary)" }}>+{lecs.length-3} more</div>}
                      {hasClash && <div style={{ position:"absolute", top:5, right:5, width:7, height:7, borderRadius:"50%", background:"#ef4444" }}/>}
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {/* ── STATS ── */}
        {view==="stats" && (
          <div>
            {/* Stats subtabs */}
            <div style={{ display:"flex", gap:4, marginBottom:14, borderBottom:"0.5px solid var(--color-border-tertiary)", paddingBottom:0 }}>
              {[["lecturer",`Lecturers (${lecStats.length})`],["class",`Classes (${classStats.length})`],["program",`Programs (${programStats.length})`]].map(([id,label])=>(
                <button key={id} onClick={()=>setStatsTab(id)} style={{ padding:"8px 16px", border:"none", borderBottom:statsTab===id?"2px solid #1d4ed8":"2px solid transparent", background:"transparent", cursor:"pointer", fontSize:13, fontWeight:statsTab===id?500:400, color:statsTab===id?"#1d4ed8":"var(--color-text-secondary)", marginBottom:-1 }}>{label}</button>
              ))}
              <div style={{ marginLeft:"auto", display:"flex", alignItems:"center", paddingBottom:8 }}>
                <div style={{ width:200 }}><input style={{...S.input, fontSize:12}} placeholder="Search…" value={search} onChange={e=>setSearch(e.target.value)} /></div>
              </div>
            </div>

            {/* By Lecturer */}
            {statsTab==="lecturer" && (
              <div style={S.card}>
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead><tr>{["Lecturer","Total","Sessions","Mid","Final","Jakarta","Bandung","Programs","Classes","Clashes"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                    <tbody>
                      {lecStats.filter(s=>!search||s.name.toLowerCase().includes(search.toLowerCase())).map((s,i) => {
                        const lc=clashes.filter(c=>c.lecturer===s.name&&!acked[c.id]);
                        return (
                          <tr key={s.name} style={{ background:i%2===0?"var(--color-background-primary)":"var(--color-background-secondary)", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>
                            <td style={{...S.td, fontWeight:500}}><button style={S.link} onClick={()=>{toggleFilter("lecturer",s.name);setView("mcp")}}>{s.name}</button></td>
                            <td style={{...S.td, fontWeight:500}}>{s.total}</td>
                            <td style={S.td}>{s.sessions}</td>
                            <td style={S.td}>{s.mid||<span style={{color:"var(--color-text-secondary)"}}>—</span>}</td>
                            <td style={S.td}>{s.final||<span style={{color:"var(--color-text-secondary)"}}>—</span>}</td>
                            <td style={S.td}>{s.jakarta?<Badge text={s.jakarta} bg="#dbeafe" color="#1e40af"/>:<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                            <td style={S.td}>{s.bandung?<Badge text={s.bandung} bg="#dcfce7" color="#166534"/>:<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                            <td style={{...S.td, fontSize:12, color:"var(--color-text-secondary)"}}>{[...s.programs].join(", ")||"—"}</td>
                            <td style={{...S.td, fontSize:11, color:"var(--color-text-secondary)", maxWidth:130, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}} title={[...s.classes].join(", ")}>{[...s.classes].join(", ")}</td>
                            <td style={S.td}>
                              {lc.length>0
                                ? <button style={{...S.link, color:"var(--color-text-danger)", fontWeight:500}} onClick={()=>{toggleFilter("lecturer",s.name);setView("clashes")}}>⚠ {lc.length}</button>
                                : <span style={{color:"var(--color-text-success)", fontWeight:500}}><Check size={13}/></span>}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* By Class */}
            {statsTab==="class" && (
              <div style={S.card}>
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead><tr>{["Class","Total Sessions","Lecturers","Programs","Jakarta","Bandung"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                    <tbody>
                      {classStats.filter(s=>!search||s.name.toLowerCase().includes(search.toLowerCase())).map((s,i) => (
                        <tr key={s.name} style={{ background:i%2===0?"var(--color-background-primary)":"var(--color-background-secondary)", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>
                          <td style={{...S.td, fontWeight:500}}><button style={S.link} onClick={()=>{toggleFilter("class",s.name);setView("mcp")}}>{s.name}</button></td>
                          <td style={{...S.td, fontWeight:500}}>{s.total}</td>
                          <td style={{...S.td, fontSize:12, color:"var(--color-text-secondary)"}}>{[...s.lecturers].length} — <span style={{fontSize:11}}>{[...s.lecturers].join(", ")}</span></td>
                          <td style={{...S.td, fontSize:12, color:"var(--color-text-secondary)"}}>{[...s.programs].join(", ")||"—"}</td>
                          <td style={S.td}>{s.jakarta?<Badge text={s.jakarta} bg="#dbeafe" color="#1e40af"/>:<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                          <td style={S.td}>{s.bandung?<Badge text={s.bandung} bg="#dcfce7" color="#166534"/>:<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* By Program */}
            {statsTab==="program" && (
              <div style={S.card}>
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse" }}>
                    <thead><tr>{["Program","Total Sessions","Lecturers","Classes","Jakarta","Bandung"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                    <tbody>
                      {programStats.filter(s=>!search||s.name.toLowerCase().includes(search.toLowerCase())).map((s,i) => (
                        <tr key={s.name} style={{ background:i%2===0?"var(--color-background-primary)":"var(--color-background-secondary)", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>
                          <td style={{...S.td, fontWeight:500}}><button style={S.link} onClick={()=>{toggleFilter("program",s.name);setView("mcp")}}>{s.name}</button></td>
                          <td style={{...S.td, fontWeight:500}}>{s.total}</td>
                          <td style={{...S.td, fontSize:12, color:"var(--color-text-secondary)"}}>{[...s.lecturers].length} lecturers</td>
                          <td style={{...S.td, fontSize:12, color:"var(--color-text-secondary)"}}>{[...s.classes].join(", ")||"—"}</td>
                          <td style={S.td}>{s.jakarta?<Badge text={s.jakarta} bg="#dbeafe" color="#1e40af"/>:<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                          <td style={S.td}>{s.bandung?<Badge text={s.bandung} bg="#dcfce7" color="#166534"/>:<span style={{color:"var(--color-border-secondary)"}}>—</span>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
}
