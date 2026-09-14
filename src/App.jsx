import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  Download, FileSpreadsheet, FileText, Image as ImageIcon, Plus, Trash2,
  Palette, RotateCcw, Save, Upload, Presentation, ChevronDown, GripVertical,
  Settings2, Copy, Check, ZoomIn, ZoomOut, FileCode2
} from "lucide-react";
import { toPng } from "html-to-image";
import jsPDF from "jspdf";
import ExcelJS from "exceljs";
import pptxgen from "pptxgenjs";

const MSPDI_NS = "http://schemas.microsoft.com/project/2007";
function startOfMonth(value){ const d=value instanceof Date ? new Date(value) : new Date(`${value}T00:00:00Z`); d.setUTCDate(1); d.setUTCHours(0,0,0,0); return d; }
function endOfMonth(value){ const d=startOfMonth(value); d.setUTCMonth(d.getUTCMonth()+1); d.setUTCDate(0); d.setUTCHours(17,0,0,0); return d; }
function monthLabel(d){ return d.toLocaleDateString("en-US",{month:"short",year:"numeric",timeZone:"UTC"}); }
function timelineGranularity(settings){
  const start=new Date(`${settings.timelineStart || settings.projectStart || "2025-11-01"}T00:00:00Z`);
  const end=new Date(`${settings.timelineEnd || "2026-09-30"}T23:59:59Z`);
  const hours=Math.max(0,(end-start)/3600000);
  const days=hours/24;
  const months=(end.getUTCFullYear()-start.getUTCFullYear())*12+(end.getUTCMonth()-start.getUTCMonth())+(end.getUTCDate()-1)/31;
  if(hours <= 72) return "hour";
  if(days <= 21) return "day";
  if(days <= 120) return "week";
  return "month";
}
function startOfWeek(value){
  const d=value instanceof Date ? new Date(value) : new Date(`${value}T00:00:00Z`);
  const day=d.getUTCDay(); const diff=day===0?-6:1-day;
  d.setUTCDate(d.getUTCDate()+diff); d.setUTCHours(0,0,0,0); return d;
}
function endOfWeek(value){const d=startOfWeek(value);d.setUTCDate(d.getUTCDate()+6);d.setUTCHours(23,59,59,999);return d;}
function startOfDay(value){const d=value instanceof Date?new Date(value):new Date(`${value}T00:00:00Z`);d.setUTCHours(0,0,0,0);return d;}
function endOfDay(value){const d=startOfDay(value);d.setUTCHours(23,59,59,999);return d;}
function startOfHour(value){const d=new Date(value);d.setUTCMinutes(0,0,0);return d;}
function endOfHour(value){const d=startOfHour(value);d.setUTCHours(d.getUTCHours()+1);d.setUTCMilliseconds(-1);return d;}
function timelineUnits(settings){
  const unit=timelineGranularity(settings);
  const startRaw=new Date(`${settings.timelineStart || settings.projectStart || "2025-11-01"}T00:00:00Z`);
  const endRaw=new Date(`${settings.timelineEnd || "2026-09-30"}T23:59:59Z`);
  let d=unit==="month"?startOfMonth(startRaw):unit==="week"?startOfWeek(startRaw):unit==="day"?startOfDay(startRaw):startOfHour(startRaw);
  const out=[];
  while(d<=endRaw && out.length<500){
    const start=new Date(d);
    let end;
    if(unit==="month") end=endOfMonth(d);
    else if(unit==="week") end=endOfWeek(d);
    else if(unit==="day") end=endOfDay(d);
    else end=endOfHour(d);
    out.push({start,end,label:timelineUnitLabel(start,unit),key:start.toISOString()});
    if(unit==="month") d.setUTCMonth(d.getUTCMonth()+1);
    else if(unit==="week") d.setUTCDate(d.getUTCDate()+7);
    else if(unit==="day") d.setUTCDate(d.getUTCDate()+1);
    else d.setUTCHours(d.getUTCHours()+1);
  }
  return out.length?out:[{start:startOfDay(startRaw),end:endOfDay(startRaw),label:timelineUnitLabel(startRaw,"day"),key:startRaw.toISOString()}];
}
function timelineUnitLabel(d,unit){
  if(unit==="month") return d.toLocaleDateString("en-US",{month:"short",year:"2-digit",timeZone:"UTC"});
  if(unit==="week") return d.toLocaleDateString("en-US",{month:"short",day:"numeric",timeZone:"UTC"});
  if(unit==="day") return d.toLocaleDateString("en-US",{day:"numeric",month:"short",timeZone:"UTC"});
  return d.toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit",timeZone:"UTC"});
}


const LINK_TYPES = { FF: 0, FS: 1, SF: 2, SS: 3 };
const LINK_LABELS = { FF: "Finish → Finish", FS: "Finish → Start", SF: "Start → Finish", SS: "Start → Start" };

const starterRows = [
  { id: 1, title: "Build / Train the model", objective: "Analyze and understand the needs of your target audience", start: 0, end: 5, startDate:"2025-11-01", finishDate:"2026-04-30", predecessorId: "", dependencyType: "FS", lag: 0 },
  { id: 2, title: "Mobile app development and deployment", objective: "Develop & test our mobile app (Pulses) and deploy in play store / apple store", start: 0, end: 5, startDate:"2025-11-01", finishDate:"2026-04-30", predecessorId: "", dependencyType: "FS", lag: 0 },
  { id: 3, title: "Conduct IoT devices Testing", objective: "Research on M5 Stick C plus, Lilygo 3 to see how they perform in sound detection and model hosting", start: 0, end: 2, startDate:"2025-11-01", finishDate:"2026-01-31", predecessorId: "", dependencyType: "FS", lag: 0 },
  { id: 4, title: "Audio sample collection for model training", objective: "Collect environmental audio samples in Benin for model training.", start: 1, end: 2, startDate:"2025-12-01", finishDate:"2026-01-31", predecessorId: "", dependencyType: "FS", lag: 0 },
  { id: 5, title: "Test for usability", objective: "Put the prototype through rigorous testing processes to ensure that it meets user requirements", start: 3, end: 6, startDate:"2026-02-01", finishDate:"2026-05-31", predecessorId: "", dependencyType: "FS", lag: 0 },
  { id: 6, title: "Analyze feedback and improvement", objective: "Analyze feedback", start: 5, end: 10, startDate:"2026-04-01", finishDate:"2026-09-30", predecessorId: "", dependencyType: "FS", lag: 0 }
];

function normalizeRow(r, projectStart="2025-11-01"){
  const deps=Array.isArray(r.dependencies) ? r.dependencies.map(d=>({
    predecessorId:d.predecessorId ?? "",
    dependencyType:d.dependencyType || "FS",
    lag:Number(d.lag || 0)
  })) : (r.predecessorId ? [{predecessorId:r.predecessorId,dependencyType:r.dependencyType || "FS",lag:Number(r.lag || 0)}] : []);
  let startDate=r.startDate, finishDate=r.finishDate;
  if(!startDate || !finishDate){
    const base=startOfMonth(projectStart); const sd=addMonths(base,Number(r.start||0)); const fd=endOfMonth(addMonths(base,Number(r.end||r.start||0)));
    startDate=sd.toISOString().slice(0,10); finishDate=fd.toISOString().slice(0,10);
  }
  const startDateTime=r.startDateTime || `${startDate}T08:00`;
  const finishDateTime=r.finishDateTime || `${finishDate}T17:00`;
  return {...r,dependencies:deps,startDate,finishDate,startDateTime,finishDateTime};
}

const themes = {
  ocean:{name:"Ocean Blue",header:"#d8d2cf",headerText:"#24272d",bar:"#6d9de3",barText:"#fff",canvas:"#fff",grid:"#d8dde5",text:"#252a33",muted:"#687080",accent:"#3f73c8"},
  slate:{name:"Slate",header:"#263142",headerText:"#fff",bar:"#5f88b9",barText:"#fff",canvas:"#fff",grid:"#d6dbe3",text:"#202733",muted:"#697386",accent:"#334b6b"},
  sage:{name:"Sage",header:"#dce3dc",headerText:"#26332a",bar:"#78947d",barText:"#fff",canvas:"#fff",grid:"#d7ded8",text:"#28332b",muted:"#6b776e",accent:"#56765e"},
  sand:{name:"Warm Sand",header:"#e8dfd1",headerText:"#3b3329",bar:"#c58c5b",barText:"#fff",canvas:"#fffdf9",grid:"#ded7cc",text:"#342f2a",muted:"#7c7369",accent:"#a86f3e"},
  lavender:{name:"Lavender",header:"#dedbea",headerText:"#312e40",bar:"#8e83bd",barText:"#fff",canvas:"#fff",grid:"#deddea",text:"#302d3a",muted:"#706b7d",accent:"#6e63a1"},
  charcoal:{name:"Charcoal",header:"#30343b",headerText:"#fff",bar:"#767f8e",barText:"#fff",canvas:"#fff",grid:"#d5d9df",text:"#24272d",muted:"#6b7280",accent:"#4b5563"}
};

const initialSettings = {
  theme:"ocean",barColor:"",barTextColor:"",headerColor:"",headerTextColor:"",
  backgroundColor:"",gridColor:"",textColor:"",showGrid:true,borderWidth:1,barRadius:2,
  fontSize:13,rowHeight:58,title:"Project Gantt Chart",subtitle:"Project plan and key initiatives",
  projectStart:"2025-11-01", timelineStart:"2025-11-01", timelineEnd:"2026-09-30", projectGuid:null
};

function uid(){return Math.random().toString(36).slice(2)+Date.now().toString(36)}
function hexArgb(h){return "FF"+h.replace("#","").padStart(6,"0").slice(0,6).toUpperCase()}
function esc(s=""){return String(s).replace(/[<>&'"]/g,c=>({"<":"&lt;",">":"&gt;","&":"&amp;","'":"&apos;",'"':"&quot;"}[c]))}
function isoDate(d){return new Date(d).toISOString().slice(0,19)}
function parseDate(s){return new Date(s.includes("Z")?s:s+"Z")}
function addMonths(d,n){const x=new Date(d); x.setUTCMonth(x.getUTCMonth()+n); return x}
function monthIndexFor(d, base){
  const a=new Date(base), b=new Date(d);
  return (b.getUTCFullYear()-a.getUTCFullYear())*12 + b.getUTCMonth()-a.getUTCMonth();
}
function workingDaysInclusive(start, finish){
  let d=new Date(start), count=0;
  while(d<=finish){ const day=d.getUTCDay(); if(day!==0&&day!==6)count++; d.setUTCDate(d.getUTCDate()+1); }
  return Math.max(1,count);
}
function xmlEl(doc,name,value,parent){
  const e=doc.createElementNS(MSPDI_NS,name);
  if(value!==undefined && value!==null)e.textContent=String(value);
  if(parent)parent.appendChild(e);
  return e;
}
function childText(node,name){
  const e=[...node.children].find(x=>x.localName===name || x.tagName===name);
  return e?.textContent || "";
}

function taskUnitSpan(r, units){
  const {start:taskStart,finish:taskFinish}=getTaskDates(r,initialSettings);
  let first=-1,last=-1;
  units.forEach((u,i)=>{ if(taskStart<=u.end && taskFinish>=u.start){if(first<0)first=i;last=i;} });
  return {first,last};
}


function getTaskDates(r, settings){
  const start=r.startDateTime ? new Date(r.startDateTime.includes("Z")?r.startDateTime:`${r.startDateTime}Z`) : r.startDate ? new Date(`${r.startDate}T08:00:00Z`) : addMonths(startOfMonth(settings.projectStart),Number(r.start||0));
  const finish=r.finishDateTime ? new Date(r.finishDateTime.includes("Z")?r.finishDateTime:`${r.finishDateTime}Z`) : r.finishDate ? new Date(`${r.finishDate}T17:00:00Z`) : endOfMonth(addMonths(startOfMonth(settings.projectStart),Number(r.end||r.start||0)));
  return {start,finish};
}
function localDateTimeInput(d){
  const x=new Date(d);
  const pad=n=>String(n).padStart(2,"0");
  return `${x.getUTCFullYear()}-${pad(x.getUTCMonth()+1)}-${pad(x.getUTCDate())}T${pad(x.getUTCHours())}:${pad(x.getUTCMinutes())}`;
}
function taskInputDate(r,which){
  const d=getTaskDates(r,initialSettings)[which];
  return localDateTimeInput(d);
}

function buildMSPDI(rows, settings){
  const doc=document.implementation.createDocument(MSPDI_NS,"Project",null);
  const root=doc.documentElement;
  root.setAttribute("xmlns",MSPDI_NS);
  root.setAttribute("xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance");

  const base = startOfMonth(settings.projectStart || settings.timelineStart || "2025-11-01");
  const taskDates=rows.map(r=>getTaskDates(r,settings));
  const starts=taskDates.map(x=>x.start);
  const finishes=taskDates.map(x=>x.finish);
  const projectStart=new Date(Math.min(...starts.map(d=>d.getTime()),base.getTime()));
  const projectFinish=new Date(Math.max(...finishes.map(d=>d.getTime()),new Date(`${settings.timelineEnd||"2026-09-30"}T17:00:00Z`).getTime()));

  xmlEl(doc,"SaveVersion",12,root);
  xmlEl(doc,"GUID",settings.projectGuid || crypto.randomUUID(),root);
  xmlEl(doc,"UID",1,root);
  xmlEl(doc,"Name",settings.title || "Gantt Project",root);
  xmlEl(doc,"Title",settings.title || "Gantt Project",root);
  xmlEl(doc,"Subject",settings.subtitle || "",root);
  xmlEl(doc,"CreationDate",isoDate(new Date()),root);
  xmlEl(doc,"Revision",1,root);
  xmlEl(doc,"LastSaved",isoDate(new Date()),root);
  xmlEl(doc,"ScheduleFromStart",1,root);
  xmlEl(doc,"StartDate",isoDate(projectStart),root);
  xmlEl(doc,"FinishDate",isoDate(projectFinish),root);
  xmlEl(doc,"CurrentDate",isoDate(new Date()),root);
  xmlEl(doc,"CurrencyDigits",2,root);
  xmlEl(doc,"CurrencySymbol","$",root);
  xmlEl(doc,"Autolink",1,root);
  xmlEl(doc,"NewTaskStartDate",0,root);
  xmlEl(doc,"DefaultStartTime","08:00:00",root);
  xmlEl(doc,"DefaultFinishTime","17:00:00",root);
  xmlEl(doc,"MinutesPerDay",480,root);
  xmlEl(doc,"MinutesPerWeek",2400,root);
  xmlEl(doc,"DaysPerMonth",20,root);
  xmlEl(doc,"DefaultTaskType",0,root);
  xmlEl(doc,"DefaultFixedCostAccrual",2,root);

  const calendars=xmlEl(doc,"Calendars",null,root);
  const cal=xmlEl(doc,"Calendar",null,calendars);
  xmlEl(doc,"UID",1,cal); xmlEl(doc,"Name","Standard",cal);
  xmlEl(doc,"IsBaseCalendar",1,cal); xmlEl(doc,"BaseCalendarUID",-1,cal);
  const weekdays=xmlEl(doc,"WeekDays",null,cal);
  for(let day=1;day<=7;day++){
    const wd=xmlEl(doc,"WeekDay",null,weekdays);
    xmlEl(doc,"DayType",day,wd);
    const working=day>=2 && day<=6;
    xmlEl(doc,"DayWorking",working?1:0,wd);
    if(working){
      const wt=xmlEl(doc,"WorkingTimes",null,wd);
      const w=xmlEl(doc,"WorkingTime",null,wt);
      xmlEl(doc,"FromTime","08:00:00",w); xmlEl(doc,"ToTime","12:00:00",w);
      const w2=xmlEl(doc,"WorkingTime",null,wt);
      xmlEl(doc,"FromTime","13:00:00",w2); xmlEl(doc,"ToTime","17:00:00",w2);
    }
  }

  const tasks=xmlEl(doc,"Tasks",null,root);
  rows.forEach((r,i)=>{
    const task=xmlEl(doc,"Task",null,tasks);
    const {start,finish}=getTaskDates(r,settings);
    const days=workingDaysInclusive(start,finish);
    xmlEl(doc,"UID",i+1,task);
    xmlEl(doc,"ID",i+1,task);
    xmlEl(doc,"Name",r.title,task);
    xmlEl(doc,"Type",1,task);
    xmlEl(doc,"IsNull",0,task);
    xmlEl(doc,"CreateDate",isoDate(new Date()),task);
    xmlEl(doc,"WBS",String(i+1),task);
    xmlEl(doc,"WBSLevel",1,task);
    xmlEl(doc,"OutlineNumber",String(i+1),task);
    xmlEl(doc,"OutlineLevel",1,task);
    xmlEl(doc,"Priority",500,task);
    xmlEl(doc,"Start",isoDate(start),task);
    xmlEl(doc,"Finish",isoDate(finish),task);
    xmlEl(doc,"Duration",`PT${days*8}H0M0S`,task);
    xmlEl(doc,"DurationFormat",7,task);
    xmlEl(doc,"Work",`PT${days*8}H0M0S`,task);
    xmlEl(doc,"CalendarUID",1,task);
    xmlEl(doc,"Notes",r.objective || "",task);
    xmlEl(doc,"PercentComplete",0,task);
    xmlEl(doc,"Milestone",0,task);
    xmlEl(doc,"Summary",0,task);
    xmlEl(doc,"Active",1,task);
    const relationships = Array.isArray(r.dependencies) ? r.dependencies : (r.predecessorId ? [{predecessorId:r.predecessorId,dependencyType:r.dependencyType || "FS",lag:r.lag || 0}] : []);
    relationships.filter(d=>d.predecessorId).forEach(d=>{
      const pred = xmlEl(doc,"PredecessorLink",null,task);
      xmlEl(doc,"PredecessorUID",Number(d.predecessorId),pred);
      xmlEl(doc,"Type",LINK_TYPES[d.dependencyType || "FS"],pred);
      xmlEl(doc,"CrossProject",0,pred);
      if (Number(d.lag || 0) !== 0) {
        xmlEl(doc,"LinkLag",Number(d.lag) * 480,pred);
        xmlEl(doc,"LagFormat",7,pred);
      }
    });
    // Preserve the Gantt Studio presentation as MSPDI custom text fields.
    const ext=xmlEl(doc,"ExtendedAttribute",null,task);
    xmlEl(doc,"UID",1,ext);
    xmlEl(doc,"FieldID","188743731",ext); // Text1
    xmlEl(doc,"Value",settings.theme || "ocean",ext);
  });

  const resources=xmlEl(doc,"Resources",null,root);
  const resource=xmlEl(doc,"Resource",null,resources);
  xmlEl(doc,"UID",1,resource); xmlEl(doc,"ID",1,resource);
  xmlEl(doc,"Name","Unassigned",resource); xmlEl(doc,"Type",1,resource);
  xmlEl(doc,"IsNull",0,resource); xmlEl(doc,"CalendarUID",1,resource);

  const assignments=xmlEl(doc,"Assignments",null,root);
  rows.forEach((r,i)=>{
    const a=xmlEl(doc,"Assignment",null,assignments);
    xmlEl(doc,"UID",i+1,a); xmlEl(doc,"TaskUID",i+1,a); xmlEl(doc,"ResourceUID",1,a);
    xmlEl(doc,"Units",1,a);
  });

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="MSProject.Project"?>\n' +
    new XMLSerializer().serializeToString(doc);
}

function parseMSPDI(text){
  const doc=new DOMParser().parseFromString(text,"application/xml");
  if(doc.querySelector("parsererror")) throw new Error("Invalid XML");
  const root=doc.documentElement;
  if(root.localName!=="Project" || root.namespaceURI!==MSPDI_NS) throw new Error("Not an MSPDI Project XML file.");
  const projectStart=childText(root,"StartDate").slice(0,10) || "2025-11-01";
  const projectFinish=childText(root,"FinishDate").slice(0,10) || projectStart;
  const base=startOfMonth(projectStart);
  const taskNodes=[...root.getElementsByTagNameNS(MSPDI_NS,"Task")];
  if(!taskNodes.length) throw new Error("The MSPDI file contains no tasks.");

  const parsed=taskNodes.map((t,i)=>{
    const start=parseDate(childText(t,"Start"));
    const finish=parseDate(childText(t,"Finish"));
    let si=Math.max(0,monthIndexFor(start,base));
    let ei=Math.max(si,monthIndexFor(finish,base));
    const endMonth=addMonths(base,ei+1); endMonth.setUTCDate(endMonth.getUTCDate()-1);
    if(finish.getUTCDate()>=20) ei=Math.max(ei,si);
    const links=[...t.getElementsByTagNameNS(MSPDI_NS,"PredecessorLink")];
    const dependencies=links.map(link=>{
      const typeNumber=Number(childText(link,"Type"));
      return {
        predecessorId:childText(link,"PredecessorUID"),
        dependencyType:Object.keys(LINK_TYPES).find(k=>LINK_TYPES[k]===typeNumber) || "FS",
        lag:childText(link,"LinkLag") ? Math.round(Number(childText(link,"LinkLag"))/480) : 0
      };
    });
    return normalizeRow({
      id:Number(childText(t,"UID")) || i+1,
      title:childText(t,"Name") || `Task ${i+1}`,
      objective:childText(t,"Notes"),
      start:Math.max(0,si),
      end:Math.max(si,ei),
      startDate:childText(t,"Start").slice(0,10) || projectStart,
      finishDate:childText(t,"Finish").slice(0,10) || projectFinish,
      dependencies
    });
  });

  const subject=childText(root,"Subject");
  const title=childText(root,"Title") || childText(root,"Name") || "Project Gantt Chart";
  let theme="ocean";
  const firstExt=taskNodes[0]?.getElementsByTagNameNS(MSPDI_NS,"ExtendedAttribute")[0];
  if(firstExt){const v=childText(firstExt,"Value"); if(themes[v])theme=v;}

  return {rows:parsed, settings:{...initialSettings, projectStart, timelineStart:projectStart, timelineEnd:projectFinish, title, subtitle:subject, theme}};
}

function downloadBlob(blob,name){
  const u=URL.createObjectURL(blob), a=document.createElement("a");
  a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);
}

function App(){
  const [rows,setRows]=useState(starterRows.map(normalizeRow));
  const [initiativeSettingsId,setInitiativeSettingsId]=useState(null);
  const [initiativeSettingsTab,setInitiativeSettingsTab]=useState("general");
  const [settings,setSettings]=useState(initialSettings);
  const [themeOpen,setThemeOpen]=useState(false),[exportOpen,setExportOpen]=useState(false);
  const [zoom,setZoom]=useState(1),[dragging,setDragging]=useState(false),[saved,setSaved]=useState(false);
  const [timelineDrag,setTimelineDrag]=useState(null);
  const chartRef=useRef(null);
  const theme=themes[settings.theme];
  const colors=useMemo(()=>({...theme,
    bar:settings.barColor||theme.bar,barText:settings.barTextColor||theme.barText,
    header:settings.headerColor||theme.header,headerText:settings.headerTextColor||theme.headerText,
    canvas:settings.backgroundColor||theme.canvas,grid:settings.gridColor||theme.grid,text:settings.textColor||theme.text
  }),[theme,settings]);
  const visibleUnits=timelineUnits(settings);

  useEffect(()=>{const raw=localStorage.getItem("gantt-studio-mspdi");if(raw)try{
    const x=JSON.parse(raw);const loadedSettings={...initialSettings,...(x.settings||{})};if(x.rows)setRows(x.rows.map(r=>normalizeRow(r,loadedSettings.projectStart)));setSettings(loadedSettings);
  }catch{}},[]);

  const set=patch=>setSettings(s=>({...s,...patch}));
  const updateRow=(id,patch)=>setRows(p=>p.map(r=>r.id===id?{...r,...patch}:r));
  const applyTimelineRange=(id,a,b)=>setRows(p=>p.map(r=>{
    if(r.id!==id)return r;
    const tl=timelineUnits(settings);
    const first=tl[Math.max(0,Math.min(a,b))];
    const last=tl[Math.max(0,Math.max(a,b))];
    if(!first||!last)return r;
    const start=new Date(first.start);
    const finish=new Date(last.end);
    return {...r,startDate:start.toISOString().slice(0,10),finishDate:finish.toISOString().slice(0,10),startDateTime:localDateTimeInput(start),finishDateTime:localDateTimeInput(finish)};
  }));
  // Timeline editing is intentionally click-based: first click sets the start,
  // then hovering previews the range, and the second click commits the finish.
  const beginTimelineSelection=(rowId,i)=>{
    if(timelineDrag?.rowId===rowId && timelineDrag.anchor!==null && timelineDrag.current!==null){
      applyTimelineRange(rowId,timelineDrag.anchor,i);
      setTimelineDrag(null);
      setDragging(false);
      return;
    }
    setDragging(false);
    setTimelineDrag({rowId,anchor:i,current:i});
  };
  const updateTimelineSelection=(rowId,i)=>{
    if(!timelineDrag || timelineDrag.rowId!==rowId || timelineDrag.anchor===null)return;
    setTimelineDrag(x=>x?{...x,current:i}:x);
  };
  const cancelTimelineSelection=()=>{setTimelineDrag(null);setDragging(false)};
  const addRow=()=>setRows(p=>[...p,normalizeRow({id:uid(),title:"New initiative",objective:"Describe the objective",start:0,end:2,dependencies:[]},settings.projectStart)]);
  const duplicateRow=id=>{const r=rows.find(x=>x.id===id);if(!r)return;const c={...r,id:uid(),title:r.title+" (copy)",dependencies:[]};setRows(p=>{const n=p.findIndex(x=>x.id===id);return[...p.slice(0,n+1),c,...p.slice(n+1)]})};
  const openInitiativeSettings=id=>{setInitiativeSettingsId(id);setInitiativeSettingsTab("general")};
  const closeInitiativeSettings=()=>setInitiativeSettingsId(null);
  const updateDependency=(rowId,index,patch)=>setRows(p=>p.map(r=>{if(r.id!==rowId)return r;const dependencies=(r.dependencies||[]).map((d,i)=>i===index?{...d,...patch}:d);return {...r,dependencies}}));
  const addDependency=rowId=>setRows(p=>p.map(r=>{if(r.id!==rowId)return r;const used=new Set((r.dependencies||[]).map(d=>String(d.predecessorId)));const candidate=p.find(x=>x.id!==r.id&&!used.has(String(x.id)));if(!candidate)return r;return {...r,dependencies:[...(r.dependencies||[]),{predecessorId:candidate.id,dependencyType:"FS",lag:0}]}}));
  const removeDependency=(rowId,index)=>setRows(p=>p.map(r=>r.id===rowId?{...r,dependencies:(r.dependencies||[]).filter((_,i)=>i!==index)}:r));
  const removeRow=id=>setRows(p=>p.filter(r=>r.id!==id));
  const applyTheme=k=>{set({...initialSettings,theme:k,title:settings.title,subtitle:settings.subtitle,projectStart:settings.projectStart,timelineStart:settings.timelineStart,timelineEnd:settings.timelineEnd});setThemeOpen(false)};
  const save=()=>{localStorage.setItem("gantt-studio-mspdi",JSON.stringify({rows,settings}));setSaved(true);setTimeout(()=>setSaved(false),1400)};
  const newProject=()=>{if(confirm("Start a new project?")){setRows(starterRows.map(normalizeRow));setSettings(initialSettings)}};

  const exportMSPDI=()=>{
    const xml=buildMSPDI(rows,settings);
    downloadBlob(new Blob([xml],{type:"application/xml;charset=utf-8"}),"gantt-project.xml");
    setExportOpen(false);
  };
  const renderFullChart=async()=>{
    const node=chartRef.current; if(!node)return null;
    const previous={transform:node.style.transform,width:node.style.width};
    const grid=node.querySelector(".gantt-grid");
    node.style.transform="none";
    node.style.width=`${Math.max(grid?.scrollWidth||0, grid?.getBoundingClientRect().width||0)}px`;
    await new Promise(requestAnimationFrame);
    const data=await toPng(node,{pixelRatio:2,cacheBust:true,backgroundColor:colors.canvas,width:node.scrollWidth,height:node.scrollHeight});
    node.style.transform=previous.transform; node.style.width=previous.width;
    return data;
  };
  const exportPNG=async()=>{const d=await renderFullChart();const r=await fetch(d);downloadBlob(await r.blob(),"gantt-chart.png");setExportOpen(false)};
  const exportPDF=async()=>{const d=await renderFullChart();const im=new Image();im.src=d;await new Promise(r=>im.onload=r);const landscape=im.width>=im.height,doc=new jsPDF({orientation:landscape?"landscape":"portrait",unit:"mm",format:"a4"});const mw=doc.internal.pageSize.getWidth()-16,mh=doc.internal.pageSize.getHeight()-16,s=Math.min(mw/im.width,mh/im.height),w=im.width*s,h=im.height*s;doc.addImage(d,"PNG",(doc.internal.pageSize.getWidth()-w)/2,(doc.internal.pageSize.getHeight()-h)/2,w,h);doc.save("gantt-chart.pdf");setExportOpen(false)};
  const exportExcel=async()=>{
    const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet("Gantt Chart");ws.views=[{showGridLines:false}];
    ws.columns=[{width:32},{width:48},{width:14},{width:14},...visibleUnits.map(()=>({width:13}))];
    ws.mergeCells(1,1,1,visibleUnits.length+4);ws.getCell(1,1).value=settings.title;ws.getCell(1,1).font={bold:true,size:18,color:{argb:hexArgb(colors.text)}};
    ws.mergeCells(2,1,2,visibleUnits.length+4);ws.getCell(2,1).value=settings.subtitle;ws.getCell(2,1).font={italic:true,size:10,color:{argb:hexArgb(colors.muted)}};
    const h=ws.getRow(4);["Initiative","Objective","Start","Finish",...visibleUnits.map(u=>u.label)].forEach((v,i)=>h.getCell(i+1).value=v);
    h.eachCell(c=>{c.fill={type:"pattern",pattern:"solid",fgColor:{argb:hexArgb(colors.header)}};c.font={bold:true,color:{argb:hexArgb(colors.headerText)}};c.alignment={vertical:"middle",horizontal:"center",wrapText:true};c.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}});
    rows.forEach((r,ri)=>{
      const row=ws.getRow(5+ri);
      row.height=64;
      row.getCell(1).value=r.title;
      row.getCell(2).value=r.objective;
      row.getCell(3).value=r.startDate;
      row.getCell(4).value=r.finishDate;
      row.getCell(1).alignment={vertical:"top",wrapText:true};
      row.getCell(2).alignment={vertical:"top",wrapText:true};
      row.getCell(3).alignment={vertical:"top"}; row.getCell(4).alignment={vertical:"top"};
      for(let i=0;i<visibleUnits.length;i++){
        const c=row.getCell(i+5);
        const span=taskUnitSpan(r,visibleUnits);
        if(i>=span.first&&i<=span.last&&span.first>=0){c.value="●";c.fill={type:"pattern",pattern:"solid",fgColor:{argb:hexArgb(colors.bar)}};c.font={color:{argb:hexArgb(colors.barText)},bold:true}}
        c.alignment={horizontal:"center",vertical:"middle"};
        c.border={top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}};
      }
    });
    const b=await wb.xlsx.writeBuffer();downloadBlob(new Blob([b],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),"gantt-chart.xlsx");setExportOpen(false)
  };
  const exportPPT=async()=>{
    const ppt=new pptxgen();ppt.layout="LAYOUT_WIDE";ppt.author="Gantt Studio";const slide=ppt.addSlide();slide.background={color:colors.canvas.slice(1)};
    slide.addText(settings.title,{x:.45,y:.25,w:12.4,h:.35,fontSize:20,bold:true,color:colors.text.slice(1)});
    slide.addText(settings.subtitle,{x:.45,y:.62,w:12.4,h:.25,fontSize:9,color:colors.muted.slice(1)});
    const x=.45,y=1.05,W=12.43,iw=2.35,ow=3.65,tw=W-iw-ow,cw=tw/visibleUnits.length,hh=.55,rh=Math.min(.65,(7.5-y-.35-hh)/Math.max(rows.length,1));
    const cell=(xx,yy,w,h,fill,text,opts={})=>{slide.addShape(ppt.ShapeType.rect,{x:xx,y:yy,w,h,fill:{color:fill.slice(1)},line:{color:colors.grid.slice(1),width:1}});if(text!==undefined)slide.addText(String(text),{x:xx+.04,y:yy+.03,w:w-.08,h:h-.06,fontSize:opts.size||8,color:(opts.color||colors.text).slice(1),bold:!!opts.bold,align:opts.align||"left",valign:"mid",margin:0,fit:"shrink"})};
    cell(x,y,iw,hh,colors.header,"Initiative",{bold:true,color:colors.headerText,align:"center"});cell(x+iw,y,ow,hh,colors.header,"Objective",{bold:true,color:colors.headerText,align:"center"});visibleUnits.forEach((u,i)=>cell(x+iw+ow+i*cw,y,cw,hh,colors.header,u.label,{bold:true,color:colors.headerText,align:"center",size:7}));
    rows.forEach((r,ri)=>{const yy=y+hh+ri*rh;const span=taskUnitSpan(r,visibleUnits);cell(x,yy,iw,rh,colors.canvas,r.title,{size:7});cell(x+iw,yy,ow,rh,colors.canvas,r.objective,{size:6.5});visibleUnits.forEach((u,i)=>cell(x+iw+ow+i*cw,yy,cw,rh,colors.canvas));if(span.first>=0)slide.addShape(ppt.ShapeType.rect,{x:x+iw+ow+span.first*cw+.04,y:yy+.08,w:(span.last-span.first+1)*cw-.08,h:rh-.16,fill:{color:colors.bar.slice(1)},line:{color:colors.bar.slice(1),transparency:100}})});
    await ppt.writeFile({fileName:"gantt-chart.pptx"});setExportOpen(false)
  };
  const importFile=e=>{const f=e.target.files?.[0];if(!f)return;const reader=new FileReader();reader.onload=ev=>{try{const x=parseMSPDI(ev.target.result);setRows(x.rows);setSettings(x.settings)}catch(err){alert("Could not open this file as MSPDI XML: "+err.message)}};reader.readAsText(f);e.target.value=""};

  return <div className="app-shell" style={{"--accent":colors.accent}}>
    <header className="topbar">
      <div className="brand"><div className="brand-mark">G</div><div><div className="brand-name">Gantt Studio</div><div className="brand-sub">MSPDI-compatible project planning in your browser.</div></div></div>
      <div className="top-actions">
        <button className="ghost-btn" onClick={newProject}><Plus size={16}/> New</button>
        <button className="ghost-btn" onClick={save}>{saved?<Check size={16}/>:<Save size={16}/>} {saved?"Saved":"Save"}</button>
        <label className="ghost-btn file-btn"><Upload size={16}/> Open XML<input type="file" accept=".xml,text/xml,application/xml" onChange={importFile}/></label>
        <div className="menu-wrap"><button className="primary-btn" onClick={()=>setExportOpen(v=>!v)}><Download size={16}/> Export <ChevronDown size={15}/></button>
          {exportOpen&&<div className="dropdown export-menu">
            <button onClick={exportMSPDI}><FileCode2 size={16}/> Microsoft Project XML (MSPDI)</button>
            <button onClick={exportPNG}><ImageIcon size={16}/> PNG image</button><button onClick={exportPDF}><FileText size={16}/> PDF</button>
            <button onClick={exportPPT}><Presentation size={16}/> PowerPoint</button><button onClick={exportExcel}><FileSpreadsheet size={16}/> Excel</button>
          </div>}
        </div>
      </div>
    </header>

    <main className="workspace">
      <section className="editor-card">
        <div className="editor-toolbar"><div><div className="eyebrow">MSPDI PROJECT EDITOR</div><h1>Build a chart that looks as good as your plan.</h1><p>The project model maps to Microsoft Project XML: tasks, dates, durations, calendars, resources and assignments.</p></div>
          <div className="toolbar-actions"><div className="menu-wrap"><button className="tool-btn" onClick={()=>setThemeOpen(v=>!v)}><Palette size={16}/> Themes</button>{themeOpen&&<div className="dropdown theme-menu">{Object.entries(themes).map(([k,t])=><button key={k} onClick={()=>applyTheme(k)}><span className="theme-swatch" style={{background:t.bar}}/>{t.name}{settings.theme===k&&<Check size={15}/>}</button>)}</div>}</div>
          <button className="tool-btn" onClick={()=>set({theme:"ocean",barColor:"",barTextColor:"",headerColor:"",headerTextColor:"",backgroundColor:"",gridColor:"",textColor:""})}><RotateCcw size={16}/> Reset style</button>
          <div className="zoom-control"><button onClick={()=>setZoom(z=>Math.max(.75,+(z-.1).toFixed(2)))}><ZoomOut size={15}/></button><span>{Math.round(zoom*100)}%</span><button onClick={()=>setZoom(z=>Math.min(1.35,+(z+.1).toFixed(2)))}><ZoomIn size={15}/></button></div></div>
        </div>

        <div className="design-panel">
          <div className="design-field wide"><label>Chart title</label><input value={settings.title} onChange={e=>set({title:e.target.value})}/></div>
          <div className="design-field wide"><label>Subtitle</label><input value={settings.subtitle} onChange={e=>set({subtitle:e.target.value})}/></div>
          <ColorControl label="Bar" value={colors.bar} onChange={v=>set({barColor:v})}/><ColorControl label="Header" value={colors.header} onChange={v=>set({headerColor:v})}/>
          <ColorControl label="Background" value={colors.canvas} onChange={v=>set({backgroundColor:v})}/><ColorControl label="Grid" value={colors.grid} onChange={v=>set({gridColor:v})}/>
          <div className="design-field"><label>Font size</label><input type="range" min="10" max="18" value={settings.fontSize} onChange={e=>set({fontSize:+e.target.value})}/></div>
          <div className="design-field"><label>Row height</label><input type="range" min="42" max="82" value={settings.rowHeight} onChange={e=>set({rowHeight:+e.target.value})}/></div>
          <label className="check-field"><input type="checkbox" checked={settings.showGrid} onChange={e=>set({showGrid:e.target.checked})}/> Show grid</label>
        </div>

        <div className="chart-scroll"><div ref={chartRef} className="chart" style={{transform:`scale(${zoom})`,transformOrigin:"top left",width:"max-content",minWidth:"100%","--header":colors.header,"--header-text":colors.headerText,"--canvas":colors.canvas,"--grid":colors.grid,"--text":colors.text,"--muted":colors.muted,"--bar":colors.bar,"--bar-text":colors.barText,"--font-size":`${settings.fontSize}px`,"--row-height":`${settings.rowHeight}px`,"--border-width":`${settings.showGrid?settings.borderWidth:0}px`,"--bar-radius":`${settings.barRadius}px`}}>
          <div className="chart-title-row"><div className="chart-title">{settings.title}</div><div className="chart-subtitle">{settings.subtitle}</div></div>
          <div className="gantt-grid" style={{gridTemplateColumns:`minmax(205px,1.05fr) minmax(280px,1.45fr) repeat(${visibleUnits.length},${timelineGranularity(settings)==="month"?"minmax(52px,52px)":timelineGranularity(settings)==="week"?"minmax(46px,46px)":timelineGranularity(settings)==="day"?"minmax(42px,42px)":"minmax(48px,48px)"})`}}><div className="cell head">Initiative</div><div className="cell head">Objective</div>{visibleUnits.map(u=><div className="cell head month-head" key={u.key} title={u.label}>{u.label}</div>)}
          {rows.map(r=><React.Fragment key={r.id}>
            <div className="cell initiative-cell"><div className="drag-handle"><GripVertical size={14}/></div><textarea rows="2" className="cell-input initiative-input" value={r.title} onChange={e=>updateRow(r.id,{title:e.target.value})}/><div className="row-actions"><button onClick={()=>openInitiativeSettings(r.id)} title="Initiative settings"><Settings2 size={13}/></button><button onClick={()=>duplicateRow(r.id)} title="Duplicate"><Copy size={13}/></button><button onClick={()=>removeRow(r.id)} title="Delete"><Trash2 size={13}/></button></div></div>
            <div className="cell objective-cell"><textarea className="cell-input objective-input" value={r.objective} onChange={e=>updateRow(r.id,{objective:e.target.value})}/></div>
            {visibleUnits.map((u,i)=>{
              const {start:taskStart,finish:taskFinish}=getTaskDates(r,settings);
              const active=taskStart<=u.end&&taskFinish>=u.start;
              const preview=timelineDrag?.rowId===r.id && i>=Math.min(timelineDrag.anchor,timelineDrag.current) && i<=Math.max(timelineDrag.anchor,timelineDrag.current);
              return <div key={u.key} className={`cell month-cell ${active?"active-cell":""} ${preview?"timeline-preview":""}`} title={timelineDrag?.rowId===r.id ? "Hover to preview the range, then click to set the finish" : "Click to set the start date, then hover and click again to set the finish date"} onClick={()=>beginTimelineSelection(r.id,i)} onMouseEnter={()=>updateTimelineSelection(r.id,i)}>{active&&<div className="bar-block"/>}{preview&&<div className="preview-block"/>}</div>
            })}
          </React.Fragment>)}</div>
        </div></div>
        <div className="chart-footer"><button className="add-row-btn" onClick={addRow}><Plus size={16}/> Add initiative</button><span>{timelineDrag ? "Start selected — move your pointer to preview the timeframe, then click to set the finish." : "Click a timeline cell to set the start, hover to preview, then click again to set the finish."}</span><button className="text-action" onClick={cancelTimelineSelection} disabled={!timelineDrag}>Cancel selection</button></div>
      </section>

      <aside className="side-panel"><div className="side-heading"><div><div className="eyebrow">STYLE</div><h2>Appearance</h2></div><Settings2 size={19}/></div>
        <div className="side-section"><h3>Timeline</h3><div className="timeline-mode"><span>Auto scale</span><strong>{timelineGranularity(settings)==="hour"?"Hours":timelineGranularity(settings)==="day"?"Days":timelineGranularity(settings)==="week"?"Weeks":"Months"}</strong></div><div className="timeline-settings"><label>Timeline start<input type="date" value={settings.timelineStart} onChange={e=>set({timelineStart:e.target.value})}/></label><label>Timeline end<input type="date" value={settings.timelineEnd} min={settings.timelineStart} onChange={e=>set({timelineEnd:e.target.value})}/></label></div><p className="side-note">The scale automatically switches to hours for very short plans, days for short plans, weeks for medium plans, and months for long plans.</p></div>
        <div className="side-section"><h3>Preset themes</h3><div className="theme-grid">{Object.entries(themes).map(([k,t])=><button className={`theme-card ${settings.theme===k?"selected":""}`} key={k} onClick={()=>applyTheme(k)}><div className="mini-theme"><div style={{background:t.header}}/><div style={{background:t.bar}}/><div style={{background:t.canvas}}/></div><span>{t.name}</span></button>)}</div></div>
        <div className="side-section"><h3>Chart styling</h3><div className="side-row"><span>Border width</span><select value={settings.borderWidth} onChange={e=>set({borderWidth:+e.target.value})}><option value="0">None</option><option value="1">1 px</option><option value="2">2 px</option></select></div><div className="side-row"><span>Bar corners</span><select value={settings.barRadius} onChange={e=>set({barRadius:+e.target.value})}><option value="0">Square</option><option value="2">Subtle</option><option value="7">Rounded</option><option value="14">Pill</option></select></div><div className="side-row"><span>Grid lines</span><button className={`toggle ${settings.showGrid?"on":""}`} onClick={()=>set({showGrid:!settings.showGrid})}><span/></button></div></div>
        <div className="side-section"><h3>Color overrides</h3><div className="color-list"><ColorControl label="Bar text" value={colors.barText} onChange={v=>set({barTextColor:v})}/><ColorControl label="Header text" value={colors.headerText} onChange={v=>set({headerTextColor:v})}/><ColorControl label="Body text" value={colors.text} onChange={v=>set({textColor:v})}/></div></div>
        <div className="privacy-card"><div className="privacy-icon">✓</div><div><strong>Browser-only processing</strong><p>MSPDI XML is generated and parsed locally. No project data needs to be uploaded.</p></div></div>
      </aside>
    </main>

    {initiativeSettingsId!==null && (()=>{
      const initiative=rows.find(r=>r.id===initiativeSettingsId);
      if(!initiative)return null;
      const dependencies=initiative.dependencies||[];
      return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)closeInitiativeSettings()}}>
        <div className="initiative-modal" role="dialog" aria-modal="true">
          <div className="modal-header"><div><div className="eyebrow">INITIATIVE SETTINGS</div><h2>{initiative.title || "Untitled initiative"}</h2><p>Configure the initiative and all task relationships without crowding the Gantt chart.</p></div><button className="modal-close" onClick={closeInitiativeSettings}>×</button></div>
          <div className="modal-tabs"><button className={initiativeSettingsTab==="general"?"active":""} onClick={()=>setInitiativeSettingsTab("general")}>Initiative</button><button className={initiativeSettingsTab==="relationships"?"active":""} onClick={()=>setInitiativeSettingsTab("relationships")}>Relationships <span>{dependencies.length}</span></button></div>
          {initiativeSettingsTab==="general" ? <div className="modal-body general-tab"><label>Initiative name<input value={initiative.title} onChange={e=>updateRow(initiative.id,{title:e.target.value})}/></label><label>Objective<textarea rows="5" value={initiative.objective} onChange={e=>updateRow(initiative.id,{objective:e.target.value})}/></label><div className="schedule-fields"><label>Start date & time<input type="datetime-local" value={taskInputDate(initiative,"start")} onChange={e=>{const v=e.target.value;updateRow(initiative.id,{startDate:v.slice(0,10),startDateTime:v,start:Math.max(0,monthIndexFor(new Date(`${v.slice(0,10)}T00:00:00Z`),startOfMonth(settings.timelineStart)))})}}/></label><label>Finish date & time<input type="datetime-local" value={taskInputDate(initiative,"finish")} min={taskInputDate(initiative,"start")} onChange={e=>{const v=e.target.value;updateRow(initiative.id,{finishDate:v.slice(0,10),finishDateTime:v})}}/></label></div></div> : <div className="modal-body relationships-tab">
            <div className="relationship-intro"><div><strong>Task relationships</strong><p>Each relationship links this initiative to a predecessor. MSPDI supports FF, FS, SF and SS links, with optional lag.</p></div><button className="primary-btn" onClick={()=>addDependency(initiative.id)} disabled={rows.filter(x=>x.id!==initiative.id).length<=dependencies.length}><Plus size={15}/> Add relationship</button></div>
            {dependencies.length===0 ? <div className="empty-relationships"><Settings2 size={22}/><strong>No relationships</strong><span>Add a predecessor relationship to control how this initiative connects to another task.</span></div> : <div className="relationship-list">{dependencies.map((d,index)=><div className="relationship-card" key={`${initiative.id}-${index}`}><div className="relationship-number">{index+1}</div><div className="relationship-fields"><label>Predecessor<select value={d.predecessorId||""} onChange={e=>updateDependency(initiative.id,index,{predecessorId:e.target.value})}><option value="">Select predecessor</option>{rows.filter(x=>x.id!==initiative.id).map((p,idx)=><option key={p.id} value={p.id}>{idx+1}. {p.title}</option>)}</select></label><label>Relationship<select value={d.dependencyType||"FS"} onChange={e=>updateDependency(initiative.id,index,{dependencyType:e.target.value})}>{Object.entries(LINK_LABELS).map(([k,v])=><option key={k} value={k}>{k} — {v}</option>)}</select></label><label>Lag (days)<input type="number" min="-365" max="365" value={d.lag||0} onChange={e=>updateDependency(initiative.id,index,{lag:Number(e.target.value)})}/></label></div><button className="icon-danger" onClick={()=>removeDependency(initiative.id,index)} title="Remove relationship"><Trash2 size={15}/></button></div>)}</div>}
            <div className="relationship-help"><strong>Relationship types</strong><div className="relationship-type-grid">{Object.entries(LINK_LABELS).map(([k,v])=><div key={k}><b>{k}</b><span>{v}</span></div>)}</div></div>
          </div>}
          <div className="modal-footer"><button className="ghost-btn" onClick={closeInitiativeSettings}>Close</button></div>
        </div>
      </div>;
    })()}
  </div>
}

function ColorControl({label,value,onChange}){
  return <div className="design-field color-control"><label>{label}</label><div className="color-input-wrap"><input type="color" value={value} onChange={e=>onChange(e.target.value)}/><input type="text" value={value} onChange={e=>onChange(e.target.value)}/></div></div>
}
export default App;
