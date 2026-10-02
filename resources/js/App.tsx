import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Screen = "dashboard" | "employees" | "attendance" | "reports" | "overtime" | "print" | "settings";
type Status = "hadir" | "telat" | "tidak-hadir" | "belum-absen";
type WorkShift = { name:string; start:string; end:string };

const workShifts:WorkShift[]=[
  {name:"10:00 - 22:00",start:"10:00",end:"22:00"},
  {name:"10:30 - 20:30",start:"10:30",end:"20:30"},
  {name:"11:00 - 21:00",start:"11:00",end:"21:00"},
  {name:"11:30 - 23:30",start:"11:30",end:"23:30"},
  {name:"15:00 - 23:30",start:"15:00",end:"23:30"},
  {name:"14:00 - 24:00",start:"14:00",end:"24:00"},
  {name:"16:00 - 02:00",start:"16:00",end:"02:00"},
];

type Employee = { id:string; name:string; dept:string; status:Status; checkIn:string; checkOut:string; avatar:string };
type Attendance = { key:string; employeeId:string; name:string; dept:string; date:string; month:string; sheet:string; shift?:string; checkIn:string; checkOut:string; status:Status; late:number; overtime:number; location:string };

const mockEmployees:Employee[]=[
 {id:"EMP001",name:"Budi Santoso",dept:"Engineering",status:"hadir",checkIn:"08:02",checkOut:"17:02",avatar:"BS"},
 {id:"EMP002",name:"Sari Dewi",dept:"Marketing",status:"telat",checkIn:"08:47",checkOut:"17:30",avatar:"SD"},
 {id:"EMP003",name:"Rizky Pratama",dept:"Finance",status:"tidak-hadir",checkIn:"-",checkOut:"-",avatar:"RP"},
 {id:"EMP004",name:"Anita Wulandari",dept:"HR",status:"hadir",checkIn:"07:58",checkOut:"18:15",avatar:"AW"},
];
const mockAttendance:Attendance[] = mockEmployees.map((e,i)=>({key:`mock-${e.id}`,employeeId:e.id,name:e.name,dept:e.dept,date:"23 September 2026",month:"September 2026",sheet:"Contoh",checkIn:e.checkIn,checkOut:e.checkOut,status:e.status,late:e.status==="telat"?47:0,overtime:i===3?75:i===1?30:0,location:"Kantor Pusat"}));

const statusLabel=(s:Status)=>({hadir:"Hadir",telat:"Telat","tidak-hadir":"Tidak Hadir","belum-absen":"Belum Absen"}[s]);
const statusClass=(s:Status)=>({hadir:"bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-300 dark:border-emerald-900",telat:"bg-orange-50 text-orange-700 border-orange-200 dark:bg-orange-950/40 dark:text-orange-300 dark:border-orange-900","tidak-hadir":"bg-red-50 text-red-700 border-red-200 dark:bg-red-950/40 dark:text-red-300 dark:border-red-900","belum-absen":"bg-slate-100 text-slate-600 border-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:border-slate-700"}[s]);
const initials=(name:string)=>name.split(" ").map(x=>x[0]).join("").slice(0,2).toUpperCase();
const normalizeKey=(v:any)=>String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]/g,"");
const monthNames=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const monthAliases=[
  ["januari","january","jan"],["februari","february","feb"],["maret","march","mar"],
  ["april","apr"],["mei","may"],["juni","june","jun"],["juli","july","jul"],
  ["agustus","august","agu","agt","aug"],["september","sep"],["oktober","october","okt","oct"],
  ["november","nov"],["desember","december","des","dec"]
];
type DateParts={year:number;month:number;day:number};
type SheetRow=Record<string,any>&{__sheet:string;__sheetMonth:string;__sheetYear:number};

const validDateParts=(year:number,month:number,day:number):DateParts|null=>{
  const date=new Date(year,month-1,day);
  return month>=1&&month<=12&&day>=1&&date.getFullYear()===year&&date.getMonth()===month-1&&date.getDate()===day?{year,month,day}:null;
};

const parseDateParts=(value:any):DateParts|null=>{
  if(value instanceof Date&&!isNaN(value.getTime())) return validDateParts(value.getFullYear(),value.getMonth()+1,value.getDate());
  if(typeof value==="number"&&Number.isFinite(value)){
    const parsed=XLSX.SSF.parse_date_code(value);
    return parsed?validDateParts(parsed.y,parsed.m,parsed.d):null;
  }
  const text=String(value??"").trim();
  let match=text.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if(match) return validDateParts(Number(match[1]),Number(match[2]),Number(match[3]));
  match=text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if(match){
    const year=Number(match[3])<100?2000+Number(match[3]):Number(match[3]);
    return validDateParts(year,Number(match[2]),Number(match[1]));
  }
  const normalized=normalizeKey(text);
  match=normalized.match(/^(\d{1,2})([a-z]+)(\d{4})$/);
  if(match){
    const month=monthAliases.findIndex(aliases=>aliases.some(alias=>alias===match![2]||alias.startsWith(match![2])));
    if(month>=0) return validDateParts(Number(match[3]),month+1,Number(match[1]));
  }
  const parsed=Date.parse(text);
  if(!isNaN(parsed)){
    const date=new Date(parsed);
    return validDateParts(date.getFullYear(),date.getMonth()+1,date.getDate());
  }
  return null;
};

const periodFromText=(value:any):{month:number;year:number}|null=>{
  if(value instanceof Date&&!isNaN(value.getTime())) return {month:value.getMonth()+1,year:value.getFullYear()};
  const text=String(value??"").trim();
  if(!text) return null;
  let match=text.match(/\b(19\d{2}|20\d{2})[./_-](0?[1-9]|1[0-2])\b/);
  if(match) return {year:Number(match[1]),month:Number(match[2])};
  match=text.match(/\b(0?[1-9]|1[0-2])[./_-](19\d{2}|20\d{2})\b/);
  if(match) return {month:Number(match[1]),year:Number(match[2])};
  match=text.match(/\b(0?[1-9]|1[0-2])\b[^a-z0-9]+(19\d{2}|20\d{2})\b/);
  if(match) return {month:Number(match[1]),year:Number(match[2])};
  const lower=text.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
  const year=Number(text.match(/(?:19|20)\d{2}/)?.[0]||0);
  const aliases=monthAliases.flat().sort((a,b)=>b.length-a.length).join("|");
  const monthMatch=lower.match(new RegExp("(?:^|[^a-z])("+aliases+")(?![a-z])"));
  const month=monthMatch?monthAliases.findIndex(group=>group.includes(monthMatch[1])):-1;
  if(month>=0) return {month:month+1,year:year||new Date().getFullYear()};
  const normalized=normalizeKey(text);
  match=normalized.match(/^(0?[1-9]|1[0-2])$/);
  if(match) return {month:Number(match[1]),year:new Date().getFullYear()};
  return null;
};

const dateFromDay=(dayValue:any,period:{month:number;year:number}|null):DateParts|null=>{
  if(typeof dayValue==="number"&&Number.isInteger(dayValue)&&dayValue>=1&&dayValue<=31&&period){
    return validDateParts(period.year,period.month,dayValue);
  }
  const text=String(dayValue??"").trim();
  if(/^\d{1,2}$/.test(text)&&period) return validDateParts(period.year,period.month,Number(text));
  return parseDateParts(dayValue);
};

const formatDate=(parts:DateParts|null)=>parts?`${String(parts.day).padStart(2,"0")} ${monthNames[parts.month-1]} ${parts.year}`:"-";
const formatMonth=(parts:DateParts|null,period:{month:number;year:number}|null=null)=>{
  const month=parts?.month||period?.month;
  const year=parts?.year||period?.year;
  return month?`${monthNames[month-1]}${year?` ${year}`:""}`:"-";
};

const monthIndex=(m:string)=>{
  const period=periodFromText(m);
  return period?period.year*100+period.month:0;
};
const sortableDate=(value:any)=>{
  const parts=parseDateParts(value);
  return parts?new Date(parts.year,parts.month-1,parts.day).getTime():0;
};
const monthFromDate=(time:number)=>{
  if(!time) return "";
  return formatMonth(parseDateParts(new Date(time)));
};
const formatDuration=(minutes:number)=>{
  if(!Number.isFinite(minutes)||minutes<=0) return "-";
  const hours=Math.floor(minutes/60);
  const remainder=Math.round(minutes%60);
  if(hours&&remainder) return `+${hours} jam ${remainder} menit`;
  if(hours) return `+${hours} jam`;
  return `+${remainder} menit`;
};

// Ubah nilai apa pun (jam Excel, teks "8.30", "0830", "8:30 PM", dsb.) menjadi menit sejak 00:00.
// Mengembalikan null bila tidak ada jam valid — dipakai agar pembacaan jam lebih akurat.
const timeToMinutes=(value:any):number|null=>{
  if(value===null||value===undefined||value==="") return null;
  if(value instanceof Date&&!isNaN(value.getTime())) return value.getHours()*60+value.getMinutes();
  if(typeof value==="number"&&Number.isFinite(value)){
    let fraction=value;
    if(value>=1&&value<86400&&Number.isInteger(value)){
      // Beberapa ekspor menyimpan jam sebagai detik sejak tengah malam.
      return Math.floor((value%86400)/60);
    }
    fraction=((value%1)+1)%1;
    const minutes=Math.round(fraction*1440)%1440;
    return minutes;
  }
  const text=String(value).trim();
  if(!text) return null;
  const meridiem=text.match(/(am|pm)/i)?.[1]?.toLowerCase();
  let match=text.match(/(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?/);
  if(match){
    let hour=Number(match[1]);
    if(meridiem==="pm"&&hour<12) hour+=12;
    if(meridiem==="am"&&hour===12) hour=0;
    if(hour>=24) hour%=24;
    return hour*60+Number(match[2]);
  }
  match=text.match(/^(\d{1,2})$/);
  if(match){
    let hour=Number(match[1]);
    if(meridiem==="pm"&&hour<12) hour+=12;
    if(meridiem==="am"&&hour===12) hour=0;
    return hour>=0&&hour<=23?hour*60:null;
  }
  match=text.match(/^(\d{3,4})$/);
  if(match){
    const digits=match[1].padStart(4,"0");
    const hour=Number(digits.slice(0,2));
    const minute=Number(digits.slice(2));
    if(hour<=23&&minute<=59) return hour*60+minute;
  }
  return null;
};


const Icons: Record<string, React.ComponentType<{ className?: string }>> = {
 Home:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="m3 10 9-7 9 7v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M9 22V12h6v10"/></svg>,
 Users:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8"/></svg>,
 Calendar:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>,
 File:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h6"/></svg>,
 Clock:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>,
 Search:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></svg>,
 Upload:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M12 16V4M7 9l5-5 5 5M4 20h16"/></svg>,
 Trash:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M3 6h18M8 6V4h8v2M19 6l-1 15H6L5 6M10 11v6M14 11v6"/></svg>,
 Download:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M12 3v12m0 0 5-5m-5 5-5-5M4 21h16"/></svg>,
 Settings:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1-1.5 1.4-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.5V20h-2v-.5a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.9.3l-.1.1-1.4-1.4.1-.1A1.7 1.7 0 0 0 7.5 15a1.7 1.7 0 0 0-1.5-1H5v-2h1a1.7 1.7 0 0 0 1.5-1 1.7 1.7 0 0 0-.3-1.9l-.1-.1 1.4-1.4.1.1a1.7 1.7 0 0 0 1.9.3 1.7 1.7 0 0 0 1-1.5V6h2v.5a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.9-.3l.1-.1 1.5 1.4-.1.1a1.7 1.7 0 0 0-.3 1.9c.2.6.8 1 1.5 1h.5v2h-.5c-.7 0-1.3.4-1.5 1z"/></svg>,
 Moon:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M21 12.8A8.5 8.5 0 1 1 11.2 3 6.6 6.6 0 0 0 21 12.8z"/></svg>,
 Sun:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>,
 LogOut:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/></svg>,
 Menu:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="M4 6h16M4 12h16M4 18h16"/></svg>,
 X:(p:any)=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" {...p}><path d="m6 6 12 12M18 6 6 18"/></svg>,
};

function Logo(){return <div className="w-10 h-10 rounded-xl bg-blue-600 text-white grid place-items-center font-black text-lg shadow-lg shadow-blue-600/20">A</div>}
function Page({children}:{children:React.ReactNode}){return <main className="p-4 sm:p-6 max-w-[1500px] mx-auto">{children}</main>}
function Card({children,className=""}:{children:React.ReactNode;className?:string}){return <div className={`bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl shadow-sm ${className}`}>{children}</div>}
function Button({children,onClick,variant="primary",disabled=false}:{children:React.ReactNode;onClick?:()=>void;variant?:"primary"|"danger"|"ghost";disabled?:boolean}){const c=variant==="primary"?"bg-blue-600 hover:bg-blue-700 text-white":variant==="danger"?"bg-red-600 hover:bg-red-700 text-white":"bg-slate-100 hover:bg-slate-200 text-slate-700 dark:bg-slate-800 dark:hover:bg-slate-700 dark:text-slate-200";return <button disabled={disabled} onClick={onClick} className={`px-3.5 py-2 rounded-xl text-sm font-semibold transition disabled:opacity-50 ${c}`}>{children}</button>}

function Login({onLogin}:{onLogin:()=>void}){const [email,setEmail]=useState("");const [password,setPassword]=useState("");return <div className="min-h-screen bg-slate-100 dark:bg-slate-950 grid place-items-center p-5"><Card className="w-full max-w-md p-7"><div className="flex items-center gap-3 mb-7"><Logo/><div><h1 className="font-black text-xl text-slate-900 dark:text-white">AbsenPro</h1><p className="text-xs text-slate-500">Admin / HRD</p></div></div><h2 className="text-2xl font-bold text-slate-900 dark:text-white">Masuk ke Dashboard</h2><p className="text-sm text-slate-500 mt-1 mb-6">Kelola data karyawan dan absensi.</p><label className="text-sm font-semibold text-slate-700 dark:text-slate-300">Email</label><input value={email} onChange={e=>setEmail(e.target.value)} placeholder="admin@perusahaan.com" className="mt-2 mb-4 w-full px-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-950 text-slate-900 dark:text-white outline-none focus:ring-2 focus:ring-blue-500/20"/><label className="text-sm font-semibold text-slate-700 dark:text-slate-300">Password</label><input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" className="mt-2 mb-6 w-full px-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-950 text-slate-900 dark:text-white outline-none focus:ring-2 focus:ring-blue-500/20"/><Button onClick={onLogin}>Masuk sebagai Admin / HRD</Button></Card></div>}

function Sidebar({screen,go,logout,open,setOpen,dark,toggle}:{screen:Screen;go:(s:Screen)=>void;logout:()=>void;open:boolean;setOpen:(v:boolean)=>void;dark:boolean;toggle:()=>void}){const nav:Array<[Screen, string, React.ComponentType<{ className?: string }>]>=[['dashboard','Dashboard',Icons.Home],['employees','Data Karyawan',Icons.Users],['attendance','Data Absensi',Icons.Calendar],['reports','Laporan',Icons.File],['overtime','Lembur & Telat',Icons.Clock],['print','Cetak Rekap',Icons.Download],['settings','Pengaturan',Icons.Settings]];return <><aside className={`fixed lg:sticky top-0 left-0 z-40 h-screen w-64 bg-slate-950 text-white p-4 transition-transform ${open?'translate-x-0':'-translate-x-full lg:translate-x-0'}`}><div className="flex items-center gap-3 px-2 py-3"><Logo/><div><div className="font-black">AbsenPro</div><div className="text-xs text-slate-400">Admin / HRD</div></div><button className="ml-auto lg:hidden" onClick={()=>setOpen(false)}><Icons.X className="w-5"/></button></div><nav className="mt-7 space-y-1">{nav.map(([id,label,Icon])=><button key={id} onClick={()=>{go(id);setOpen(false)}} className={`w-full flex items-center gap-3 px-3 py-3 rounded-xl text-sm ${screen===id?'bg-blue-600 text-white':'text-slate-300 hover:bg-slate-800'}`}><Icon className="w-5 h-5"/>{label}</button>)}</nav><div className="absolute bottom-4 left-4 right-4 space-y-2"><button onClick={toggle} className="w-full flex items-center gap-3 px-3 py-3 rounded-xl text-sm text-slate-300 hover:bg-slate-800">{dark?<Icons.Sun className="w-5"/>:<Icons.Moon className="w-5"/>}{dark?'Mode Light':'Mode Dark'}</button><button onClick={logout} className="w-full flex items-center gap-3 px-3 py-3 rounded-xl text-sm text-red-300 hover:bg-red-950/40"><Icons.LogOut className="w-5"/>Logout</button></div></aside>{open&&<div onClick={()=>setOpen(false)} className="fixed inset-0 bg-black/50 z-30 lg:hidden"/>}</>}
function Header({screen,menu,dark,toggle}:{screen:Screen;menu:()=>void;dark:boolean;toggle:()=>void}){const titles:Record<Screen,string>={dashboard:'Dashboard',employees:'Data Karyawan',attendance:'Data Absensi',reports:'Laporan',overtime:'Lembur & Telat',print:'Cetak Rekap',settings:'Pengaturan'};return <header className="h-16 sticky top-0 z-20 bg-white/90 dark:bg-slate-950/90 backdrop-blur border-b border-slate-200 dark:border-slate-800 flex items-center px-4 sm:px-6 gap-3"><button onClick={menu} className="lg:hidden"><Icons.Menu className="w-5 text-slate-700 dark:text-slate-200"/></button><div><h1 className="font-bold text-slate-900 dark:text-white">{titles[screen]}</h1><p className="text-xs text-slate-500 hidden sm:block">Manajemen Sistem Absensi</p></div><div className="ml-auto flex items-center gap-3"><button onClick={toggle} className="p-2 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-200">{dark?<Icons.Sun className="w-5"/>:<Icons.Moon className="w-5"/>}</button><span className="hidden sm:block text-xs font-semibold text-slate-500">Admin / HRD</span></div></header>}

function Dashboard({employees,attendance,go}:{employees:Employee[];attendance:Attendance[];go:(s:Screen)=>void}){const late=attendance.filter(a=>a.status==='telat').length;const overtime=attendance.reduce((n,a)=>n+a.overtime,0);return <Page><div className="mb-6"><h2 className="text-2xl font-black text-slate-900 dark:text-white">Selamat datang 👋</h2><p className="text-sm text-slate-500">Ringkasan kondisi absensi saat ini.</p></div><div className="grid sm:grid-cols-2 xl:grid-cols-4 gap-4 mb-6">{[["Total Karyawan",employees.length,"text-blue-600"],["Hadir",attendance.filter(a=>a.status==='hadir').length,"text-emerald-600"],["Telat",late,"text-orange-600"],["Total Lembur",formatDuration(overtime),"text-purple-600"]].map(([a,b,c])=><Card key={String(a)} className="p-5"><div className="text-sm text-slate-500">{a}</div><div className={`text-3xl font-black mt-2 ${c}`}>{b}</div></Card>)}</div><Card><div className="p-5 flex items-center justify-between border-b border-slate-100 dark:border-slate-800"><div><h3 className="font-bold text-slate-900 dark:text-white">Absensi Terbaru</h3><p className="text-xs text-slate-500">Tanggal, jam masuk, telat dan lembur.</p></div><Button variant="ghost" onClick={()=>go('attendance')}>Lihat semua</Button></div><div className="overflow-x-auto"><table className="w-full"><thead><tr className="text-left text-xs text-slate-500 bg-slate-50 dark:bg-slate-800/60"><th className="px-5 py-3">Nama</th><th>Tanggal</th><th>Masuk</th><th>Telat</th><th>Lembur</th></tr></thead><tbody className="divide-y divide-slate-100 dark:divide-slate-800">{attendance.slice(0,6).map(a=><tr key={a.key} className="text-sm"><td className="px-5 py-3 font-semibold text-slate-800 dark:text-slate-100">{a.name}</td><td>{a.date}</td><td>{a.checkIn}</td><td>{a.late?`${a.late} menit`:'-'}</td><td>{formatDuration(a.overtime)}</td></tr>)}</tbody></table></div></Card></Page>}

function Employees({employees,setEmployees,attendance,setAttendance}:{employees:Employee[];setEmployees:React.Dispatch<React.SetStateAction<Employee[]>>;attendance:Attendance[];setAttendance:React.Dispatch<React.SetStateAction<Attendance[]>>}){const [q,setQ]=useState("");const [selected,setSelected]=useState<string[]>([]);const filtered=employees.filter(e=>`${e.id} ${e.name} ${e.dept}`.toLowerCase().includes(q.toLowerCase()));const del=(id:string)=>{setEmployees(es=>es.filter(e=>e.id!==id));setAttendance(as=>as.filter(a=>a.employeeId!==id));setSelected(s=>s.filter(x=>x!==id))};return <Page><div className="flex flex-col sm:flex-row gap-3 justify-between mb-5"><div><h2 className="text-xl font-bold text-slate-900 dark:text-white">Data Karyawan</h2><p className="text-sm text-slate-500">Kelola data karyawan.</p></div><div className="flex gap-2"><Button variant="danger" disabled={!selected.length} onClick={()=>selected.forEach(del)}>Hapus Terpilih</Button></div></div><div className="mb-4 relative max-w-md"><Icons.Search className="absolute left-3 top-3.5 w-4 text-slate-400"/><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Cari nama, ID, departemen..." className="w-full pl-10 pr-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-white"/></div><Card className="overflow-hidden"><div className="overflow-x-auto"><table className="w-full"><thead><tr className="bg-slate-50 dark:bg-slate-800/60 text-xs text-slate-500 text-left"><th className="p-4"><input type="checkbox" checked={filtered.length>0&&filtered.every(e=>selected.includes(e.id))} onChange={e=>setSelected(e.target.checked?filtered.map(x=>x.id):[])}/></th><th>Nama</th><th>ID</th><th>Departemen</th><th>Status</th><th>Aksi</th></tr></thead><tbody className="divide-y divide-slate-100 dark:divide-slate-800">{filtered.map(e=><tr key={e.id} className="text-sm"><td className="p-4"><input type="checkbox" checked={selected.includes(e.id)} onChange={x=>setSelected(s=>x.target.checked?[...s,e.id]:s.filter(id=>id!==e.id))}/></td><td className="font-semibold text-slate-900 dark:text-white">{e.name}</td><td className="font-mono text-xs">{e.id}</td><td>{e.dept}</td><td><span className={`px-2.5 py-1 rounded-full border text-xs ${statusClass(e.status)}`}>{statusLabel(e.status)}</span></td><td><Button variant="danger" onClick={()=>del(e.id)}><Icons.Trash className="w-4"/></Button></td></tr>)}</tbody></table></div></Card></Page>}

function AttendancePage({attendance,setAttendance,setEmployees}:{attendance:Attendance[];setAttendance:React.Dispatch<React.SetStateAction<Attendance[]>>;setEmployees:React.Dispatch<React.SetStateAction<Employee[]>>}){
  const [q,setQ]=useState("");
  const [month,setMonth]=useState("Semua");
  const [sheet,setSheet]=useState("Semua");
  const [shift,setShift]=useState("Semua");
  const [selected,setSelected]=useState<string[]>([]);
  const [page,setPage]=useState(1);
  const [importNotice,setImportNotice]=useState<{type:"success"|"error";message:string}|null>(null);
  const [detail,setDetail]=useState<Attendance|null>(null);

  const parseTime=(value:any):string=>{
    const minutes=timeToMinutes(value);
    if(minutes===null) return "-";
    return String(Math.floor(minutes/60)).padStart(2,"0")+":"+String(minutes%60).padStart(2,"0");
  };
  const num=(value:any)=>{
    if(value===null||value===undefined||value==="") return 0;
    // Nilai pecahan (0..1) biasanya durasi jam/hari Excel — konversi ke menit.
    if(typeof value==="number"&&Number.isFinite(value)) return value>0&&value<1?Math.round(value*1440):Math.round(value);
    const text=String(value).trim().toLowerCase();
    const hours=Number(text.match(/(\d+(?:[.,]\d+)?)\s*(?:jam|hours?|hrs?|h)\b/)?.[1]?.replace(",",".")||0);
    const minutes=Number(text.match(/(\d+(?:[.,]\d+)?)\s*(?:menit|minutes?|mins?|m)\b/)?.[1]?.replace(",",".")||0);
    if(hours||minutes) return Math.round(hours*60+minutes);
    const clock=text.match(/^(\d{1,2}):(\d{2})$/);
    if(clock) return Number(clock[1])*60+Number(clock[2]);
    const number=Number(text.replace(",","."));
    return Number.isFinite(number)?Math.round(number):0;
  };
  const stableEmployeeId=(name:string)=>"EMP-"+(normalizeKey(name).slice(0,18).toUpperCase()||"TANPA-NAMA");
  const attendanceStatus=(raw:any,checkIn:string,checkOut:string):Status=>{
    const value=normalizeKey(raw);
    if(value.includes("belum")) return "belum-absen";
    if(value.includes("telat")||value.includes("terlambat")||value==="12"||value==="05") return "telat";
    if(value.includes("tidak")||value.includes("absen")||value.includes("alpa")||value==="alpha"||value==="a"||value==="0"||value.includes("izin")||value.includes("ijin")||value.includes("sakit")) return "tidak-hadir";
    if(value.includes("hadir")||value.includes("masuk")||value==="h"||value==="1"||checkIn!=="-"||checkOut!=="-") return "hadir";
    return "belum-absen";
  };
  const headerAliases:Record<string,string[]> = {
    date:["tanggal","tanggal absen","tanggal absensi","tgl","date","attendance date","absen date","dated"],
    month:["bulan","month","periode","period"],
    year:["tahun","year"],
    name:["nama","nama karyawan","karyawan","employee name","name","employee","full name"],
    id:["id","id karyawan","nik","employee id","empid","nip","id pegawai","kode karyawan"],
    dept:["departemen","department","dept","unit","bagian","division"],
    checkIn:["jam masuk","masuk","check in","in time","clock in","in"],
    checkOut:["jam pulang","pulang","check out","out time","clock out","out"],
    status:["status","attendance status","ket","keterangan","kehadiran"],
    late:["terlambat","keterlambatan","late","minutes late","telat","delay"],
    overtime:["lembur","overtime","jam lembur","minutes overtime","ot"],
    shift:["shift","shift name","nama shift","shift kerja","regu"],
    location:["lokasi","location","address","alamat"]
  };

  const resolveHeaderField=(value:any):string|undefined=>{
    const normalized=normalizeKey(value);
    for(const [field,aliases] of Object.entries(headerAliases)){
      if(aliases.some(alias=>{
        const key=normalizeKey(alias);
        return normalized===key||(key.length>=5&&normalized.includes(key));
      })) return field;
    }
    return undefined;
  };

  const collectSheetRows=(workbook:any):SheetRow[]=>{
    const rows:SheetRow[]=[];
    const stats={sheets:0,sheetsRead:0,sheetsSkipped:0,raw:0,invalid:0,deduped:0,duplicateConflicts:0,nightWrapped:0,clockValues:0};
    (workbook.SheetNames as string[]).forEach((sheetName:string)=>{
      const worksheet=workbook.Sheets[sheetName];
      if(!worksheet) return;
      stats.sheets++;
      const sheetRowStart=rows.length;
      const matrix=XLSX.utils.sheet_to_json<any[]>(worksheet,{header:1,raw:true,blankrows:false,defval:""});
      const topText=matrix.slice(0,30).flat().filter(value=>String(value??"").trim()!=="").join(" ");
      const period=periodFromText(sheetName+" "+topText);

      // Fingerprint reports store each employee in a block, with days across columns.
      const isFingerprintReport=matrix.slice(0,10).some(row=>(row||[]).some((value:any)=>normalizeKey(value).includes("employeeattendancerecord")));
      const userRows:number[]=[];
      if(isFingerprintReport){
        matrix.forEach((row:any[],rowIndex:number)=>{
          if((row||[]).some((value:any)=>normalizeKey(value)==="userid")) userRows.push(rowIndex);
        });
      }
      if(isFingerprintReport&&userRows.length&&period){
        const minutesOf=(time:string)=>{
          const [hour,minute]=time.split(":").map(Number);
          return hour*60+minute;
        };
        const shiftWindow=(date:DateParts,shift:WorkShift)=>{
          const midnight=new Date(date.year,date.month-1,date.day).getTime();
          const start=minutesOf(shift.start);
          const end=minutesOf(shift.end);
          const endOffset=end<=start?end+1440:end;
          return {start:midnight+start*60000,end:midnight+endOffset*60000};
        };
        const nearestShift=(timestamp:number,date:DateParts,edge:"start"|"end")=>{
          const candidates=workShifts.map(candidate=>{
            const expected=shiftWindow(date,candidate)[edge];
            return {shift:candidate,expected,difference:Math.abs(timestamp-expected)/60000};
          });
          if(edge==="start"){
            const upcoming=candidates.filter(candidate=>candidate.expected>=timestamp&&candidate.difference<=30).sort((a,b)=>a.difference-b.difference)[0];
            if(upcoming) return upcoming;
          }
          return candidates.reduce<typeof candidates[number]|null>((best,candidate)=>!best||candidate.difference<best.difference?candidate:best,null);
        };
        const findPair=(start:{timestamp:number;time:string;date:DateParts},end:{timestamp:number;time:string;date:DateParts})=>{
          const duration=(end.timestamp-start.timestamp)/60000;
          if(duration<300||duration>1200) return null;
          const arrivalMatch=nearestShift(start.timestamp,start.date,"start");
          if(!arrivalMatch||arrivalMatch.difference>240) return null;
          const previousShiftWithOnTimeDeparture=workShifts.reduce<{shift:WorkShift;endDifference:number}|null>((best,candidate)=>{
            const window=shiftWindow(start.date,candidate);
            if(window.start>start.timestamp) return best;
            const startDifference=(start.timestamp-window.start)/60000;
            const endDifference=Math.abs(end.timestamp-window.end)/60000;
            if(startDifference>240||endDifference>30) return best;
            return !best||endDifference<best.endDifference?{shift:candidate,endDifference}:best;
          },null);
          const selected=previousShiftWithOnTimeDeparture?.shift||arrivalMatch.shift;
          const window=shiftWindow(start.date,selected);
          const endDifference=Math.abs(end.timestamp-window.end)/60000;
          if(endDifference>360) return null;
          return {
            shift:selected,
            overtime:Math.max(0,Math.floor((end.timestamp-window.end)/60000)),
            score:arrivalMatch.difference+endDifference
          };
        };
        const valueAfterLabel=(row:any[],label:string)=>{
          const index=(row||[]).findIndex((value:any)=>normalizeKey(value)===label);
          if(index<0) return "";
          for(let column=index+1;column<row.length;column++){
            const value=String(row[column]??"").trim();
            if(value) return value;
          }
          return "";
        };
        userRows.forEach((userRowIndex,userIndex)=>{
          const userRow=matrix[userRowIndex]||[];
          const name=valueAfterLabel(userRow,"name");
          if(!name) return;
          const rawId=valueAfterLabel(userRow,"userid");
          const id=rawId?`EMP-${/^\d+$/.test(rawId)?rawId.padStart(3,"0"):rawId}`:stableEmployeeId(name);
          const dept=valueAfterLabel(userRow,"department")||"-";
          const dayRowIndex=userRowIndex+1;
          const dayRow=matrix[dayRowIndex]||[];
          const dayColumns:{column:number;day:number}[]=[];
          dayRow.forEach((value:any,column:number)=>{
            const day=Number(String(value??"").trim());
            if(Number.isInteger(day)&&day>=1&&day<=31) dayColumns.push({column,day});
          });
          if(!dayColumns.length) return;

          const nextUserRowIndex=userRows[userIndex+1]??matrix.length;
          const events=new Map<number,{timestamp:number;time:string;date:DateParts}>();
          const daySet=new Set<string>();
          dayColumns.forEach(({column,day})=>{
            const date=validDateParts(period.year,period.month,day);
            if(!date) return;
            daySet.add(`${date.year}-${date.month}-${date.day}`);
            for(let rowIndex=dayRowIndex+1;rowIndex<nextUserRowIndex;rowIndex++){
              const rawCell=matrix[rowIndex]?.[column];
              if(typeof rawCell==="number"&&rawCell>0&&rawCell<1){
                const time=parseTime(rawCell);
                const [hour,minute]=time.split(":").map(Number);
                const timestamp=new Date(date.year,date.month-1,date.day,hour,minute).getTime();
                events.set(timestamp,{timestamp,time,date});
                continue;
              }
              const cell=String(rawCell??"");
              const matches=cell.match(/\d{1,2}[:.]\d{2}(?::\d{2})?/g)||[];
              matches.forEach((value:string)=>{
                const time=parseTime(value);
                if(time!=="-"){
                  const [hour,minute]=time.split(":").map(Number);
                  const timestamp=new Date(date.year,date.month-1,date.day,hour,minute).getTime();
                  events.set(timestamp,{timestamp,time,date});
                }
              });
            }
          });

          const orderedEvents=Array.from(events.values()).sort((a,b)=>a.timestamp-b.timestamp);
          const consumed=new Set<number>();
          // Kunci punch agar baris kembar dari shift malam tidak dihitung dua kali.
          const punchKeys=new Set<string>();
          const punchKey=(shiftName:string,checkIn:string,checkOut:string)=>`${normalizeKey(shiftName)}|${checkIn}|${checkOut}`;
          const pushAttendance=(date:DateParts,shiftName:string,checkIn:string,checkOut:string,overtime=0)=>{
            const key=punchKey(shiftName,checkIn,checkOut);
            if(punchKeys.has(key)) return;
            punchKeys.add(key);
            rows.push({
              name,id,dept,shift:shiftName,date:new Date(date.year,date.month-1,date.day),month:formatMonth(date),
              checkIn,checkOut,status:checkIn||checkOut?"Hadir":"Tidak Hadir",
              late:0,overtime,location:"",__sheet:sheetName,
              __sheetMonth:formatMonth(date),__sheetYear:date.year
            });
          };

          // Match arrivals against configured shifts and allow the checkout to fall on the next day.
          orderedEvents.forEach((start,startIndex)=>{
            if(consumed.has(start.timestamp)) return;
            let best:{end:typeof start;shift:WorkShift;overtime:number;score:number;endIndex:number}|null=null;
            for(let endIndex=startIndex+1;endIndex<orderedEvents.length;endIndex++){
              const end=orderedEvents[endIndex];
              if((end.timestamp-start.timestamp)/60000>1200) break;
              if(consumed.has(end.timestamp)) continue;
              const pair=findPair(start,end);
              if(!pair) continue;
              const skipped=orderedEvents.slice(startIndex+1,endIndex).filter(event=>!consumed.has(event.timestamp)).length;
              const score=pair.score+skipped*180;
              if(!best||score<best.score) best={end,shift:pair.shift,overtime:pair.overtime,score,endIndex};
            }
            if(!best) return;
            pushAttendance(start.date,best.shift.name,start.time,best.end.time,best.overtime);
            for(let index=startIndex;index<=best.endIndex;index++) consumed.add(orderedEvents[index].timestamp);
          });

          // Keep unmatched punches visible and infer missing arrivals or departures from shift times.
          orderedEvents.forEach(event=>{
            if(consumed.has(event.timestamp)) return;
            const startMatch=nearestShift(event.timestamp,event.date,"start");
            const sameDayEnd=nearestShift(event.timestamp,event.date,"end");
            const previousDay=new Date(event.date.year,event.date.month-1,event.date.day-1);
            const previousParts=validDateParts(previousDay.getFullYear(),previousDay.getMonth()+1,previousDay.getDate());
            const previousEnd=previousParts?nearestShift(event.timestamp,previousParts,"end"):null;
            const endMatch=previousEnd&&(!sameDayEnd||previousEnd.difference<sameDayEnd.difference)?previousEnd:sameDayEnd;
            if(endMatch&&endMatch.difference<=240&&endMatch.difference<=(startMatch?.difference??Infinity)){
              const isNightWrap=endMatch===previousEnd&&Boolean(previousParts);
              const workDate=isNightWrap&&previousParts?previousParts:event.date;
              if(isNightWrap) stats.nightWrapped++;
              const overtime=Math.max(0,Math.floor((event.timestamp-endMatch.expected)/60000));
              pushAttendance(workDate,endMatch.shift.name,"",event.time,overtime);
            }else if(startMatch&&startMatch.difference<=240){
              pushAttendance(event.date,startMatch.shift.name,event.time,"");
            }else{
              pushAttendance(event.date,"Shift tidak terdeteksi",event.time,"");
            }
            consumed.add(event.timestamp);
          });

          daySet.forEach(dateKey=>{
            const [year,month,day]=dateKey.split("-").map(Number);
            const date=validDateParts(year,month,day);
            if(!date) return;
            const hasPunch=orderedEvents.some(event=>event.date.year===year&&event.date.month===month&&event.date.day===day);
            if(!hasPunch) pushAttendance(date,"Belum ditentukan","","");
          });
        });
        if(rows.length>sheetRowStart) stats.sheetsRead++; else stats.sheetsSkipped++;
        return;
      }

      // Expand monthly calendars where each day is a separate column.
      for(let headerIndex=0;headerIndex<Math.min(matrix.length,30);headerIndex++){
        const header=matrix[headerIndex]||[];
        const nameColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="name");
        if(nameColumn<0) continue;
        const dayColumns:{column:number;day:number;date:DateParts|null}[]=[];
        header.forEach((value:any,column:number)=>{
          if(column===nameColumn) return;
          if(typeof value==="number"&&Number.isInteger(value)&&value>=1&&value<=31){
            dayColumns.push({column,day:value,date:null});
          }else{
            const date=parseDateParts(value);
            if(date) dayColumns.push({column,day:date.day,date});
          }
        });
        if(dayColumns.length<5) continue;
        const idColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="id");
        const deptColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="dept");
        const shiftColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="shift");
        const monthColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="month");
        const yearColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="year");
        for(let rowIndex=headerIndex+1;rowIndex<matrix.length;rowIndex++){
          const row=matrix[rowIndex]||[];
          const name=String(row[nameColumn]??"").trim();
          if(!name||["nama","name","karyawan","jumlah","total","no","nomor"].includes(normalizeKey(name))) continue;
          const employeeId=String(idColumn>=0?row[idColumn]??"":"").trim()||stableEmployeeId(name);
          const dept=String(deptColumn>=0?row[deptColumn]??"":"").trim()||"-";
          const rowPeriod=periodFromText(String(monthColumn>=0?row[monthColumn]??"":"")+" "+String(yearColumn>=0?row[yearColumn]??"":""))||period;
          for(const item of dayColumns){
            const value=row[item.column];
            if(value===null||value===undefined||String(value).trim()===""||String(value).trim()==="-") continue;
            const date=item.date||dateFromDay(item.day,rowPeriod);
            const raw=String(value).trim();
            const times=raw.match(/\d{1,2}[:.]\d{2}(?::\d{2})?/g)||[];
            rows.push({
              name,id:employeeId,dept,date:date?new Date(date.year,date.month-1,date.day):null,
              shift:String(shiftColumn>=0?row[shiftColumn]??"":"").trim(),
              month:formatMonth(date,rowPeriod),checkIn:times[0]||"",checkOut:times[1]||"",
              status:raw,late:0,overtime:0,location:"",__sheet:sheetName,
              __sheetMonth:formatMonth(null,rowPeriod),__sheetYear:rowPeriod?.year||0
            });
          }
        }
        if(rows.length>sheetRowStart) stats.sheetsRead++; else stats.sheetsSkipped++;
        return;
      }

      // Read workbooks with one date column and employee IN/OUT columns across the sheet.
      for(let headerIndex=0;headerIndex<Math.min(matrix.length,20);headerIndex++){
        const header=matrix[headerIndex]||[];
        const dateColumn=header.findIndex((value:any)=>resolveHeaderField(value)==="date");
        if(dateColumn<0) continue;
        let ioRow=-1;
        for(let candidate=headerIndex+1;candidate<=Math.min(headerIndex+5,matrix.length-1);candidate++){
          const markers=(matrix[candidate]||[]).map((value:any)=>normalizeKey(value));
          if(markers.filter((value:string)=>value==="in"||value==="out").length>=2){ioRow=candidate;break;}
        }
        if(ioRow<0) continue;

        const employees=new Map<string,{name:string;inColumn?:number;outColumn?:number}>();
        let currentName="";
        header.forEach((value:any,column:number)=>{
          const text=String(value??"").trim();
          if(text&&column!==dateColumn&&!resolveHeaderField(text)&&!/^\d+$/.test(text)) currentName=text;
          const marker=normalizeKey(matrix[ioRow]?.[column]);
          if((marker!=="in"&&marker!=="out")||!currentName) return;
          const employee=employees.get(currentName)||{name:currentName};
          if(marker==="in") employee.inColumn=column;
          else employee.outColumn=column;
          employees.set(currentName,employee);
        });
        if(!employees.size) continue;

        for(let rowIndex=ioRow+1;rowIndex<matrix.length;rowIndex++){
          const date=dateFromDay(row[dateColumn],period);
          if(!date) continue;
          for(const employee of employees.values()){
            const rawIn=employee.inColumn===undefined?null:row[employee.inColumn];
            const rawOut=employee.outColumn===undefined?null:row[employee.outColumn];
            const checkIn=parseTime(rawIn);
            const checkOut=parseTime(rawOut);
            const note=[rawIn,rawOut].map(value=>String(value??"").trim()).find(value=>value&&value!=="-"&&!/\d{1,2}[:.]\d{2}/.test(value))||"";
            if(checkIn==="-"&&checkOut==="-"&&!note) continue;
            rows.push({
              name:employee.name,id:stableEmployeeId(employee.name),dept:"-",
              date:new Date(date.year,date.month-1,date.day),month:formatMonth(date),
              checkIn,checkOut,status:note,late:0,overtime:0,location:"",__sheet:sheetName,
              __sheetMonth:formatMonth(date),__sheetYear:date.year
            });
          }
        }
        if(rows.length>sheetRowStart) stats.sheetsRead++; else stats.sheetsSkipped++;
        return;
      }

      let headerRowIndex=-1;
      let headerMap:Record<number,string>={};
      for(let rowIndex=0;rowIndex<Math.min(matrix.length,30);rowIndex++){
        const detected:Record<number,string>={};
        (matrix[rowIndex]||[]).forEach((cell:any,column:number)=>{
          const field=resolveHeaderField(cell);
          if(field&&!Object.values(detected).includes(field)) detected[column]=field;
        });
        const fields=Object.values(detected);
        const hasIdentity=fields.includes("name")||fields.includes("id");
        const hasAttendance=fields.some(field=>["date","month","checkIn","checkOut","status","late","overtime"].includes(field));
        if(hasIdentity&&hasAttendance){headerRowIndex=rowIndex;headerMap=detected;break;}
      }
      if(headerRowIndex<0) return;

      for(let rowIndex=headerRowIndex+1;rowIndex<matrix.length;rowIndex++){
        const values=matrix[rowIndex]||[];
        if(!values.some((value:any)=>String(value??"").trim()!=="")) continue;
        const row:SheetRow={__sheet:sheetName,__sheetMonth:formatMonth(null,period),__sheetYear:period?.year||0};
        Object.entries(headerMap).forEach(([column,field])=>{
          let value=values[Number(column)];
          if(field==="date"&&period&&((typeof value==="number"&&Number.isInteger(value)&&value>=1&&value<=31)||/^\d{1,2}$/.test(String(value??"").trim()))){
            value=new Date(period.year,period.month-1,Number(value));
          }
          row[field]=value;
        });
        const explicitRowPeriod=periodFromText(String(row.month??"")+" "+String(row.year??""));
        const rowPeriod=explicitRowPeriod||period;
        if(row.date!==undefined&&explicitRowPeriod){
          const dateParts=dateFromDay(row.date,rowPeriod);
          if(dateParts) row.date=new Date(dateParts.year,dateParts.month-1,dateParts.day);
        }
        row.__sheetMonth=formatMonth(null,rowPeriod);
        row.__sheetYear=rowPeriod?.year||0;
        rows.push(row);
      }
      if(rows.length>sheetRowStart) stats.sheetsRead++; else stats.sheetsSkipped++;
    });
    return {rows,stats};
  };

  const sortedAll=useMemo(()=>[...attendance].sort((a,b)=>{
    const d=sortableDate(a.date)-sortableDate(b.date);
    if(d!==0) return d;
    const ma=monthIndex(a.month)-monthIndex(b.month);
    if(ma!==0) return ma;
    return a.name.localeCompare(b.name,"id");
  }),[attendance]);

  const months=useMemo(()=>{
    return Array.from(new Set(sortedAll.map(a=>a.month).filter(value=>value&&value!=="-"))).sort((a,b)=>monthIndex(a)-monthIndex(b));
  },[sortedAll]);

  const sheets=useMemo(()=>Array.from(new Set(sortedAll.map(a=>a.sheet).filter(Boolean)).values()).sort((a,b)=>a.localeCompare(b,"id")),[sortedAll]);
  const shifts=useMemo(()=>{
    const extraShifts=Array.from(new Set(sortedAll.map(a=>a.shift).filter((value):value is string=>Boolean(value&&value!=="-"&&value!=="Belum ditentukan"&&!workShifts.some(option=>option.name===value))))).sort((a,b)=>a.localeCompare(b,"id"));
    return [...workShifts.map(option=>option.name),...extraShifts];
  },[sortedAll]);

  const selectedShift=workShifts.find(option=>option.name===shift);
  const filtered=useMemo(()=>sortedAll.filter(a=>{
    const checkInMatch=a.checkIn.match(/^(\d{1,2}):(\d{2})$/);
    const checkInMinutes=checkInMatch?Number(checkInMatch[1])*60+Number(checkInMatch[2]):-1;
    const [startHour,startMinute]=selectedShift?.start.split(":").map(Number)||[];
    const minutesBeforeShift=selectedShift&&checkInMinutes>=0?startHour*60+startMinute-checkInMinutes:0;
    const matchesShift=shift==="Semua"||a.shift===shift||(minutesBeforeShift>0&&minutesBeforeShift<=30);
    return (month==="Semua"||a.month===month) &&
      (sheet==="Semua"||a.sheet===sheet) &&
      matchesShift &&
      `${a.name} ${a.dept} ${a.date} ${a.month} ${a.sheet} ${a.shift??""}`.toLowerCase().includes(q.toLowerCase());
  },[sortedAll,month,sheet,shift,q,selectedShift]));
  const pageSize=50;
  const pageCount=Math.max(1,Math.ceil(filtered.length/pageSize));
  const currentPage=Math.min(page,pageCount);
  const visibleRows=filtered.slice((currentPage-1)*pageSize,currentPage*pageSize);

  const importExcel=async(e:React.ChangeEvent<HTMLInputElement>)=>{
    const file=e.target.files?.[0];
    if(!file) return;
    setImportNotice(null);
    try{
      const workbook=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
      const {rows,stats}=collectSheetRows(workbook);
      const imported=rows.map((row,index):Attendance|null=>{
        const rawDate=row.date;
        const explicitMonth=String(row.month??row.__sheetMonth??"").trim();
        const period=periodFromText(String(row.month??"")+" "+String(row.year??""))||periodFromText(row.__sheetMonth);
        const dateParts=dateFromDay(rawDate,period);
        const date=formatDate(dateParts);
        const dateTime=dateParts?new Date(dateParts.year,dateParts.month-1,dateParts.day).getTime():sortableDate(rawDate);
        const month=formatMonth(dateParts,period)||monthFromDate(dateTime)||explicitMonth||"-";
        const name=String(row.name??"").trim();
        const employeeId=String(row.id??row.employeeId??"").trim()||stableEmployeeId(name);
        const dept=String(row.dept??"").trim()||"-";
        const checkIn=parseTime(row.checkIn);
        const checkOut=parseTime(row.checkOut);
        const statusRaw=row.status;
        const late=num(row.late);
        const overtime=num(row.overtime);
        const hasAttendance=date!=="-"||checkIn!=="-"||checkOut!=="-"||String(statusRaw??"").trim()!==""||late>0||overtime>0;
        stats.raw++;
        if((!name&&!row.id)||!hasAttendance){stats.invalid++;return null;}
        const employeeName=name||"Karyawan "+employeeId;
        return {
          key:"excel-"+Date.now()+"-"+index+"-"+normalizeKey(row.__sheet),
          employeeId,name:employeeName,dept,date,month,sheet:row.__sheet||"Sheet 1",
          shift:String(row.shift??row.shiftName??"").trim()||"Belum ditentukan",
          checkIn,checkOut,status:attendanceStatus(statusRaw,checkIn,checkOut),
          late,overtime,location:String(row.location??"").trim()||"-"
        };
      }).filter((record):record is Attendance=>record!==null);

      if(!imported.length){
        setImportNotice({type:"error",message:"Belum ada data absensi yang terbaca. Pastikan periode bulan/tahun tersedia dan sheet memakai header standar, tanggal per kolom, atau format laporan mesin dengan User ID dan Name."});
        return;
      }

      // Buang baris kembar (karyawan + tanggal + shift + jam sama) agar rekap tidak dihitung dobel.
      // Bila jam berbeda pada kunci sama, data dipertahankan (kecuali satu baris kosong tanpa jam).
      const byKey=new Map<string,Attendance>();
      imported.forEach(record=>{
        const key=[record.employeeId,normalizeKey(record.date),normalizeKey(record.shift||""),record.checkIn,record.checkOut].join("|");
        const bareKey=[record.employeeId,normalizeKey(record.date),normalizeKey(record.shift||"")].join("|");
        const hasClock=record.checkIn!=="-"||record.checkOut!=="-";
        const existing=byKey.get(key);
        if(!existing){byKey.set(key,record);return;}
        stats.deduped++;
        if(existing.name==="Karyawan "+existing.employeeId&&record.name!=="Karyawan "+record.employeeId) byKey.set(key,record);
      });
      const deduped=Array.from(byKey.values());
      // Baris tanpa jam kembar dengan baris berjam pada hari & shift sama → buang yang kosong.
      const withClock=new Set(deduped.filter(r=>r.checkIn!=="-"||r.checkOut!=="-").map(r=>[r.employeeId,normalizeKey(r.date),normalizeKey(r.shift||"")].join("|")));
      const cleaned=deduped.filter(r=>{
        if(!workShifts.some(w=>w.name===r.shift)) return true;
        const bare=[r.employeeId,normalizeKey(r.date),normalizeKey(r.shift||"")].join("|");
        const noClock=r.checkIn==="-"&&r.checkOut==="-"&&r.status!=="telat"&&r.late===0&&r.overtime===0;
        if(noClock&&withClock.has(bare)){stats.deduped++;return false;}
        return true;
      });

      setAttendance(cleaned);
      setSelected([]);
      setPage(1);
      setDetail(null);
      setMonth("Semua");
      setSheet("Semua");
      setShift("Semua");
      const employeeMap=new Map<string,Employee>();
      cleaned.forEach(record=>{
        if(!employeeMap.has(record.employeeId)){
          employeeMap.set(record.employeeId,{
            id:record.employeeId,name:record.name,dept:record.dept,status:record.status,
            checkIn:record.checkIn,checkOut:record.checkOut,avatar:initials(record.name)
          });
        }
      });
      setEmployees(Array.from(employeeMap.values()));
      const sheetCount=new Set(cleaned.map(record=>record.sheet)).size;
      const detailParts=[
        cleaned.length+" data dari "+sheetCount+" sheet",
        stats.nightWrapped?stats.nightWrapped+" shift malam digabung ke tanggal kerja":"",
        stats.deduped?stats.deduped+" baris kembar dibuang":"",
        stats.invalid?stats.invalid+" baris dilewati (tanpa nama/tanggal/jam)":"",
        stats.sheetsSkipped?stats.sheetsSkipped+" sheet dilewati (format tidak dikenali)":""
      ].filter(Boolean);
      setImportNotice({type:"success",message:file.name+" berhasil dibaca: "+detailParts.join(" · ")+"."});
    }catch(error){
      console.error(error);
      setImportNotice({type:"error",message:"File Excel tidak dapat dibaca. Pastikan formatnya .xlsx, .xls, atau .csv."});
    }finally{
      e.target.value="";
    }
  };
  const del=(key:string)=>{
    setAttendance(a=>a.filter(x=>x.key!==key));
    setSelected(s=>s.filter(k=>k!==key));
    if(detail?.key===key) setDetail(null);
  };

  return <Page>
    <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-5">
      <div>
        <h2 className="text-xl font-bold text-slate-900 dark:text-white">Data Absensi</h2>
        <p className="text-sm text-slate-500">Impor beberapa sheet sekaligus, lalu saring data berdasarkan bulan, sheet, atau jam shift.</p>
      </div>
      <div className="flex gap-2">
        <label className="inline-flex cursor-pointer items-center gap-2 rounded-xl bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-blue-700 focus-within:ring-4 focus-within:ring-blue-500/20">
          <input type="file" accept=".xlsx,.xls,.csv" className="sr-only" onChange={importExcel} aria-label="Pilih file Excel absensi"/>
            <Icons.Upload className="w-4"/>Import Excel
        </label>
        <Button variant="danger" disabled={!selected.length} onClick={()=>setAttendance(a=>a.filter(x=>!selected.includes(x.key)))}>Hapus Terpilih</Button>
      </div>
    </div>

    {importNotice&&<div role="status" className={`mb-4 rounded-xl border px-4 py-3 text-sm ${importNotice.type==="success"?"border-emerald-200 bg-emerald-50 text-emerald-700 dark:border-emerald-900 dark:bg-emerald-950/30 dark:text-emerald-300":"border-rose-200 bg-rose-50 text-rose-700 dark:border-rose-900 dark:bg-rose-950/30 dark:text-rose-300"}`}>
      {importNotice.message}
    </div>}

    <div className="flex flex-col sm:flex-row gap-3 mb-4">
      <div className="relative flex-1">
        <Icons.Search className="absolute left-3 top-3.5 w-4 text-slate-400"/>
        <input value={q} onChange={e=>{setPage(1);setQ(e.target.value)}} placeholder="Cari nama, tanggal, bulan..." className="w-full pl-10 pr-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-white"/>
      </div>
      <select value={month} onChange={e=>{setPage(1);setMonth(e.target.value)}} className="px-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-white">
        <option>Semua</option>{months.map(m=><option key={m}>{m}</option>)}
      </select>
      <select value={sheet} onChange={e=>{setPage(1);setSheet(e.target.value)}} aria-label="Filter sheet Excel" className="px-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-white">
        <option>Semua</option>{sheets.map(value=><option key={value}>{value}</option>)}
      </select>
      <select value={shift} onChange={e=>{setPage(1);setShift(e.target.value)}} aria-label="Filter shift" className="px-4 py-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-white">
        <option value="Semua">Semua Shift</option>{shifts.map(value=><option key={value}>{value}</option>)}
      </select>
    </div>

    <Card className="overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full min-w-[1600px]">
          <thead><tr className="bg-slate-50 dark:bg-slate-800/60 text-xs text-slate-500 text-left whitespace-nowrap">
            <th className="p-4"><input type="checkbox" checked={filtered.length>0&&filtered.every(a=>selected.includes(a.key))} onChange={e=>setSelected(e.target.checked?filtered.map(a=>a.key):[])}/></th>
            <th className="px-4 py-4 min-w-[150px]">Nama</th>
            <th className="px-4 py-4 min-w-[170px]">Tanggal</th>
            <th className="px-4 py-4 min-w-[150px]">Bulan</th>
            <th className="px-4 py-4 min-w-[210px]">Sheet</th>
            <th className="px-4 py-4 min-w-[140px]">Shift</th>
            <th className="px-4 py-4 min-w-[135px] text-center border-l border-blue-100 bg-blue-50 text-blue-700 dark:border-blue-900/70 dark:bg-blue-950/40 dark:text-blue-300">Jam Masuk</th>
            <th className="px-4 py-4 min-w-[135px] text-center border-l border-emerald-100 bg-emerald-50 text-emerald-700 dark:border-emerald-900/70 dark:bg-emerald-950/40 dark:text-emerald-300">Jam Pulang</th>
            <th className="px-4 py-4 min-w-[105px]">Status</th>
            <th className="px-4 py-4 min-w-[85px]">Telat</th>
            <th className="px-4 py-4 min-w-[130px]">Lembur</th>
            <th className="px-4 py-4 min-w-[90px]">Detail</th>
            <th className="px-4 py-4 min-w-[75px]">Aksi</th>
          </tr></thead>
          <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
            {visibleRows.map(a=><tr key={a.key} onClick={()=>setDetail(a)} className="text-sm whitespace-nowrap hover:bg-slate-50 dark:hover:bg-slate-800/50 cursor-pointer">
              <td className="p-4" onClick={e=>e.stopPropagation()}><input type="checkbox" checked={selected.includes(a.key)} onChange={e=>setSelected(s=>e.target.checked?[...s,a.key]:s.filter(k=>k!==a.key))}/></td>
              <td className="px-4 py-4 min-w-[150px] font-semibold text-slate-900 dark:text-white">{a.name}</td>
              <td className="px-4 py-4 min-w-[170px]">{a.date}</td>
              <td className="px-4 py-4 min-w-[150px]">{a.month}</td>
              <td className="px-4 py-4 min-w-[210px]"><span className="rounded-lg bg-slate-100 px-2 py-1 text-xs text-slate-600 dark:bg-slate-800 dark:text-slate-300">{a.sheet}</span></td>
              <td className="px-4 py-4 min-w-[140px]">{a.shift||"-"}</td>
              <td className="px-4 py-3 min-w-[135px] text-center border-l border-slate-100 dark:border-slate-800">
                <span className={`inline-flex min-w-[82px] justify-center rounded-lg px-3 py-1.5 text-base font-bold tabular-nums ring-1 ${a.checkIn==="-"?"bg-slate-50 text-slate-400 ring-slate-200 dark:bg-slate-800 dark:text-slate-500 dark:ring-slate-700":"bg-blue-50 text-blue-700 ring-blue-200 dark:bg-blue-950/40 dark:text-blue-300 dark:ring-blue-900"}`}>{a.checkIn}</span>
              </td>
              <td className="px-4 py-3 min-w-[135px] text-center border-l border-slate-100 dark:border-slate-800">
                <span className={`inline-flex min-w-[82px] justify-center rounded-lg px-3 py-1.5 text-base font-bold tabular-nums ring-1 ${a.checkOut==="-"?"bg-slate-50 text-slate-400 ring-slate-200 dark:bg-slate-800 dark:text-slate-500 dark:ring-slate-700":"bg-emerald-50 text-emerald-700 ring-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-300 dark:ring-emerald-900"}`}>{a.checkOut}</span>
              </td>
              <td className="px-4 py-4 min-w-[105px]"><span className={`px-2.5 py-1 rounded-full border text-xs ${statusClass(a.status)}`}>{statusLabel(a.status)}</span></td>
              <td className="px-4 py-4 min-w-[85px]">{a.late?a.late+" m":"-"}</td>
              <td className="px-4 py-4 min-w-[130px]">{formatDuration(a.overtime)}</td>
              <td className="px-4 py-4 min-w-[90px]"><Button variant="ghost" onClick={()=>setDetail(a)}>Lihat</Button></td>
              <td className="px-4 py-4 min-w-[75px]" onClick={e=>e.stopPropagation()}><button onClick={()=>del(a.key)} className="p-2 text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30 rounded-lg"><Icons.Trash className="w-4"/></button></td>
            </tr>)}
            {!filtered.length&&<tr><td colSpan={13} className="p-10 text-center text-sm text-slate-400">Tidak ada data untuk filter ini. Impor file Excel untuk memuat absensi.</td></tr>}
          </tbody>
        </table>
      </div>
    </Card>

    <div className="mt-3 flex flex-col gap-3 rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-500 dark:border-slate-800 dark:bg-slate-900 sm:flex-row sm:items-center sm:justify-between">
      <span>Menampilkan {filtered.length?(currentPage-1)*pageSize+1:0}–{Math.min(currentPage*pageSize,filtered.length)} dari {filtered.length} data</span>
      <div className="flex items-center gap-3">
        <button type="button" onClick={()=>setPage(value=>Math.max(1,value-1))} disabled={currentPage===1} className="rounded-lg px-3 py-2 font-medium text-slate-700 transition hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-40 dark:text-slate-200 dark:hover:bg-slate-800">Sebelumnya</button>
        <span>Halaman {currentPage} dari {pageCount}</span>
        <button type="button" onClick={()=>setPage(value=>Math.min(pageCount,value+1))} disabled={currentPage===pageCount} className="rounded-lg px-3 py-2 font-medium text-slate-700 transition hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-40 dark:text-slate-200 dark:hover:bg-slate-800">Berikutnya</button>
      </div>
    </div>

    {detail&&<div className="fixed inset-0 z-50 bg-black/50 grid place-items-center p-4" onClick={()=>setDetail(null)}>
      <div className="w-full max-w-lg bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden" onClick={e=>e.stopPropagation()}>
        <div className="p-5 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
          <div><h3 className="font-bold text-lg text-slate-900 dark:text-white">Detail Absensi</h3><p className="text-xs text-slate-500">{detail.name}</p></div>
          <button onClick={()=>setDetail(null)} className="p-2 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800"><Icons.X className="w-5"/></button>
        </div>
        <div className="p-5 grid grid-cols-2 gap-3">
          {[
            ["Tanggal",detail.date],["Bulan",detail.month],["Sheet",detail.sheet],["Shift",detail.shift||"-"],["Departemen",detail.dept],["Jam Masuk",detail.checkIn],
            ["Jam Pulang",detail.checkOut],["Status",statusLabel(detail.status)],["Terlambat",detail.late?`${detail.late} menit`:"-"],
            ["Lembur",formatDuration(detail.overtime)],["Lokasi",detail.location]
          ].map(([k,v])=><div key={k} className="rounded-xl bg-slate-50 dark:bg-slate-800 p-3"><div className="text-xs text-slate-500">{k}</div><div className="font-semibold text-slate-900 dark:text-white mt-1">{v}</div></div>)}
        </div>
      </div>
    </div>}
  </Page>
}
const dayNames=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];

function PrintRekap({employees,attendance}:{employees:Employee[];attendance:Attendance[]}){
  const months=useMemo(()=>{
    const set=new Set<string>();
    attendance.forEach(a=>{if(a.month&&a.month!=="-") set.add(a.month);});
    if(!set.size) set.add(formatMonth(parseDateParts(new Date())));
    return Array.from(set).sort((a,b)=>monthIndex(a)-monthIndex(b));
  },[attendance]);
  const EMPTY_SENTINEL="__empty__";
  const matchesShift=(a:Attendance,shiftValue:string)=>{
    if(shiftValue==="Semua") return true;
    if(shiftValue===EMPTY_SENTINEL) return !a.shift||a.shift==="-"||a.shift==="Belum ditentukan";
    if(a.shift===shiftValue) return true;
    const option=workShifts.find(w=>w.name===shiftValue);
    if(!option) return false;
    const minutes=timeToMinutes(a.checkIn);
    if(minutes===null) return false;
    const before=(timeToMinutes(option.start)??0)-minutes;
    return before>0&&before<=30;
  };
  const shifts=useMemo(()=>{
    const extraShifts=Array.from(new Set(attendance.map(a=>a.shift).filter((value):value is string=>Boolean(value&&value!=="-"&&value!=="Belum ditentukan"&&!workShifts.some(option=>option.name===value))))).sort((a,b)=>a.localeCompare(b,"id"));
    return [...workShifts.map(option=>option.name),...extraShifts];
  },[attendance]);
  const hasEmptyShift=useMemo(()=>attendance.some(a=>!a.shift||a.shift==="-"||a.shift==="Belum ditentukan"),[attendance]);
  const [month,setMonth]=useState("Semua");
  const [monthMode,setMonthMode]=useState<"single"|"range"|"all">("single");
  const [monthFrom,setMonthFrom]=useState("");
  const [monthTo,setMonthTo]=useState("");
  const [shift,setShift]=useState("Semua");
  const [employeeId,setEmployeeId]=useState("Semua");
  const effectiveMonth=monthMode==="single"?month:(monthMode==="range"?(monthFrom||months[months.length-1]||""):(months[months.length-1]||""));
  const period=periodFromText(effectiveMonth);
  const selectedMonths=useMemo(()=>{
    if(monthMode==="all") return months;
    if(monthMode==="single") return month==="Semua"?months:[month];
    const a=monthIndex(monthFrom||months[0]);
    const b=monthIndex(monthTo||months[months.length-1]);
    const lo=Math.min(a,b);const hi=Math.max(a,b);
    return months.filter(m=>{const pi=monthIndex(m);return pi>=lo&&pi<=hi;});
  },[monthMode,month,monthFrom,monthTo,months]);
  const employeeList=useMemo(()=>{
    const seen=new Set<string>();
    const list:Array<{id:string;name:string;dept:string}>=[];
    employees.forEach(e=>{if(!seen.has(e.id)){seen.add(e.id);list.push({id:e.id,name:e.name,dept:e.dept});}});
    attendance.forEach(a=>{if(a.employeeId&&!seen.has(a.employeeId)){seen.add(a.employeeId);list.push({id:a.employeeId,name:a.name,dept:a.dept});}});
    return list.sort((a,b)=>a.name.localeCompare(b.name));
  },[employees,attendance]);
  const targets=employeeId==="Semua"?employeeList:employeeList.filter(e=>e.id===employeeId);
  const parseDay=(a:Attendance)=>{const p=parseDateParts(a.date);return p?p.day:0;};
  const minutesFrom=(v:string)=>{const m=String(v||"").match(/^(\d{1,2}):(\d{2})$/);return m?Number(m[1])*60+Number(m[2]):null;};
  const clock=(min:number|null)=>min===null?"-":String(Math.floor(min/60)).padStart(2,"0")+":"+String(min%60).padStart(2,"0");
  const summarize=(id:string,day:number,monthValue:string)=>{
    const recs=attendance.filter(a=>a.employeeId===id&&parseDay(a)===day&&a.month===monthValue&&matchesShift(a,shift));
    if(!recs.length) return null;
    const ins=recs.map(r=>minutesFrom(r.checkIn)).filter((x):x is number=>x!==null);
    const outs=recs.map(r=>minutesFrom(r.checkOut)).filter((x):x is number=>x!==null);
    const late=Math.max(0,...recs.map(r=>r.late||0));
    const overtime=recs.reduce((n,r)=>n+(r.overtime||0),0);
    const status:Status=late>0?"telat":(recs.some(r=>r.status==="hadir")?"hadir":recs[0].status);
    const shiftNames=Array.from(new Set(recs.map(r=>r.shift||"Belum ditentukan"))).join(", ");
    return {checkIn:clock(ins.length?Math.min(...ins):null),checkOut:clock(outs.length?Math.max(...outs):null),late,overtime,status,shiftNames};
  };
  const totals=(id:string,monthValue:string,monthPeriod:{month:number;year:number}|null)=>{
    let late=0,ot=0,hadir=0,telat=0;
    const dim=monthPeriod?new Date(monthPeriod.year,monthPeriod.month,0).getDate():31;
    for(let d=1;d<=dim;d++){const r=summarize(id,d,monthValue);if(!r) continue;late+=r.late;ot+=r.overtime;if(r.status==="hadir")hadir++;if(r.status==="telat")telat++;}
    return {late,ot,hadir,telat};
  };
  const selectCls="mt-1 block w-full rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-white px-3 py-2 text-sm";
  const shiftLabel=shift===EMPTY_SENTINEL?"Tanpa shift":shift;
  return <Page>
    <div className="no-print mb-5 flex flex-col lg:flex-row lg:items-end gap-3">
      <div><h2 className="text-xl font-bold text-slate-900 dark:text-white">Cetak Rekap Absensi</h2><p className="text-sm text-slate-500">Rekap per karyawan tanggal 1 s/d akhir bulan — telat &amp; lembur terlihat jelas.</p></div>
      <div className="lg:ml-auto flex flex-wrap items-end gap-3">
        <label className="text-xs font-semibold text-slate-500">Mode Bulan<select value={monthMode} onChange={e=>setMonthMode(e.target.value as any)} className={selectCls}><option value="single">Per Bulan</option><option value="range">Rentang Bulan</option><option value="all">Semua Bulan</option></select></label>
        {monthMode==="single"&&<label className="text-xs font-semibold text-slate-500">Bulan<select value={month} onChange={e=>setMonth(e.target.value)} className={selectCls}><option value="Semua">Semua Bulan</option>{months.map(m=><option key={m} value={m}>{m}</option>)}</select></label>}
        {monthMode==="range"&&<><label className="text-xs font-semibold text-slate-500">Dari bulan<select value={monthFrom} onChange={e=>setMonthFrom(e.target.value)} className={selectCls}>{months.map(m=><option key={m} value={m}>{m}</option>)}</select></label><label className="text-xs font-semibold text-slate-500">s/d bulan<select value={monthTo} onChange={e=>setMonthTo(e.target.value)} className={selectCls}>{months.map(m=><option key={m} value={m}>{m}</option>)}</select></label></>}
        {monthMode==="all"&&<span className="self-center rounded-lg bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-700 dark:bg-blue-950/40 dark:text-blue-300">Mencetak {months.length} bulan</span>}
        <label className="text-xs font-semibold text-slate-500">Shift<select value={shift} onChange={e=>setShift(e.target.value)} className={selectCls}><option value="Semua">Semua Shift</option>{shifts.map(sv=><option key={sv} value={sv}>{sv}</option>)}{hasEmptyShift&&<option value={EMPTY_SENTINEL}>Tanpa shift</option>}</select></label>
        <label className="text-xs font-semibold text-slate-500">Karyawan<select value={employeeId} onChange={e=>setEmployeeId(e.target.value)} className={selectCls}><option value="Semua">Semua Karyawan</option>{employeeList.map(e=><option key={e.id} value={e.id}>{e.name}</option>)}</select></label>
        <Button onClick={()=>window.print()}>🖨️ Cetak / Simpan PDF</Button>
      </div>
    </div>
    <div className="print-area">
      {selectedMonths.map(mo=>{
        const moPeriod=periodFromText(mo);
        const moDays=moPeriod?new Date(moPeriod.year,moPeriod.month,0).getDate():31;
        return <div key={mo} className="rekap-month-group">
          {targets.map(emp=>{const t=totals(emp.id,mo,moPeriod);return <div key={mo+"-"+emp.id} className="rekap-card">
            <div className="rekap-title">REKAP ABSENSI KARYAWAN</div>
            <div className="rekap-sub">Periode: {mo}{shift!=="Semua"?" - Shift: "+shiftLabel:""}</div>
            <div className="rekap-meta"><span><b>Nama:</b> {emp.name}</span><span><b>ID:</b> {emp.id}</span><span><b>Departemen:</b> {emp.dept||"Umum"}</span></div>
            <table className="rekap-table">
              <thead><tr><th>Tgl</th><th>Hari</th><th>Jam Masuk</th><th>Jam Pulang</th><th>Shift</th><th>Telat</th><th>Lembur</th><th>Status</th></tr></thead>
              <tbody>{Array.from({length:moDays},(_,k)=>k+1).map(d=>{const r=summarize(emp.id,d,mo);const dow=moPeriod?new Date(moPeriod.year,moPeriod.month-1,d).getDay():0;const weekend=dow===0||dow===6;return <tr key={d} style={weekend?{color:"#b91c1c"}:undefined}>
                <td>{String(d).padStart(2,"0")}</td><td>{dayNames[dow]}</td><td>{r?.checkIn||"-"}</td><td>{r?.checkOut||"-"}</td><td>{r?.shiftNames||"-"}</td>
                <td>{r&&r.late>0?r.late+" menit":"-"}</td><td>{r&&r.overtime>0?formatDuration(r.overtime):"-"}</td><td>{r?statusLabel(r.status):"-"}</td>
              </tr>;})}</tbody>
              <tfoot><tr><td colSpan={4}>TOTAL ({t.hadir} hadir, {t.telat} hari telat)</td><td></td><td>{t.late} menit</td><td>{formatDuration(t.ot)}</td><td></td></tr></tfoot>
            </table>
            <div className="rekap-sign"><div>Disetujui,<div className="rekap-line">HRD / Manager</div></div><div>Dibuat oleh,<div className="rekap-line">Admin Absensi</div></div></div>
          </div>;})}
        </div>;
      })}
      {!targets.length&&<Card className="p-8 text-center text-slate-400">Belum ada data. Import file Excel di menu <b>Data Absensi</b> terlebih dahulu.</Card>}
    </div>
  </Page>;
}

function Reports({attendance}:{attendance:Attendance[]}){const totalLate=attendance.reduce((n,a)=>n+a.late,0);const totalOvertime=attendance.reduce((n,a)=>n+a.overtime,0);return <Page><div className="mb-5"><h2 className="text-xl font-bold text-slate-900 dark:text-white">Laporan</h2><p className="text-sm text-slate-500">Ringkasan berdasarkan data absensi yang tersedia.</p></div><div className="grid md:grid-cols-3 gap-4">{[["Total Record",attendance.length],["Total Menit Telat",totalLate],["Total Lembur",formatDuration(totalOvertime)]].map(([a,b])=><Card key={String(a)} className="p-5"><div className="text-sm text-slate-500">{a}</div><div className="text-3xl font-black text-slate-900 dark:text-white mt-2">{b}</div></Card>)}</div></Page>}

function OvertimeLate({attendance}:{attendance:Attendance[]}){const late=useMemo(()=>attendance.filter(a=>a.late>0),[attendance]);const over=useMemo(()=>attendance.filter(a=>a.overtime>0),[attendance]);return <Page><div className="mb-5"><h2 className="text-xl font-bold text-slate-900 dark:text-white">Lembur & Telat</h2><p className="text-sm text-slate-500">Pantau keterlambatan dan durasi lembur dari data Excel.</p></div><div className="grid lg:grid-cols-2 gap-5"><Card className="overflow-hidden"><div className="p-5 border-b border-slate-100 dark:border-slate-800"><h3 className="font-bold text-orange-600">Keterlambatan</h3><p className="text-xs text-slate-500">{late.length} record terlambat</p></div><div className="overflow-x-auto"><table className="w-full"><thead className="bg-slate-50 dark:bg-slate-800/60 text-xs text-slate-500"><tr><th className="p-4 text-left">Nama</th><th className="text-left">Tanggal</th><th className="text-left">Telat</th></tr></thead><tbody className="divide-y divide-slate-100 dark:divide-slate-800">{late.map(a=><tr key={a.key} className="text-sm"><td className="p-4 font-semibold text-slate-900 dark:text-white">{a.name}</td><td>{a.date}</td><td className="font-bold text-orange-600">{a.late} menit</td></tr>)}{!late.length&&<tr><td colSpan={3} className="p-8 text-center text-slate-400">Tidak ada data telat.</td></tr>}</tbody></table></div></Card><Card className="overflow-hidden"><div className="p-5 border-b border-slate-100 dark:border-slate-800"><h3 className="font-bold text-purple-600">Lembur</h3><p className="text-xs text-slate-500">{over.length} record lembur</p></div><div className="overflow-x-auto"><table className="w-full"><thead className="bg-slate-50 dark:bg-slate-800/60 text-xs text-slate-500"><tr><th className="p-4 text-left">Nama</th><th className="text-left">Tanggal</th><th className="text-left">Lembur</th></tr></thead><tbody className="divide-y divide-slate-100 dark:divide-slate-800">{over.map(a=><tr key={a.key} className="text-sm"><td className="p-4 font-semibold text-slate-900 dark:text-white">{a.name}</td><td>{a.date}</td><td className="font-bold text-purple-600">{formatDuration(a.overtime)}</td></tr>)}{!over.length&&<tr><td colSpan={3} className="p-8 text-center text-slate-400">Tidak ada data lembur.</td></tr>}</tbody></table></div></Card></div></Page>}

function Settings({dark,toggle}:{dark:boolean;toggle:()=>void}){return <Page><div className="mb-5"><h2 className="text-xl font-bold text-slate-900 dark:text-white">Pengaturan</h2><p className="text-sm text-slate-500">Atur tampilan dashboard.</p></div><Card className="p-5 max-w-2xl"><div className="flex items-center justify-between py-3"><div><div className="font-semibold text-slate-900 dark:text-white">Tema Tampilan</div><div className="text-sm text-slate-500">Pilih Light Mode atau Dark Mode.</div></div><button onClick={toggle} className="flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-200 font-semibold text-sm">{dark?<Icons.Sun className="w-4"/>:<Icons.Moon className="w-4"/>}{dark?'Light Mode':'Dark Mode'}</button></div></Card></Page>}

export default function App(){const [loggedIn,setLoggedIn]=useState(false);const [screen,setScreen]=useState<Screen>('dashboard');const [employees,setEmployees]=useState<Employee[]>(()=>{try{const s=localStorage.getItem('absenpro-employees');return s?JSON.parse(s):mockEmployees;}catch{return mockEmployees;}});const [attendance,setAttendance]=useState<Attendance[]>(()=>{try{const s=localStorage.getItem('absenpro-attendance');return s?JSON.parse(s):mockAttendance;}catch{return mockAttendance;}});const [menu,setMenu]=useState(false);const [dark,setDark]=useState<boolean>(()=>{if(typeof window==='undefined') return false;try{return localStorage.getItem('absenpro-theme')==='dark';}catch{return false;}});useEffect(()=>{const root=document.documentElement;root.classList.toggle('dark',dark);root.style.colorScheme=dark?'dark':'light';try{localStorage.setItem('absenpro-theme',dark?'dark':'light');}catch{}},[dark]);useEffect(()=>{try{localStorage.setItem('absenpro-employees',JSON.stringify(employees));}catch{}},[employees]);useEffect(()=>{try{localStorage.setItem('absenpro-attendance',JSON.stringify(attendance));}catch{}},[attendance]);const toggle=()=>setDark(value=>!value);if(!loggedIn)return <Login onLogin={()=>setLoggedIn(true)}/>;const logout=()=>{setLoggedIn(false);setScreen('dashboard')};return <div className="min-h-screen bg-slate-100 dark:bg-slate-950 text-slate-800 dark:text-slate-200"><div className="flex min-h-screen"><Sidebar screen={screen} go={setScreen} logout={logout} open={menu} setOpen={setMenu} dark={dark} toggle={toggle}/><div className="flex-1 min-w-0"><Header screen={screen} menu={()=>setMenu(true)} dark={dark} toggle={toggle}/>{screen==='dashboard'&&<Dashboard employees={employees} attendance={attendance} go={setScreen}/>} {screen==='employees'&&<Employees employees={employees} setEmployees={setEmployees} attendance={attendance} setAttendance={setAttendance}/>} {screen==='attendance'&&<AttendancePage attendance={attendance} setAttendance={setAttendance} setEmployees={setEmployees}/>} {screen==='reports'&&<Reports attendance={attendance}/>} {screen==='overtime'&&<OvertimeLate attendance={attendance}/>} {screen==='print'&&<PrintRekap employees={employees} attendance={attendance}/>} {screen==='settings'&&<Settings dark={dark} toggle={toggle}/>}</div></div></div>}
