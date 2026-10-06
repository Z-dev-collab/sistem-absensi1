import * as XLSX from "xlsx";
import { normalizeKey } from "./dateUtils";

export const EMPTY_CLOCK = "-";
export type AttendanceStatus = "hadir" | "telat" | "tidak-hadir" | "belum-absen";

/** Jam Excel/teks apa pun -> menit sejak 00:00, atau null bila bukan jam. */
export function parseClockToMinutes(value: any): number | null {
  if (value === null || value === undefined) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value.getHours() * 60 + value.getMinutes();
  if (typeof value === "number" && Number.isFinite(value)) {
    if (Number.isInteger(value) && value >= 0 && value < 24) return value * 60;
    if (Number.isInteger(value) && value >= 100 && value <= 2359) {
      const h = Math.floor(value / 100);
      const m = value % 100;
      if (h <= 23 && m <= 59) return h * 60 + m;
    }
    if (value > 0 && value < 1) return Math.round(value * 1440) % 1440;
    if (value >= 1 && value < 60000) {
      const frac = ((value % 1) + 1) % 1;
      if (frac > 0.0001 && frac < 0.9999) return Math.round(frac * 1440) % 1440;
      const d = XLSX.SSF.parse_date_code(value);
      if (d) return Math.round((value - Math.floor(value)) * 1440) % 1440;
    }
    return null;
  }
  let text = String(value ?? "").trim().toLowerCase();
  if (!text || text === "-") return null;
  text = text.replace(/[\s_]+/g, " ");
  if (/^(libur|cuti|off|izin|ijin|sakit|alpa|alpha|tanpa keterangan)\b/.test(text)) return null;
  const mer = text.match(/\b(am|pm)\b/)?.[1];
  text = text.replace(/\b(am|pm)\b/g, "").trim().replace(",", ":").replace(".", ":");
  // 08:30[:ss], 8:30, 8 : 30
  let m = text.match(/(\d{1,2})\s*[:]\s*(\d{2})(?:\s*[:]\s*(\d{2}))?/);
  if (m) {
    let h = Number(m[1]);
    if (mer === "pm" && h < 12) h += 12;
    if (mer === "am" && h === 12) h = 0;
    if (h === 24) h = 0;
    if (h >= 0 && h <= 23 && Number(m[2]) <= 59) return h * 60 + Number(m[2]);
    return null;
  }
  m = text.match(/^(\d{1,2})$/);
  if (m) {
    const h = Number(m[1]);
    if (h >= 0 && h <= 23) return h * 60;
    return null;
  }
  m = text.match(/^(\d{3,4})$/);
  if (m) {
    const d = m[1].padStart(4, "0");
    const h = Number(d.slice(0, 2));
    const mm = Number(d.slice(2));
    if (h <= 23 && mm <= 59) return h * 60 + mm;
    return null;
  }
  m = text.match(/^(\d{1,2})\s+(\d{2})$/);
  if (m) {
    const h = Number(m[1]);
    const mm = Number(m[2]);
    if (h <= 23 && mm <= 59) return h * 60 + mm;
  }
  return null;
}

export const formatClock = (minutes: number | null): string =>
  minutes === null ? "-" : String(Math.floor(minutes / 60)).padStart(2, "0") + ":" + String(minutes % 60).padStart(2, "0");

/** Normalisasi ID karyawan agar konsisten antar perangkat & file. */
export function normalizeEmployeeId(name: string, raw: any): string {
  const r = String(raw ?? "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (r) return r.startsWith("EMP") ? r : "EMP-" + r.padStart(3, "0").slice(-12);
  const key = normalizeKey(name).slice(0, 18).toUpperCase() || "TANPA-NAMA";
  return "EMP-" + key;
}

/** Durasi ("1 jam 30 mnt", "01:30", 90, 0.0625 hari Excel) -> menit. */
export function parseDurationMinutes(value: any): number {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value > 0 && value < 1) return Math.round(value * 1440);
    return Math.max(0, Math.round(value));
  }
  const text = String(value).trim().toLowerCase().replace(",", ".");
  if (!text || text === "-") return 0;
  const h = Number(text.match(/(\d+(?:\.\d+)?)\s*(?:jam|hours?|hrs?|h)\b/)?.[1] || 0);
  const mm = Number(text.match(/(\d+(?:\.\d+)?)\s*(?:menit|minutes?|mins?|mnt|m)\b/)?.[1] || 0);
  if (h || mm) return Math.round(h * 60 + mm);
  const c = text.match(/^(\d{1,3})\s*[:.]\s*(\d{2})(?::\d{2})?$/);
  if (c) return Number(c[1]) * 60 + Number(c[2]);
  const n = Number(text.replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(n) ? Math.max(0, Math.round(n)) : 0;
}

/** Klasifikasi status tanpa magic number. */
export function classifyStatus(raw: any, checkIn: string, checkOut: string, late = 0): AttendanceStatus {
  const v = normalizeKey(raw);
  const hasClock = checkIn !== "-" || checkOut !== "-";
  if (!v && !hasClock && late <= 0) return "belum-absen";
  if (v.includes("belum")) return hasClock ? (late > 0 ? "telat" : "hadir") : "belum-absen";
  if (v.includes("telat") || v.includes("terlambat") || v.includes("late")) return "telat";
  if (v.includes("izin") || v.includes("ijin") || v.includes("sakit") || v.includes("cuti") || v.includes("libur")) return "tidak-hadir";
  if (v.includes("tidak") || v.includes("alpa") || v === "alpha" || v === "a") return "tidak-hadir";
  if (v === "0" && !hasClock) return "tidak-hadir";
  if (v.includes("hadir") || v.includes("masuk") || v.includes("present") || v === "h" || v === "1" || v === "v" || v.includes("check")) return late > 0 ? "telat" : "hadir";
  if (hasClock) return late > 0 ? "telat" : "hadir";
  return "belum-absen";
}

/** Kunci stabil agar dedup antar perangkat & impor ulang konsisten. */
export const stableRecordKey = (employeeId: string, dateLabel: string, shift: string, ci: string, co: string): string =>
  [String(employeeId).trim().toUpperCase(), normalizeKey(dateLabel), normalizeKey(shift || ""), ci, co].join("|");
