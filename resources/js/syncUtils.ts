/** Sinkronisasi mobile <-> desktop untuk build statis (Vercel) tanpa backend.
 *  Satu-satunya sumber kebenaran: localStorage per perangkat, sehingga perlu
 *  mekanisme pindah/gabung data: file JSON + kode sinkron (base64).
 *  Kunci stabil + merge by key membuat data gabungan identik di semua perangkat.
 */

export const SYNC_VERSION = 1;
export const LS_EMP = "absenpro-employees";
export const LS_ATT = "absenpro-attendance";

export type SyncPayload = {
  version: number;
  exportedAt: string;
  employees: any[];
  attendance: any[];
};

export function buildPayload(employees: any[], attendance: any[]): SyncPayload {
  return { version: SYNC_VERSION, exportedAt: new Date().toISOString(), employees, attendance };
}

export function downloadJson(filename: string, data: unknown): void {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

export function encodeShareCode(payload: SyncPayload): string {
  const json = JSON.stringify(payload);
  const utf8 = new TextEncoder().encode(json);
  let bin = "";
  utf8.forEach((b) => (bin += String.fromCharCode(b)));
  return btoa(bin);
}

export function decodeShareCode(code: string): SyncPayload {
  const clean = code.trim().replace(/\s+/g, "");
  const bin = atob(clean);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const json = new TextDecoder().decode(bytes);
  const obj = JSON.parse(json);
  if (!obj || !Array.isArray(obj.attendance) || !Array.isArray(obj.employees)) throw new Error("Kode tidak valid");
  return obj as SyncPayload;
}

/** Gabung 2 daftar berdasar key unik -> hasil sama di semua perangkat. */
export function mergeByKey<T extends { key?: string; employeeId?: string }>(oldList: T[], incoming: T[], keyOf: (r: T) => string): T[] {
  const map = new Map<string, T>();
  oldList.forEach((r) => map.set(keyOf(r), r));
  incoming.forEach((r) => {
    const k = keyOf(r);
    if (!map.has(k)) map.set(k, r);
    else {
      // Pertahankan nama paling lengkap, jangan duplikat.
      const cur: any = map.get(k);
      const nxt: any = r;
      if (String(cur?.name || "").startsWith("Karyawan ") && !String(nxt?.name || "").startsWith("Karyawan ")) map.set(k, r);
    }
  });
  return Array.from(map.values());
}
