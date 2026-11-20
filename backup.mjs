// بک‌آپ روزانه‌ی خودکار: وارد حساب می‌شود، داده‌ها را می‌خواند، فایل Excel (همان خروجی تنظیمات) می‌سازد و به ایمیل می‌فرستد.
import fs from "node:fs";

const p2 = n => String(n).padStart(2, "0");
const isOpen = e => !e.end;
function hoursBetween(a, b) {
  if (!a || !b) return null;
  const h = (new Date(b) - new Date(a)) / 36e5;
  return isNaN(h) || h < 0 ? null : h;
}

const enc = new TextEncoder(), dec = new TextDecoder();
const crcT = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = crcT[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zipStore(files) {
  const parts = [], central = []; let offset = 0;
  files.forEach(f => {
    const nm = enc.encode(f.name), len = f.data.length, crc = crc32(f.data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(12, 0x21, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, len, true); lh.setUint32(22, len, true); lh.setUint16(26, nm.length, true);
    parts.push(new Uint8Array(lh.buffer), nm, f.data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(14, 0x21, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, len, true); ch.setUint32(24, len, true); ch.setUint16(28, nm.length, true); ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), nm);
    offset += 30 + nm.length + len;
  });
  const cdSize = central.reduce((s, c) => s + c.length, 0), end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
const xmlEsc = s => String(s ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const isoToSerial = iso => { const m = /^(\d{4})-(\d\d)-(\d\d)[T ](\d\d):(\d\d)(?::(\d\d))?/.exec(iso || ""); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)) / 864e5 + 25569 : null; };
const serialToISO = n => { const d = new Date(Math.round((n - 25569) * 86400) * 1000); return `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}T${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())}`; };
const cs_ = (ref, v, s) => v === "" || v == null ? "" : `<c r="${ref}" t="inlineStr"${s ? ` s="${s}"` : ""}><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
const cn_ = (ref, v, s) => v == null ? "" : `<c r="${ref}"${s ? ` s="${s}"` : ""}><v>${v}</v></c>`;
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
function sheetXml(rows, widths, nCols) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS}><sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${rows.join("")}</sheetData><autoFilter ref="A1:${String.fromCharCode(64 + nCols)}${rows.length}"/></worksheet>`;
}
function buildXlsx(list, codes, withSummary) {
  const L = "ABCDEFGHI";
  const head = r => `<row r="1">${r.map((h, i) => cs_(L[i] + "1", h, 1)).join("")}</row>`;
  list = list.slice().sort((a, b) => a.id - b.id);
  const ids = new Set(list.map(e => e.id));
  const rows1 = [head(["ردیف", "کد پروژه", "کار", "زمان شروع", "وضعیت", "زمان پایان", "شرح", "پاسخ", "ادامه یافته در ردیف"])];
  list.forEach((e, i) => {
    const r = i + 2, sv = isoToSerial(e.start), ev = isoToSerial(e.end);
    rows1.push(`<row r="${r}">${cn_("A" + r, e.id)}${/^[1-9]\d{0,14}$/.test(e.project) ? cn_("B" + r, e.project) : cs_("B" + r, e.project)}${cs_("C" + r, e.work)}${cn_("D" + r, sv, 2)}${cs_("E" + r, e.status)}${cn_("F" + r, ev, 2)}${cs_("G" + r, e.desc, 3)}${cs_("H" + r, e.reply, 3)}${cn_("I" + r, e.resumedBy && ids.has(e.resumedBy) ? e.resumedBy : null)}</row>`);
  });
  const rows2 = [head(["کد", "شرح", "نوع محاسبه"])];
  codes.forEach((c, i) => { const r = i + 2; rows2.push(`<row r="${r}">${cs_("A" + r, c.code)}${cs_("B" + r, c.desc)}${cs_("C" + r, c.type === "count" ? "تعداد" : "ساعت")}</row>`); });
  const rows3 = [head(["کد پروژه", "کار", "واحد", "مجموع", "تعداد ردیف", "باز"])];
  if (withSummary) {
    const typeOf = w => ((codes.find(c => c.code === w) || {}).type) || "hours", order = codes.map(c => c.code);
    const rk = w => order.indexOf(w) < 0 ? 99 : order.indexOf(w), groups = new Map();
    list.forEach(e => {
      const k = e.project + "\u0000" + e.work; let g = groups.get(k);
      if (!g) groups.set(k, g = { p: e.project, w: e.work, n: 0, open: 0, h: 0 });
      g.n++; if (isOpen(e)) g.open++; else g.h += hoursBetween(e.start, e.end) || 0;
    });
    [...groups.values()].sort((a, b) => a.p.localeCompare(b.p, undefined, { numeric: true }) || rk(a.w) - rk(b.w)).forEach((g, i) => {
      const r = i + 2, cnt = typeOf(g.w) === "count";
      rows3.push(`<row r="${r}">${/^[1-9]\d{0,14}$/.test(g.p) ? cn_("A" + r, g.p) : cs_("A" + r, g.p)}${cs_("B" + r, g.w)}${cs_("C" + r, cnt ? "عدد" : "ساعت")}${cn_("D" + r, cnt ? g.n : Math.round(g.h * 100) / 100)}${cn_("E" + r, g.n)}${cn_("F" + r, g.open || null)}</row>`);
    });
  }
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>${withSummary ? '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' : ""}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="ثبت روزانه" sheetId="1" r:id="rId1"/><sheet name="کدهای کار" sheetId="2" r:id="rId2"/>${withSummary ? '<sheet name="خلاصه" sheetId="3" r:id="rId4"/>' : ""}</sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${withSummary ? '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>' : ""}</Relationships>`],
    ["xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/></numFmts><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`],
    ["xl/worksheets/sheet1.xml", sheetXml(rows1, [8, 12, 9, 20, 10, 20, 50, 50, 14], 9)],
    ["xl/worksheets/sheet2.xml", sheetXml(rows2, [12, 36, 14], 3)],
  ].concat(withSummary ? [["xl/worksheets/sheet3.xml", sheetXml(rows3, [12, 10, 8, 12, 12, 8], 6)]] : []).map(([name, text]) => ({ name, data: enc.encode(text) }));
  return zipStore(files);
}

// ---------------------------------------------------------------- اجرا
const env = process.env;
const need = ["SUPABASE_URL", "SUPABASE_ANON_KEY", "APP_EMAIL", "APP_PASSWORD"];
if (!env.BACKUP_DRY_RUN) need.push("GMAIL_USER", "GMAIL_APP_PASSWORD", "BACKUP_TO");
const missing = need.filter(k => !env[k]);
if (missing.length) { console.error("Missing secrets: " + missing.join(", ")); process.exit(1); }

const base = env.SUPABASE_URL.replace(/\/$/, "");
const auth = await fetch(`${base}/auth/v1/token?grant_type=password`, {
  method: "POST", headers: { apikey: env.SUPABASE_ANON_KEY, "Content-Type": "application/json" },
  body: JSON.stringify({ email: env.APP_EMAIL, password: env.APP_PASSWORD }),
});
if (!auth.ok) { console.error("Login failed: HTTP " + auth.status); process.exit(1); }
const { access_token, user } = await auth.json();

const res = await fetch(`${base}/rest/v1/app_state?select=data,updated_at&user_id=eq.${user.id}`, {
  headers: { apikey: env.SUPABASE_ANON_KEY, Authorization: "Bearer " + access_token },
});
if (!res.ok) { console.error("Read failed: HTTP " + res.status); process.exit(1); }
const row = (await res.json())[0];
if (!row) { console.error("No data found on the server."); process.exit(1); }

const entries = row.data.entries || [], codes = row.data.codes || [];
const buf = Buffer.from(await buildXlsx(entries, codes, false).arrayBuffer());

const parts = Object.fromEntries(new Intl.DateTimeFormat("en-GB", {
  timeZone: "Asia/Tehran", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23",
}).formatToParts(new Date()).map(p => [p.type, p.value]));
const name = `task-tracker_${parts.year}-${parts.month}-${parts.day}_${parts.hour}-${parts.minute}.xlsx`;
fs.mkdirSync("backup", { recursive: true });
fs.writeFileSync(`backup/${name}`, buf);
console.log(`Built ${name}: ${entries.length} entries, ${buf.length} bytes`);

if (env.BACKUP_DRY_RUN) { console.log("Dry run: email skipped."); process.exit(0); }

const { default: nodemailer } = await import("nodemailer");
const transport = nodemailer.createTransport({ host: "smtp.gmail.com", port: 465, secure: true, auth: { user: env.GMAIL_USER, pass: env.GMAIL_APP_PASSWORD } });
await transport.sendMail({
  from: `Task Tracker Backup <${env.GMAIL_USER}>`,
  to: env.BACKUP_TO,
  subject: `Task Tracker backup ${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}`,
  text: `Daily backup of Task Tracker.\nFile: ${name}\nEntries: ${entries.length}\nLast change on server: ${row.updated_at}\n`,
  attachments: [{ filename: name, content: buf }],
});
console.log("Email sent to the configured address.");
