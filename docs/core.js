/* CHECK EDI Able — ตรรกะตรวจ MANIFEST (.xls/.xlsx หรือ PDF "CARGO MANIFEST") เทียบ ENTER (.pdf ฟอร์มของ Able)
   พอร์ตจาก build_edi.py (_pdf_manifest_rows + layout B แบบ Able) ทำงานในเบราว์เซอร์ล้วน ๆ ไม่มีการอัปโหลด
   ฐานเดียวกับ Check EDI MAC-NELS / Penanshin / Quality ต่างกันที่ตัวอ่าน ENTER และรับ MANIFEST เป็น PDF ได้ */
(function (root) {
  "use strict";

  // ───────────── regex / ตัวช่วยร่วม ─────────────
  const CONT_RE = /[A-Z]{4}\d{6,7}/;
  const TRANSIT_RE = /IN[\s-]?TRANSIT|TRANS[\s-]?S?HIP(?:MENT)?|TRANSIT\s*PORT/i;
  const DG_RE = /\bCLASS\s*[0-9]|\bUN\s*[- ]?\d{3,4}\b/i;
  const REEFER_RE = /REEFER|TEMP|อุณหภูมิ/i;
  const WARN = "⚠", OK = "✔";

  function findDg(...txts) {
    for (const t of txts) {
      if (t && DG_RE.test(t)) {
        const cl = t.match(/CLASS\s*([0-9](?:\.[0-9])?)/i), un = t.match(/UN\s*[- ]?\s*(\d{3,4})/i);
        const s = [cl ? `CLASS ${cl[1]}` : "", un ? `UN ${un[1]}` : ""].filter(Boolean).join(" ");
        return s || t.trim();
      }
    }
    return "";
  }
  function findTemp(...txts) {
    for (const t of txts) {
      if (!t) continue;
      if (REEFER_RE.test(t)) {
        const m = t.match(/(-?\d{1,2}(?:\.\d)?)\s*['°]?\s*(?:deg\.?)?\s*C\b/i);
        return m ? m[1].replace(/^\+/, "") + "°C" : t.replace(/\s+/g, " ").trim();
      }
    }
    return "";
  }
  function extractTransitPhrase(txt) {
    if (!txt) return "";
    const m = txt.match(/(IN\s*TRANSIT|TRANS\s?S?HIP(?:MENT)?)/i);
    if (!m) return txt.replace(/\s+/g, " ").trim();
    const kw = m[1].toUpperCase().replace(/ /g, "").includes("TRANSIT") ? "INTRANSIT" : "TRANSHIPMENT";
    const rest = txt.slice(m.index + m[0].length);
    const tos = [...rest.matchAll(/\bTO\s+([A-Za-z][A-Za-z,.\s]*?)(?=\s+(?:VIA|BY|FROM)\b|["\.\-]|$)/gi)];
    if (tos.length) return `${kw} TO ${tos[tos.length - 1][1].replace(/\s+/g, " ").trim().replace(/^[ ,.]+|[ ,.]+$/g, "")}`;
    return kw;
  }

  // ───────────── 4 กฎเพิ่มเติม: SHED / ประเทศปลายทาง / CARGO MOVEMENT / TAX ID ─────────────
  const DEST_CODE = { "LAOS": "LA", "LAO PDR": "LA", "LAO": "LA", "MYANMAR": "MM", "CHINA": "CN",
    "MALAYSIA": "MY", "PHILIPPINES": "PH", "CAMBODIA": "KH", "MARSHALL ISLAND": "MH", "MARSHALL ISLANDS": "MH", "INDIA": "IN" };
  const LAOS_NAMES = ["LAOS", "LAO PDR", "LAO"];

  function findDestCountry(txt) {
    if (!txt) return "";
    const u = txt.toUpperCase();
    for (const name of Object.keys(DEST_CODE).sort((a, b) => b.length - a.length)) if (u.includes(name)) return name;
    return "";
  }
  function checkShed(portDischarge, hasDg, hasUsedEngine, destCountry, shedRaw, expectedShed) {
    const pd = (portDischarge || "").toUpperCase();
    const m = (shedRaw || "").match(/\d{3,4}/);
    const actual = m ? m[0].padStart(4, "0") : "";
    let expected = "", label = "";
    if (expectedShed && !LAOS_NAMES.includes(destCountry)) return [actual === expectedShed, expectedShed, "SHED ที่แจ้ง"];
    if (pd.includes("THLKR")) [expected, label] = ["0332", "THLKR"];
    else if (pd.includes("BMT")) [expected, label] = ["0110", "BMT"];
    else if (pd.includes("SCT")) [expected, label] = ["0302", "SCT"];
    else if (pd.includes("UNITHAI")) [expected, label] = ["0113", "UNITHAI"];
    else if (pd.includes("LAEM CHABANG") && hasDg) [expected, label] = ["2826", "DG ที่ DISCHARGE LAEM CHABANG"];
    else if (pd.includes("BANGKOK") && LAOS_NAMES.includes(destCountry)) [expected, label] = ["0124", "LAOS ที่ DISCHARGE BANGKOK"];
    else if (pd.includes("BANGKOK") && hasUsedEngine) [expected, label] = ["0126", "USED ENGINE ที่ DISCHARGE BANGKOK"];
    if (!expected) return [null, "", ""];
    return [actual === expected, expected, label];
  }
  function checkDestCode(transitText, destCodeRaw) {
    const country = findDestCountry(transitText);
    if (!country) return [null, "", ""];
    const expected = DEST_CODE[country] || "";
    if (!expected) return [null, "", ""];
    const m = (destCodeRaw || "").match(/\(([A-Za-z]{2})\)/);
    const actual = m ? m[1].toUpperCase() : "";
    return [actual === expected, expected, actual];
  }
  function checkMovement(transitText, destCountry, mvManifest, mvEnter) {
    if (!transitText) return [null, "", ""];
    const mv = (mvManifest || "").trim().toUpperCase();
    if (LAOS_NAMES.includes(destCountry)) {
      const exp = (mvEnter || "").trim().toUpperCase();
      if (!exp) return ["review", "", mv];
      return [(exp === "7" && mv.startsWith("7")) || (exp.startsWith("L") && mv.startsWith("L")), `ตาม ENTER = ${exp}`, mv];
    }
    return [mv.startsWith("7"), "7-TRANSIT", mv];
  }
  function checkTaxid(taxIdRaw, notifyName, enterNotify) {
    const m = (taxIdRaw || "").match(/\(([^)]*)\)/);
    const name = (m ? m[1] : "").trim();
    if (!name) return [null, "", notifyName || ""];
    const a = normCo(name), b = normCo(notifyName || "");
    let ok = !!(a && b && comatch(a, b));
    if (!ok && enterNotify) { const c = normCo(enterNotify); if (a && c && comatch(a, c)) ok = true; }
    return [ok, name, notifyName || ""];
  }
  function computeExtraChecks(m, e, expectedShed) {
    const hasUsedEngine = (m.desc || "").toUpperCase().includes("USED ENGINE");
    const destCountry = findDestCountry(m.transit);
    const out = {};
    let [ok, exp, lbl] = checkShed(m.port_discharge || "", !!m.dg, hasUsedEngine, destCountry, m.shed_no || "", expectedShed);
    const raw = m.shed_no || "";
    if (ok === null) out.shed = { ok: null, text: "-", note: null };
    else if (ok && lbl === "SHED ที่แจ้ง") out.shed = { ok: true, text: "-", note: null };
    else {
      const txt = raw + (ok ? "" : lbl === "SHED ที่แจ้ง" ? `  (ต้องเป็น ${exp})` : `  (ต้องเป็น ${exp} — ${lbl})`);
      out.shed = { ok, text: txt, note: ok ? null : `SHED NO. ไม่ตรง — ${lbl} ต้องเป็น ${exp} (MANIFEST: ${raw || "(ว่าง)"})` };
    }
    let act;
    [ok, exp, act] = checkDestCode(m.transit, m.dest_code_raw || "");
    const noRule = m.no_rule_lines;   // MANIFEST แบบ PDF ไม่พิมพ์บรรทัด 1)-4) เลย — ว่าง ≠ ผิด ต้องดูในระบบเอง
    if (ok === null) out.dest = { ok: null, text: "-", note: null };
    else if (!ok && noRule && !act) out.dest = { ok: "review", text: `ไม่มีใน MANIFEST PDF  (ต้องเป็น (${exp}) — ${destCountry})`,
      note: `ประเทศปลายทาง — MANIFEST PDF ไม่พิมพ์ช่องนี้ ต้องตรวจในระบบว่าเป็น (${exp})` };
    else {
      const disp = act ? `(${act})` : "(ว่าง)";
      out.dest = { ok, text: ok ? "-" : disp + `  (ต้องเป็น (${exp}) — ${destCountry})`,
        note: ok ? null : `ประเทศปลายทางไม่ตรง — ${destCountry} ต้องเป็น (${exp}) (MANIFEST: ${disp})` };
    }
    [ok, exp, act] = checkMovement(m.transit, destCountry, m.movement_raw || "", (e || {}).movement_raw || "");
    const rawMv = m.movement_raw || "", entMv = (e || {}).movement_raw || "";
    if (ok === null) out.mv = { ok: null, text: "-", note: null };
    else if (!ok && noRule && !rawMv) {
      const expD = exp || `ตาม ENTER = ${entMv || "?"}`;
      out.mv = { ok: "review", text: `ไม่มีใน MANIFEST PDF  (ต้องเป็น ${expD}` + (entMv ? ` ; ENTER = ${entMv})` : ")"),
        note: `CARGO MOVEMENT — MANIFEST PDF ไม่พิมพ์ช่องนี้ ต้องตรวจในระบบว่าเป็น ${expD}` };
    }
    else if (ok === "review") out.mv = { ok: "review", text: (rawMv || "(ว่าง)") + "  (ต้องเทียบ ENTER)",
      note: "CARGO MOVEMENT ไปลาว — ไม่พบค่า CARGO MOVEMENT ใน ENTER ของ B/L นี้ ต้องตรวจสอบด้วยคน" };
    else out.mv = { ok, text: ok ? "-" : (rawMv || "(ว่าง)") + `  (ต้องเป็น ${exp})`,
      note: ok ? null : `CARGO MOVEMENT ไม่ตรง — ต้องเป็น ${exp} (MANIFEST: ${rawMv || "(ว่าง)"})` };
    const enterNotify = m.transit ? (e || {}).notify : null;
    let name, notify;
    [ok, name, notify] = checkTaxid(m.tax_id_raw || "", m.notify, enterNotify);
    if (ok === null) out.tax = { ok: null, text: "-", note: null };
    // ตรงตามกฎที่ล็อคไว้ → '-' (เหมือน SHED) โชว์ค่าจริงเฉพาะตอนไม่ตรง
    else out.tax = { ok, text: ok ? "-" : name + `  (N:/NOTIFY = ${notify || "(ว่าง)"})`,
      note: ok ? null : `TAX ID ไม่ตรงกับ N:/NOTIFY PARTY — TAX ID: ${name} / N: ${notify || "(ว่าง)"}` + (enterNotify ? ` / ENTER NOTIFY: ${enterNotify}` : "") };
    return out;
  }

  // ───────────── ตัวช่วยเทียบข้อมูล ─────────────
  const PKG_KIND = { CS: "CASE", PX: "PALLET", PL: "PALLET", CT: "CARTON", CTN: "CARTON", BX: "BOX", PK: "PACKAGE", PKG: "PACKAGE",
    SX: "SET", DR: "DRUM", RO: "ROLL", RL: "ROLL", BG: "BAG", BE: "BALE", BL: "BALE", CR: "CRATE", UN: "UNIT", PC: "PIECE" };
  function canonPkg(s) {
    if (!s) return "";
    const u = String(s).toUpperCase().replace(/\bCTNS?\b/g, "CARTON").replace(/\bPKGS?\b/g, "PACKAGE")
      .replace(/\bPLTS?\b/g, "PALLET");   // MAC-NELS: '2 PLTS'
    for (const w of ["WOODEN CASE", "PALLET", "CARTON", "PACKAGE", "CASE", "BOX", "SET", "DRUM", "ROLL", "BAG", "BALE", "CRATE", "UNIT", "PIECE", "SKID", "BUNDLE"])
      if (u.includes(w)) return w === "WOODEN CASE" ? "CASE" : w;
    let m = u.match(/^\s*[\d,]*\s*([A-Z]{2,3})\b/);
    if (m && PKG_KIND[m[1]]) return PKG_KIND[m[1]];
    m = u.match(/\(([A-Z]+?)\(S\)\)/) || u.match(/\(([A-Z ]+?)\)/);
    if (m) return m[1].trim().replace(/S+$/, "") || m[1].trim();
    return u.replace(/[\d,()]/g, "").trim();
  }
  function canonStatus(s) {
    if (!s) return "";
    const th = String(s);
    if (th.includes("เปิดตู้")) return "LCL/CFS";
    if (th.includes("ลากตู้")) return "CY";
    if (th.includes("ขน")) return "LCL";
    const u = th.replace(/STATUS/gi, "").replace(/[^A-Za-z/]/g, "").toUpperCase();
    if (!u) return th.trim();
    if (u.includes("CFS")) return "LCL/CFS";
    if (u === "LCL") return "LCL";
    if (u.includes("CY") || u.includes("FCL")) return "CY";
    if (u.includes("LCL")) return "LCL";
    return u;
  }
  function num(s) {
    const m = String(s == null ? "" : s).match(/-?\d[\d,]*\.?\d*/);
    return m ? parseFloat(m[0].replace(/,/g, "")) : null;
  }
  function pkgnum(s) {
    const m = String(s == null ? "" : s).match(/\d[\d,]*/);
    return m ? parseInt(m[0].replace(/,/g, ""), 10) : null;
  }
  function normCo(s) {
    return String(s || "").toUpperCase().replace(/[^A-Z0-9 ]/g, " ")
      .replace(/\b(CO|LTD|COMPANY|LIMITED|PUBLIC|CORP|CORPORATION|INC|GROUP|THE)\b/g, " ").replace(/\s+/g, " ").trim();
  }
  function comatch(a, b) {
    const sa = new Set(a.split(" ").filter(Boolean)), sb = new Set(b.split(" ").filter(Boolean));
    if (!sa.size || !sb.size) return false;
    const inter = [...sa].filter(x => sb.has(x));
    const subA = [...sa].every(x => sb.has(x)), subB = [...sb].every(x => sa.has(x));
    if (sa.size === sb.size && subA) return true;
    if (inter.length && (subA || subB) && Math.min(sa.size, sb.size) >= 2) return true;
    return inter.length >= 2;
  }
  function sameDetail(a, b) {
    // 'CARGO MOVEMENT n' เป็นรหัสศุลกากร ไม่ใช่รายละเอียดสินค้า — MANIFEST บางไฟล์พิมพ์ไว้ท้าย DESC ไม่นับในการเทียบ
    const mv = s => String(s || "").toUpperCase().replace(/\*?\s*CARGO\s*MOVEMENT\s*[0-9A-Z\-]+\s*\*?/g, " ");
    const na = mv(a).replace(/[^A-Z0-9]/g, ""), nb = mv(b).replace(/[^A-Z0-9]/g, "");
    if (!na && !nb) return null;
    return na === nb;
  }
  const fmt3 = x => x.toLocaleString("en-US", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const r2 = x => Math.round(x * 100) / 100, r3 = x => Math.round(x * 1000) / 1000;
  const blSort = (a, b) => a.length - b.length || (a < b ? -1 : a > b ? 1 : 0);

  // ───────────── 1) MANIFEST.xls ─────────────
  function cell(row, i) {
    if (i >= row.length) return "";
    let v = row[i];
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && Number.isInteger(v)) v = String(v);
    return String(v).replace(/[\x00-\x1f]+/g, " ").trim();
  }
  function detectBlPrefix(rows) {
    const cand = [];
    for (const row of rows) {
      const v = row && row[0];
      if (typeof v === "number" || v == null) continue;
      const m = String(v).trim().match(/^([A-Z]{3,8})[A-Z0-9]{5,}$/);
      if (m) cand.push(m[1]);
    }
    if (cand.length) {
      let p = cand[0];
      for (const c of cand) { let i = 0; while (i < p.length && i < c.length && p[i] === c[i]) i++; p = p.slice(0, i); }
      p = (p.match(/^[A-Z]*/) || [""])[0];
      if (p.length < 3) {
        const cnt = {}; cand.forEach(c => cnt[c] = (cnt[c] || 0) + 1);
        p = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
      }
      return new RegExp(`^${p}[A-Z0-9]{5,}$`);
    }
    return /^[A-Z]{3,8}[A-Z0-9]{6,}$/;
  }
  function detectMarksCol(rows) {
    const counts = {};
    for (const row of rows) row.forEach((v, idx) => { if (/^\d+\.\s*[A-Z]{4}\d{6,7}/.test(String(v == null ? "" : v).trim())) counts[idx] = (counts[idx] || 0) + 1; });
    const k = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
    return k === undefined ? 5 : +k;
  }

  // MANIFEST แบบ PDF (CARGO MANIFEST พิมพ์แนวนอน) → แถว/คอลัมน์เลขเดียวกับ .xls ตามขอบซ้ายของข้อความ:
  // B/L,S:,C:,N: ≈35 | ค่าหัวตาราง ≈123 | MARKS ≈219 | NO. OF PKGS/DESC ≈390 | KGM/MTQ ≈638 (เหมือน _pdf_manifest_rows ใน build_edi.py)
  const PDF_MAN_COLS = [[82, 0], [202, 1], [322, 5], [592, 10], [702, 16], [1e9, 18]];
  function isManifestPdf(pages) {
    const all = pages.slice(0, 2).flatMap(pg => pg.items.map(it => it.t)).join(" ");
    return /CARGO MANIFEST/i.test(all) && /BILL OF LANDING NO/i.test(all);
  }
  function manifestRowsFromPdf(pages) {
    const rows = [];
    for (const pg of pages) {
      const its = pg.items.slice().sort((a, b) => a.y - b.y || a.x - b.x);
      const groups = [];
      for (const it of its) {
        const g = groups[groups.length - 1];
        if (g && it.y - g.y <= 5) g.items.push(it); else groups.push({ y: it.y, items: [it] });
      }
      for (const g of groups) {
        const row = new Array(19).fill("");
        for (const it of g.items.sort((a, b) => a.x - b.x)) {
          const col = PDF_MAN_COLS.find(([lim]) => it.x < lim)[1];
          row[col] = (row[col] + " " + it.t).trim();
        }
        rows.push(row);
      }
    }
    rows.fromPdf = true;
    return rows;
  }

  function parseManifest(rowsIn) {
    const fromPdf = !!rowsIn.fromPdf;
    const rows = rowsIn.map(r => r.map(v => (v === null || v === undefined ? "" : v)));
    const blRe = detectBlPrefix(rows), MARKS_COL = detectMarksCol(rows);
    const n = rows.length;
    let vessel = "";
    const declared = {};
    const MAN = {};
    const END = /^(PORT TOTAL|GRAND TOTAL|TOTAL\b|PortOfDischarge)/i;
    const ST_TOK = ["CY/CY", "CY-CY", "LCL/CFS", "CFS/CFS", "FCL/CFS", "CY", "LCL", "CFS", "FCL"];
    const pdAt = new Array(n).fill("");
    let curPd = "";
    rows.forEach((row, ridx) => {
      const v0 = cell(row, 0);
      if (v0.toUpperCase() === "VESSEL & VOYAGE" && !vessel) {
        for (const c of row.slice(1)) { const cs = String(c).trim(); if (cs) { vessel = cs; break; } }
      }
      for (let k = 0; k < row.length; k++) {
        if (cell(row, k).toUpperCase() === "PORT OF DISCHARGE") {
          for (let k2 = k + 1; k2 < row.length; k2++) { const v = cell(row, k2); if (v) { curPd = v; break; } }
          break;
        }
      }
      pdAt[ridx] = curPd;
      const rowTxt = row.map((_, k) => cell(row, k)).join(" ");
      if (/GRAND TOTAL/i.test(rowTxt)) {
        const joined = rows.slice(ridx, ridx + 4).map(r2_ => r2_.map((_, k) => cell(r2_, k)).join(" ")).join(" ");
        const mpk = joined.match(/TOTAL PACK\s*:\s*([\d,]+)/i), mgw = joined.match(/TOTAL GROSSWEIGHT\s*:\s*([\d,.]+)/i),
          mme = joined.match(/TOTAL MEASUREMENT\s*:\s*([\d,.]+)/i);
        if (mpk) declared.pkg = parseInt(mpk[1].replace(/,/g, ""), 10);
        if (mgw) declared.gw = parseFloat(mgw[1].replace(/,/g, ""));
        if (mme) declared.meas = parseFloat(mme[1].replace(/,/g, ""));
      }
    });

    let i = 0;
    while (i < n) {
      const bl = cell(rows[i], 0);
      if (!blRe.test(bl)) { i++; continue; }
      const rec = { bl, pkg_hdr: cell(rows[i], 10), gw: cell(rows[i], 16), cons: "", notify: "", cont: [], status: "", meas: "",
        marks: [], desc: [], transit: "", dg: "", reefer: "", port_discharge: pdAt[i], shed_no: "", dest_code_raw: "", movement_raw: "", tax_id_raw: "", no_rule_lines: fromPdf };
      const h5 = cell(rows[i], MARKS_COL);
      if (h5 && !/^\d+\.\s*([A-Z]{4}\d{6,7}|\s*$)/.test(h5)) rec.marks.push(h5);
      let transitOpen = false;
      let j = i + 1;
      while (j < n) {
        const a = cell(rows[j], 0);
        if (blRe.test(a)) break;
        if (rows[j].some((_, k) => END.test(cell(rows[j], k)))) { j++; break; }
        const numbered = /^[123]\./.test(a);
        if (a.startsWith("C :")) rec.cons = a.slice(3).trim();
        else if (a.startsWith("N :")) rec.notify = a.slice(3).trim();
        else if (a.startsWith("S :")) { /* shipper */ }
        else if (rec.cons && !rec.notify && !numbered) rec.cons += " " + a;
        else if (rec.notify && !numbered) rec.notify += " " + a;
        const c5 = cell(rows[j], MARKS_COL);
        if (c5) {
          const mm = c5.match(/^\d+\.\s*([A-Z]{4}\d{6,7})/);
          if (mm) rec.cont.push(mm[1]);
          else if (!/^\d+\.\s*$/.test(c5)) rec.marks.push(c5);
        }
        for (const cc of [cell(rows[j], MARKS_COL), cell(rows[j], 10)]) {
          if (cc && DG_RE.test(cc) && !rec.dg) rec.dg = findDg(cc);
          if (cc && REEFER_RE.test(cc) && !rec.reefer) rec.reefer = findTemp(cc);
        }
        const c10 = cell(rows[j], 10);
        if (c10) {
          const u = c10.toUpperCase();
          if (DG_RE.test(c10) || REEFER_RE.test(c10)) rec.desc.push(c10);
          else if (transitOpen || TRANSIT_RE.test(u)) { rec.transit = (rec.transit + " " + c10).trim(); transitOpen = true; rec.desc.push(c10); }
          else if (ST_TOK.includes(u) && !rec.status) rec.status = c10;
          // รหัสประเภทตู้ ISO (45G0X1) — ห้ามกินเลข invoice อย่าง 'V266411' (เหมือน build_edi.py)
          else if (/^\d{0,2}[A-Z]\d[A-Z0-9]{1,3}(?:\s|$)/.test(u) || u.includes("PART CNTR") || /^[\d,]+\s*(PK|CS|PX|CT|BX|SX|BE|DR|RO|BG)\b/.test(u)) { /* โครงสร้าง */ }
          else rec.desc.push(c10);
        }
        const c16 = cell(rows[j], 16);
        if (c16 && c16.toUpperCase().includes("MTQ") && !rec.meas) rec.meas = c16;
        const c15 = cell(rows[j], 15);
        if (c15) {
          if (c15.startsWith("1)") && !rec.shed_no) { const msh = c15.match(/\d{3,4}/); rec.shed_no = msh ? msh[0] : ""; }
          else if (c15.startsWith("2)") && !rec.dest_code_raw) rec.dest_code_raw = c15;
          else if (c15.startsWith("3)") && !rec.movement_raw) { const mmv = c15.match(/\(([^)]*)\)/); rec.movement_raw = mmv ? mmv[1].trim() : ""; }
          else if (c15.startsWith("4)") && !rec.tax_id_raw) rec.tax_id_raw = c15;
        }
        j++;
      }
      rec.marks = rec.marks.join(" ");
      rec.desc = rec.desc.join(" ");
      rec.cont = [...new Set(rec.cont)].sort();
      MAN[bl] = rec;
      i = j;
    }
    return { MAN, vessel, declared, blRe };
  }

  // ───────────── 2) ENTER.pdf (ฟอร์มของ Able) ─────────────
  // หัวตาราง: MARKS & NOS | DESCRIPTIONS OF GOODS | GROSS WT.& MEAS ; ต่อบล็อก:
  //   '<n> PALLETS'  ...  '6,405.000 KGS.'      ← แถวจำนวน + น้ำหนัก (ขวา)
  //                       '13.450 CBM'          ← แถวถัดไป: ปริมาตร
  //   marks (ซ้าย) | description (กลาง)  — หมายเหตุ '" CARGO MOVEMENT = 7 "' / '"... CODE 7 "' = รหัส CARGO MOVEMENT
  //   'CONSIGNEE :' ชื่อ (หลายบรรทัด) [ 'Notify Party :' ชื่อ (หลายบรรทัด) ]
  //   'CONT NO.  :' เลขตู้/ขนาด
  //   'B/L  NO.  :' เลข B/L + ตัวอักษรต่อท้ายแยกเป็นอีกคำทางขวา ('HASLK…293' 'A')  ← จบบล็อก
  // บล็อกตัดข้ามหน้าได้ — ต่อแถวทุกหน้าเป็นลำดับเดียว (พอร์ตจาก layout B แบบ Able ใน build_edi.py)
  const PAGELN = /^Page \d+ of \d+$/i;
  const LB_CONS = /^CONSIGNEE\s*:?$/i, LB_NOTIFY = /^NOTIFY\s*PARTY\s*:?$/i, LB_CONT = /^CONT\.?\s*NO\.?\s*:?$/i, LB_BL = /^B\s*\/\s*L\s*NO\.?\s*:?$/i;
  const QTY_RE = /^([\d,]+)\s+([A-Za-z].*)$/;
  const MV_NOTE = /CARGO\s*MOVEMENT\s*=?\s*([0-9A-Z])\b|\bCODE\s*([0-9])\b/i;
  const ALL_BL = /ทุก\s*-?\s*B\s*\/?\s*L|\bALL\s+B\s*\/?\s*LS?\b|\bEVERY\s+B\s*\/?\s*L\b/i;
  const MID = 155, RIGHT = 330;   // ระยะจากขอบป้าย CONSIGNEE: < MID = MARKS/ป้าย, MID..RIGHT = DESC/ค่า, ≥ RIGHT = น้ำหนัก/ตัวอักษร B/L

  // แถวของหน้า → segment (คำที่ติดกัน ช่องว่าง ≤ 20pt และไม่ข้ามเส้นแบ่งคอลัมน์ MARKS|DESC)
  function pageRows(pg) {
    const items = pg.items.filter(it => it.t && !PAGELN.test(it.t) && !(/^\d+\s*\/\s*\d+$/.test(it.t) && it.x > 400))
      .sort((a, b) => a.y - b.y || a.x - b.x);
    const lab = items.find(it => /^CONSIGNEE/i.test(it.t));
    const L = lab ? lab.x : 57;
    const rows = [];
    for (const it of items) {
      const r = rows[rows.length - 1];
      if (r && Math.abs(it.y - r.y) <= 3) r.items.push(it); else rows.push({ y: it.y, items: [it] });
    }
    return rows.map(r => {
      const segs = [];
      for (const it of r.items.sort((a, b) => a.x - b.x)) {
        const s = segs[segs.length - 1];
        const cross = s && s.x - L < MID && it.x - L >= MID;
        if (s && !cross && it.x - s.end <= 20) { s.t += " " + it.t; s.end = it.x + it.w; }
        else segs.push({ x: it.x, rel: it.x - L, t: it.t, end: it.x + it.w });
      }
      segs.forEach(s => { s.t = s.t.replace(/\s+/g, " ").trim(); });
      return { y: r.y, segs };
    });
  }

  // ฟอร์มของ Able: หัวตาราง MARKS & NOS + ป้าย 'CONT NO.' (มีวรรค ไม่ใช่ 'CONT.NO.' ของ Quality) + 'B/L NO.'
  function isAbleForm(pages) {
    const rows = pages.flatMap(pageRows);
    const lead = re => rows.some(r => r.segs.length && re.test(r.segs[0].t));
    const all = pages.flatMap(pg => pg.items.map(it => it.t)).join(" ");
    return /MARKS\s*&\s*NOS/i.test(all) && lead(/^CONT\s+NO\.?\s*:?$/i) && lead(LB_BL) && lead(LB_CONS);
  }

  // pages: [{items:[{x,y,w,t}], annots:[{y,text}]}]  (y วัดจากบนลงล่าง) — ENTER หลายไฟล์ให้ต่อ pages กันมาได้เลย
  function parseEnter(pages, blRe) {
    if (!isAbleForm(pages)) return { ENT: {}, declared: {}, notAble: true };

    const blocks = [];
    let cur = null;
    const fresh = () => ({ rows: [], cons: [], notify: [], cont: "", state: "goods" });
    pages.forEach((pg, p) => {
      for (const r of pageRows(pg)) {
        if (!r.segs.length) continue;
        if (!cur) cur = fresh();
        const s0 = r.segs[0], lead = s0.rel < MID ? s0.t : "";
        const rest = r.segs.filter(s => s !== s0 || !lead);
        if (LB_BL.test(lead)) {
          const mid = rest.filter(s => s.rel < RIGHT).map(s => s.t).join(" ");
          const sfx = rest.find(s => s.rel >= RIGHT && /^[A-Z]{1,2}$/.test(s.t));
          let bl = (mid.split(/\s+/).find(t => blRe.test(t)) || "");
          if (bl && sfx && blRe.test(bl + sfx.t)) bl += sfx.t;
          if (bl) blocks.push(Object.assign(cur, { bl, p, y: r.y }));
          cur = null;
          continue;
        }
        if (LB_CONS.test(lead)) { cur.state = "cons"; cur.cons.push(...rest.map(s => s.t)); continue; }
        if (LB_NOTIFY.test(lead)) { cur.state = "notify"; cur.notify.push(...rest.map(s => s.t)); continue; }
        if (LB_CONT.test(lead)) { cur.state = "cont"; cur.cont = rest.map(s => s.t).join(" "); continue; }
        if (cur.state === "cons" && !lead) { cur.cons.push(...rest.map(s => s.t)); continue; }
        if (cur.state === "notify" && !lead) { cur.notify.push(...rest.map(s => s.t)); continue; }
        if (cur.state === "goods") cur.rows.push(r);
      }
    });

    // FreeText ที่ลูกค้าพิมพ์ทับ: 'ทุก BL' = ทุก B/L ; ช่วง '<ฐาน>-<ตัวจบ>' ; อื่น ๆ = บล็อกที่แถว B/L อยู่ใต้ข้อความ
    const annot = {}, docWide = [];
    const applyAnnot = (bl, txt) => {
      const r = annot[bl] || (annot[bl] = { status: "", transit: "", dg: "", reefer: "" });
      if (DG_RE.test(txt)) r.dg = findDg(txt);
      else if (REEFER_RE.test(txt)) r.reefer = findTemp(txt);
      else if (TRANSIT_RE.test(txt)) r.transit = txt;
      else r.status = (r.status + " " + txt).trim();
    };
    const RANGE1 = /^([A-Z0-9]{6,}?)-([A-Z])\s+([\s\S]+)$/, RANGE2 = /^([\s\S]+?)\s+([A-Z0-9]{6,}?)-([A-Z])$/;
    pages.forEach((pg, p) => {
      const pageBlocks = blocks.filter(b => b.p === p).sort((a, b) => a.y - b.y);
      for (const a of pg.annots || []) {
        const txt = (a.text || "").replace(/\s+/g, " ").trim();
        if (!txt) continue;
        if (ALL_BL.test(txt)) { docWide.push(txt.replace(new RegExp(ALL_BL.source, "gi"), "").replace(/^[\s-]+|[\s-]+$/g, "")); continue; }
        const r1 = txt.match(RANGE1), r2 = r1 ? null : txt.match(RANGE2);
        const rg = r1 && blRe.test(r1[1]) ? { base: r1[1], endL: r1[2], rest: r1[3] }
          : r2 && blRe.test(r2[2]) ? { base: r2[2], endL: r2[3], rest: r2[1] } : null;
        if (rg) {
          for (let c = 64; c <= rg.endL.charCodeAt(0); c++) applyAnnot(rg.base + (c === 64 ? "" : String.fromCharCode(c)), rg.rest);
          continue;
        }
        const tgt = pageBlocks.find(b => b.y > a.y - 5) || pageBlocks[pageBlocks.length - 1];
        if (tgt) applyAnnot(tgt.bl, txt);
      }
    });

    const ENT = {};
    for (const b of blocks) {
      // แถวจำนวน: segment กลางเป็น '<n> <ชนิด>' — ก่อนหน้านั้น (เช่น '= AS PER ATTACHED SHEET =') ไม่ใช่ของบล็อกนี้
      const qi = b.rows.findIndex(r => r.segs.some(s => s.rel >= MID && s.rel < RIGHT && QTY_RE.test(s.t)));
      let pkgs = null, pkgtype = "", gw = null, meas = null;
      const marks = [], desc = [];
      let movement = "";
      if (qi >= 0) {
        const q = b.rows[qi].segs.find(s => s.rel >= MID && s.rel < RIGHT && QTY_RE.test(s.t)).t.match(QTY_RE);
        pkgs = parseInt(q[1].replace(/,/g, ""), 10);
        pkgtype = q[2].trim().replace(/\bPLTS?\b/gi, "PALLET").replace(/\bPKGS?\b/gi, "PACKAGE");
        const w = b.rows[qi].segs.find(s => s.rel >= RIGHT);
        gw = w ? num(w.t) : null;
        let start = qi + 1;
        const mrow = b.rows[start];
        const mt = mrow ? mrow.segs.filter(s => s.rel >= RIGHT).map(s => s.t).join(" ") : "";
        if (/^[\d,]+(\.\d+)?\s*(CBM\.?)?$/i.test(mt)) { meas = num(mt); start++; }
        for (const r of b.rows.slice(start)) {
          for (const s of r.segs) {
            let t = s.t;
            const mv = t.includes('"') ? t.match(MV_NOTE) : null;
            if (mv) {
              movement = (mv[1] || mv[2]).toUpperCase();
              t = t.replace(new RegExp(MV_NOTE.source, "gi"), "").replace(/\s+/g, " ").trim();
              if (!t.replace(/["\s]/g, "")) continue;
            }
            (s.rel < MID ? marks : desc).push(t);
          }
        }
      }
      const marksT = marks.join(" "), descT = desc.join(" "), combined = `${marksT} ${descT}`;
      // หมายเหตุ MOVEMENT ที่พิมพ์ใต้ Notify Party ก็นับ (และไม่ปนเข้าชื่อ notify)
      const notifyL = [];
      for (const t of b.notify) {
        const mv = t.includes('"') ? t.match(MV_NOTE) : null;
        if (mv) movement = (mv[1] || mv[2]).toUpperCase(); else notifyL.push(t);
      }
      const an = annot[b.bl] || {};
      let status = an.status || "", transit = an.transit || "", dg = an.dg || findDg(combined), reefer = an.reefer || findTemp(combined);
      for (const t of docWide) {   // คำสั่ง "ทุก BL" ใช้เมื่อ B/L นั้นไม่มีคำสั่งเฉพาะตัวเอง
        if (DG_RE.test(t)) dg = dg || findDg(t);
        else if (REEFER_RE.test(t)) reefer = reefer || findTemp(t);
        else if (TRANSIT_RE.test(t)) transit = transit || t;
        else if (!status) status = t;
      }
      const cm = (b.cont || "").match(CONT_RE);
      ENT[b.bl] = { bl: b.bl, cons: b.cons.join(" ").replace(/\s+/g, " ").trim(), notify: notifyL.join(" ").replace(/\s+/g, " ").trim(),
        cont: cm ? cm[0] : (b.cont || "").trim(), pkgs, pkgtype, gw, meas, marks: marksT, desc: descT, combined_text: true,
        status_raw: status, transit_raw: transit, dg, reefer, movement_raw: movement };
    }

    // ยอดรวมที่เอกสารระบุเอง: 'TOTAL' / '119 PACKAGES' / '47,948.880' / 'KGS.' / '109.940' / 'CBM'
    const declared = {};
    const full = pages.flatMap(pageRows).map(r => r.segs.map(s => s.t).join(" ")).join("\n");
    const NUM = "(\\d[\\d,]*\\.?\\d*)";
    const tm = full.match(new RegExp(`\\bTOTAL\\s*${NUM}\\s*PACKAGES\\s*${NUM}\\s*KGS\\.?\\s*${NUM}\\s*CBM`, "i"));
    if (tm) {
      declared.pkg = parseInt(tm[1].replace(/,/g, ""), 10);
      declared.gw = parseFloat(tm[2].replace(/,/g, ""));
      declared.meas = parseFloat(tm[3].replace(/,/g, ""));
    }
    return { ENT, declared };
  }

  // ───────────── 3) รายงาน ─────────────
  const KEYS = ["urgent", "bl", "cneeE", "cneeM", "tax", "cntE", "cntM", "stE", "stM", "tpE", "pkE", "tpM", "pkM",
    "gwE", "gwM", "meE", "meM", "mkE", "mkM", "dsE", "dsM", "reefer", "dg", "transit", "shed", "dest", "mv", "note"];
  const HEADERS = ["สรุปผลเร่งด่วน", "B/L NO.", "CNEE. (ENTER)", "CNEE. (MANIFEST)", "TAX ID vs NOTIFY", "CNTRS NO. (ENTER)", "CNTRS NO. (MANIFEST)",
    "STATUS (ENTER)", "STATUS (MANIFEST)", "TOTAL PKG (ENTER)", "PKG. (ENTER)", "TOTAL PKG (MANIFEST)", "PKG. (MANIFEST)",
    "G.W. KGM (ENTER)", "G.W. KGM (MANIFEST)", "MEAS. MTQ (ENTER)", "MEAS. MTQ (MANIFEST)", "MARKS (ENTER)", "MARKS (MANIFEST)",
    "DESC. (ENTER)", "DESC. (MANIFEST)", "REEFER TEMP", "DG", "TRANSIT/TRANSHIPMENT", "SHED NO.", "ประเทศปลายทาง", "CARGO MOVEMENT", "หมายเหตุ / จุดที่ไม่ตรงกัน"];
  const NUMKEYS = new Set(["gwE", "gwM", "meE", "meM"]), WIDEKEYS = new Set(["mkE", "mkM", "dsE", "dsM"]);
  const LBL = { cneeE: "CONSIGNEE", cneeM: "CONSIGNEE", cntE: "CONTAINER", cntM: "CONTAINER", stE: "STATUS", stM: "STATUS",
    tpE: "PACKAGE", tpM: "PACKAGE", pkE: "PACKAGING", pkM: "PACKAGING", gwE: "G.W.", gwM: "G.W.", meE: "MEAS.", meM: "MEAS.",
    mkE: "MARKS", mkM: "MARKS", dsE: "DESC", dsM: "DESC", reefer: "REEFER", dg: "DG", transit: "TRANSIT",
    shed: "SHED", dest: "ปลายทาง", mv: "MOVEMENT", tax: "TAX ID" };
  const OWN_URGENT = new Set(["reefer", "dg", "transit", "shed", "dest", "mv", "tax"]);

  function buildReport(MAN, ENT, vessel, manDecl, entDecl, opts) {
    opts = opts || {};
    const expectedShed = (opts.expectedShed || "").replace(/\D/g, "").padStart(4, "0").slice(-4);
    const exShed = opts.expectedShed && /\d{3,4}/.test(opts.expectedShed) ? opts.expectedShed.match(/\d{3,4}/)[0].padStart(4, "0") : "";
    const rows = [], crit = [];
    let issueRows = 0;
    const TOT = { pe: 0, pm: 0, ge: 0, gm: 0, me: 0, mm: 0, pke: {}, pkm: {}, cte: {}, ctm: {}, st_bad: 0, transit: 0, dg: 0, reefer: 0,
      shed_bad: 0, dest_bad: 0, mv_bad: 0, mv_review: 0, tax_bad: 0 };
    const sub = Object.keys(MAN).sort(blSort);
    const mkCell = (v, s, bad) => ({ v: v === null || v === undefined ? "" : v, s: s || "ok", bad: !!bad });

    for (const bl of sub) {
      const m = MAN[bl], e = ENT[bl];
      const notes = [], urgent = [], cs = {};
      const mk = (keys, ok) => keys.forEach(k => { cs[k] = ok ? "ok" : "diff"; });

      if (!e) {
        const pm_ = pkgnum(m.pkg_hdr), gm_ = num(m.gw), mm_ = num(m.meas);
        const xc = computeExtraChecks(m, null, exShed);
        const extraNotes = ["shed", "dest", "mv", "tax"].map(k => xc[k].note).filter(Boolean);
        const vals = { urgent: `${WARN} ไม่มีเอกสาร ENTER เทียบ`, bl: `${WARN} ${bl}`, cneeE: "—", cneeM: m.cons || m.notify, tax: xc.tax.text,
          cntE: "—", cntM: m.cont.join(";"), stE: "—", stM: canonStatus(m.status) || m.status || "(ว่าง)",
          tpE: "—", pkE: "—", tpM: pm_ !== null ? pm_ : (m.pkg_hdr || "-"), pkM: canonPkg(m.pkg_hdr) || "-",
          gwE: "—", gwM: gm_ !== null ? fmt3(gm_) : "-", meE: "—", meM: mm_ !== null ? fmt3(mm_) : "-",
          mkE: "—", mkM: m.marks, dsE: "—", dsM: m.desc, reefer: "-", dg: "-", transit: extractTransitPhrase(m.transit) || "-",
          shed: xc.shed.text, dest: xc.dest.text, mv: xc.mv.text, note: "ไม่พบ B/L นี้ในเอกสาร ENTER" + (extraNotes.length ? " ; " + extraNotes.join(" ; ") : "") };
        const okx = { shed: xc.shed.ok, dest: xc.dest.ok, mv: xc.mv.ok, tax: xc.tax.ok };
        if (xc.shed.ok === false) TOT.shed_bad++;
        if (xc.dest.ok === false) TOT.dest_bad++;
        if (xc.mv.ok === false) TOT.mv_bad++;
        if (xc.mv.ok === "review") TOT.mv_review++;
        if (xc.tax.ok === false) TOT.tax_bad++;
        const cells = {};
        for (const k of KEYS) {
          const o = okx[k];
          cells[k] = mkCell(vals[k], o === false ? "diff" : o === "review" ? "warn" : "yel", ["urgent", "bl", "note"].includes(k) || o === false || o === "review");
        }
        if (pm_) TOT.pm += pm_; if (gm_) TOT.gm += gm_; if (mm_) TOT.mm += mm_;
        rows.push({ kind: "noenter", bl, cells, bad: true });
        issueRows++;
        crit.push(`${bl}: ไม่มีเอกสาร ENTER`);
        extraNotes.forEach(x => crit.push(`${bl}: ${x}`));
        continue;
      }

      // CONSIGNEE
      const ne = normCo(e.cons), mc = normCo(m.cons), mn = normCo(m.notify);
      const hitC = !!(ne && mc && comatch(ne, mc)), hitN = !!(ne && mn && comatch(ne, mn));
      const consOk = hitC || hitN;
      mk(["cneeE", "cneeM"], consOk);
      if (!consOk) notes.push(`CONSIGNEE ไม่ตรง (ENTER: ${e.cons} / MANIFEST C:${m.cons} N:${m.notify})`);
      else if (hitN && !hitC) { cs.cneeE = "warn"; notes.push(`CONSIGNEE(MANIFEST) ช่อง C: = '${m.cons}' ไม่ครบ — ชื่อเต็มตรงกับ Notify`); }

      // CONTAINER
      const ecm = (e.cont || "").match(CONT_RE), eC = ecm ? ecm[0] : "";
      const contOk = !!(eC && m.cont.includes(eC));
      mk(["cntE", "cntM"], contOk);
      if (!contOk) notes.push(`CONTAINER ไม่ตรง (ENTER: ${eC || "-"} / MANIFEST: ${m.cont.join(";") || "-"})`);

      // STATUS
      const seRaw = e.status_raw || "", smRaw = m.status || "";
      const se = canonStatus(seRaw), sm = canonStatus(smRaw);
      const stOk = se && sm ? se === sm : true;
      mk(["stE", "stM"], stOk);
      const dispE = seRaw ? seRaw + (se && seRaw.trim() !== se ? `  (= ${se})` : "") : "(ไม่ระบุใน ENTER)";
      const dispM = smRaw ? smRaw + (sm && smRaw !== sm ? `  (= ${sm})` : "") : "(ว่าง)";
      if (!stOk) { TOT.st_bad++; notes.push(`STATUS ไม่ตรง (ENTER ${seRaw} = ${se} / MANIFEST ${smRaw} = ${sm})`); }

      // PACKAGE
      const pe = e.pkgs, pm = pkgnum(m.pkg_hdr);
      const pkOk = pe !== null && pm !== null && pe === pm;
      mk(["tpE", "tpM"], pkOk);
      if (!pkOk) notes.push(`TOTAL PACKAGE ไม่ตรง (ENTER ${pe !== null ? pe : "-"} / MANIFEST ${pm !== null ? pm : "(ไม่มี)"})`);
      const ke = canonPkg(e.pkgtype), km = canonPkg(m.pkg_hdr);
      const pgOk = !!(ke && km && (ke === km || km.includes(ke) || ke.includes(km)));
      mk(["pkE", "pkM"], pgOk);
      if (!pgOk) notes.push(`PACKAGING ไม่ตรง (ENTER ${ke || e.pkgtype || "-"} / MANIFEST ${km || m.pkg_hdr || "(ไม่มี)"})`);

      // GW / MEAS
      const ge = e.gw, gm = num(m.gw);
      let gwOk = ge !== null && gm !== null && r2(ge) === r2(gm);
      if ((gm === 0 || gm === null) && ge) gwOk = false;
      mk(["gwE", "gwM"], gwOk);
      if (!gwOk) notes.push(`GROSS WEIGHT ไม่ตรง (ENTER ${ge} / MANIFEST ${gm})` + ((gm === 0 || gm === null) && ge ? " — MANIFEST = 0.000 (ค่าหาย) แต่ ENTER มีค่าจริง" : ""));
      const me = e.meas, mm = num(m.meas);
      let msOk = me !== null && mm !== null && r3(me) === r3(mm);
      if ((mm === 0 || mm === null) && me) msOk = false;
      mk(["meE", "meM"], msOk);
      if (!msOk) notes.push(`MEASUREMENT ไม่ตรง (ENTER ${me} / MANIFEST ${mm})` + ((mm === 0 || mm === null) && me ? " — MANIFEST = 0.000 (ค่าหาย) แต่ ENTER มีค่าจริง" : ""));

      // MARKS / DESCRIPTION — เทียบทีละคู่ (MARKS กับ MARKS, DESC กับ DESC) ตัดช่องว่าง/วรรคตอนทิ้ง ไม่จับผิดช่องว่าง แต่คำหาย/เกินต้องเตือน
      // ฟอร์ม Able (layout B) ตั้ง combined_text → ไม่ตรง = ส้ม ทั้งสองฝั่ง (เหมือน build_edi.py)
      const mdLvl = e.combined_text ? "warn" : "diff";
      const mdTxt = (lbl, a, b) => e.combined_text ? `${lbl} ไม่ตรงกัน (ไม่นับช่องว่าง) — (ENTER: ${a || "(ว่าง)"} / MANIFEST: ${b || "(ว่าง)"})`
        : `${lbl} ไม่เหมือนกัน (ENTER: ${a || "(ว่าง)"} / MANIFEST: ${b || "(ว่าง)"})`;
      if (sameDetail(e.marks, m.marks) === false) { cs.mkE = cs.mkM = mdLvl; notes.push(mdTxt("MARKS", e.marks, m.marks)); }
      if (sameDetail(e.desc, m.desc) === false) { cs.dsE = cs.dsM = mdLvl; notes.push(mdTxt("DESCRIPTION", e.desc, m.desc)); }

      // REEFER / DG
      const reE = e.reefer || "", reM = m.reefer || "", dgE = e.dg || "", dgM = m.dg || "";
      const both = (a, b) => [a ? `ENTER: ${a}` : "", b ? `MANIFEST: ${b}` : ""].filter(Boolean).join(" | ") || "-";
      const reefer = both(reE, reM), dg = both(dgE, dgM);
      if (reE || reM) {
        mk(["reefer"], false); TOT.reefer++;
        notes.push(`REEFER TEMP ${reefer} — ${!!reE !== !!reM ? "อีกฝั่งไม่ระบุ ; " : ""}ต้องสำแดงตู้เย็น/ยืนยันอุณหภูมิในใบขน`);
        urgent.push(`${WARN} REEFER ${reE || reM}`);
      }
      if (dgE || dgM) {
        mk(["dg"], false); TOT.dg++;
        notes.push(`DG ${dg} — ${!!dgE !== !!dgM ? "อีกฝั่งไม่ระบุ ; " : ""}ต้องสำแดงวัตถุอันตราย (CLASS/UN) + แนบ MSDS/DGD`);
        urgent.push(`${WARN} DG ${dgE || dgM}`);
      }

      // TRANSIT
      const parts = [extractTransitPhrase(e.transit_raw || ""), extractTransitPhrase(m.transit || "")].filter(Boolean);
      const transitDisp = parts.length ? [...new Set(parts)].join(" / ") : "-";
      if (parts.length) {
        mk(["transit"], false); TOT.transit++;
        notes.push(`พบคำ transit/transhipment — ${transitDisp} (ไม่ใช่ LOCAL, ตรวจประเภทใบขน: ถ่ายลำ/ผ่านแดน)`);
        urgent.push(`${WARN} ${transitDisp}`);
      }

      // 4 กฎเพิ่มเติม
      const xc = computeExtraChecks(m, e, exShed);
      for (const [key, tag] of [["shed", "SHED"], ["dest", "ปลายทาง"], ["mv", "MOVEMENT"], ["tax", "TAX ID"]]) {
        const o = xc[key].ok;
        if (o === null) continue;
        if (o === "review") {
          cs[key] = "warn"; TOT.mv_review++;
          notes.push(xc[key].note); urgent.push(`${WARN} ${tag} ต้องตรวจสอบด้วยคน`);
          crit.push(`${bl}: ${tag} ต้องตรวจสอบด้วยคน — ${xc[key].note}`);
        } else {
          cs[key] = o ? "ok" : "diff";
          if (!o) {
            if (key === "shed") TOT.shed_bad++; else if (key === "dest") TOT.dest_bad++; else if (key === "mv") TOT.mv_bad++; else TOT.tax_bad++;
            notes.push(xc[key].note); urgent.push(`${WARN} ${tag} ${xc[key].text}`);
            crit.push(`${bl}: ${xc[key].note}`);
          }
        }
      }

      if (e.low_confidence) {
        for (const k of Object.keys(cs)) if (cs[k] === "diff" && !["reefer", "dg", "transit", "shed", "dest", "tax"].includes(k)) cs[k] = "warn";
        notes.push("ENTER เอกสารนี้ดึงข้อมูลด้วยความมั่นใจต่ำ (layout ไม่ชัดเจน/พบ anchor น้อยเกินไป) — ตรวจสอบทุกช่องที่ไฮไลต์ส้มด้วยคน");
      }

      const diff = Object.keys(cs).filter(k => cs[k] === "diff"), warnCols = Object.keys(cs).filter(k => cs[k] === "warn");
      if (diff.length || warnCols.length) issueRows++;
      if (diff.length) {
        const tags = [...new Set(diff.filter(k => !OWN_URGENT.has(k)).map(k => LBL[k]))].sort();
        if (tags.length) { urgent.unshift(`${WARN} ไม่ตรง: ` + tags.join(", ")); crit.push(`${bl}: ไม่ตรง ` + tags.join(", ")); }
      } else if (warnCols.length) {
        const wt = [...new Set(warnCols.filter(k => LBL[k] && !OWN_URGENT.has(k)).map(k => LBL[k]))].sort();
        urgent.unshift(wt.length ? `${WARN} ต้องตรวจสอบ: ` + wt.join(", ") : `${WARN} ต้องตรวจสอบด้วยคน`);
      }
      if (urgent.length && !urgent[0].startsWith(WARN)) urgent[0] = `${WARN} ${urgent[0]}`;
      const urgentTxt = urgent.length ? [...new Set(urgent)].join(" ; ") : `${OK} ผ่าน`;

      if (pe !== null) TOT.pe += pe; if (pm !== null) TOT.pm += pm;
      if (ge !== null) TOT.ge += ge; if (gm !== null) TOT.gm += gm;
      if (me !== null) TOT.me += me; if (mm !== null) TOT.mm += mm;
      if (ke) TOT.pke[ke] = (TOT.pke[ke] || 0) + (pe || 0);
      if (km) TOT.pkm[km] = (TOT.pkm[km] || 0) + (pm || 0);
      if (eC) TOT.cte[eC] = (TOT.cte[eC] || 0) + 1;
      m.cont.forEach(x => { TOT.ctm[x] = (TOT.ctm[x] || 0) + 1; });

      const vals = { urgent: urgentTxt, bl, cneeE: e.cons, cneeM: m.cons || m.notify, tax: xc.tax.text, cntE: eC, cntM: m.cont.join(";"),
        stE: dispE, stM: dispM, tpE: pe !== null ? pe : "-", pkE: ke || e.pkgtype || "-", tpM: pm !== null ? pm : "(ไม่มี)", pkM: km || m.pkg_hdr || "(ไม่มี)",
        gwE: ge !== null ? fmt3(ge) : "-", gwM: gm !== null ? fmt3(gm) : "-", meE: me !== null ? fmt3(me) : "-", meM: mm !== null ? fmt3(mm) : "-",
        mkE: e.marks, mkM: m.marks, dsE: e.desc, dsM: m.desc, reefer, dg, transit: transitDisp,
        shed: xc.shed.text, dest: xc.dest.text, mv: xc.mv.text, note: notes.length ? notes.join(" ; ") : "-" };
      const hasDiff = diff.length > 0, hasWarn = warnCols.length > 0;
      const cells = {};
      for (const k of KEYS) {
        const s = cs[k] || "ok";
        let v = vals[k];
        cells[k] = mkCell(v, s, s === "diff" || s === "warn");
        if ((s === "diff" || s === "warn") && v !== "" && v != null && !String(v).startsWith(WARN)) cells[k].v = `${WARN} ${v}`;
      }
      cells.urgent = hasDiff ? mkCell(urgentTxt, "diff", true) : hasWarn ? mkCell(urgentTxt, "warn", true) : mkCell(urgentTxt, "grn", true);
      const nv = vals.note;
      cells.note = (nv === "-" || nv === "") ? mkCell(nv, "ok", false)
        : mkCell((hasDiff || hasWarn) && !String(nv).startsWith(WARN) ? `${WARN} ${nv}` : nv, hasDiff ? "diff" : hasWarn ? "warn" : "ok", hasDiff || hasWarn);
      rows.push({ kind: "normal", bl, cells, bad: hasDiff || hasWarn });
    }

    // B/L ที่อยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย
    for (const bl of Object.keys(ENT).filter(b => !(b in MAN)).sort(blSort)) {
      const e = ENT[bl];
      const generic = !!e.low_confidence;
      const head = generic ? `${WARN} พบข้อความคล้ายเลข B/L นี้ใน ENTER (ความมั่นใจต่ำ) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคนว่าเป็น B/L จริงหรือขยะจากการอ่านข้อความ`
        : `${WARN} มีใน ENTER แต่ไม่มีใน MANIFEST — เสี่ยงตกหล่นจากใบขน`;
      const note = generic ? `${WARN} B/L นี้ไม่พบใน MANIFEST — ดึงข้อมูลมาด้วยความมั่นใจต่ำ ตรวจสอบด้วยคนก่อนแจ้งสายเรือ/ลูกค้า`
        : `${WARN} B/L นี้ไม่มีอยู่ใน MANIFEST เลย — ของอาจตกหล่นจากใบขนทั้งรายการ ต้องแจ้งสายเรือ/ลูกค้าด่วน`;
      const vals = { urgent: head, bl: `${WARN} ${bl}`, cneeE: e.cons, cneeM: "—", tax: "—", cntE: e.cont || "", cntM: "—",
        stE: e.status_raw || "(ไม่ระบุใน ENTER)", stM: "—", tpE: e.pkgs !== null ? e.pkgs : "-", pkE: canonPkg(e.pkgtype) || e.pkgtype || "-", tpM: "—", pkM: "—",
        gwE: e.gw !== null ? fmt3(e.gw) : "-", gwM: "—", meE: e.meas !== null ? fmt3(e.meas) : "-", meM: "—",
        mkE: e.marks, mkM: "—", dsE: e.desc, dsM: "—", reefer: e.reefer || "-", dg: e.dg || "-",
        transit: extractTransitPhrase(e.transit_raw) || "-", shed: "—", dest: "—", mv: "—", note };
      const cells = {};
      for (const k of KEYS) cells[k] = mkCell(vals[k], generic ? "warn" : "diff", true);
      rows.push({ kind: "extra", bl, cells, bad: true });
      if (e.pkgs) TOT.pe += e.pkgs; if (e.gw) TOT.ge += e.gw; if (e.meas) TOT.me += e.meas;
      issueRows++;
      crit.unshift(`${bl}: ` + (generic ? "⚠ อยู่ใน ENTER (layout ไม่รู้จัก) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคน" : "⚠⚠ อยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย (วิกฤต)"));
    }

    // แถวยอดรวม
    const peOk = TOT.pe === TOT.pm, geOk = r2(TOT.ge) === r2(TOT.gm), meOk = r3(TOT.me) === r3(TOT.mm);
    const pk = d => Object.keys(d).sort().map(k => `${k} ${d[k]}`).join(" / ") || "-";
    const ct = d => Object.keys(d).sort().join(", ") || "-";
    const na = sub.length;
    const allOk = peOk && geOk && meOk && !TOT.st_bad && !TOT.transit && !TOT.dg && !TOT.reefer && !TOT.shed_bad && !TOT.dest_bad && !TOT.mv_bad && !TOT.mv_review && !TOT.tax_bad;
    const dbits = [];
    const dtxt = d => ["pkg", "gw", "meas"].filter(k => k in d).map(k => (k === "pkg" ? String(d[k]) : fmt3(d[k]))).join(" / ");
    if (entDecl && Object.keys(entDecl).length) dbits.push("ENTER: " + dtxt(entDecl));
    if (manDecl && Object.keys(manDecl).length) dbits.push("MANIFEST(GRAND TOTAL): " + dtxt(manDecl));
    const declTxt = dbits.length ? "ยอดที่ระบุในเอกสาร — " + dbits.join("  |  ") + "  ||  " : "";
    const T = {
      urgent: allOk ? `${OK} ยอดรวม ENTER = MANIFEST` : `${WARN} ยอดรวมไม่ตรง — ตรวจด่วน`, bl: "TOTAL / ยอดรวม",
      cneeE: `${na} B/L ย่อย`, cneeM: `${na} B/L ย่อย`, tax: `TAX ID ไม่ตรง ${TOT.tax_bad} B/L`,
      cntE: `${Object.keys(TOT.cte).length} ตู้: ${ct(TOT.cte)}`, cntM: `${Object.keys(TOT.ctm).length} ตู้: ${ct(TOT.ctm)}`,
      stE: `STATUS ไม่ตรง ${TOT.st_bad} B/L`, stM: `STATUS ไม่ตรง ${TOT.st_bad} B/L`,
      tpE: TOT.pe, pkE: pk(TOT.pke), tpM: TOT.pm, pkM: pk(TOT.pkm),
      gwE: fmt3(TOT.ge), gwM: fmt3(TOT.gm), meE: fmt3(TOT.me), meM: fmt3(TOT.mm),
      mkE: "", mkM: "", dsE: "", dsM: "", reefer: `${TOT.reefer} B/L`, dg: `${TOT.dg} B/L`, transit: `${TOT.transit} B/L`,
      shed: `SHED ไม่ตรง ${TOT.shed_bad} B/L`, dest: `ปลายทางไม่ตรง ${TOT.dest_bad} B/L`, mv: `MOVEMENT ไม่ตรง ${TOT.mv_bad} / ต้องตรวจ ${TOT.mv_review} B/L`,
      note: declTxt + `ตรวจยอดรวมจากรายการข้างบน: PACKAGE ${peOk ? "ตรง" : "ไม่ตรง"}, GROSS WEIGHT ${geOk ? "ตรง" : "ไม่ตรง"}, MEASUREMENT ${meOk ? "ตรง" : "ไม่ตรง"}` };
    const bad = {};
    if (!peOk) bad.tpE = bad.tpM = 1; if (!geOk) bad.gwE = bad.gwM = 1; if (!meOk) bad.meE = bad.meM = 1;
    if (TOT.st_bad) bad.stE = bad.stM = 1;
    if (TOT.reefer) bad.reefer = 1; if (TOT.dg) bad.dg = 1; if (TOT.transit) bad.transit = 1;
    if (TOT.shed_bad) bad.shed = 1; if (TOT.dest_bad) bad.dest = 1; if (TOT.mv_bad + TOT.mv_review) bad.mv = 1; if (TOT.tax_bad) bad.tax = 1;
    if (!(peOk && geOk && meOk)) bad.note = 1;
    const tcells = {};
    for (const k of KEYS) {
      const isBad = !!bad[k];
      tcells[k] = { v: isBad ? `${WARN} ${T[k]}` : T[k], s: isBad ? "diff" : "ok", bad: isBad, total: true };
    }
    tcells.urgent = { v: T.urgent, s: allOk ? "grn" : "diff", bad: !allOk, total: true };
    tcells.bl = { v: T.bl, s: "navy", total: true };

    const masterBl = sub.length ? sub.slice().sort((a, b) => a.length - b.length)[0] : "";
    return { rows, total: { cells: tcells }, crit, issueRows, count: na, vessel, masterBl, TOT,
      stats: { bl: na, issue_rows: issueRows, allOk }, expectedShed: exShed };
  }

  // ───────────── อ่าน PDF ด้วย pdf.js → pages [{items:[{x,y,t}], annots:[{y,text}]}] ─────────────
  // รวม text item ที่อยู่แถวเดียวกันและติดกันเป็นบรรทัดเดียว; หน้าที่ข้อความถูกวาดหมุน 90° (เนื้อหาแนวนอนบนหน้าแนวตั้ง
  // ทั้งที่ /Rotate = 0 — พบใน ENTER ของ AGN บางไฟล์) จะหมุนพิกัดกลับเป็นแนวอ่านปกติก่อน
  async function pdfPages(pdfjsLib, buf) {
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const pages = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 1 }), W = vp.width, H = vp.height;
      const src = (await page.getTextContent()).items.filter(i => i.str && i.str.trim() !== "");
      let up = 0, down = 0, flat = 0;
      for (const i of src) {
        const [a, b] = i.transform;
        if (Math.abs(b) > Math.abs(a)) (b > 0 ? up++ : down++); else flat++;
      }
      const rot = up > flat && up >= down ? "up" : down > flat ? "down" : "";
      const raw = src.map(i => {
        const px = i.transform[4], py = i.transform[5];
        const pos = rot === "up" ? { x: py, y: px } : rot === "down" ? { x: H - py, y: W - px } : { x: px, y: H - py };
        return { x: pos.x, y: pos.y, w: i.width || 0, t: i.str };
      });
      raw.sort((a, b) => a.y - b.y || a.x - b.x);
      const items = [];
      let cur = null;
      for (const it of raw) {
        // ต่อเฉพาะชิ้นที่ชิดกันจริง (ช่องว่าง ≤ 1.5) — เลข B/L ยาวบางใบชิดคอลัมน์ MARKS ถัดไปไม่ถึง 6 pt ห้ามรวมข้ามคอลัมน์
        if (cur && Math.abs(it.y - cur.y) <= 1.5 && it.x - (cur.x + cur.w) <= 1.5 && it.x >= cur.x) {
          const gap = it.x - (cur.x + cur.w);
          cur.t += (gap > 1.5 && !cur.t.endsWith(" ") ? " " : "") + it.t;
          cur.w = it.x + it.w - cur.x;
        } else { cur = { x: it.x, y: it.y, w: it.w, t: it.t }; items.push(cur); }
      }
      items.forEach(i => { i.t = i.t.replace(/\s+/g, " ").trim(); });
      const annots = [];
      try {
        for (const a of await page.getAnnotations()) {
          const text = a.contents || (a.contentsObj && a.contentsObj.str);
          if (a.subtype !== "FreeText" || !text) continue;
          const [x1, y1, x2, y2] = a.rect;
          const top = rot === "up" ? { x: y1, y: x1 } : rot === "down" ? { x: H - y2, y: W - x2 } : { x: x1, y: H - y2 };
          annots.push({ y: top.y, text });
        }
      } catch (e) { /* ไม่มี annotation */ }
      pages.push({ items: items.filter(i => i.t), annots });
    }
    return pages;
  }

  root.CheckAble = { pdfPages, HEADERS, KEYS, NUMKEYS, WIDEKEYS, parseManifest, parseEnter, isAbleForm, isManifestPdf, manifestRowsFromPdf, buildReport, WARN, OK,
    _t: { canonPkg, canonStatus, normCo, comatch, sameDetail, extractTransitPhrase, findDestCountry } };
})(typeof window !== "undefined" ? window : globalThis);
