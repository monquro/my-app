/* 데이터에 묻는 창 — 화면.
 *
 * DuckDB-WASM 이 브라우저 안에서 parquet 을 읽는다. 표는 SQL에 나오는 것만 그때그때 받는다.
 * AI(api/sql)에는 질문 문장과 표 설명만 보낸다. 데이터는 보내지 않는다.
 * 실행하는 SQL은 전부 «읽기 전용» 검사를 거친다 — AI가 쓴 것이든 사람이 쓴 것이든.
 */
import * as duckdb from "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.32.0/+esm";

const $ = s => document.querySelector(s);
const esc = v => String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const S = { db: null, conn: null, loaded: new Set(), schema: null, questions: [], files: [], mine: {} };

// ── 엔진 ─────────────────────────────────────────────
async function startEngine() {
  const bundle = await duckdb.selectBundle(duckdb.getJsDelivrBundles());
  const workerUrl = URL.createObjectURL(new Blob([`importScripts("${bundle.mainWorker}");`], { type: "text/javascript" }));
  const db = new duckdb.AsyncDuckDB(new duckdb.VoidLogger(), new Worker(workerUrl));
  await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
  S.db = db;
  S.conn = await db.connect();
}

// 실행 버튼을 연달아 눌러도 표를 두 번 만들지 않게, 표 준비는 한 번에 하나씩 한다
let tableLock = Promise.resolve();
function ensureTables(sql) {
  const p = tableLock.then(() => prepareTables(sql));
  tableLock = p.catch(() => {});
  return p;
}

async function prepareTables(sql) {
  const names = Object.keys(S.schema.tables).filter(t => new RegExp(`\\b${t}\\b`, "i").test(sql));
  // 내 파일을 하나라도 넣었으면, 파일로 바꿀 수 있는 표는 모두 내 파일이어야 한다.
  // 빠진 표를 가상 데이터로 채우면 내 숫자와 지어낸 숫자가 소리 없이 섞인다 — 그래서 멈춘다.
  if (Object.keys(S.mine).length) {
    const missing = S.files.filter(f => names.includes(f.table) && !S.mine[f.table]);
    if (missing.length) {
      throw new Error(`${missing.map(f => `«${f.match}»`).join(" · ")} 파일을 아직 넣지 않았습니다. `
        + `내 파일과 가상 데이터가 섞이지 않게 멈췄습니다 — «내 파일로»에 마저 넣거나, «원래 데이터로 돌아가기»를 누르십시오.`);
    }
  }
  for (const t of names) {
    if (S.loaded.has(t)) continue;
    $("#engine").textContent = `표 «${t}» 받는 중…`;
    const buf = new Uint8Array(await (await fetch(`data/${t}.parquet`)).arrayBuffer());
    await S.db.registerFileBuffer(`${t}.parquet`, buf);
    await S.conn.query(`CREATE VIEW ${t} AS SELECT * FROM '${t}.parquet'`);
    S.loaded.add(t);
  }
  engineLine();
}

function engineLine() {
  const mine = Object.keys(S.mine);
  // 파일이 많으면 이름을 다 늘어놓지 않고 개수만 — 이름은 마우스를 올리면 보인다
  const full = `내 파일로 묻는 중 · ${mine.map(t => `${t} ← ${S.mine[t].file}`).join(" · ")}`;
  const nFiles = mine.reduce((s, t) => s + (S.mine[t].count || 1), 0);
  $("#engine").textContent = !mine.length
    ? `준비됨 · 불러온 표 ${S.loaded.size}개 / ${Object.keys(S.schema.tables).length}개`
    : nFiles > 3 ? `내 파일로 묻는 중 · 파일 ${nFiles}개 · 표 ${mine.length}개` : full;
  $("#engine").title = mine.length ? full : "";
  $("#engine").classList.add("ready");
}

// ── 내 파일로 — 엑셀 · CSV 를 이 브라우저 안에서 표로 올린다 ─────────────
// 파일 이름에 questions.json 의 files[].match 가 들어 있으면 그 표를 대신한다.
// 맥은 파일 이름의 한글을 자모로 쪼개(NFD) 넘긴다 — 붙인 꼴(NFC)로 맞춘 뒤 비교한다
function tableFor(name) {
  const n = name.normalize("NFC");
  const hit = S.files.find(f => n.includes(f.match.normalize("NFC")));
  return hit ? hit.table : null;
}

// 파일 하나를 CSV 글자로 바꾼다. 셀은 sheet_to_csv 의 글자 그대로 읽는다 —
// 날짜 셀이 46265 같은 일련번호로 바뀌지 않게 (줄 수를 셀 때 sheet_to_json 을 써도 raw:false)
async function readMineFile(f) {
  const buf = new Uint8Array(await f.arrayBuffer());
  const isCsv = /\.csv$/i.test(f.name);
  const wb = isCsv
    ? XLSX.read(new TextDecoder("utf-8").decode(buf).replace(/^﻿/, ""), { type: "string", raw: true })
    : XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return {
    read: XLSX.utils.sheet_to_json(ws, { defval: "", raw: false }).length,
    csv: XLSX.utils.sheet_to_csv(ws),
    sheets: wb.SheetNames.length,
  };
}

// 이 컴퓨터의 날짜(로컬) — toISOString 은 UTC 라서 하루 어긋날 수 있다
function localDate(ms) {
  const d = new Date(ms);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// 파일 이름에서 날짜(YYYY-MM-DD 또는 YYYYMMDD)를 찾는다. 없으면 null.
// 숫자만 찾으므로 계정 모양(A_IG)을 찾는 정규식과 겹치지 않는다. 8자리보다 긴 숫자 덩어리는 날짜로 보지 않는다.
function dateInName(name) {
  const n = name.normalize("NFC");
  const m = n.match(/(?:^|\D)(\d{4})-(\d{2})-(\d{2})(?!\d)/) || n.match(/(?:^|\D)(\d{4})(\d{2})(\d{2})(?!\d)/);
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3], t = new Date(y, mo - 1, d);
  if (t.getFullYear() !== y || t.getMonth() !== mo - 1 || t.getDate() !== d)
    throw new Error(`파일 이름의 날짜 «${m[0].replace(/^\D/, "")}» 는 달력에 없는 날입니다`);
  return `${m[1]}-${m[2]}-${m[3]}`;
}

function mineItem(html) {
  const el = document.createElement("div");
  el.className = "mine";
  el.innerHTML = html;
  $("#mine-list").appendChild(el);
  return el;
}

// 같은 표로 가는 파일들을 하나로 합쳐 그 표를 통째로 대신한다. 파일이 하나면 예전과 같다.
//  - 열 이름 기준으로 맞춘다(UNION ALL BY NAME) — 어느 파일에 없는 열은 빈칸
//  - 표의 열 목록(schema.json)에 account 가 있는데 파일에 없으면 파일 이름에서 A_IG 꼴을 찾아 넣는다
//    (\b 는 쓰지 않는다 — _ 가 단어 글자라서 «게시물_A_IG» 에서 못 찾는다)
//  - 표의 열 목록에 recorded_on 이 있는데 파일에 없으면, 파일 이름에 날짜가 있으면 그것을, 없을 때만 파일 수정 날짜(로컬)를 DATE 로 넣는다
async function mergeIntoTable(table, files) {
  const cols = (S.schema.tables[table]?.columns || []).map(c => c.name.toLowerCase());
  const parts = [], used = [];
  let sumRead = 0;
  for (const [i, f] of files.entries()) {
    const item = mineItem(`${esc(f.name)} 읽는 중…`);
    try {
      const r = await readMineFile(f);
      const tmp = `mine_${table}_${i}.csv`;
      await S.db.registerFileText(tmp, r.csv);
      const src = `read_csv('${tmp}', header = true, auto_detect = true)`;
      const have = (await S.conn.query(`DESCRIBE SELECT * FROM ${src}`)).toArray()
        .map(x => String(x.column_name).toLowerCase());
      const extra = [], notes = [];
      if (cols.includes("account") && !have.includes("account")) {
        const m = f.name.normalize("NFC").match(/[A-Z]+_(?:IG|TT|YT)/);
        if (!m) throw new Error("account 열이 없고, 파일 이름에서 «A_IG» 같은 계정(영문 대문자_IG · TT · YT)을 찾지 못했습니다");
        extra.push(`'${m[0]}' AS account`);
        notes.push(`account = ${esc(m[0])} (파일 이름에서)`);
      }
      if (cols.includes("recorded_on") && !have.includes("recorded_on")) {
        // 파일 이름에 날짜가 있으면 그것을, 없을 때만 파일 수정 날짜를 쓴다
        const named = dateInName(f.name);
        const day = named || localDate(f.lastModified);
        extra.push(`CAST('${day}' AS DATE) AS recorded_on`);
        notes.push(`기록일 = ${named ? "파일 이름" : "파일 받은 날"} ${day}`);
      }
      parts.push(`SELECT *${extra.length ? ", " + extra.join(", ") : ""} FROM ${src}`);
      used.push({ f, item, read: r.read, sheets: r.sheets, notes });
      sumRead += r.read;
    } catch (e) {
      // 파일 하나가 실패해도 나머지는 계속 넣는다 — 실패한 파일은 ✕와 이유를 남기고 합치지 않는다
      item.innerHTML = `<span class="no">✕</span> <b>${esc(f.name)}</b> — 합치지 않았습니다: ${esc(e.message)}`;
    }
  }
  if (!parts.length) return;
  try {
    // 새 표를 옆 이름으로 먼저 만들고, 성공하면 옛 표를 지우고 이름을 바꾼다 — 실패해도 옛 표가 남는다
    await S.conn.query(`DROP TABLE IF EXISTS ${table}__new`);
    await S.conn.query(`CREATE TABLE ${table}__new AS ${parts.join(" UNION ALL BY NAME ")}`);
    const kind = (await S.conn.query(
      `SELECT table_type FROM information_schema.tables WHERE table_name = '${table}'`)).toArray()[0]?.table_type;
    if (kind === "VIEW") await S.conn.query(`DROP VIEW ${table}`);
    else if (kind) await S.conn.query(`DROP TABLE ${table}`);
    await S.conn.query(`ALTER TABLE ${table}__new RENAME TO ${table}`);
  } catch (e) {
    used.forEach(u => { u.item.innerHTML = `<span class="no">✕</span> <b>${esc(u.f.name)}</b> — 넣지 못했습니다: ${esc(e.message)}`; });
    return;
  }
  const n = Number((await S.conn.query(`SELECT count(*) AS n FROM ${table}`)).toArray()[0].n);
  S.loaded.add(table);
  const again = S.mine[table] ? ` · 앞서 넣은 «${esc(S.mine[table].file)}»를 대신합니다` : "";
  S.mine[table] = { file: used.map(u => u.f.name).join(" + "), count: used.length, rows: n };
  const ok = n === sumRead;
  const many = used.length > 1;
  used.forEach((u, k) => {
    u.item.innerHTML = `<span class="${ok ? "ok" : "no"}">${ok ? "✓" : "✕"}</span> <b>${esc(u.f.name)}</b> → 표 <code>${table}</code> · `
      + (many
        ? `읽은 줄 ${u.read.toLocaleString("ko-KR")} · 합친 표 ${n.toLocaleString("ko-KR")}줄 (파일 줄 합 ${sumRead.toLocaleString("ko-KR")})`
        : `파일에서 읽은 줄 ${u.read.toLocaleString("ko-KR")} · 표에 올라간 줄 ${n.toLocaleString("ko-KR")}`)
      + (u.notes.length ? ` · ${u.notes.join(" · ")}` : "")
      + (u.sheets > 1 ? ` · ⚠ 시트 ${u.sheets}개 중 첫 시트만 읽었습니다` : "")
      + (k === 0 ? again : "");
  });
}

async function loadMine(fileList) {
  // 같은 표로 가는 파일은 한 묶음 — 여러 개면 하나로 합친다
  const groups = new Map();
  for (const f of fileList) {
    const table = tableFor(f.name);
    if (!table) {
      mineItem(`<span class="no">✕</span> ${esc(f.name)} — 이 파일에 맞는 표가 없습니다. 파일 이름에 ${S.files.map(x => `«${esc(x.match)}»`).join(" · ")} 가운데 하나가 들어 있어야 합니다`);
      continue;
    }
    if (!groups.has(table)) groups.set(table, []);
    groups.get(table).push(f);
  }
  for (const [table, files] of groups) await mergeIntoTable(table, files);
  $("#btn-mine-reset").hidden = !Object.keys(S.mine).length;
  engineLine();
  document.querySelectorAll(".q .out").forEach(o => { o.innerHTML = ""; });
}

// ── 읽기 전용 검사 (api/sql.py 와 같은 규칙) ───────────────
const FORBIDDEN = /\b(insert|update|delete|drop|create|alter|attach|detach|copy|pragma|install|load|export|import|call|set)\b/i;
function readOnly(sql) {
  const s = String(sql || "").replace(/--[^\n]*/g, "").trim().replace(/;\s*$/, "").trim();
  if (!s || s.includes(";")) return null;
  if (!/^(select|with)\b/i.test(s)) return null;
  if (FORBIDDEN.test(s.replace(/'[^']*'/g, "''"))) return null;
  return s;
}

// ── 실행 ─────────────────────────────────────────────
async function run(sql) {
  const safe = readOnly(sql);
  if (!safe) throw new Error("읽기(SELECT · WITH)만 실행합니다. 한 문장만, 바꾸는 명령 없이.");
  await ensureTables(safe);
  let t = await S.conn.query(safe);
  // 소수 자리가 정해진 숫자(DECIMAL)는 화면에서 읽기 어려운 모양으로 온다 — 보통 숫자로 바꿔 다시 받는다
  const dec = t.schema.fields.filter(f => String(f.type).startsWith("Decimal"));
  if (dec.length) {
    // 열 이름 안의 큰따옴표는 두 번 써야 한다 — sum("views") 같은 이름에서 깨지지 않게
    const q = n => `"${String(n).replace(/"/g, '""')}"`;
    const cols = t.schema.fields.map(f => dec.includes(f) ? `${q(f.name)}::DOUBLE AS ${q(f.name)}` : q(f.name));
    t = await S.conn.query(`SELECT ${cols.join(", ")} FROM (${safe})`);
  }
  const fields = t.schema.fields.map(f => f.name);
  // 날짜 열은 밀리초 숫자로 온다 — 숫자처럼 보이지 않게 YYYY-MM-DD 글자로 바꾼다
  const kind = Object.fromEntries(t.schema.fields.map(f => [f.name, String(f.type)]));
  const day = v => new Date(v).toISOString().slice(0, 10);
  const rows = t.toArray().map(r => {
    const o = {};
    fields.forEach(f => {
      let v = r[f];
      if (typeof v === "bigint") v = Number(v);
      if (v != null && kind[f].startsWith("Date")) v = day(v);
      else if (v != null && kind[f].startsWith("Timestamp")) v = new Date(v).toISOString().slice(0, 19).replace("T", " ");
      o[f] = v;
    });
    return o;
  });
  return { fields, rows };
}

function fmt(v) {
  if (typeof v === "number") return Number.isInteger(v) ? v.toLocaleString("ko-KR") : String(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return esc(v);
}

function render({ fields, rows }) {
  const head = `<tr>${fields.map(f => `<th>${esc(f)}</th>`).join("")}</tr>`;
  const body = rows.slice(0, 50).map(r => `<tr>${fields.map(f =>
    `<td class="${typeof r[f] === "number" ? "n" : ""}">${fmt(r[f])}</td>`).join("")}</tr>`).join("");
  let html = `<div class="res"><table><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
  if (rows.length > 50) html += `<p class="small muted">… ${rows.length}줄 중 50줄만 보여줍니다</p>`;
  // 막대 — 첫 열이 글자이고 마지막 열이 숫자일 때
  const last = fields[fields.length - 1];
  if (rows.length > 1 && rows.length <= 20 && typeof rows[0][fields[0]] === "string" && typeof rows[0][last] === "number") {
    const max = Math.max(...rows.map(r => Math.abs(r[last]))) || 1;
    html += `<div class="bars">` + rows.map(r => `<div class="bar"><span class="l">${esc(r[fields[0]])}</span>`
      + `<span class="t"><i style="width:${(Math.abs(r[last]) / max) * 100}%"></i></span><span class="v">${fmt(r[last])}</span></div>`).join("") + `</div>`;
  }
  return html;
}

function checks(q, res) {
  return `<div class="checks">` + q.check.map(c => {
    const hit = res.rows.find(r => Object.entries(c.where || {}).every(([k, v]) => r[k] === v));
    const got = hit ? hit[c.col] : undefined;
    const ok = typeof got === "number" && Math.abs(got - c.want) < 1e-9;
    const where = Object.values(c.where || {}).join(" ");
    return `<span class="${ok ? "ok" : "no"}">${ok ? "✓" : "✕"} ${esc(where)} ${esc(c.col)} ${fmt(c.want)}</span>`;
  }).join("") + (Object.keys(S.mine).some(t => new RegExp(`\\b${t}\\b`, "i").test(q.sql))
    ? `<span class="muted">← 확인할 값은 가상 데이터의 값입니다. 내 파일에서는 달라도 됩니다</span>`
    : `<span class="muted">← 가상 데이터 기댓값과 대조</span>`) + `</div>`;
}

// ── 준비한 질문 ───────────────────────────────────────
function renderQuestions() {
  $("#qlist").innerHTML = S.questions.map(q => `
    <div class="q" id="q-${q.id}">
      <h4 data-q="${q.id}">${esc(q.q)}</h4>
      <div class="body" hidden>
        <div class="judge">${esc(q.judge)}</div>
        <pre class="sql">${esc(q.sql)}</pre>
        <div class="row"><button class="primary" data-run="${q.id}">실행</button>
          <button class="ghost" data-edit="${q.id}">고쳐서 묻기</button></div>
        <div class="out"></div>
      </div>
    </div>`).join("");
  document.querySelectorAll("[data-q]").forEach(h => h.addEventListener("click", () => {
    const b = h.parentElement.querySelector(".body"); b.hidden = !b.hidden;
  }));
  document.querySelectorAll("[data-run]").forEach(b => b.addEventListener("click", async () => {
    const q = S.questions.find(x => x.id === b.dataset.run);
    const out = b.closest(".q").querySelector(".out");
    b.disabled = true; b.textContent = "실행 중…";
    try { const res = await run(q.sql); out.innerHTML = render(res) + checks(q, res); out.dataset.done = "1"; }
    catch (e) { out.innerHTML = `<p class="err">${esc(e.message)}</p>`; }
    b.disabled = false; b.textContent = "실행";
  }));
  document.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => {
    $("#sql").value = S.questions.find(x => x.id === b.dataset.edit).sql;
    document.querySelector("[data-tab=ask]").click();
  }));
}

// ── 직접 묻기 ─────────────────────────────────────────
$("#btn-run").addEventListener("click", async () => {
  const b = $("#btn-run");
  b.disabled = true; b.textContent = "실행 중…";
  try { $("#ask-out").innerHTML = render(await run($("#sql").value)); }
  catch (e) { $("#ask-out").innerHTML = `<p class="err">${esc(e.message)}</p>`; }
  b.disabled = false; b.textContent = "실행";
});

$("#btn-nl").addEventListener("click", async () => {
  const q = $("#nl").value.trim();
  if (!q) return;
  const b = $("#btn-nl"); b.disabled = true; b.textContent = "AI가 SQL을 쓰는 중…";
  let r;
  try {
    const res = await fetch("api/sql", { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question: q, schema: S.schema }) });
    r = res.ok ? await res.json() : { sql: null, reason: `서버 함수가 없습니다 (${res.status}) — 배포하면 생깁니다` };
  } catch (e) { r = { sql: null, reason: "서버 함수에 닿지 못했습니다" }; }
  $("#nl-msg").hidden = false;
  if (r.sql) {
    $("#sql").value = r.sql;
    $("#nl-msg").textContent = `AI가 쓴 SQL입니다 — ${r.why || ""} · 실행 전에 읽어 보고, 결과 하나는 손으로 검산하십시오.`;
  } else {
    $("#nl-msg").textContent = `SQL을 만들지 못했습니다 — ${r.reason}. 아래 SQL을 직접 고쳐서 실행할 수 있습니다.`;
  }
  b.disabled = false; b.textContent = "SQL 만들기 (AI)";
});

// ── 표 보기 ───────────────────────────────────────────
function renderTables() {
  $("#schema-note").textContent = S.schema.note;
  $("#tables").innerHTML = Object.entries(S.schema.tables).map(([t, i]) => `
    <div class="tbl"><b>${esc(t)}</b> · <span class="g">한 줄 = ${esc(i.grain)}</span> · ${i.rows.toLocaleString("ko-KR")}행<br>
      <span class="muted">${i.columns.map(c => esc(c.name) + (c.values ? ` <small>(${c.values.map(esc).join("·")})</small>` : "")).join(", ")}</span></div>`).join("");
}

// ── 탭 ───────────────────────────────────────────────
document.querySelectorAll(".tabs button").forEach(b => b.addEventListener("click", () => {
  document.querySelectorAll(".tabs button").forEach(x => x.setAttribute("aria-selected", x === b));
  ["preset", "ask", "tables", "mine"].forEach(k => { $("#tab-" + k).hidden = k !== b.dataset.tab; });
}));

// ── 내 파일로 — 넣기 ─────────────────────────────────────
// 고른 뒤 값을 비운다 — 같은 파일을 다시 골라도 반응하게
$("#mine-file").addEventListener("change", e => { const fs = [...e.target.files]; e.target.value = ""; if (fs.length) loadMine(fs); });
const drop = $("#drop");
["dragenter", "dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => { if (e.dataTransfer.files.length) loadMine([...e.dataTransfer.files]); });
$("#btn-mine-reset").addEventListener("click", () => location.reload());

// ── 시작 ─────────────────────────────────────────────
(async () => {
  const [sch, qs] = await Promise.all([fetch("schema.json").then(r => r.json()), fetch("questions.json").then(r => r.json())]);
  S.schema = sch; S.questions = qs.questions; S.files = qs.files || [];
  renderQuestions(); renderTables();
  $("#mine-rule").textContent = S.files.length
    ? "파일 이름으로 표를 고릅니다 — " + S.files.map(f => `«${f.match}»가 든 파일 → ${f.table}`).join(" · ")
    : "questions.json 에 files 가 없어 내 파일을 받을 수 없습니다";
  try {
    await startEngine();
    $("#engine").textContent = "준비됨 · 표는 질문에 나올 때 받습니다";
    $("#engine").classList.add("ready");
    document.body.dataset.engine = "ready";
  } catch (e) {
    $("#engine").textContent = "데이터 엔진을 불러오지 못했습니다 — 인터넷 연결을 확인하십시오";
    document.body.dataset.engine = "fail";
  }
})();
