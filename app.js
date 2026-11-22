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
  // 맨 위 상태 바: 데이터 출처 · 불러온 표 · 엔진
  $("#st-src").textContent = mine.length ? `내 파일 ${nFiles}개 · 표 ${mine.length}개` : "가상 데이터";
  $("#st-src").title = mine.length ? full : "";
  $("#st-src").classList.toggle("mine", mine.length > 0);
  $("#src-badge").textContent = mine.length ? "내 파일" : "가상 데이터";
  $("#src-badge").classList.toggle("mine", mine.length > 0);
  $("#st-tables").textContent = `${S.loaded.size} / ${Object.keys(S.schema.tables).length}`;
  $("#engine").textContent = mine.length ? "내 파일로 묻는 중" : "준비됨";
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
  // 새 파일이 들어오면 대시보드를 그 파일의 가장 늦은 주로 다시 그린다
  if (Object.keys(S.mine).length) { D.week = null; renderDash(); }
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
// 결과가 먼저, SQL 은 «SQL 보기»를 눌러야 펼쳐진다. 카드 머리에 주기 · 보고서 칸, 안에 판단 기준과 지표 정의.
const defLine = d => { const k = d.indexOf(" — "); return k < 0 ? esc(d) : `<b>${esc(d.slice(0, k))}</b> — ${esc(d.slice(k + 3))}`; };
function renderQuestions() {
  $("#qlist").innerHTML = S.questions.map((q, i) => `
    <div class="q" id="q-${q.id}">
      <button class="qhead" data-q="${q.id}" aria-expanded="false">
        <span class="qno">Q${i + 1}</span>
        <span class="qmain">${(q.tags || []).length ? `<span class="qtags">${q.tags.map(t => `<span class="qtag">${esc(t)}</span>`).join("")}</span>` : ""}
          <span class="qtitle">${esc(q.q)}</span></span>
        <svg class="chev" viewBox="0 0 16 16" aria-hidden="true"><path d="M6 3.5 10.5 8 6 12.5"/></svg>
      </button>
      <div class="body" hidden>
        ${q.judge ? `<div class="judge"><span class="i" aria-hidden="true">i</span><span>${esc(q.judge)}</span></div>` : ""}
        ${(q.defs || []).length ? `<ul class="defs">${q.defs.map(d => `<li>${defLine(d)}</li>`).join("")}</ul>` : ""}
        <div class="row"><button class="primary" data-run="${q.id}">실행</button>
          <button class="ghost" data-edit="${q.id}">고쳐서 묻기</button></div>
        <div class="out"></div>
        <details class="sqlbox"><summary>SQL 보기</summary><pre class="sql">${esc(q.sql)}</pre></details>
      </div>
    </div>`).join("");
  document.querySelectorAll("[data-q]").forEach(h => h.addEventListener("click", () => {
    const b = h.closest(".q").querySelector(".body"); b.hidden = !b.hidden;
    h.setAttribute("aria-expanded", String(!b.hidden));
  }));
  document.querySelectorAll("[data-run]").forEach(b => b.addEventListener("click", () => runQ(b)));
  document.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => {
    $("#sql").value = S.questions.find(x => x.id === b.dataset.edit).sql;
    document.querySelector("[data-tab=ask]").click();
  }));
}

async function runQ(b) {
  const q = S.questions.find(x => x.id === b.dataset.run);
  const out = b.closest(".q").querySelector(".out");
  b.disabled = true; b.textContent = "실행 중…";
  try { const res = await run(q.sql); out.innerHTML = render(res) + checks(q, res); out.dataset.done = "1"; }
  catch (e) { out.innerHTML = `<p class="err">${esc(e.message)}</p>`; }
  b.disabled = false; b.textContent = "실행";
}

// 모두 실행 — 카드를 펼치고 위에서부터 차례로
$("#btn-run-all").addEventListener("click", async e => {
  const all = e.currentTarget;
  all.disabled = true;
  for (const q of document.querySelectorAll(".q")) {
    q.querySelector(".body").hidden = false;
    q.querySelector("[data-q]").setAttribute("aria-expanded", "true");
    await runQ(q.querySelector("[data-run]"));
  }
  all.disabled = false;
});

// ── 직접 묻기 ─────────────────────────────────────────
$("#btn-run").addEventListener("click", async () => {
  const b = $("#btn-run");
  b.disabled = true; b.textContent = "실행 중…";
  try { $("#ask-out").innerHTML = render(await run($("#sql").value)); }
  catch (e) { $("#ask-out").innerHTML = `<p class="err">${esc(e.message)}</p>`; }
  b.disabled = false; b.textContent = "실행";
});
// SQL 칸에서 Ctrl+Enter(맥은 ⌘+Enter)로 실행
$("#sql").addEventListener("keydown", e => {
  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); if (!$("#btn-run").disabled) $("#btn-run").click(); }
});
if (/Mac|iPhone|iPad/.test(navigator.platform)) $("#run-hint").textContent = "⌘+Enter로 실행";

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

// 표 사이의 연결 — 한 줄이 무엇인지와 잇는 키. 카드 숫자를 게시물에 붙여 더하면 불어난다는 경고까지.
function renderErd() {
  const T = S.schema.tables, file = t => S.files.find(f => f.table === t)?.match || "";
  const box = t => T[t] ? `<div class="ebox"><b>${esc(t)}</b><span>${esc(file(t))} 파일 · 한 줄 = ${esc(T[t].grain)}</span>
    <em>${T[t].rows.toLocaleString("ko-KR")}행 (가상 데이터)</em></div>` : "";
  const link = (key, card) => `<div class="elink" aria-hidden="true"><span>${esc(key)}</span><i>1 : N</i><small>${esc(card)}</small></div>`;
  $("#erd").innerHTML = box("targets") + link("account", "계정마다 0~N주")
    + box("account_weekly") + link("account + 그 주", "그 주에 올린 게시물")
    + box("posts");
  $("#erd").insertAdjacentHTML("afterend", `<p class="cap warnline"><span class="wi" aria-hidden="true">!</span>
    account_weekly 의 카드 숫자(조회 · 인터랙션)를 posts 에 붙인 채로 더하면 게시물 수만큼 불어납니다.
    표마다 계정별로 먼저 집계한 뒤 붙입니다.</p>`);
}

// ── 대시보드 ─────────────────────────────────────────
// 숫자는 «준비한 질문»의 weekly_kpi · top_posts SQL 에서 «가장 늦은 주» 한 줄만 고른 주로 바꿔 돌린다.
// 새 정의를 만들지 않는다. IG 와 TT, 카드 조회와 게시물 조회는 더하지 않는다.
const LATEST = "SELECT MAX(CAST(week_start AS DATE)) AS wk FROM account_weekly";
const D = { week: null, plat: "all", kpi: null, top: null, trend: null, tgt: null, sql: {} };
const nf = v => v == null ? "—" : Number(v).toLocaleString("ko-KR");
const isDay = s => typeof s === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s);
const addDays = (day, n) => { const [y, m, d] = day.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); };
// «n월 n주차» — 그 주 월요일이 속한 달에서 몇 번째 월요일인지
const weekName = day => { const [, m, d] = day.split("-").map(Number); return `${m}월 ${Math.ceil(d / 7)}주차`; };
const atWeek = (sql, wk) => isDay(wk) && sql.includes(LATEST) ? sql.replace(LATEST, `SELECT CAST('${wk}' AS DATE) AS wk`) : null;
const PLAT = { IG: "Instagram", TT: "TikTok", YT: "YouTube" };
const platOf = a => String(a).split("_").pop();

function dashMsg(text, kind) {
  const m = $("#dash-msg");
  m.hidden = !text; m.className = `dash-msg ${kind || ""}`; m.textContent = text || "";
}

async function renderDash() {
  dashMsg("");
  $("#tab-dash").classList.add("loading");
  $("#tab-dash").setAttribute("aria-busy", "true");
  try {
    const weeks = (await run("SELECT DISTINCT CAST(week_start AS DATE) AS w FROM account_weekly ORDER BY w DESC")).rows.map(r => r.w).filter(isDay);
    if (!weeks.length) throw new Error("계정주간 표에 주(week_start)가 없습니다");
    D.weeks = weeks;   // 데이터가 시작하는 주 — 그보다 앞 주는 «0»이 아니라 «데이터 없음»
    if (!weeks.includes(D.week)) D.week = weeks[0];
    $("#dash-week").innerHTML = weeks.map(w => `<option value="${w}"${w === D.week ? " selected" : ""}>`
      + `${w} ~ ${addDays(w, 6).slice(5)} · ${weekName(w)}${w === weeks[0] ? " · 가장 늦은 주" : ""}</option>`).join("");
    await drawWeek();
  } catch (e) {
    dashMsg(e.message, "err");
    $("#dash-cards").innerHTML = ""; $("#dash-top").innerHTML = ""; $("#dash-yt").textContent = "";
    $("#btn-report").disabled = true;
  }
  $("#tab-dash").classList.remove("loading");
  $("#tab-dash").removeAttribute("aria-busy");
}

async function drawWeek() {
  const wk = D.week;
  const kq = S.questions.find(q => q.id === "weekly_kpi"), tq = S.questions.find(q => q.id === "top_posts");
  const kSql = kq && atWeek(kq.sql, wk), tSql = tq && atWeek(tq.sql, wk);
  if (!kSql || !tSql) throw new Error("questions.json 의 weekly_kpi · top_posts SQL 모양이 바뀌어 대시보드가 주를 고르지 못합니다");
  const trendSql = `WITH w AS (SELECT CAST('${wk}' AS DATE) AS wk)
SELECT account, CAST(week_start AS DATE) AS week_start, views
FROM account_weekly, w
WHERE CAST(week_start AS DATE) BETWEEN w.wk - 77 AND w.wk
ORDER BY account, week_start`;
  const recSql = `WITH w AS (SELECT CAST('${wk}' AS DATE) AS wk)
SELECT MIN(CAST(recorded_on AS DATE)) AS lo, MAX(CAST(recorded_on AS DATE)) AS hi
FROM posts, w
WHERE CAST(date_trunc('week', CAST(posted_date AS DATE)) AS DATE) = w.wk`;
  // 차례로 돌린다 — 연결 하나에서 한 번에 하나씩
  const kpi = await run(kSql), top = await run(tSql), trend = await run(trendSql), rec = await run(recSql);
  const tgt = await run("SELECT account, platform, weekly_post_target FROM targets ORDER BY account");
  // 주별 업로드 수 — weekly_kpi 의 «업로드수»와 같은 방법(게시물 파일에서 그 주에 올린 게시물을 센다)
  const upsSql = `WITH w AS (SELECT CAST('${wk}' AS DATE) AS wk)
SELECT p.account, CAST(date_trunc('week', CAST(p.posted_date AS DATE)) AS DATE) AS week_start, COUNT(*) AS n
FROM posts p, w
WHERE CAST(date_trunc('week', CAST(p.posted_date AS DATE)) AS DATE) BETWEEN w.wk - 77 AND w.wk
GROUP BY ALL
ORDER BY 1, 2`;
  const ups = await run(upsSql);
  Object.assign(D, { kpi, top, trend, tgt, ups, sql: { "ⓐ 계정 카드 (weekly_kpi)": kSql, "ⓒ 상위 콘텐츠 (top_posts)": tSql,
    "12주 추이 — 조회 (카드)": trendSql, "12주 추이 — 업로드": upsSql } });

  const end = addDays(wk, 6);
  $("#st-week").textContent = `${wk} ~ ${end.slice(5)}`;
  $("#dash-range").textContent = `${wk} ~ ${end} (월~일) · ${weekName(wk)}`;
  const r0 = rec.rows[0] || {};
  $("#st-rec").textContent = !r0.hi ? "—" : r0.lo === r0.hi ? r0.hi : `${r0.lo} ~ ${r0.hi}`;
  paintDash();
  $("#report-text").value = reportText();
  $("#btn-report").disabled = false;
  $("#dash-sql").innerHTML = Object.entries(D.sql).map(([k, s]) => `<p class="label">${esc(k)}</p><pre class="sql">${esc(s)}</pre>`).join("");
  $("#dash-weekly").innerHTML = weeklyTable();
}

// 계정마다 최근 12주(고른 주까지) — 빠진 주는 null 로 둔다(선을 끊는다)
function seriesOf(account) {
  return Array.from({ length: 12 }, (_, i) => {
    const w = addDays(D.week, (i - 11) * 7);
    const hit = D.trend.rows.find(r => r.account === account && r.week_start === w);
    return { w, v: hit ? hit.views : null };
  });
}

function spark(series, account) {
  const W = 240, H = 46, P = 5, n = series.length, step = (W - 2 * P) / (n - 1);
  const vals = series.map(s => s.v).filter(v => v != null);
  if (!vals.length) return `<p class="spark-empty">최근 12주 카드 기록 없음</p>`;
  const lo = Math.min(...vals), hi = Math.max(...vals), span = hi - lo || 1;
  const x = i => P + i * step, y = v => H - P - ((v - lo) / span) * (H - 2 * P);
  let d = "", pen = false;
  series.forEach((s, i) => { if (s.v == null) { pen = false; return; } d += `${pen ? "L" : "M"}${x(i).toFixed(1)} ${y(s.v).toFixed(1)}`; pen = true; });
  const lone = series.map((s, i) => s.v != null && i < n - 1 && series[i - 1]?.v == null && series[i + 1]?.v == null
    ? `<circle cx="${x(i).toFixed(1)}" cy="${y(s.v).toFixed(1)}" r="2.5" class="dot"/>` : "").join("");
  const cur = series[n - 1];
  const hits = series.map((s, i) => `<rect class="hit" x="${(x(i) - step / 2).toFixed(1)}" y="0" width="${step.toFixed(1)}" height="${H}"
    data-tip="${s.w} 주 · ${s.v == null ? "카드 기록 없음" : "조회(카드) " + nf(s.v)}"></rect>`).join("");
  const gaps = series.filter(s => s.v == null).map(s => s.w.slice(5));
  // 키보드로 와도 같은 내용을 볼 수 있게 — 포커스하면 요약 말풍선, 화면 읽기에는 같은 문장
  const sum = `${account} 최근 12주 조회(카드) · 최저 ${nf(lo)} · 최고 ${nf(hi)} · 이번 주 ${cur.v == null ? "기록 없음" : nf(cur.v)}`
    + (gaps.length ? ` · 기록 없는 주 ${gaps.length}개` : "");
  return `<svg class="spark" viewBox="0 0 ${W} ${H}" role="img" tabindex="0"
      aria-label="${esc(sum)}" data-focus-tip="${esc(sum)}">
      <path d="${d}" class="line"/>${lone}
      ${cur.v != null ? `<circle cx="${x(n - 1).toFixed(1)}" cy="${y(cur.v).toFixed(1)}" r="4" class="cur"/>` : ""}${hits}</svg>
    <p class="spark-cap">최근 12주 조회(카드)${gaps.length ? ` · 기록 없는 주 ${gaps.length}개 (${gaps.join(", ")})` : ""}</p>`;
}

const delta = (v, unit) => v == null
  ? `<span class="delta none">전주 없음</span>`
  : `<span class="delta ${v > 0 ? "up" : v < 0 ? "down" : "flat"}"><span aria-hidden="true">${v > 0 ? "▲" : v < 0 ? "▼" : "■"}</span> ${v > 0 ? "+" : v < 0 ? "−" : ""}${Math.abs(v).toLocaleString("ko-KR")}${unit}</span>`;

function card(r) {
  const a = r.계정, p = platOf(a), noCard = r.조회_카드 == null;
  // 비고(카드 기록 없음 · 전주 기록 없음)는 숫자 자리에 바로 적으므로 깃발로 다시 달지 않는다
  const flags = [];
  if (r.posts_count_차이) flags.push(["warn", `카드의 업로드 수(${nf(r.업로드수 + r.posts_count_차이)})와 게시물 파일(${nf(r.업로드수)}개)이 다릅니다 — Asana 업로드 기록과 대조`]);
  const pct = r.달성_pct;
  // 업로드 대 목표 — 불릿 막대: 막대 = 업로드, 세로선 = 목표. 축은 목표의 1.5배(넘으면 업로드)까지라 넘친 것도 보인다
  const up = r.업로드수 || 0, goal = r.목표;
  const axis = Math.max(goal ? goal * 1.5 : 0, up, 1);
  const bullet = `<div class="bullet" role="img" aria-label="업로드 ${up}개, 목표 ${goal ?? "없음"}개${pct == null ? "" : `, 달성 ${pct}%`}">
      <i style="width:${(up / axis * 100).toFixed(1)}%"></i>${goal ? `<b style="left:${(goal / axis * 100).toFixed(1)}%"></b>` : ""}</div>
    <p class="bullet-cap" aria-hidden="true"><span>막대 = 업로드</span><span>세로선 = 목표 ${nf(goal)}개</span></p>`;
  return `<article class="kcard" data-plat="${esc(p)}">
    <header><b>${esc(a)}</b><span class="pl pl-${esc(p)}">${esc(PLAT[p] || p)}</span></header>
    <div class="kmain">
      <span class="klbl">조회 (카드)</span>
      <span class="kbig">${noCard ? "이번 주 기록 없음" : nf(r.조회_카드)}</span>
      ${noCard ? "" : r.조회_전주대비_pct == null ? `<span class="delta none">전주 기록 없음</span>`
        : `${delta(r.조회_전주대비_pct, "%")}<span class="vs">전주 대비</span>`}
    </div>
    ${spark(seriesOf(a), a)}
    <dl class="kv">
      <div><dt>팔로워</dt><dd>${nf(r.팔로워)} ${r.팔로워_증감 == null ? "" : delta(r.팔로워_증감, "")}</dd></div>
      <div><dt>인터랙션 (카드)</dt><dd>${nf(r.인터랙션_카드)} ${r.인터랙션_전주대비_pct == null ? "" : delta(r.인터랙션_전주대비_pct, "%")}</dd></div>
      <div><dt>업로드</dt><dd><b>${nf(r.업로드수)}</b> / 목표 ${nf(r.목표)} · ${pct == null ? "—" : pct + "%"}</dd></div>
    </dl>
    ${bullet}
    ${flags.length ? `<ul class="flags">${flags.map(([k, t]) => `<li class="${k}"><span class="fi" aria-hidden="true">${k === "warn" ? "!" : "i"}</span>${esc(t)}</li>`).join("")}</ul>` : ""}
  </article>`;
}

function topBox(account) {
  const rows = D.top.rows.filter(r => r.계정 === account);
  const p = platOf(account);
  return `<section class="topbox" data-plat="${esc(p)}">
    <h3><b>${esc(account)}</b><span class="pl pl-${esc(p)}">${esc(PLAT[p] || p)}</span></h3>
    ${rows.length ? `<table class="mini"><thead><tr><th>#</th><th>콘텐츠</th><th class="n">조회</th><th class="n">좋아요</th><th class="n">나이</th></tr></thead><tbody>
      ${rows.map(r => `<tr><td class="rk">${r.순위}</td>
        <td><span class="tt">${esc(r.콘텐츠)}</span><small>${esc(r.분류)} · ${esc(r.카테고리)}</small></td>
        <td class="n">${nf(r.조회)}</td><td class="n">${nf(r.좋아요)}</td>
        <td class="n">${r.나이_일 == null ? "—" : r.나이_일 + "일"}${r.나이_일 != null && r.나이_일 <= 3 ? `<span class="young">덜 자람</span>` : ""}</td></tr>`).join("")}
    </tbody></table>` : `<p class="cap">이번 주에 올린 게시물이 없습니다.</p>`}
  </section>`;
}

function paintDash() {
  const show = r => D.plat === "all" || platOf(r) === D.plat;
  const rows = D.kpi.rows.filter(r => show(r.계정));
  $("#dash-cards").innerHTML = rows.map(card).join("") || `<p class="cap">이 플랫폼의 계정이 없습니다.</p>`;
  $("#dash-top").innerHTML = rows.map(r => topBox(r.계정)).join("");
  const seen = new Set(D.kpi.rows.map(r => r.계정));
  const quiet = D.tgt.rows.filter(t => !seen.has(t.account));
  $("#dash-yt").textContent = quiet.length
    ? `${quiet.map(t => `${t.account}(${PLAT[t.platform] || t.platform}, 주 ${t.weekly_post_target}개 목표)`).join(" · ")}는 목표만 있고 아직 계정주간 기록이 없어 카드에서 뺐습니다.`
    : "";
  drawTrend();
}

// 기준 주까지 12주 — 데이터가 시작하기 전 주는 null
const last12 = () => Array.from({ length: 12 }, (_, i) => addDays(D.week, (i - 11) * 7));
const before = w => D.weeks?.length && w < D.weeks[D.weeks.length - 1];
const upsOf = (a, w) => before(w) ? null : (D.ups.rows.find(r => r.account === a && r.week_start === w)?.n ?? 0);

function weeklyTable() {
  const accts = [...new Set([...D.trend.rows.map(r => r.account), ...D.ups.rows.map(r => r.account)])].sort();
  const weeks = last12().reverse();
  const grid = (title, cell) => `<p class="label" style="margin-top:12px">${title}</p>
    <div class="res"><table><thead><tr><th>주 시작</th>${accts.map(a => `<th class="n">${esc(a)}</th>`).join("")}</tr></thead><tbody>
    ${weeks.map(w => `<tr><td>${w}</td>${accts.map(a => `<td class="n">${cell(a, w)}</td>`).join("")}</tr>`).join("")}
    </tbody></table></div>`;
  return grid("주별 조회 (카드) — 기록 없는 주는 —", (a, w) => { const h = D.trend.rows.find(r => r.account === a && r.week_start === w); return h ? nf(h.views) : "—"; })
    + grid("주별 업로드 (게시물 파일) — 데이터 시작 전은 —", (a, w) => { const v = upsOf(a, w); return v == null ? "—" : nf(v); });
}

// ── 12주 추이 — Chart.js ──────────────────────────────
// 메트리쿨에 없는 기준으로만 그린다: 월~일 주, 우리 업로드 목표, 빠진 주는 선을 끊는다.
// 계정마다 따로 — IG 와 TT 를 한 축에 올리거나 더하지 않는다. 축은 하나씩(두 눈금 차트 없음).
const C = { acct: null, charts: [] };
const INK = "#1A1B23", MUTED = "#6B6F80", GRID = "#E7E9F1", AXIS = "#C9CCD8", ACCENT = "#1D4ED8";
const kcompact = new Intl.NumberFormat("ko-KR", { notation: "compact", maximumFractionDigits: 1 });

function drawTrend() {
  const box = $("#trend");
  C.charts.forEach(c => c.destroy()); C.charts = [];
  if (!window.Chart) {
    box.innerHTML = `<p class="cap">차트 라이브러리(Chart.js)를 불러오지 못했습니다 — 아래 «이 화면의 SQL과 주별 숫자»의 표로 같은 값을 볼 수 있습니다.</p>`;
    return;
  }
  const accts = D.kpi.rows.map(r => r.계정).filter(a => D.plat === "all" || platOf(a) === D.plat);
  if (!accts.includes(C.acct)) C.acct = accts[0] || null;
  $("#trend-acct").innerHTML = accts.map(a => `<button data-acct="${esc(a)}" aria-pressed="${a === C.acct}">${esc(a)}</button>`).join("");
  $("#trend-acct").querySelectorAll("[data-acct]").forEach(b => b.addEventListener("click", () => { C.acct = b.dataset.acct; drawTrend(); }));
  $("#trend-range").textContent = `${last12()[0]} ~ ${addDays(D.week, 6)} · 계정마다 따로`;
  if (!C.acct) return;

  const a = C.acct, weeks = last12(), labels = weeks.map(w => w.slice(5));
  const views = weeks.map(w => D.trend.rows.find(r => r.account === a && r.week_start === w)?.views ?? null);
  const ups = weeks.map(w => upsOf(a, w));
  const goal = D.tgt.rows.find(t => t.account === a)?.weekly_post_target ?? null;
  const motion = matchMedia("(prefers-reduced-motion: reduce)").matches ? false : { duration: 250 };
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.font.size = 12;
  Chart.defaults.color = MUTED;
  const base = {
    responsive: true, maintainAspectRatio: false, animation: motion,
    interaction: { mode: "index", intersect: false },   // 선이 아니라 주(X)를 찾는다
    layout: { padding: { top: 4, right: 6 } },
  };
  const xAxis = { grid: { display: false }, border: { color: AXIS }, ticks: { maxRotation: 0, autoSkipPadding: 8 } };
  const tipTitle = items => `${weeks[items[0].dataIndex]} 주 (월~일)`;

  // ① 주별 조회 (카드) — 선 하나, 이번 주에만 점
  const gaps = weeks.filter((w, i) => views[i] == null && !before(w)).map(w => w.slice(5));
  C.charts.push(new Chart($("#ch-views"), {
    type: "line",
    data: { labels, datasets: [{ label: "조회 (카드)", data: views, borderColor: ACCENT, backgroundColor: ACCENT, borderWidth: 2,
      spanGaps: false, tension: 0, borderJoinStyle: "round", borderCapStyle: "round",
      // 이번 주에 점, 양옆이 비어 홀로 남은 주에도 점(선이 없으면 안 보이므로)
      pointRadius: views.map((v, i) => v == null ? 0 : i === 11 ? 4 : (views[i - 1] == null && views[i + 1] == null ? 3 : 0)),
      pointBorderColor: "#fff", pointBorderWidth: 2,
      pointHoverRadius: 5, pointHitRadius: 14 }] },
    options: { ...base,
      plugins: { legend: { display: false },
        tooltip: { callbacks: { title: tipTitle, label: it => it.raw == null ? " 카드 기록 없음" : ` 조회(카드) ${nf(it.raw)}` } } },
      scales: { x: xAxis, y: { beginAtZero: true, grid: { color: GRID }, border: { display: false },
        ticks: { maxTicksLimit: 6, callback: v => kcompact.format(v) } } } },
  }));
  const vv = views.filter(v => v != null);
  $("#ch-views").setAttribute("aria-label", `${a} 주별 조회(카드) 12주: ${vv.length ? `최저 ${nf(Math.min(...vv))}, 최고 ${nf(Math.max(...vv))}, 이번 주 ${nf(views[11])}` : "기록 없음"}`);
  $("#ch-views-cap").textContent = gaps.length ? `카드 기록이 없는 주는 선을 끊었습니다: ${gaps.join(", ")}` : "12주 모두 카드 기록이 있습니다.";

  // ② 주별 업로드 + 목표 선 — 메트리쿨은 우리 목표를 모른다
  const ds = [{ type: "bar", label: "업로드", data: ups, backgroundColor: ACCENT, hoverBackgroundColor: "#0037B0",
    borderRadius: { topLeft: 4, topRight: 4 }, borderSkipped: "bottom", maxBarThickness: 24, pointStyle: "rect", order: 2 }];
  if (goal) ds.push({ type: "line", label: `목표 (주 ${goal}개)`, data: labels.map(() => goal), borderColor: INK, backgroundColor: INK,
    borderWidth: 2, pointRadius: 0, pointHitRadius: 0, pointStyle: "line", order: 1 });
  C.charts.push(new Chart($("#ch-ups"), {
    type: "bar",
    data: { labels, datasets: ds },
    options: { ...base,
      plugins: { legend: { display: true, position: "top", align: "end",
          labels: { usePointStyle: true, boxWidth: 14, boxHeight: 8, padding: 12, sort: (x, y) => x.datasetIndex - y.datasetIndex } },
        tooltip: { itemSort: (x, y) => x.datasetIndex - y.datasetIndex, callbacks: { title: tipTitle, label: it => it.dataset.type === "line" ? ` 목표 ${nf(it.raw)}개` : ` 업로드 ${it.raw == null ? "데이터 없음" : nf(it.raw) + "개"}` } } },
      scales: { x: xAxis, y: { beginAtZero: true, suggestedMax: goal ? Math.ceil(goal * 1.5) : undefined, grid: { color: GRID }, border: { display: false },
        ticks: { precision: 0, maxTicksLimit: 6 } } } },
  }));
  const known = ups.filter(v => v != null);
  const hit = goal ? known.filter(v => v >= goal).length : null;
  $("#ch-ups").setAttribute("aria-label", `${a} 주별 업로드 12주${goal ? `, 목표 주 ${goal}개, 목표를 채운 주 ${hit}/${known.length}` : ""}`);
  $("#ch-ups-cap").textContent = goal
    ? `목표를 채운 주 ${hit} / ${known.length} · 업로드 0개인 주 ${known.filter(v => v === 0).length}개 · 업로드는 게시물 파일에서 센 수입니다.`
    : "이 계정은 목표가 없습니다.";
}

// Asana «SNS 성과 분석 - n월 n주차»에 그대로 붙일 글. 숫자는 위 카드 · 표와 같은 결과에서만 가져온다.
function reportText() {
  const wk = D.week, sign = v => v == null ? "" : (v > 0 ? "+" : v < 0 ? "−" : "±") + Math.abs(v).toLocaleString("ko-KR");
  const src = Object.keys(S.mine).length ? "내 파일" : "가상 데이터(연습용)";
  const rec = $("#st-rec").textContent;
  const L = [];
  L.push(`SNS 성과 분석 - ${weekName(wk)}`);
  L.push(`보고 기간: ${wk} ~ ${addDays(wk, 6)} (월~일)`);
  L.push(`원본: Metricool 계정 화면 · 게시물 목록${rec !== "—" ? ` (게시물 기록일 ${rec})` : ""} · 데이터: ${src}`);
  L.push(`※ IG와 TT 조회는 세는 방식이 달라 더하지 않았습니다. 카드 조회와 게시물 조회도 더하지 않았습니다.`);
  L.push("", "ⓐ 전체 성과");
  for (const r of D.kpi.rows) {
    const p = PLAT[platOf(r.계정)] || platOf(r.계정);
    const parts = r.조회_카드 == null ? ["이번 주 카드 기록 없음"] : [
      `팔로워 ${nf(r.팔로워)}${r.팔로워_증감 == null ? "" : ` (${sign(r.팔로워_증감)})`}`,
      `조회 ${nf(r.조회_카드)}${r.조회_전주대비_pct == null ? "" : ` (전주 대비 ${sign(r.조회_전주대비_pct)}%)`}`,
      `인터랙션 ${nf(r.인터랙션_카드)}${r.인터랙션_전주대비_pct == null ? "" : ` (${sign(r.인터랙션_전주대비_pct)}%)`}`];
    parts.push(`업로드 ${nf(r.업로드수)} / 목표 ${nf(r.목표)} (${r.달성_pct == null ? "—" : r.달성_pct + "%"})`);
    L.push(`- ${r.계정} (${p}): ${parts.join(" · ")}`);
    const notes = [];
    if (r.비고) notes.push(r.비고);
    if (r.posts_count_차이) notes.push(`카드 업로드 수와 게시물 파일이 ${Math.abs(r.posts_count_차이)}개 다름 — Asana 업로드 기록과 대조 필요`);
    if (notes.length) L.push(`  확인: ${notes.join(" · ")}`);
  }
  L.push("", "ⓑ 분석: (직접 작성)", "", "ⓒ 상위 콘텐츠 (이번 주에 올린 게시물 · 조회 순)");
  for (const a of D.kpi.rows.map(r => r.계정)) {
    const rows = D.top.rows.filter(r => r.계정 === a);
    L.push(`${a}${rows.length ? "" : " — 이번 주에 올린 게시물 없음"}`);
    rows.forEach(r => L.push(`  ${r.순위}. ${r.콘텐츠} · ${r.분류} · ${r.카테고리} — 조회 ${nf(r.조회)} · 좋아요 ${nf(r.좋아요)} · 나이 ${r.나이_일 ?? "—"}일${r.나이_일 != null && r.나이_일 <= 3 ? " (덜 자람)" : ""}`));
  }
  L.push("", "ⓓ 다음 액션: (직접 작성)");
  return L.join("\n");
}

$("#dash-week").addEventListener("change", async e => {
  D.week = e.target.value;
  $("#tab-dash").classList.add("loading");
  try { await drawWeek(); dashMsg(""); } catch (err) { dashMsg(err.message, "err"); }
  $("#tab-dash").classList.remove("loading");
});
document.querySelectorAll(".seg [data-plat]").forEach(b => b.addEventListener("click", () => {
  D.plat = b.dataset.plat;
  document.querySelectorAll(".seg [data-plat]").forEach(x => x.setAttribute("aria-pressed", String(x === b)));
  if (D.kpi) paintDash();
}));
$("#btn-report").addEventListener("click", async () => {
  const text = $("#report-text").value;
  let ok = false;
  try { await navigator.clipboard.writeText(text); ok = true; } catch (e) { ok = false; }
  if (!ok) {   // 클립보드 권한이 없으면 미리 보기를 펼쳐 골라 둔다
    $("#report-box").open = true;
    const ta = $("#report-text"); ta.focus(); ta.select();
    try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
  }
  dashMsg(ok ? `복사했습니다 — Asana «SNS 성과 분석 - ${weekName(D.week)}»에 붙여 넣고 ⓑ · ⓓ를 채우세요.`
    : "복사하지 못했습니다 — 아래 «보고서 요약 미리 보기»에서 골라 직접 복사하세요.", ok ? "ok" : "err");
});

// 차트 말풍선 — data-tip 이 있는 곳 위에서. 글자는 textContent 로만 넣는다
const tip = $("#tip");
document.addEventListener("pointermove", e => {
  const t = e.target.closest?.("[data-tip]");
  if (!t) { tip.hidden = true; return; }
  tip.textContent = t.dataset.tip; tip.hidden = false;
  const x = Math.min(e.clientX + 12, innerWidth - tip.offsetWidth - 8);
  tip.style.left = `${Math.max(8, x)}px`; tip.style.top = `${e.clientY - tip.offsetHeight - 10}px`;
});
document.addEventListener("pointerleave", () => { tip.hidden = true; });
// 키보드 포커스에도 같은 말풍선 (마우스 올리기에만 기대지 않는다)
document.addEventListener("focusin", e => {
  const t = e.target.closest?.("[data-focus-tip]");
  if (!t) return;
  tip.textContent = t.dataset.focusTip; tip.hidden = false;
  const r = t.getBoundingClientRect();
  tip.style.left = `${Math.max(8, Math.min(r.left, innerWidth - tip.offsetWidth - 8))}px`;
  tip.style.top = `${Math.min(innerHeight - tip.offsetHeight - 8, r.bottom + 6)}px`;   // 아래에 — 큰 숫자를 가리지 않게
});
document.addEventListener("focusout", e => { if (e.target.closest?.("[data-focus-tip]")) tip.hidden = true; });

// ── 탭 — 누르기 + 키보드(←/→ · Home/End), 고른 탭만 Tab 키로 들어간다 ─────────
const tabBtns = [...document.querySelectorAll(".tabs [role=tab]")];
tabBtns.forEach(b => {
  b.id = "tabbtn-" + b.dataset.tab;
  b.setAttribute("aria-controls", "tab-" + b.dataset.tab);
  b.tabIndex = b.getAttribute("aria-selected") === "true" ? 0 : -1;
  const p = $("#tab-" + b.dataset.tab);
  p.setAttribute("role", "tabpanel"); p.setAttribute("aria-labelledby", b.id);
  b.addEventListener("click", () => {
    tabBtns.forEach(x => { x.setAttribute("aria-selected", String(x === b)); x.tabIndex = x === b ? 0 : -1; });
    tabBtns.forEach(x => { $("#tab-" + x.dataset.tab).hidden = x !== b; });
  });
});
$(".tabs").addEventListener("keydown", e => {
  const i = tabBtns.indexOf(document.activeElement);
  if (i < 0) return;
  const j = e.key === "ArrowRight" ? (i + 1) % tabBtns.length : e.key === "ArrowLeft" ? (i - 1 + tabBtns.length) % tabBtns.length
    : e.key === "Home" ? 0 : e.key === "End" ? tabBtns.length - 1 : -1;
  if (j < 0) return;
  e.preventDefault(); tabBtns[j].focus(); tabBtns[j].click();
});

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
  renderQuestions(); renderTables(); renderErd();
  $("#mine-rule").textContent = S.files.length
    ? "파일 이름으로 표를 고릅니다 — " + S.files.map(f => `«${f.match}»가 든 파일 → ${f.table}`).join(" · ")
    : "questions.json 에 files 가 없어 내 파일을 받을 수 없습니다";
  try {
    await startEngine();
    engineLine();
    document.body.dataset.engine = "ready";
    renderDash();
  } catch (e) {
    $("#engine").textContent = "데이터 엔진을 불러오지 못했습니다 — 인터넷 연결을 확인하십시오";
    document.body.dataset.engine = "fail";
    $("#dash-cards").innerHTML = ""; $("#dash-top").innerHTML = "";
    dashMsg("데이터 엔진을 불러오지 못했습니다 — 인터넷 연결을 확인하고 새로고침하십시오", "err");
  }
})();
