/* Order Scanner — scan products into an order built from an Impact order-form .xlsx. Works offline. */
"use strict";
const APP_VERSION = "1.1.0";

// ---------------- storage (IndexedDB) ----------------
const DB = (() => {
  let dbp;
  const open = () => dbp || (dbp = new Promise((res, rej) => {
    const r = indexedDB.open("order-scanner", 1);
    r.onupgradeneeded = () => {
      const d = r.result;
      d.createObjectStore("forms", { keyPath: "id" });
      d.createObjectStore("orders", { keyPath: "id" });
      d.createObjectStore("aliases", { keyPath: "code" });
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  }));
  const tx = async (store, mode, fn) => {
    const d = await open();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode);
      const s = t.objectStore(store);
      const out = fn(s);
      t.oncomplete = () => res(out && "result" in out ? out.result : out);
      t.onerror = () => rej(t.error);
    });
  };
  return {
    all: (s) => tx(s, "readonly", (st) => st.getAll()),
    get: (s, k) => tx(s, "readonly", (st) => st.get(k)),
    put: (s, v) => tx(s, "readwrite", (st) => st.put(v)),
    del: (s, k) => tx(s, "readwrite", (st) => st.delete(k)),
  };
})();

// ---------------- helpers ----------------
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const money = (n) => (typeof n === "number" ? "$" + n.toFixed(2) : "call");
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };

function toast(msg, ms = 2200) {
  const t = $("toast");
  t.textContent = msg; t.hidden = false;
  clearTimeout(toast._h); toast._h = setTimeout(() => (t.hidden = true), ms);
}

/** Normalise any scanned / typed barcode to compare against the form. */
function normCode(s) {
  const d = String(s ?? "").replace(/\D/g, "");
  if (d.length === 13 && d[0] === "0") return d.slice(1);   // EAN-13 form of a UPC-A
  return d;
}

/** Turn the Order/status cell text into tags: NEW, OUT, back-in date. */
function parseStatus(s) {
  const t = String(s ?? "").toLowerCase();
  const eta = t.match(/eta\s*([\d/]+)/);
  return { isNew: t.includes("new"), out: t.includes("temp out") || !!eta, eta: eta ? eta[1] : "" };
}

// ---------------- form import ----------------
/** Parse one of her park order forms: B UPC, C description, E product, G price, H min, I order/status, J Impact #. */
function parseForm(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hdr = rows.findIndex((r) => r && String(r[9] ?? "").trim() === "Impact #");
  if (hdr < 0) throw new Error("Couldn't find the 'Impact #' header in column J. Is this one of the order forms?");
  const items = [];
  for (let i = hdr + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const item = String(r[9] ?? "").trim();
    if (!item || /\s/.test(item) || /^(total|freight)$/i.test(item)) continue;
    let moq = r[7], status = r[8];
    if (typeof moq === "string") { status = moq; moq = null; }        // e.g. "NEW" typed in the MOQ column
    const price = typeof r[6] === "number" ? Math.round(r[6] * 100) / 100 : null;
    items.push({
      row: i + 1, item, upc: normCode(r[1]),
      desc: String(r[2] ?? "").replace(/\s+/g, " ").trim(),
      product: String(r[4] ?? "").replace(/\s+/g, " ").trim(),
      price, moq: typeof moq === "number" ? moq : null, ...parseStatus(status),
    });
  }
  const name = fileName.replace(/\.xls[xm]$/i, "");
  return { id: name.toLowerCase(), name, sheet: wb.SheetNames[0], loadedAt: new Date().toISOString(), items };
}

// ---------------- app state ----------------
const S = { forms: [], orders: [], aliases: {}, order: null, form: null, index: null, pending: null, view: "home" };

function buildIndex(form) {
  const byCode = new Map(), byItem = new Map();
  for (const it of form.items) {
    byItem.set(it.item.toUpperCase(), it);
    if (it.upc && it.upc.length >= 8) {
      if (!byCode.has(it.upc)) byCode.set(it.upc, []);
      byCode.get(it.upc).push(it);            // some barcodes are shared (e.g. every mint tin design)
    }
  }
  return { byCode, byItem };
}

/** Items matching a scanned code: usually one; several when designs share a barcode; empty if unknown. */
function lookup(code) {
  const raw = String(code).trim();
  const n = normCode(raw);
  const { byCode, byItem } = S.index;
  const al = S.aliases[n || raw.toUpperCase()];
  if (al && al.form === S.form.id && byItem.has(al.item.toUpperCase())) return [byItem.get(al.item.toUpperCase())];
  if (n && byCode.has(n)) return byCode.get(n);
  if (byItem.has(raw.toUpperCase())) return [byItem.get(raw.toUpperCase())];   // Code-128 of an item #
  return [];
}

async function loadAll() {
  S.forms = (await DB.all("forms")).sort((a, b) => a.name.localeCompare(b.name));
  S.orders = (await DB.all("orders")).sort((a, b) => b.updated.localeCompare(a.updated));
  S.aliases = Object.fromEntries((await DB.all("aliases")).map((a) => [a.code, a]));
}

function orderTotals(o, form) {
  const byItem = new Map(form.items.map((i) => [i.item, i]));
  let qty = 0, total = 0, lines = 0;
  for (const [k, q] of Object.entries(o.lines)) {
    if (!q) continue;
    lines++; qty += q;
    const it = byItem.get(k);
    if (it && typeof it.price === "number") total += it.price * q;
  }
  return { qty, total, lines };
}

const saveOrder = async () => { S.order.updated = new Date().toISOString(); await DB.put("orders", S.order); };

// ---------------- navigation ----------------
function show(view) {
  S.view = view;
  for (const v of ["Home", "Header", "Scan", "Review"]) $("v" + v).hidden = v.toLowerCase() !== view;
  $("btnBack").hidden = view === "home";
  $("title").textContent = { home: "Order Scanner", header: "Order details", scan: S.order?.account || "Scan",
                             review: "Review order" }[view];
  if (view !== "scan") Cam.stop();
  if (view === "home") renderHome();
  if (view === "scan") renderScan();
  if (view === "review") renderReview();
}
$("btnBack").onclick = () => {
  if (S.view === "scan") show("review");
  else if (S.view === "review") show("home");
  else show("home");
};

// ---------------- home ----------------
function renderHome() {
  const fl = $("formList");
  fl.innerHTML = S.forms.map((f) => {
    const withCode = f.items.filter((i) => i.upc.length >= 8).length;
    return `<li><div class="grow"><div class="t">${esc(f.name)}</div>
      <div class="s">${f.items.length} items · ${withCode} with barcodes · loaded ${new Date(f.loadedAt).toLocaleDateString()}</div></div>
      <button class="btn danger" data-delform="${esc(f.id)}">Remove</button></li>`;
  }).join("");
  $("noForms").hidden = S.forms.length > 0;
  const ol = $("orderList");
  ol.innerHTML = S.orders.map((o) => {
    const f = S.forms.find((x) => x.id === o.form);
    const t = f ? orderTotals(o, f) : { lines: Object.keys(o.lines).length, total: 0 };
    return `<li><div class="grow"><div class="t">${esc(o.account || "(no account)")}${o.exported ? ' <span class="tag new">EXPORTED</span>' : ""}</div>
      <div class="s">${esc(f ? f.name : "form removed")} · ${t.lines} items · ${money(t.total)} · ${esc(o.date)}</div></div>
      <button class="btn" data-open="${o.id}">Open</button><button class="btn danger" data-delorder="${o.id}">Delete</button></li>`;
  }).join("");
  $("noOrders").hidden = S.orders.length > 0;
  $("ver").textContent = `Version ${APP_VERSION}`;
}

$("fileForm").onchange = async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  try {
    const form = parseForm(await file.arrayBuffer(), file.name);
    await DB.put("forms", form);
    await loadAll(); renderHome();
    const withCode = form.items.filter((i) => i.upc.length >= 8).length;
    toast(`Loaded ${form.name}: ${form.items.length} items (${withCode} with barcodes)`, 3500);
  } catch (err) { toast("Couldn't read that file: " + err.message, 5000); }
};

document.addEventListener("click", async (e) => {
  const b = e.target.closest("button");
  if (!b) return;
  if (b.dataset.delform && confirm("Remove this order form from the iPad? Orders using it can't be opened until it's loaded again.")) {
    await DB.del("forms", b.dataset.delform); await loadAll(); renderHome();
  }
  if (b.dataset.delorder && confirm("Delete this order?")) {
    await DB.del("orders", b.dataset.delorder); await loadAll(); renderHome();
  }
  if (b.dataset.open) {
    const o = S.orders.find((x) => x.id === b.dataset.open);
    if (!openOrder(o)) return;
    show(Object.keys(o.lines).length ? "review" : "scan");
  }
});

function openOrder(o) {
  const f = S.forms.find((x) => x.id === o.form);
  if (!f) { toast("Load the order form for this order first."); return false; }
  S.order = o; S.form = f; S.index = buildIndex(f); S.pending = null;
  return true;
}

// ---------------- header ----------------
function fillHeader(o) {
  $("hForm").innerHTML = S.forms.map((f) => `<option value="${esc(f.id)}">${esc(f.name)}</option>`).join("");
  if (o) $("hForm").value = o.form;
  $("hForm").disabled = !!(o && Object.keys(o.lines).length);
  $("hAccount").value = o?.account || ""; $("hAcct").value = o?.acct || ""; $("hPO").value = o?.po || "";
  $("hShip").value = o?.ship || ""; $("hBuyer").value = o?.buyer || ""; $("hDate").value = o?.date || today();
  $("hNotes").value = o?.notes || "";
  $("btnStartScan").textContent = o ? "Save" : "Start scanning";
}
$("btnNewOrder").onclick = () => {
  if (!S.forms.length) return toast("Load an order form first.");
  S.order = null; fillHeader(null); show("header"); $("hAccount").focus();
};
$("btnEditHeader").onclick = () => { fillHeader(S.order); show("header"); };
$("btnStartScan").onclick = async () => {
  const isNew = !S.order;
  const o = S.order || { id: uid(), lines: {}, created: new Date().toISOString() };
  Object.assign(o, { form: $("hForm").value, account: $("hAccount").value.trim(), acct: $("hAcct").value.trim(),
    po: $("hPO").value.trim(), ship: $("hShip").value.trim(), buyer: $("hBuyer").value.trim(),
    date: $("hDate").value || today(), notes: $("hNotes").value.trim() });
  openOrder(o); await saveOrder(); await loadAll();
  show(isNew ? "scan" : "review");
};

// ---------------- scanning ----------------
let audio;
function beep(ok = true) {
  try {
    audio = audio || new (window.AudioContext || window.webkitAudioContext)();
    const o = audio.createOscillator(), g = audio.createGain();
    o.frequency.value = ok ? 1650 : 330; g.gain.value = 0.15;
    o.connect(g); g.connect(audio.destination); o.start(); o.stop(audio.currentTime + (ok ? 0.09 : 0.3));
  } catch (_) { /* no audio */ }
}

// ---- image clean-up passes for hard frames (bars are vertical, so work along rows) ----
function toGray(id) {
  const d = id.data, n = id.width * id.height, g = new Uint8ClampedArray(n);
  for (let i = 0, j = 0; j < n; i += 4, j++) g[j] = (d[i] * 77 + d[i + 1] * 150 + d[i + 2] * 29) >> 8;
  return g;
}
function grayToImageData(g, w, h) {
  const o = new ImageData(w, h), d = o.data;
  for (let j = 0, i = 0; j < g.length; j++, i += 4) { d[i] = d[i + 1] = d[i + 2] = g[j]; d[i + 3] = 255; }
  return o;
}
/** Horizontal sliding-window mean of each row. */
function rowMean(g, w, h, r) {
  const o = new Float32Array(g.length), span = 2 * r + 1;
  for (let y = 0, base = 0; y < h; y++, base += w) {
    let acc = 0;
    for (let x = -r; x <= r; x++) acc += g[base + (x < 0 ? 0 : x >= w ? w - 1 : x)];
    for (let x = 0; x < w; x++) {
      o[base + x] = acc / span;
      const add = x + r + 1, sub = x - r;
      acc += g[base + (add >= w ? w - 1 : add)] - g[base + (sub < 0 ? 0 : sub)];
    }
  }
  return o;
}
/** Remove glare / uneven light: subtract the local average, then stretch contrast. */
function flattenLight(id) {
  const w = id.width, h = id.height, g = toGray(id), m = rowMean(g, w, h, 40), o = new Uint8ClampedArray(g.length);
  for (let i = 0; i < g.length; i++) o[i] = 128 + (g[i] - m[i]) * 2.2;
  return grayToImageData(o, w, h);
}
/** Sharpen slightly blurred bars (unsharp mask along rows). */
function sharpenBars(id) {
  const w = id.width, h = id.height, g = toGray(id), m = rowMean(g, w, h, 2), o = new Uint8ClampedArray(g.length);
  for (let i = 0; i < g.length; i++) o[i] = g[i] + 2.2 * (g[i] - m[i]);
  return grayToImageData(o, w, h);
}

// zxing-cpp compiled to WebAssembly: far faster than the JS port, so we can try ~10+ frames a second
const READ_OPTS = { formats: ["UPCA", "UPCE", "EAN13", "EAN8", "Code128"], tryHarder: true, tryRotate: true,
                    tryInvert: false, maxNumberOfSymbols: 1 };
const PASSES = ["full", "aim", "aim-flatten", "aim-sharpen"];

const Cam = {
  on: false, stream: null, track: null, caps: {}, timer: null, step: 0, tries: 0, last: "", lastAt: 0, ready: null,
  cv: null, cx: null, zoom: 1, torch: false,
  init() {
    return this.ready || (this.ready = ZXingWASM.prepareZXingModule({
      overrides: { locateFile: (p, prefix) => (p.endsWith(".wasm") ? new URL("vendor/" + p, location.href).href : prefix + p) },
      fireImmediately: true,
    }));
  },
  async start() {
    if (this.on) return;
    try {
      await this.init();
      this.stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: {
        facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 }, frameRate: { ideal: 30 } } });
      const v = $("video");
      v.srcObject = this.stream;
      await v.play();
      this.track = this.stream.getVideoTracks()[0];
      this.caps = (this.track.getCapabilities && this.track.getCapabilities()) || {};
      try { if ((this.caps.focusMode || []).includes("continuous")) await this.track.applyConstraints({ advanced: [{ focusMode: "continuous" }] }); } catch (_) {}
      this.on = true; $("camOff").hidden = true; beep(true);
      // zoom in a little so she can hold items further back, where the camera can focus
      if (this.caps.zoom) await this.setZoom(Math.min(2, this.caps.zoom.max));
      this.renderTools();
      this.statTimer = setInterval(() => this.renderStat(), 1000);
      this.loop();
    } catch (err) {
      this.stop();
      toast("Camera unavailable: " + (err.message || err.name), 4500);
    }
  },
  stop() {
    clearTimeout(this.timer); clearInterval(this.statTimer);
    if (this.stream) this.stream.getTracks().forEach((t) => t.stop());
    this.stream = this.track = null; this.on = false;
    $("video").srcObject = null; $("camOff").hidden = false; $("camTools").hidden = true;
  },
  async setZoom(z) {
    try { await this.track.applyConstraints({ advanced: [{ zoom: z }] }); this.zoom = z; } catch (_) {}
    this.renderTools();
  },
  async toggleTorch() {
    try { await this.track.applyConstraints({ advanced: [{ torch: !this.torch }] }); this.torch = !this.torch; } catch (_) {}
    this.renderTools();
  },
  renderTools() {
    const t = $("camTools");
    t.hidden = !this.on;
    if (!this.on) return;
    const z = this.caps.zoom ? [1, 2, 3].filter((x) => x >= this.caps.zoom.min && x <= this.caps.zoom.max) : [];
    $("camZoom").innerHTML = z.map((x) => `<button data-zoom="${x}" class="${x === this.zoom ? "on" : ""}">${x}×</button>`).join("");
    $("camTorch").hidden = !this.caps.torch;
    $("camTorch").classList.toggle("on", this.torch);
  },
  renderStat() {
    const s = this.track && this.track.getSettings ? this.track.getSettings() : {};
    $("camStat").textContent = `${this.tries} tries/sec · ${s.width || "?"}×${s.height || "?"}`;
    this.tries = 0;
  },
  loop() {
    if (!this.on) return;
    const v = $("video");
    const go = async () => {
      if (v.readyState >= 2 && v.videoWidth) { try { await this.tick(v); } catch (_) {} }
      if (this.on) this.timer = setTimeout(() => this.loop(), 30);
    };
    go();
  },
  /** One decode attempt; passes rotate so each frame gets a different chance. */
  async tick(src) {
    const code = await decodeFrame(src, PASSES[this.step++ % PASSES.length], this);
    this.tries++;
    if (code) this.found(code);
  },
  found(code) {
    const now = Date.now();
    if (code === this.last && now - this.lastAt < 2500) return;   // same label still in view
    this.last = code; this.lastAt = now;
    $("camBox").classList.add("hit"); setTimeout(() => $("camBox").classList.remove("hit"), 400);
    onCode(code, "camera");
  },
};

/** Grab one pass from a video/image and decode it. Returns the barcode text or "". */
async function decodeFrame(src, pass, buf) {
  const W = src.videoWidth || src.naturalWidth || src.width, H = src.videoHeight || src.naturalHeight || src.height;
  let sx = 0, sy = 0, sw = W, sh = H;
  if (pass !== "full") { sw = Math.round(W * 0.8); sh = Math.round(H * 0.5); sx = (W - sw) >> 1; sy = (H - sh) >> 1; }
  if (!buf.cv) { buf.cv = document.createElement("canvas"); buf.cx = buf.cv.getContext("2d", { willReadFrequently: true }); }
  if (buf.cv.width !== sw || buf.cv.height !== sh) { buf.cv.width = sw; buf.cv.height = sh; }
  buf.cx.drawImage(src, sx, sy, sw, sh, 0, 0, sw, sh);
  let id = buf.cx.getImageData(0, 0, sw, sh);
  if (pass === "aim-flatten") id = flattenLight(id);
  else if (pass === "aim-sharpen") id = sharpenBars(id);
  const out = await ZXingWASM.readBarcodes(id, READ_OPTS);
  const hit = out.find((o) => o.isValid && o.text);
  return hit ? hit.text : "";
}
$("camTools").onclick = (e) => {
  const b = e.target.closest("button");
  if (!b) return;
  if (b.dataset.zoom) Cam.setZoom(+b.dataset.zoom);
  if (b.id === "camTorch") Cam.toggleTorch();
};
$("btnCam").onclick = () => Cam.start();
$("btnCamToggle").onclick = () => (Cam.on ? Cam.stop() : Cam.start());

// Bluetooth / USB scanners type the code fast and finish with Enter.
const kb = { buf: "", t: 0 };
document.addEventListener("keydown", (e) => {
  if (S.view !== "scan") return;
  const tag = (e.target.tagName || "").toLowerCase();
  if (tag === "input" || tag === "textarea" || tag === "select") return;
  const now = Date.now();
  if (now - kb.t > 120) kb.buf = "";
  kb.t = now;
  if (e.key === "Enter") {
    if (kb.buf.length >= 6) { e.preventDefault(); onCode(kb.buf, "scanner"); kb.buf = ""; return; }
    if (S.pending) { e.preventDefault(); commitPending(); }
    return;
  }
  if (e.key.length === 1) {
    kb.buf += e.key;
    if (S.pending && /\d/.test(e.key) && kb.buf.length === 1) setTimeout(() => {   // a lone keypress = typed qty
      if (kb.buf === e.key && Date.now() - kb.t >= 100) { padKey(e.key); kb.buf = ""; }
    }, 110);
  } else if (e.key === "Backspace" && S.pending) padKey("⌫");
});

function onCode(code, source) {
  if (S.pending && S.pending.qty > 0) commitPending();     // scanning the next item confirms the last one
  const hits = lookup(code);
  S.unknownCode = null;
  if (!hits.length) {
    beep(false);
    S.unknownCode = normCode(code) || String(code).trim().toUpperCase();
    $("search").value = "";
    showResults([], `Barcode ${esc(code)} isn't on ${esc(S.form.name)}. Search for the item to link it:`);
    $("search").focus();
    return;
  }
  beep(true);
  if (hits.length > 1) {      // shared barcode: let her tap the design
    S.pending = null; renderPending();
    showResults(hits, `This barcode is on ${hits.length} items — tap the one you're ordering:`);
    return;
  }
  startPending(hits[0]);
}

function startPending(it) {
  S.pending = { it, qty: 0 };
  $("results").hidden = true;
  renderPending();
}

function renderPending() {
  const p = S.pending;
  $("pending").hidden = !p; $("idle").hidden = !!p;
  if (!p) return;
  const it = p.it;
  $("pDesc").innerHTML = esc(it.desc) + (it.isNew ? '<span class="tag new">NEW</span>' : "");
  $("pMeta").textContent = `${it.product} · ${it.item} · ${money(it.price)}${it.moq ? " · min " + it.moq : ""}`;
  $("pWarn").hidden = !it.out;
  $("pWarn").textContent = it.out ? (it.eta ? `Temporarily out — back ${it.eta}` : "Temporarily out — no back-in date yet") : "";
  const have = S.order.lines[it.item] || 0;
  $("pAlready").textContent = have ? `Already ${have} on this order — this adds to it.` : "";
  $("pQty").textContent = p.qty;
  const m = it.moq || 1;
  const chips = [...new Set([m, m * 2, m * 3, m * 4])];
  $("pChips").innerHTML = chips.map((c) => `<button data-chip="${c}">${c}</button>`).join("");
  $("btnAdd").textContent = p.qty ? `Add ${p.qty} to order` : "Add to order";
}

$("pad").innerHTML = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "C", "0", "⌫"].map((k) => `<button data-k="${k}">${k}</button>`).join("");
$("pad").onclick = (e) => { const k = e.target.closest("button")?.dataset.k; if (k) padKey(k); };
$("pChips").onclick = (e) => { const c = e.target.closest("button")?.dataset.chip; if (c) { S.pending.qty = +c; renderPending(); } };
function padKey(k) {
  if (!S.pending) return;
  let q = String(S.pending.qty || "");
  if (k === "C") q = "";
  else if (k === "⌫") q = q.slice(0, -1);
  else if (q.length < 5) q += k;
  S.pending.qty = parseInt(q || "0", 10);
  renderPending();
}
$("btnAdd").onclick = () => commitPending();
$("btnCancelPending").onclick = () => { S.pending = null; renderPending(); };

async function commitPending() {
  const p = S.pending;
  if (!p) return;
  if (!p.qty) { toast("Enter a quantity first."); return; }
  S.order.lines[p.it.item] = (S.order.lines[p.it.item] || 0) + p.qty;
  S.order.exported = false;
  S.pending = null;
  await saveOrder();
  renderPending(); renderFoot(`Added ${p.qty} × ${p.it.desc} (${p.it.product})`);
  if (p.it.moq && S.order.lines[p.it.item] < p.it.moq) toast(`Below the minimum of ${p.it.moq}`);
}

function renderScan() {
  $("title").textContent = S.order.account || S.form.name;
  renderPending(); renderFoot();
  $("results").hidden = true; $("search").value = "";
}
function renderFoot(last) {
  const t = orderTotals(S.order, S.form);
  $("fLines").textContent = `${t.lines} item${t.lines === 1 ? "" : "s"}`;
  $("fTotal").textContent = money(t.total);
  if (last !== undefined) $("fLast").textContent = last;
}

// search
function showResults(list, note) {
  const r = $("results");
  r.hidden = false;
  r.innerHTML = (note ? `<li class="muted" style="cursor:default">${note}</li>` : "") +
    list.slice(0, 40).map((it) => `<li data-item="${esc(it.item)}"><b>${esc(it.desc)}</b>${it.out ? '<span class="tag out">OUT</span>' : ""}
      <div class="i">${esc(it.product)} · ${esc(it.item)} · ${money(it.price)}</div></li>`).join("") +
    (note || list.length ? "" : `<li class="muted">No matches.</li>`);
}
$("search").oninput = () => {
  const q = $("search").value.trim().toLowerCase();
  if (!q) { if (!S.unknownCode) $("results").hidden = true; return; }
  const words = q.split(/\s+/);
  const n = normCode(q);
  const hits = S.form.items.filter((it) => {
    const hay = `${it.desc} ${it.product} ${it.item}`.toLowerCase();
    return words.every((w) => hay.includes(w)) || (n.length >= 5 && it.upc.includes(n));
  });
  showResults(hits, S.unknownCode ? `Linking barcode ${esc(S.unknownCode)} — pick the item:` : "");
};
$("results").onclick = async (e) => {
  const li = e.target.closest("li[data-item]");
  if (!li) return;
  const it = S.index.byItem.get(li.dataset.item.toUpperCase());
  if (S.unknownCode) {   // remember this barcode for next time
    await DB.put("aliases", { code: S.unknownCode, form: S.form.id, item: it.item });
    S.aliases[S.unknownCode] = { code: S.unknownCode, form: S.form.id, item: it.item };
    toast("Barcode linked — it'll scan straight to this item next time.");
    S.unknownCode = null;
  }
  $("search").value = ""; $("search").blur();
  startPending(it);
};
$("btnReview").onclick = () => { if (S.pending?.qty) commitPending(); show("review"); };

// ---------------- review ----------------
function orderedItems() {
  return S.form.items.filter((it) => S.order.lines[it.item] > 0);
}
function renderReview() {
  const its = orderedItems();
  const t = orderTotals(S.order, S.form);
  $("rTitle").textContent = S.order.account || "(no account)";
  $("rSub").textContent = `${S.form.name} · ${S.order.date}${S.order.po ? " · PO " + S.order.po : ""}`;
  const low = its.filter((it) => it.moq && S.order.lines[it.item] < it.moq);
  const outs = its.filter((it) => it.out);
  const w = [];
  if (low.length) w.push(`${low.length} line${low.length > 1 ? "s are" : " is"} below the order minimum (shaded).`);
  if (outs.length) w.push(`${outs.length} item${outs.length > 1 ? "s are" : " is"} temporarily out.`);
  $("rWarn").hidden = !w.length; $("rWarn").textContent = w.join(" ");
  $("rBody").innerHTML = its.map((it) => {
    const q = S.order.lines[it.item];
    const lt = typeof it.price === "number" ? it.price * q : null;
    return `<tr class="${it.moq && q < it.moq ? "low" : ""}">
      <td><div class="d">${esc(it.desc)}${it.out ? `<span class="tag out">${it.eta ? "BACK " + esc(it.eta) : "OUT"}</span>` : ""}</div>
          <div class="i">${esc(it.product)} · ${esc(it.item)}</div></td>
      <td class="num">${money(it.price)}</td><td class="num">${it.moq ?? ""}</td>
      <td class="num"><span class="qbtn"><button data-dq="${esc(it.item)}" data-d="-1">−</button>
        <input inputmode="numeric" data-q="${esc(it.item)}" value="${q}"><button data-dq="${esc(it.item)}" data-d="1">+</button></span></td>
      <td class="num">${lt === null ? "" : money(lt)}</td>
      <td><button class="btn danger" data-rm="${esc(it.item)}">✕</button></td></tr>`;
  }).join("") || `<tr><td colspan="6" class="muted">Nothing on this order yet.</td></tr>`;
  $("rQty").textContent = t.qty; $("rTotal").textContent = money(t.total);
}
$("rBody").onclick = async (e) => {
  const b = e.target.closest("button");
  if (!b) return;
  if (b.dataset.rm) delete S.order.lines[b.dataset.rm];
  if (b.dataset.dq) {
    const it = S.index.byItem.get(b.dataset.dq.toUpperCase());
    const step = it.moq && it.moq > 1 ? it.moq : 1;                 // +/- by the order minimum
    S.order.lines[it.item] = Math.max(0, (S.order.lines[it.item] || 0) + step * +b.dataset.d);
    if (!S.order.lines[it.item]) delete S.order.lines[it.item];
  }
  S.order.exported = false; await saveOrder(); renderReview();
};
$("rBody").onchange = async (e) => {
  const k = e.target.dataset.q;
  if (!k) return;
  const v = parseInt(e.target.value || "0", 10);
  if (v > 0) S.order.lines[k] = v; else delete S.order.lines[k];
  S.order.exported = false; await saveOrder(); renderReview();
};
$("btnBackToScan").onclick = () => show("scan");

// ---------------- export ----------------
function buildOrderWorkbook() {
  const o = S.order, its = orderedItems();
  const aoa = [
    [`${S.form.name} — Order`], [],
    ["Account:", o.account, "", "Date:", o.date], ["Acct #:", o.acct, "", "PO #:", o.po],
    ["Ship to:", o.ship, "", "Buyer / phone:", o.buyer], ["Notes:", o.notes], [],
    ["Impact #", "UPC #", "Description", "Product", "Price", "MOQ", "Order", "Total"],
  ];
  const first = aoa.length + 1;
  its.forEach((it, k) => {
    const r = first + k;
    aoa.push([it.item, it.upc, it.desc, it.product, it.price ?? "", it.moq ?? "", o.lines[it.item],
              { f: `IF(ISNUMBER(E${r}),E${r}*G${r},0)` }]);
  });
  const last = first + its.length - 1;
  aoa.push(["", "", "", "", "", "Total", { f: `SUM(G${first}:G${last})` }, { f: `SUM(H${first}:H${last})` }]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 14 }, { wch: 14 }, { wch: 40 }, { wch: 26 }, { wch: 9 }, { wch: 7 }, { wch: 8 }, { wch: 11 }];
  for (let r = first; r <= last + 1; r++) {
    for (const c of ["E", "H"]) if (ws[c + r]) ws[c + r].z = "$#,##0.00";
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Order");
  return wb;
}
$("btnExport").onclick = async () => {
  if (!orderedItems().length) return toast("Nothing to export yet.");
  const safe = (s) => String(s || "").replace(/[\\/:*?"<>|]+/g, " ").trim();
  const name = `Order - ${safe(S.order.account) || "store"} - ${safe(S.form.name)} - ${S.order.date}.xlsx`;
  const data = XLSX.write(buildOrderWorkbook(), { bookType: "xlsx", type: "array" });
  const file = new File([data], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  S.order.exported = true; await saveOrder();
  try {
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: name });
      return;
    }
  } catch (err) { if (err.name === "AbortError") return; }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(file); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
};

// ---------------- boot ----------------
let reloaded = false;
function netStatus() { $("net").textContent = navigator.onLine ? "" : "Offline ✓"; }
addEventListener("online", netStatus); addEventListener("offline", netStatus);
(async () => {
  netStatus();
  if (navigator.storage?.persist) navigator.storage.persist().catch(() => {});
  if ("serviceWorker" in navigator) {
    const hadController = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register("sw.js").catch(() => {});
    // a new version took over: reload once so she's on it (skip on very first install)
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (hadController && S.view === "home" && !reloaded) { reloaded = true; location.reload(); }
    });
  }
  await loadAll();
  show("home");
})();

// test hooks (used by automated checks; harmless in normal use)
window.__os = { decodeFrame, PASSES, parseForm, onCode, normCode, S, DB, loadAll, buildOrderWorkbook, show, commitPending, padKey };
