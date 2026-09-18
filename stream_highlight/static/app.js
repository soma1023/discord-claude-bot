/* 配信ハイライト抽出ツール - 画面側 */
"use strict";

const state = {
  result: null,
  videoKey: null,
  tab: "all",
  moments: [],
  visible: [],
  selected: null,
  view: { start: 0, end: 0 },
  overlays: new Set(),
  keyword: null,
  hover: null,
  drag: null,
  audioOverlayTouched: false,
  poll: null,
};

const CATEGORY_COLORS = {};
const AUDIO_KEY = "__audio__";
// これ未満の跳ねは「声が大きくなった」と言うには弱いので、バッジを出さない
const AUDIO_BADGE_DB = 3.0;
const $ = (id) => document.getElementById(id);

/* ------------------------------------------------------------ 小物 */

function fmtTime(sec) {
  sec = Math.max(0, Math.round(sec));
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = sec % 60;
  const mm = h > 0 ? String(m).padStart(2, "0") : String(m);
  return (h > 0 ? h + ":" : "") + mm + ":" + String(s).padStart(2, "0");
}

function fmtDuration(sec) {
  const h = Math.floor(sec / 3600);
  const m = Math.round((sec % 3600) / 60);
  return h > 0 ? `${h}時間${m}分` : `${m}分`;
}

function num(n) {
  return Number(n).toLocaleString("ja-JP");
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text == null ? "" : String(text);
  return div.innerHTML;
}

async function api(path, options) {
  const res = await fetch(path, options);
  let payload = null;
  try {
    payload = await res.json();
  } catch (e) {
    payload = null;
  }
  if (!res.ok) {
    throw new Error((payload && payload.detail) || `通信に失敗しました (${res.status})`);
  }
  return payload;
}

function showError(message) {
  const panel = $("errorPanel");
  panel.textContent = message;
  panel.classList.remove("hidden");
}

function clearError() {
  $("errorPanel").classList.add("hidden");
}

/* ------------------------------------------------------------ 解析の実行 */

function currentParams() {
  return {
    min_z: parseFloat($("minZ").value),
    merge_sec: parseInt($("mergeSec").value, 10),
    lead_sec: parseInt($("leadSec").value, 10),
    tail_sec: parseInt($("tailSec").value, 10),
    window_sec: parseInt($("windowSec").value, 10),
    top_n: parseInt($("topN").value, 10),
    skip_start_sec: parseInt($("skipStart").value, 10),
    audio_min_z: parseFloat($("audioMinZ").value),
    chat_support_z: parseFloat($("chatSupportZ").value),
  };
}

function setBusy(busy, message) {
  $("runBtn").disabled = busy;
  $("progressPanel").classList.toggle("hidden", !busy);
  if (message) $("progressMessage").textContent = message;
  if (!busy) $("progressBar").style.width = "0%";
}

async function startAnalyze(url, refresh) {
  clearError();
  setBusy(true, "準備中…");
  try {
    const { job_id } = await api("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        url, refresh: !!refresh, with_audio: $("audioInput").checked,
        params: currentParams(),
      }),
    });
    pollJob(job_id);
  } catch (err) {
    setBusy(false);
    showError(err.message);
  }
}

function pollJob(jobId) {
  if (state.poll) clearInterval(state.poll);
  state.poll = setInterval(async () => {
    let job;
    try {
      job = await api(`/api/job/${jobId}`);
    } catch (err) {
      clearInterval(state.poll);
      setBusy(false);
      showError(err.message);
      return;
    }
    $("progressMessage").textContent = job.message || "処理中…";
    $("progressBar").style.width = Math.round((job.progress || 0) * 100) + "%";
    if (job.status === "done") {
      clearInterval(state.poll);
      setBusy(false);
      applyResult(job.result);
      loadHistory();
    } else if (job.status === "error") {
      clearInterval(state.poll);
      setBusy(false);
      showError(job.error || "解析に失敗しました。");
    }
  }, 600);
}

async function reanalyze() {
  if (!state.videoKey) return;
  $("applyBtn").disabled = true;
  try {
    const result = await api("/api/reanalyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        video_key: state.videoKey, with_audio: $("audioInput").checked,
        params: currentParams(),
      }),
    });
    state.keyword = null;
    applyResult(result);
  } catch (err) {
    showError(err.message);
  } finally {
    $("applyBtn").disabled = false;
  }
}

/* ------------------------------------------------------------ 結果の反映 */

function applyResult(result) {
  if (state.videoKey && state.videoKey !== result.video_key) {
    // 別の配信に切り替わったら、前の検索結果は捨てる
    state.keyword = null;
    state.overlays.clear();
    $("keywordInfo").textContent = "";
  }
  state.result = result;
  state.videoKey = result.video_key;
  state.selected = null;
  state.view = { start: 0, end: result.stats.duration || 1 };
  (result.categories || []).forEach((c) => (CATEGORY_COLORS[c.id] = c.color));
  PAD.r = result.audio && result.audio.available ? 46 : 14;   // 右軸のラベル用
  // 音声を解析したなら、わざわざ消すまでは重ねて見せる
  if (result.audio && result.audio.available && !state.audioOverlayTouched) {
    state.overlays.add(AUDIO_KEY);
  }

  // サーバ側で丸められた値をUIに戻す
  const p = result.params;
  $("minZ").value = p.min_z;
  $("mergeSec").value = p.merge_sec;
  $("leadSec").value = p.lead_sec;
  $("tailSec").value = p.tail_sec;
  $("windowSec").value = p.window_sec;
  $("topN").value = p.top_n;
  $("audioMinZ").value = p.audio_min_z;
  $("chatSupportZ").value = p.chat_support_z;
  $("skipStart").value = p.skip_start_sec;
  syncOutputs();

  $("resultPanel").classList.remove("hidden");
  renderVideo(result);
  renderAudioSummary(result);
  renderOverlayToggles(result);
  renderTabs();
  selectTab(state.keyword ? "keyword" : "all");
  resizeCanvas();
}

function renderVideo(result) {
  const v = result.video;
  const s = result.stats;
  const thumb = $("videoThumb");
  if (v.thumbnail) {
    thumb.src = v.thumbnail;
    thumb.classList.remove("hidden");
  } else {
    thumb.classList.add("hidden");
  }
  $("videoTitle").textContent = v.title || v.url;
  $("videoTitle").href = v.url;
  $("videoChannel").textContent = v.channel || "";
  $("videoStats").innerHTML = [
    `長さ <b>${fmtDuration(s.duration)}</b>`,
    `コメント <b>${num(s.messages)}</b>件`,
    `平均 <b>${s.per_minute}</b>コメ/分`,
    `参加 <b>${num(s.authors)}</b>人`,
    s.superchats ? `スパチャ <b>${num(s.superchats)}</b>件` : "",
    s.system_notices ? `<span title="サブスク告知などは見せ場の判定から除いています">通知 <b>${num(s.system_notices)}</b>件を除外</span>` : "",
    `候補 <b>${result.moments.length}</b>件`,
  ].filter(Boolean).map((text) => `<span>${text}</span>`).join("");

  const notice = $("noticePanel");
  notice.textContent = result.notice || "";
  notice.classList.toggle("hidden", !result.notice);
}

function renderAudioSummary(result) {
  const box = $("audioSummary");
  const filter = $("audioFilterLabel");
  const audio = result.audio || {};
  if (!audio.available) {
    box.classList.add("hidden");
    filter.classList.add("hidden");
    return;
  }
  const st = audio.stats;
  if (!st.raw_peaks) {
    // 音量の変化が緩やかな配信では、跳ねた箇所が無いのが正しい答えになる
    box.innerHTML =
      `<b>音量が大きく跳ねた箇所は見つかりませんでした。</b>` +
      `<br><span class="small">この配信は音量の変化が緩やかなようです。` +
      `「解析の設定」の<b>音量のしきい値</b>を下げると候補が増えます。` +
      `各候補に付く「声 +◯dB」のバッジはしきい値と無関係に出ます。</span>`;
  } else if (st.chat_too_sparse) {
    // コメントがほとんど流れない配信では、照合しようにも材料がない
    box.innerHTML =
      `音量が跳ねた箇所を <b>${st.raw_peaks}</b>件 見つけました。` +
      `ただしこの配信はコメントが <b>${st.chat_per_minute}</b>件/分 と少なく、` +
      `<b>音声とチャットの照合ができません</b>。` +
      `<br><span class="small">照合は「音量が上がった直後にコメントが増えたか」で判定するため、` +
      `もともとコメントがほとんど流れない配信では成立しません。` +
      `音量の山はグラフのピンクの線で確認できます。` +
      `各候補の「声 +◯dB」バッジはそのまま使えます。</span>`;
  } else {
    box.innerHTML =
      `音量が跳ねた箇所 <b>${st.raw_peaks}</b>件 → チャットも反応していたのは <b>${st.confirmed}</b>件` +
      `（<b>${st.rejected}</b>件を除外）。` +
      `うち <b>${st.already_in_chat_list}</b>件はチャット側で既に拾えていたので、` +
      `「音声」タブには残り <b>${result.audio_moments.length}</b>件を出しています。` +
      `<br><span class="small">除外した分の多くはゲームのSEなど、音量だけ大きい箇所です。</span>`;
  }
  box.classList.remove("hidden");
  filter.classList.remove("hidden");
}

const AUDIO_OVERLAY = { id: "__audio__", label: "音量", color: "#db61a2" };

function renderOverlayToggles(result) {
  const box = $("overlayToggles");
  box.innerHTML = "";
  const items = (result.categories || []).slice();
  if (result.audio && result.audio.available) items.unshift(AUDIO_OVERLAY);
  items.forEach((cat) => {
    const btn = document.createElement("button");
    btn.className = "toggle" + (state.overlays.has(cat.id) ? " on" : "");
    btn.textContent = cat.label;
    btn.style.borderColor = cat.color;
    btn.style.color = state.overlays.has(cat.id) ? "#06111f" : cat.color;
    btn.style.background = state.overlays.has(cat.id) ? cat.color : "transparent";
    btn.onclick = () => {
      if (cat.id === AUDIO_KEY) state.audioOverlayTouched = true;
      if (state.overlays.has(cat.id)) state.overlays.delete(cat.id);
      else state.overlays.add(cat.id);
      renderOverlayToggles(result);
      draw();
    };
    box.appendChild(btn);
  });
}

function renderTabs() {
  const tabs = $("tabs");
  tabs.innerHTML = "";
  const items = [{ id: "all", label: "総合" }];
  (state.result.categories || []).forEach((cat) => {
    const list = state.result.category_moments[cat.id];
    if (list && list.length) items.push({ id: cat.id, label: `${cat.label} (${list.length})` });
  });
  const audioList = state.result.audio_moments || [];
  if (audioList.length) items.push({ id: "audio", label: `音声 (${audioList.length})` });
  if (state.keyword) items.push({ id: "keyword", label: `「${state.keyword.keyword}」` });

  items.forEach((item) => {
    const btn = document.createElement("button");
    btn.className = "tab" + (state.tab === item.id ? " active" : "");
    btn.textContent = item.label;
    btn.onclick = () => selectTab(item.id);
    tabs.appendChild(btn);
  });
}

function momentsForTab(tab) {
  if (!state.result) return [];
  if (tab === "all") return state.result.moments;
  if (tab === "audio") return state.result.audio_moments || [];
  if (tab === "keyword") return state.keyword ? state.keyword.moments : [];
  return state.result.category_moments[tab] || [];
}

function selectTab(tab) {
  // 中身が無いタブを選んでしまったら総合に戻す
  state.tab = tab === "all" || momentsForTab(tab).length ? tab : "all";
  state.moments = momentsForTab(state.tab);
  state.selected = null;
  renderTabs();
  renderMoments();
  draw();
}

$("audioFilter").addEventListener("change", () => {
  state.selected = null;
  renderMoments();
  draw();
});

function renderMoments() {
  const list = $("momentList");
  list.innerHTML = "";
  const audioOnly = $("audioFilter").checked && !$("audioFilterLabel").classList.contains("hidden");
  const moments = audioOnly
    ? state.moments.filter((m) => m.audio && m.audio.excess_db >= AUDIO_BADGE_DB)
    : state.moments;
  state.visible = moments;

  let summary;
  if (state.tab === "keyword" && state.keyword) {
    summary = `「${state.keyword.keyword}」は ${num(state.keyword.hits)} 回。よく飛んでいた順に ${moments.length} 箇所。`;
  } else if (!moments.length) {
    summary = audioOnly
      ? "音声も跳ねた候補はありませんでした。絞り込みを外してください。"
      : "該当する候補がありません。感度を下げてみてください。";
  } else {
    summary = `盛り上がった順に ${moments.length} 箇所。時刻をクリックすると配信のその場面が開きます。`;
  }
  $("listSummary").textContent = summary;

  moments.forEach((m, index) => {
    const li = document.createElement("li");
    li.className = "moment";
    li.dataset.index = index;

    const chips = (m.tags || []).map((t) => {
      const color = CATEGORY_COLORS[t.id] || "#8b949e";
      return `<span class="chip" style="background:${color}">${escapeHtml(t.label)} ${t.count}</span>`;
    }).join("") + (
      m.audio && m.audio.excess_db >= AUDIO_BADGE_DB
        ? `<span class="chip audio" title="平常の音量より ${m.audio.excess_db}dB 大きい">声 +${m.audio.excess_db}dB</span>`
        : ""
    );

    // 普段ほぼ出ないワードでは「普段の◯倍」が意味を持たないので出し分ける
    const hasBaseline = m.baseline_rate >= 1;
    const metrics = [
      hasBaseline ? `普段の <b>${m.ratio}倍</b>` : "",
      `スコア <b>${m.score}</b>`,
      hasBaseline
        ? `<b>${m.rate}</b> コメ/分（平常 ${m.baseline_rate}）`
        : `<b>${m.rate}</b> コメ/分`,
      `<b>${num(m.messages)}</b>コメント / <b>${num(m.authors)}</b>人`,
      m.superchats ? `スパチャ <b>${m.superchats}</b>` : "",
    ].filter(Boolean).map((text) => `<span>${text}</span>`).join("");

    const comments = (m.top_comments || []).slice(0, 4).map((c) =>
      `<div class="comment-line">${escapeHtml(c.text)}<span class="count">×${c.count}</span></div>`
    ).join("");

    li.innerHTML = `
      <div class="rank">${m.rank}</div>
      <div class="body">
        <div>
          <a class="time" href="${escapeHtml(m.url)}" target="_blank" rel="noopener">${fmtTime(m.clip_start)}</a>
          <span class="range">${fmtTime(m.clip_start)} 〜 ${fmtTime(m.clip_end)}（ピーク ${fmtTime(m.peak_sec)}）</span>
          ${m.source === "audio" ? '<span class="source-audio">音声から検出</span>' : ""}
        </div>
        <div class="metrics">${metrics}</div>
        <div class="chips">${chips}</div>
        <div class="comments">${comments}</div>
      </div>
      <div class="actions">
        <button class="ghost" type="button" data-act="open">▶ 開く</button>
        <button class="ghost" type="button" data-act="copy">時刻コピー</button>
      </div>`;

    li.onclick = (ev) => {
      const act = ev.target.dataset ? ev.target.dataset.act : null;
      if (act === "open") {
        window.open(m.url, "_blank", "noopener");
        return;
      }
      if (act === "copy") {
        navigator.clipboard.writeText(`${fmtTime(m.clip_start)} ${m.url}`);
        ev.target.textContent = "コピーした";
        setTimeout(() => (ev.target.textContent = "時刻コピー"), 1200);
        return;
      }
      if (ev.target.classList.contains("time")) return;
      selectMoment(index, true);
    };
    list.appendChild(li);
  });
  highlightSelected();
}

function selectMoment(index, zoom) {
  state.selected = index;
  const m = state.visible[index];
  if (m && zoom) {
    const pad = Math.max(120, (m.clip_end - m.clip_start) * 4);
    setView(m.peak_sec - pad, m.peak_sec + pad);
  }
  highlightSelected();
  draw();
}

function highlightSelected() {
  document.querySelectorAll(".moment").forEach((el) => {
    el.classList.toggle("selected", Number(el.dataset.index) === state.selected);
  });
}

/* ------------------------------------------------------------ グラフ */

const canvas = $("graph");
const ctx = canvas.getContext("2d");
const PAD = { l: 46, r: 14, t: 26, b: 24 };   // r は音量の右軸がある時だけ広げる
let plot = { w: 0, h: 0, width: 0, height: 0 };

function resizeCanvas() {
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  if (!rect.width) return;
  canvas.width = Math.round(rect.width * dpr);
  canvas.height = Math.round(rect.height * dpr);
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  plot.width = rect.width;
  plot.height = rect.height;
  plot.w = rect.width - PAD.l - PAD.r;
  plot.h = rect.height - PAD.t - PAD.b;
  draw();
}

function span() {
  return Math.max(1, state.view.end - state.view.start);
}
function timeToX(t) {
  return PAD.l + ((t - state.view.start) / span()) * plot.w;
}
function xToTime(x) {
  return state.view.start + ((x - PAD.l) / plot.w) * span();
}

function setView(start, end) {
  const duration = state.result ? state.result.stats.duration : 1;
  let s = Math.max(0, start);
  let e = Math.min(duration, end);
  if (e - s < 20) {
    const mid = (s + e) / 2;
    s = Math.max(0, mid - 10);
    e = Math.min(duration, mid + 10);
  }
  state.view = { start: s, end: e };
}

/* 1ピクセル幅ごとに最大値を取る。間引いてもピークが消えないようにする。 */
function columns(values, binSec) {
  const width = Math.max(1, Math.floor(plot.w));
  const out = new Array(width).fill(0);
  for (let px = 0; px < width; px++) {
    const t0 = xToTime(PAD.l + px);
    const t1 = xToTime(PAD.l + px + 1);
    const i0 = Math.max(0, Math.floor(t0 / binSec));
    const i1 = Math.min(values.length, Math.max(i0 + 1, Math.ceil(t1 / binSec)));
    let best = 0;
    for (let i = i0; i < i1; i++) if (values[i] > best) best = values[i];
    out[px] = best;
  }
  return out;
}

function niceMax(value) {
  if (value <= 0) return 1;
  const exp = Math.pow(10, Math.floor(Math.log10(value)));
  const scaled = value / exp;
  const step = scaled <= 1 ? 1 : scaled <= 2 ? 2 : scaled <= 5 ? 5 : 10;
  return step * exp;
}

const TICK_STEPS = [10, 30, 60, 120, 300, 600, 900, 1800, 3600, 7200];

function draw() {
  if (!state.result || !plot.w) return;
  const series = state.result.series;
  const bin = series.bin_sec;

  const rate = columns(series.rate, bin);
  const base = columns(series.baseline, bin);
  const overlays = [];
  state.overlays.forEach((cid) => {
    if (cid === AUDIO_KEY) return;   // 音量は単位が違うので右軸で別に描く
    const values = series.categories[cid];
    if (values) overlays.push({ color: CATEGORY_COLORS[cid] || "#fff", cols: columns(values, bin) });
  });
  const audioOn = state.overlays.has(AUDIO_KEY) &&
                  state.result.audio && state.result.audio.available;
  // 見せ場と判定される水準を超えた分だけ描く。
  // 平常どおりの音量まで描くと線が全面を埋めて、山が見えなくなる。
  const audioFloor = audioOn ? (state.result.audio.threshold_db || 0) : 0;
  const audioCols = audioOn
    ? columns(state.result.audio.excess, bin).map((v) => Math.max(0, v - audioFloor))
    : null;
  if (state.tab === "keyword" && state.keyword && state.keyword.series.length) {
    overlays.push({ color: "#ffd866", cols: columns(state.keyword.series, state.keyword.bin_sec) });
  }

  let top = 0;
  rate.forEach((v) => (top = Math.max(top, v)));
  overlays.forEach((o) => o.cols.forEach((v) => (top = Math.max(top, v))));
  top = niceMax(top * 1.1);
  const yOf = (v) => PAD.t + plot.h - (v / top) * plot.h;

  ctx.clearRect(0, 0, plot.width, plot.height);

  // 目盛り
  ctx.strokeStyle = "#2a3140";
  ctx.fillStyle = "#8b949e";
  ctx.font = "11px system-ui, sans-serif";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {
    const v = (top / 4) * i;
    const y = Math.round(yOf(v)) + 0.5;
    ctx.beginPath();
    ctx.moveTo(PAD.l, y);
    ctx.lineTo(PAD.l + plot.w, y);
    ctx.stroke();
    ctx.textAlign = "right";
    ctx.fillText(Math.round(v), PAD.l - 6, y + 4);
  }
  ctx.textAlign = "left";
  ctx.fillText("コメ/分", 4, PAD.t - 12);

  // 時刻ラベル
  const step = TICK_STEPS.find((s) => span() / s <= 9) || TICK_STEPS[TICK_STEPS.length - 1];
  ctx.textAlign = "center";
  for (let t = Math.ceil(state.view.start / step) * step; t <= state.view.end; t += step) {
    const x = Math.round(timeToX(t)) + 0.5;
    ctx.strokeStyle = "#20262f";
    ctx.beginPath();
    ctx.moveTo(x, PAD.t);
    ctx.lineTo(x, PAD.t + plot.h);
    ctx.stroke();
    ctx.fillStyle = "#8b949e";
    const label = fmtTime(t);
    const half = ctx.measureText(label).width / 2;
    ctx.textAlign = x + half > plot.width - 2 ? "right"
                  : x - half < 2 ? "left" : "center";
    ctx.fillText(label, x, plot.height - 8);
    ctx.textAlign = "center";
  }

  // 選択中の見せ場の範囲
  const selected = state.visible[state.selected];
  if (selected) {
    const x0 = timeToX(selected.clip_start);
    const x1 = timeToX(selected.clip_end);
    ctx.fillStyle = "rgba(88,166,255,0.18)";
    ctx.fillRect(x0, PAD.t, Math.max(2, x1 - x0), plot.h);
  }

  // コメント密度（塗り）
  ctx.beginPath();
  ctx.moveTo(PAD.l, PAD.t + plot.h);
  rate.forEach((v, px) => ctx.lineTo(PAD.l + px, yOf(v)));
  ctx.lineTo(PAD.l + rate.length, PAD.t + plot.h);
  ctx.closePath();
  const grad = ctx.createLinearGradient(0, PAD.t, 0, PAD.t + plot.h);
  grad.addColorStop(0, "rgba(88,166,255,0.55)");
  grad.addColorStop(1, "rgba(88,166,255,0.05)");
  ctx.fillStyle = grad;
  ctx.fill();

  ctx.beginPath();
  rate.forEach((v, px) => (px ? ctx.lineTo(PAD.l + px, yOf(v)) : ctx.moveTo(PAD.l, yOf(v))));
  ctx.strokeStyle = "#58a6ff";
  ctx.lineWidth = 1.2;
  ctx.stroke();

  // 平常ライン
  ctx.beginPath();
  base.forEach((v, px) => (px ? ctx.lineTo(PAD.l + px, yOf(v)) : ctx.moveTo(PAD.l, yOf(v))));
  ctx.strokeStyle = "#8b949e";
  ctx.setLineDash([4, 4]);
  ctx.stroke();
  ctx.setLineDash([]);

  // 重ね表示（カテゴリ・キーワード）
  overlays.forEach((o) => {
    ctx.beginPath();
    o.cols.forEach((v, px) => (px ? ctx.lineTo(PAD.l + px, yOf(v)) : ctx.moveTo(PAD.l, yOf(v))));
    ctx.strokeStyle = o.color;
    ctx.lineWidth = 1.6;
    ctx.stroke();
  });
  ctx.lineWidth = 1;

  // 音量（平常からの差・dB）。コメント数とは単位が違うので右側の軸に振る
  if (audioOn) {
    let aMax = 6;
    audioCols.forEach((v) => {
      if (v > aMax) aMax = v;
    });
    aMax = niceMax(aMax);
    const aMin = 0;
    const yAudio = (v) => PAD.t + plot.h - ((v - aMin) / (aMax - aMin)) * plot.h;

    ctx.strokeStyle = "rgba(219,97,162,0.35)";
    ctx.setLineDash([2, 3]);
    ctx.beginPath();
    ctx.moveTo(PAD.l, yAudio(0));
    ctx.lineTo(PAD.l + plot.w, yAudio(0));
    ctx.stroke();
    ctx.setLineDash([]);

    ctx.beginPath();
    audioCols.forEach((v, px) => (px ? ctx.lineTo(PAD.l + px, yAudio(v))
                                     : ctx.moveTo(PAD.l, yAudio(v))));
    ctx.strokeStyle = AUDIO_OVERLAY.color;
    ctx.lineWidth = 1.4;
    ctx.stroke();
    ctx.lineWidth = 1;

    ctx.fillStyle = AUDIO_OVERLAY.color;
    ctx.textAlign = "left";
    [aMax, 0].forEach((v) => ctx.fillText((v > 0 ? "+" : "") + v + "dB",
                                          PAD.l + plot.w + 5, yAudio(v) + 4));
    ctx.fillText("音量超過", PAD.l + plot.w + 5, PAD.t - 12);
    ctx.textAlign = "center";
  }

  // 候補のマーカー
  state.visible.forEach((m, index) => {
    if (m.peak_sec < state.view.start || m.peak_sec > state.view.end) return;
    const x = Math.round(timeToX(m.peak_sec)) + 0.5;
    const isSel = index === state.selected;
    ctx.strokeStyle = isSel ? "#ffffff" : "rgba(255,255,255,0.35)";
    ctx.beginPath();
    ctx.moveTo(x, PAD.t);
    ctx.lineTo(x, PAD.t + plot.h);
    ctx.stroke();
    if (m.rank <= 20 || isSel) {
      const label = String(m.rank);
      const w = ctx.measureText(label).width + 8;
      ctx.fillStyle = isSel ? "#ffffff" : "#58a6ff";
      ctx.fillRect(x - w / 2, PAD.t - 1, w, 14);
      ctx.fillStyle = "#06111f";
      ctx.textAlign = "center";
      ctx.font = "bold 10px system-ui, sans-serif";
      ctx.fillText(label, x, PAD.t + 9);
      ctx.font = "11px system-ui, sans-serif";
    }
  });

  // ドラッグ中の範囲
  if (state.drag && Math.abs(state.drag.x1 - state.drag.x0) > 2) {
    const x0 = Math.min(state.drag.x0, state.drag.x1);
    const x1 = Math.max(state.drag.x0, state.drag.x1);
    ctx.fillStyle = "rgba(255,255,255,0.12)";
    ctx.fillRect(x0, PAD.t, x1 - x0, plot.h);
  }

  // マウス位置の縦線
  if (state.hover != null) {
    const x = Math.round(timeToX(state.hover)) + 0.5;
    ctx.strokeStyle = "rgba(255,255,255,0.5)";
    ctx.beginPath();
    ctx.moveTo(x, PAD.t);
    ctx.lineTo(x, PAD.t + plot.h);
    ctx.stroke();
  }
}

/* --- グラフの操作 --- */

function eventX(ev) {
  const rect = canvas.getBoundingClientRect();
  return Math.min(PAD.l + plot.w, Math.max(PAD.l, ev.clientX - rect.left));
}

function audioFloorForTip() {
  const audio = state.result && state.result.audio;
  return audio && audio.available ? (audio.threshold_db || 0) : 0;
}

function nearestMoment(t) {
  let best = null;
  let bestDist = Infinity;
  state.visible.forEach((m, index) => {
    const dist = Math.abs(m.peak_sec - t);
    if (dist < bestDist) {
      bestDist = dist;
      best = index;
    }
  });
  const pxPerSec = plot.w / span();
  return bestDist * pxPerSec <= 8 ? best : null;
}

canvas.addEventListener("mousemove", (ev) => {
  if (!state.result) return;
  const x = eventX(ev);
  state.hover = xToTime(x);
  if (state.drag) state.drag.x1 = x;

  const series = state.result.series;
  const idx = Math.min(series.rate.length - 1, Math.max(0, Math.floor(state.hover / series.bin_sec)));
  const tip = $("tooltip");
  const near = nearestMoment(state.hover);
  const lines = [
    `<b>${fmtTime(state.hover)}</b>`,
    `${Math.round(series.rate[idx])} コメ/分（平常 ${Math.round(series.baseline[idx])}）`,
  ];
  if (state.overlays.has(AUDIO_KEY) && state.result.audio && state.result.audio.available) {
    const excess = state.result.audio.excess;
    const ai = Math.min(excess.length - 1, Math.max(0, Math.floor(state.hover / series.bin_sec)));
    const over = excess[ai] - audioFloorForTip();
    lines.push(`<span style="color:${AUDIO_OVERLAY.color}">音量 ${excess[ai] > 0 ? "+" : ""}${excess[ai]}dB`
      + (over >= 0 ? `（しきい値+${over.toFixed(1)}）` : "") + `</span>`);
  }
  if (near != null) {
    const m = state.visible[near];
    const tags = (m.tags || []).map((t) => `${t.label}${t.count}`).join(" ");
    lines.push(`<span style="color:#58a6ff">#${m.rank} 普段の${m.ratio}倍 ${escapeHtml(tags)}</span>`);
  }
  tip.innerHTML = lines.join("<br>");
  tip.style.left = x + "px";
  tip.style.top = PAD.t + plot.h - 6 + "px";
  tip.classList.remove("hidden");
  draw();
});

canvas.addEventListener("mouseleave", () => {
  state.hover = null;
  state.drag = null;
  $("tooltip").classList.add("hidden");
  draw();
});

canvas.addEventListener("mousedown", (ev) => {
  if (!state.result) return;
  const x = eventX(ev);
  state.drag = { x0: x, x1: x };
});

canvas.addEventListener("mouseup", (ev) => {
  if (!state.result || !state.drag) return;
  const { x0 } = state.drag;
  const x1 = eventX(ev);
  state.drag = null;
  if (Math.abs(x1 - x0) > 8) {
    setView(xToTime(Math.min(x0, x1)), xToTime(Math.max(x0, x1)));
    draw();
    return;
  }
  const t = xToTime(x1);
  const near = nearestMoment(t);
  if (near != null) {
    selectMoment(near, false);
    document.querySelector(`.moment[data-index="${near}"]`)
      .scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  const v = state.result.video;
  const base = v.platform === "youtube"
    ? `https://www.youtube.com/watch?v=${v.video_id}&t=${Math.round(t)}s`
    : `https://www.twitch.tv/videos/${v.video_id}?t=${Math.floor(t / 3600)}h${Math.floor((t % 3600) / 60)}m${Math.floor(t % 60)}s`;
  window.open(base, "_blank", "noopener");
});

canvas.addEventListener("dblclick", () => {
  if (!state.result) return;
  setView(0, state.result.stats.duration);
  draw();
});

canvas.addEventListener("wheel", (ev) => {
  if (!state.result) return;
  ev.preventDefault();
  const t = xToTime(eventX(ev));
  const factor = ev.deltaY > 0 ? 1.25 : 0.8;
  const newSpan = span() * factor;
  const ratio = (t - state.view.start) / span();
  setView(t - newSpan * ratio, t + newSpan * (1 - ratio));
  draw();
}, { passive: false });

$("resetZoom").onclick = () => {
  if (!state.result) return;
  setView(0, state.result.stats.duration);
  draw();
};

/* ------------------------------------------------------------ キーワード検索 */

async function searchKeyword() {
  const word = $("keywordInput").value.trim();
  if (!word || !state.videoKey) return;
  $("keywordBtn").disabled = true;
  $("keywordInfo").textContent = "検索中…";
  try {
    const found = await api("/api/keyword", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ video_key: state.videoKey, keyword: word, params: currentParams() }),
    });
    state.keyword = found;
    $("keywordInfo").textContent = found.hits
      ? `「${word}」を含むコメント ${num(found.hits)} 件`
      : `「${word}」を含むコメントはありませんでした`;
    renderTabs();
    if (found.hits) selectTab("keyword");
  } catch (err) {
    $("keywordInfo").textContent = "";
    showError(err.message);
  } finally {
    $("keywordBtn").disabled = false;
  }
}

/* ------------------------------------------------------------ 書き出し */

function copyText(text, button, label) {
  navigator.clipboard.writeText(text).then(() => {
    const original = button.textContent;
    button.textContent = label;
    setTimeout(() => (button.textContent = original), 1400);
  });
}

$("copyBtn").onclick = (ev) => {
  const lines = state.visible.map((m) => {
    const tags = (m.tags || []).map((t) => `${t.label}${t.count}`).join(" ");
    return `${fmtTime(m.clip_start)}  #${m.rank} 普段の${m.ratio}倍 ${tags}  ${m.url}`;
  });
  copyText(lines.join("\n"), ev.target, "コピーした");
};

$("chapterBtn").onclick = (ev) => {
  const lines = state.visible
    .slice()
    .sort((a, b) => a.clip_start - b.clip_start)
    .map((m) => {
      const tag = (m.tags || [])[0];
      return `${fmtTime(m.clip_start)} 見せ場${m.rank}${tag ? "（" + tag.label + "）" : ""}`;
    });
  copyText(["0:00 オープニング", ...lines].join("\n"), ev.target, "コピーした");
};

$("csvBtn").onclick = () => {
  const header = ["順位", "開始", "終了", "ピーク", "秒", "倍率", "スコア", "コメ/分",
                  "コメント数", "人数", "タグ", "代表コメント", "URL"];
  const rows = state.visible.map((m) => [
    m.rank, fmtTime(m.clip_start), fmtTime(m.clip_end), fmtTime(m.peak_sec),
    m.peak_sec, m.ratio, m.score, m.rate, m.messages, m.authors,
    (m.tags || []).map((t) => `${t.label}${t.count}`).join(" "),
    (m.top_comments || []).map((c) => `${c.text}×${c.count}`).join(" / "),
    m.url,
  ]);
  const csv = [header, ...rows]
    .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(","))
    .join("\r\n");
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  const title = (state.result.video.title || "highlights").replace(/[\\/:*?"<>|]/g, "_").slice(0, 60);
  link.download = `${title}_見せ場.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
};

/* ------------------------------------------------------------ 履歴 */

async function loadHistory() {
  try {
    const { entries } = await api("/api/history");
    const box = $("historyList");
    box.innerHTML = "";
    $("historyPanel").classList.toggle("hidden", !entries.length);
    entries.forEach((entry) => {
      const item = document.createElement("div");
      item.className = "history-item";
      item.innerHTML = `
        <span title="${escapeHtml(entry.title)}">${escapeHtml(entry.title.slice(0, 34))}</span>
        <span class="muted small">${num(entry.messages)}コメ</span>
        <button type="button">開く</button>
        <button type="button" class="del" title="削除">×</button>`;
      const [open, del] = item.querySelectorAll("button");
      open.onclick = () => {
        $("urlInput").value = entry.url;
        startAnalyze(entry.url, false);
      };
      del.onclick = async () => {
        await api(`/api/history/${entry.platform}/${entry.video_id}`, { method: "DELETE" });
        loadHistory();
      };
      box.appendChild(item);
    });
  } catch (err) {
    /* 履歴が出せなくても本体の邪魔はしない */
  }
}

/* ------------------------------------------------------------ 初期化 */

async function loadCapabilities() {
  try {
    const caps = await api("/api/capabilities");
    if (caps.version) {
      // サーバを再起動し忘れると画面だけ新しくなるので、動作中の版を出しておく
      $("codeVersion").textContent = caps.version;
    }
    if (!caps.ffmpeg) {
      $("audioInput").disabled = true;
      $("audioLabel").classList.add("disabled");
      $("audioLabel").title =
        "音声解析には ffmpeg が必要です。`pip install imageio-ffmpeg` で使えるようになります。";
    }
  } catch (err) {
    /* 判定できなければ触らずそのままにする */
  }
}

function syncOutputs() {
  const pairs = [["minZ", "minZOut"], ["mergeSec", "mergeSecOut"], ["leadSec", "leadSecOut"],
                 ["tailSec", "tailSecOut"], ["windowSec", "windowSecOut"], ["topN", "topNOut"],
                 ["audioMinZ", "audioMinZOut"], ["chatSupportZ", "chatSupportZOut"],
                 ["skipStart", "skipStartOut"]];
  pairs.forEach(([input, out]) => {
    const value = $(input).value;
    $(out).value = out === "skipStartOut"
      ? (Number(value) ? `${Math.round(Number(value) / 60)}分` : "なし")
      : value;
  });
}

["minZ", "mergeSec", "leadSec", "tailSec", "windowSec", "topN",
 "audioMinZ", "chatSupportZ", "skipStart"].forEach((id) => {
  $(id).addEventListener("input", syncOutputs);
});

$("urlForm").onsubmit = (ev) => {
  ev.preventDefault();
  const url = $("urlInput").value.trim();
  if (url) startAnalyze(url, $("refreshInput").checked);
};
$("applyBtn").onclick = reanalyze;
$("keywordBtn").onclick = searchKeyword;
$("keywordInput").addEventListener("keydown", (ev) => {
  if (ev.key === "Enter") searchKeyword();
});
window.addEventListener("resize", resizeCanvas);

// 既定値（サーバ側 Params と合わせる）
$("minZ").value = 3;
$("mergeSec").value = 90;
$("leadSec").value = 30;
$("tailSec").value = 15;
$("windowSec").value = 30;
$("topN").value = 30;
$("audioMinZ").value = 3;
$("chatSupportZ").value = 2;
$("skipStart").value = 0;
syncOutputs();
loadHistory();
loadCapabilities();
