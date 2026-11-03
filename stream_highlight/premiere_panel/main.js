// オートカットパネル: 選んだクリップの無音とフィラーを、stream_highlight のサーバーで判定し、
// タイムライン上でカットしてリップル削除する。
//
// カットの方式は2つあり、STRATEGIES から選ぶ（chooseStrategy）。
//
//   入れ直し（reinsert）: Premiere 26 でも使える。UXP にはクリップを分割する操作（レーザー）が
//     無いので、元のクリップをリップル削除し、残す部分ごとに素材のイン・アウトを設定して
//     同じ位置へインサートする。全てを1つのトランザクションで行うので Ctrl+Z 1回で戻る。
//     素材から作り直すため、エフェクト・キーフレーム・速度変更などは失われる。
//     失うものがあるクリップは、cutplan.cutLosses で見つけて実行を止める。
//
//   レーザー（razor）: Premiere 27.1 以降の createRazorTrackAction を使う。元のクリップを
//     切り分けてから要らない部分をリップル削除するので、エフェクトなどは残る。
//     切った後の断片を探してから消すため、トランザクションは2つ（Ctrl+Z 2回）。
//     手元の Premiere（26.0）で試せていないので、画面のチェックを入れたときだけ使う。

const ppro = require("premierepro");
const plan = require("./cutplan.js");
const align = require("./captionalign.js");

const SERVER = "http://127.0.0.1:8765";
const MARKER_NAME = "オートカット";
const POLL_MS = 700;

const $ = (id) => document.getElementById(id);
let profiles = [];
let busy = false;
let stage = "";          // 今している処理。失敗したときに、どこで失敗したかを出す

function at(label) {
  stage = label;
}

// ---------------------------------------------------------------- 画面

function setStatus(message, isError) {
  $("status").textContent = message;
  $("status").className = isError ? "error" : "";
}

function setProgress(frac) {
  $("barFill").style.width = Math.round(Math.max(0, Math.min(1, frac)) * 100) + "%";
}

function log(message) {
  $("log").value = message + "\n" + $("log").value;
}

function setBusy(value) {
  busy = value;
  for (const id of ["previewBtn", "cutBtn", "captionBtn", "alignBtn", "clearBtn", "diagnoseBtn"]) $(id).disabled = value;
}

function applyProfile(name) {
  const p = profiles.find((row) => row.name === name);
  if (!p) return;
  $("gapCut").value = p.gap_cut_sec;
  $("pad").value = p.pad_sec;
  $("fillers").checked = !!p.fillers;
}

function clock(ticks) {
  const total = Math.floor(ticks / plan.TICKS_PER_SECOND);
  const pad = (n) => String(n).padStart(2, "0");
  return Math.floor(total / 3600) + ":" + pad(Math.floor(total / 60) % 60) + ":" + pad(total % 60);
}

// ---------------------------------------------------------------- サーバー

async function request(path, body) {
  let response;
  try {
    response = await fetch(SERVER + path, body === undefined ? {} : {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  } catch (err) {
    throw new Error("サーバーに接続できません（" + SERVER + "）。python -m stream_highlight.server を" +
                    "起動してください。詳しい原因: " + (err && err.message ? err.message : err));
  }
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.detail || ("サーバーのエラー（" + response.status + "）"));
  return data;
}

async function loadProfiles() {
  try {
    const data = await request("/api/panel/profiles");
    profiles = data.profiles;
    const select = $("profile");
    select.innerHTML = "";
    for (const p of profiles) {
      const option = document.createElement("option");
      option.value = p.name;
      option.textContent = p.label === p.name ? p.name : p.name + "（" + p.label + "）";
      select.appendChild(option);
    }
    select.value = data.default;
    applyProfile(data.default);
    setStatus("準備できました。クリップを選んでください。");
  } catch (err) {
    setStatus(err.message, true);
    setTimeout(loadProfiles, 5000);      // サーバーを後から起動しても、そのまま使えるように
  }
}

// クリップ1つ分を判定する。戻り値はサーバーの結果（keep / removed は素材の秒）。
async function analyzeClip(clip, portion, report) {
  const started = await request("/api/panel/analyze", {
    media_path: clip.path,
    source_in: portion.sourceIn,
    source_out: portion.sourceOut,
    profile: $("profile").value,
    gap_cut_sec: parseFloat($("gapCut").value),
    pad_sec: parseFloat($("pad").value),
    fillers: $("fillers").checked,
  });
  return waitJob(started.job_id, report, "判定に失敗しました。");
}

// サーバーのジョブが終わるまで待ち、結果を返す。report(割合, 文言) で途中経過を伝える。
async function waitJob(jobId, report, failure) {
  for (;;) {
    await new Promise((resolve) => setTimeout(resolve, POLL_MS));
    const job = await request("/api/job/" + jobId);
    report(job.progress || 0, job.message || "");
    if (job.status === "done") return job.result;
    if (job.status === "error") throw new Error(job.error || failure);
  }
}

// ---------------------------------------------------------------- タイムライン

function ticksOf(tickTime) {
  return Number(tickTime.ticks);
}

function tickTime(ticks) {
  return ppro.TickTime.createWithTicks(String(Math.round(ticks)));
}

async function describe(item) {
  let sourceItem = null;     // 元の ProjectItem（インサートにはこちらを渡す。ClipProjectItem は受け付けない）
  let projectItem = null;    // ClipProjectItem（素材のパスとイン・アウトに使う）
  let path = "";
  try {
    sourceItem = await item.getProjectItem();
    projectItem = ppro.ClipProjectItem.cast(sourceItem);
    path = projectItem ? await projectItem.getMediaFilePath() : "";
  } catch (err) {
    // グラフィックやカラーマットなど、素材ファイルを持たないクリップ
    projectItem = null;
  }
  return {
    item,
    sourceItem,
    projectItem,
    path,
    startTicks: ticksOf(await item.getStartTime()),
    endTicks: ticksOf(await item.getEndTime()),
    inTicks: ticksOf(await item.getInPoint()),
    speed: await item.getSpeed(),
    trackIndex: await item.getTrackIndex(),
  };
}

async function audioClips(sequence) {
  const out = [];
  const count = await sequence.getAudioTrackCount();
  for (let i = 0; i < count; i++) {
    const track = await sequence.getAudioTrack(i);
    for (const item of await track.getTrackItems(ppro.Constants.TrackItemType.CLIP, false)) {
      try {
        out.push(await describe(item));
      } catch (err) {
        log("音声トラック A" + (i + 1) + " の読めないクリップを飛ばしました: " + (err.message || err));
      }
    }
  }
  return out;
}

function partnersOf(clip, audios) {
  return audios.filter((a) => a.path === clip.path && a.startTicks === clip.startTicks &&
                              a.endTicks === clip.endTicks);
}

// 対象のクリップ（映像）と、リンクした音声。範囲指定なら range（シーケンス上のティック）も使う。
async function gatherTargets(sequence) {
  let videos = [];
  let range = null;
  if ($("targetRange").checked) {
    at("シーケンスのイン・アウトの読み取り");
    const inTicks = ticksOf(await sequence.getInPoint());
    const outTicks = ticksOf(await sequence.getOutPoint());
    if (!(outTicks > inTicks)) throw new Error("シーケンスにインとアウトを設定してください。");
    range = { inTicks, outTicks };
    at("映像トラック V" + $("rangeTrack").value + " のクリップの読み取り");
    const track = await sequence.getVideoTrack(parseInt($("rangeTrack").value, 10) - 1);
    if (!track) throw new Error("その番号の映像トラックはありません。");
    videos = await track.getTrackItems(ppro.Constants.TrackItemType.CLIP, false);
  } else {
    at("選んだクリップの読み取り");
    const selection = await sequence.getSelection();
    videos = (await selection.getTrackItems()).filter((item) => item instanceof ppro.VideoClipTrackItem);
    if (!videos.length) throw new Error("タイムラインで映像のクリップを選んでください。");
  }
  at("音声トラックのクリップの読み取り");
  const audios = await audioClips(sequence);
  const targets = [];
  for (const item of videos) {
    at("映像クリップの情報の読み取り");
    const clip = await describe(item);
    const portion = plan.analysisPortion(clip, range);
    if (!portion) continue;
    if (!clip.path) { log("飛ばしました: 素材ファイルの無いクリップ（" + clock(clip.startTicks) + "）"); continue; }
    clip.audios = partnersOf(clip, audios);
    targets.push({ clip, portion });
  }
  targets.sort((a, b) => a.clip.startTicks - b.clip.startTicks);
  return targets;
}

// ---------------------------------------------------------------- 失うものの点検

// クリップに付いている効果を { matchName, displayName, params: [{name, keyframes, value}] } で集める。
async function componentInfo(project, item) {
  const chain = await item.getComponentChain();
  const components = [];
  project.lockedAccess(() => {
    const count = chain.getComponentCount();
    for (let i = 0; i < count; i++) components.push(chain.getComponentAtIndex(i));
  });
  const out = [];
  for (const component of components) {
    const entry = { matchName: await component.getMatchName(),
                    displayName: await component.getDisplayName(), params: [] };
    const count = component.getParamCount();
    for (let i = 0; i < count; i++) {
      const param = component.getParam(i);
      let keyframes = 0;
      let value = null;
      try {
        if (param.isTimeVarying()) keyframes = param.getKeyframeListAsTickTimes().length;
      } catch (err) { /* キーフレームを持てない値 */ }
      try {
        let v = await param.getStartValue();
        if (v && typeof v === "object" && "value" in v) v = v.value;
        if (v && typeof v === "object" && "value" in v) v = v.value;
        value = typeof v === "number" ? v : null;
      } catch (err) { /* 読めない値は比べない */ }
      entry.params.push({ name: param.displayName, keyframes, value });
    }
    out.push(entry);
  }
  return out;
}

async function transitionsAtEdges(sequence, clip) {
  try {
    const track = await sequence.getVideoTrack(clip.trackIndex);
    const items = await track.getTrackItems(ppro.Constants.TrackItemType.TRANSITION, false);
    let count = 0;
    for (const t of items) {
      const start = ticksOf(await t.getStartTime());
      const end = ticksOf(await t.getEndTime());
      if ((start <= clip.startTicks && clip.startTicks < end) ||
          (start < clip.endTicks && clip.endTicks <= end)) count++;
    }
    return count;
  } catch (err) {
    return 0;
  }
}

// cutplan.cutLosses に渡す情報を集める。
async function inspectClip(project, sequence, clip) {
  at("クリップの点検（エフェクト・速度など）");
  const audio = [];
  for (const a of clip.audios) audio.push(...await componentInfo(project, a.item));
  return {
    speed: clip.speed,
    reversed: !!(await clip.item.isSpeedReversed()),
    disabled: await clip.item.isDisabled(),
    isSequence: await clip.projectItem.isSequence(),
    name: await clip.item.getName(),
    projectName: clip.projectItem.name,
    audioTracks: clip.audios.map((a) => a.trackIndex),
    transitions: await transitionsAtEdges(sequence, clip),
    video: await componentInfo(project, clip.item),
    audio,
  };
}

// 全ての対象を点検する。戻り値は止める理由の一覧（クリップごとに1行）。
async function checkLosses(ctx) {
  const lines = [];
  for (const t of ctx.targets) {
    const found = plan.cutLosses(await inspectClip(ctx.project, ctx.sequence, t.clip));
    const where = "V" + (t.clip.trackIndex + 1) + " " + clock(t.clip.startTicks);
    for (const note of found.notes) log("注意（" + where + "）: " + note);
    if (found.blocks.length) lines.push(where + ": " + found.blocks.join("、"));
  }
  return lines;
}

// ---------------------------------------------------------------- カットの方式

const STRATEGIES = {
  reinsert: {
    label: "入れ直し方式（エフェクト等の付いたクリップには使えない。戻すときは複製したシーケンスを使う）",
    needsCleanClips: true,
    // インサートは部分ごとに別の取り消し単位になり、Ctrl+Z では戻しきれない。複製は必ず作る
    forceBackup: true,
    available: () => true,
    async apply(ctx, work) {
      // 素材のイン・アウトは一時的に書き換えるので、元の値を控えて最後に戻す
      const originals = new Map();
      for (const t of work) {
        if (!originals.has(t.clip.path)) {
          originals.set(t.clip.path, {
            projectItem: t.clip.projectItem,
            inPoint: await t.clip.projectItem.getInPoint(ppro.Constants.MediaType.VIDEO),
            outPoint: await t.clip.projectItem.getOutPoint(ppro.Constants.MediaType.VIDEO),
          });
        }
      }
      // Premiere はインサートの命令を作った時点の素材のイン・アウトを使う。同じトランザクションの中で
      // イン・アウトを変えても効かず、全ての部分が素材全体になった。そこで
      //   [元のクリップを消す + 最後の部分のイン・アウト] → [インサート + 次の部分のイン・アウト] → …
      // と、イン・アウトを1つ前のトランザクションで済ませてからインサートを作る。
      const total = work.reduce((n, t) => n + t.pieces.length, 0);
      let done = 0;
      const transact = (name, build) => {
        let ok = false;
        ctx.project.lockedAccess(() => {
          ok = ctx.project.executeTransaction((compound) => build(compound), name);
        });
        if (!ok) {
          throw new Error("途中で失敗しました（" + done + "/" + total + " 部分まで）。" +
                          "複製したシーケンス（「…のコピー」）が元の状態です。");
        }
      };
      const setInOut = (clip, piece) => clip.projectItem.createSetInOutPointsAction(
        tickTime(piece.inTicks), tickTime(piece.outTicks));

      // 後ろのクリップから直す。前を先に詰めると、後ろのクリップの位置がずれるため
      for (const t of work.slice().reverse()) {
        const { clip } = t;
        const pieces = t.pieces.slice().reverse();      // 同じ位置へ後ろの部分から入れると、前から並ぶ
        const audioTrack = Math.min(...clip.audios.map((a) => a.trackIndex));
        transact("オートカット（元のクリップを外す）", (compound) => {
          let selection = null;
          ppro.TrackItemSelection.createEmptySelection((s) => { selection = s; });
          selection.addItem(clip.item, true);
          for (const a of clip.audios) selection.addItem(a.item, true);
          compound.addAction(ctx.editor.createRemoveItemsAction(
            selection, true, ppro.Constants.MediaType.VIDEO, false));
          compound.addAction(setInOut(clip, pieces[0]));
        });
        for (let k = 0; k < pieces.length; k++) {
          transact("オートカット（部分を入れる）", (compound) => {
            compound.addAction(ctx.editor.createInsertProjectItemAction(
              clip.sourceItem, tickTime(clip.startTicks), clip.trackIndex, audioTrack, false));
            if (k + 1 < pieces.length) compound.addAction(setInOut(clip, pieces[k + 1]));
          });
          done++;
          if (done % 10 === 0 || done === total) {
            setProgress(0.9 + 0.1 * done / total);
            setStatus("タイムラインをカットしています… " + done + " / " + total);
            await new Promise((resolve) => setTimeout(resolve, 0));     // 画面を更新させる
          }
        }
      }
      // 素材のイン・アウトを戻す。控えた値をそのまま設定し直すと効かないことがあった（実際に残った）ので、
      // まず消去し、消去後の値（素材の全体）と違うときだけ、元のイン・アウトを設定し直す
      transact("オートカット（素材のイン・アウトを消去）", (compound) => {
        for (const o of originals.values()) compound.addAction(o.projectItem.createClearInOutPointsAction());
      });
      for (const o of originals.values()) {
        const fullIn = await o.projectItem.getInPoint(ppro.Constants.MediaType.VIDEO);
        const fullOut = await o.projectItem.getOutPoint(ppro.Constants.MediaType.VIDEO);
        if (ticksOf(fullIn) !== ticksOf(o.inPoint) || ticksOf(fullOut) !== ticksOf(o.outPoint)) {
          transact("オートカット（素材のイン・アウトを戻す）", (compound) => {
            compound.addAction(o.projectItem.createSetInOutPointsAction(
              tickTime(ticksOf(o.inPoint)), tickTime(ticksOf(o.outPoint))));
          });
        }
        const nowIn = ticksOf(await o.projectItem.getInPoint(ppro.Constants.MediaType.VIDEO));
        const nowOut = ticksOf(await o.projectItem.getOutPoint(ppro.Constants.MediaType.VIDEO));
        if (nowIn !== ticksOf(o.inPoint) || nowOut !== ticksOf(o.outPoint)) {
          log("注意: 素材「" + o.projectItem.name + "」のイン・アウトを元に戻せませんでした。" +
              "ソースモニターで開いて Ctrl+Shift+X で消去してください。");
        }
      }
      return "複製したシーケンス（「…のコピー」）を使う";
    },
  },

  razor: {
    label: "レーザー方式（Premiere 27.1 以降。エフェクトは残る。Ctrl+Z 2回）",
    needsCleanClips: false,
    available: (editor) => typeof editor.createRazorTrackAction === "function",
    async apply(ctx, work) {
      const VIDEO = ppro.Constants.MediaType.VIDEO;
      const cuts = work.map((t) => ({ t, ...plan.timelineCuts(t.clip, t.pieces) }));
      // 1. 切る（リンクした音声も一緒に切れる）
      let ok = false;
      ctx.project.lockedAccess(() => {
        const actions = [];
        for (const c of cuts) {
          for (const at of c.boundaries) {
            actions.push(ctx.editor.createRazorTrackAction(VIDEO, c.t.clip.trackIndex, tickTime(at), true, VIDEO));
          }
        }
        ok = ctx.project.executeTransaction((compound) => {
          actions.forEach((action) => compound.addAction(action));
        }, "オートカット（分割）");
      });
      if (!ok) throw new Error("クリップを分割できませんでした（タイムラインは変更していません）。");

      // 2. 切った後の断片のうち、消す範囲に収まるものを探してリップル削除する
      const half = ctx.frameTicks / 2;
      const audios = await audioClips(ctx.sequence);
      const doomed = [];
      for (const c of cuts) {
        const track = await ctx.sequence.getVideoTrack(c.t.clip.trackIndex);
        for (const item of await track.getTrackItems(ppro.Constants.TrackItemType.CLIP, false)) {
          const piece = await describe(item);
          if (piece.path !== c.t.clip.path) continue;
          if (c.removed.some(([a, b]) => piece.startTicks >= a - half && piece.endTicks <= b + half)) {
            doomed.push(piece, ...partnersOf(piece, audios));
          }
        }
      }
      ok = false;
      ctx.project.lockedAccess(() => {
        let selection = null;
        ppro.TrackItemSelection.createEmptySelection((s) => { selection = s; });
        for (const piece of doomed) selection.addItem(piece.item, true);
        ok = ctx.project.executeTransaction((compound) => {
          compound.addAction(ctx.editor.createRemoveItemsAction(selection, true, VIDEO, false));
        }, "オートカット（間を詰める）");
      });
      if (!ok) throw new Error("分割はしましたが、間を詰められませんでした。Ctrl+Z で分割を戻してください。");
      return "Ctrl+Z 2回";
    },
  },
};

function chooseStrategy(editor) {
  const razor = STRATEGIES.razor;
  return $("useRazor").checked && razor.available(editor) ? razor : STRATEGIES.reinsert;
}

// レーザー方式が使える版かを画面に出す（シーケンスが開いていなければ、実行時に判断する）
async function refreshStrategy() {
  try {
    const project = await ppro.Project.getActiveProject();
    const sequence = project && await project.getActiveSequence();
    const usable = sequence && STRATEGIES.razor.available(ppro.SequenceEditor.getEditor(sequence));
    $("useRazor").disabled = !usable;
    if (!usable) $("useRazor").checked = false;
    showStrategy(chooseStrategy(usable ? ppro.SequenceEditor.getEditor(sequence) : {}));
  } catch (err) {
    showStrategy(STRATEGIES.reinsert);
  }
}

function showStrategy(strategy) {
  $("strategy").textContent = strategy.label;
  // 複製が欠かせない方式では、チェックを外せないようにする
  $("backup").disabled = !!strategy.forceBackup;
  if (strategy.forceBackup) $("backup").checked = true;
}

// ---------------------------------------------------------------- 流れ

async function prepare() {
  at("プロジェクトとシーケンスの読み取り");
  const project = await ppro.Project.getActiveProject();
  const sequence = project && await project.getActiveSequence();
  if (!sequence) throw new Error("シーケンスを開いてください。");
  const editor = ppro.SequenceEditor.getEditor(sequence);
  at("シーケンスのフレームレートの読み取り");
  const frameTicks = Number(await sequence.getTimebase());
  const targets = await gatherTargets(sequence);
  if (!targets.length) throw new Error("処理できるクリップがありません。");
  return { project, sequence, editor, frameTicks, targets, strategy: chooseStrategy(editor) };
}

// 全ての対象を判定する（サーバー）。結果を targets に書き足す。
async function analyzeAll(ctx, share) {
  const { targets, frameTicks } = ctx;
  for (let k = 0; k < targets.length; k++) {
    const t = targets[k];
    const head = "クリップ " + (k + 1) + "/" + targets.length + ": ";
    at("サーバーでの判定（" + head + "）");
    t.result = await analyzeClip(t.clip, t.portion, (frac, message) => {
      setProgress(share * (k + frac) / targets.length);
      setStatus(head + message);
    });
    at("判定結果の時刻の変換");
    t.pieces = plan.keepPieces(t.clip, t.portion, t.result.keep, frameTicks);
    t.removed = plan.removedOnTimeline(t.portion, t.result.removed, frameTicks);
    log(head + "残す " + t.result.kept_sec.toFixed(1) + " 秒 / " + t.result.original_sec.toFixed(1) +
        " 秒（消す所 " + t.removed.length + " 件" + (t.result.cached ? "・キャッシュ" : "") + "）");
  }
}

async function preview() {
  const ctx = await prepare();
  if (ctx.strategy.needsCleanClips) {
    // プレビューは止めない。カット実行で止まることだけ先に伝える
    for (const line of await checkLosses(ctx)) log("このクリップはカット実行で止まります — " + line);
  }
  await analyzeAll(ctx, 0.95);
  at("プレビューのマーカーを打つ");
  const markers = await ppro.Markers.getMarkers(ctx.sequence);
  const removed = ctx.targets.flatMap((t) => t.removed);
  let ok = false;
  ctx.project.lockedAccess(() => {
    ok = ctx.project.executeTransaction((compound) => {
      for (const r of removed) {
        compound.addAction(markers.createAddMarkerAction(
          MARKER_NAME, ppro.Marker.MARKER_TYPE_COMMENT, tickTime(r.startTicks),
          tickTime(r.durationTicks), r.reasons));
      }
    }, "オートカットのプレビュー");
  });
  if (!ok) throw new Error("マーカーを打てませんでした。");
  setProgress(1);
  setStatus("消す所 " + removed.length + " 件にマーカーを打ちました（カットはしていません）。");
}

async function clearMarkers() {
  at("マーカーを消す");
  const project = await ppro.Project.getActiveProject();
  const sequence = project && await project.getActiveSequence();
  if (!sequence) throw new Error("シーケンスを開いてください。");
  const markers = await ppro.Markers.getMarkers(sequence);
  const ours = markers.getMarkers().filter((m) => m.getName() === MARKER_NAME);
  if (!ours.length) { setStatus("消すマーカーはありません。"); return; }
  project.lockedAccess(() => {
    project.executeTransaction((compound) => {
      for (const m of ours) compound.addAction(markers.createRemoveMarkerAction(m));
    }, "オートカットのマーカーを消す");
  });
  setStatus("マーカー " + ours.length + " 件を消しました。");
}

async function cut() {
  const ctx = await prepare();
  log("方式: " + ctx.strategy.label);
  if (ctx.strategy.needsCleanClips) {
    // 文字起こしで待たせる前に点検する。1本でも失うものがあれば何もしない
    setStatus("クリップを点検しています…");
    const blocked = await checkLosses(ctx);
    if (blocked.length) {
      for (const line of blocked) log("止めた理由 — " + line);
      throw new Error("カットを中止しました。入れ直し方式では失われるものが付いたクリップがあります" +
                      "（" + blocked.length + " 本。詳しくは下の記録）。エフェクト等を外すか、" +
                      "その部分は手でカットしてください。タイムラインは変更していません。");
    }
  }
  await analyzeAll(ctx, 0.9);
  const work = ctx.targets.filter((t) => !plan.unchanged(t.clip, t.pieces));
  if (!work.length) { setProgress(1); setStatus("消す所はありませんでした。"); return; }

  if ($("backup").checked || ctx.strategy.forceBackup) {
    at("シーケンスの複製（バックアップ）");
    setStatus("シーケンスを複製しています（バックアップ）…");
    let cloned = false;
    ctx.project.lockedAccess(() => {
      cloned = ctx.project.executeTransaction((compound) => {
        compound.addAction(ctx.sequence.createCloneAction());
      }, "オートカット前の複製");
    });
    if (!cloned) throw new Error("シーケンスを複製できませんでした。カットは行っていません。");
  }

  setStatus("タイムラインをカットしています…");
  at("タイムラインのカット（" + ctx.strategy.label.split("（")[0] + "）");
  const undo = await ctx.strategy.apply(ctx, work);

  const before = work.reduce((s, t) => s + (t.clip.endTicks - t.clip.startTicks), 0);
  const after = work.reduce((s, t) => s + plan.keptTicks(t.pieces), 0);
  setProgress(1);
  setStatus("カットしました: " + (before / plan.TICKS_PER_SECOND).toFixed(1) + " 秒 → " +
            (after / plan.TICKS_PER_SECOND).toFixed(1) + " 秒（クリップ " + work.length + " 本）。" +
            "元に戻すには " + undo + "。");
  if (ctx.strategy.forceBackup) log("この方式は Ctrl+Z では戻しきれません。戻すときは複製したシーケンスを開いてください。");
}

// 対象のクリップの今の位置に合わせてテロップ（SRT）を作り、プロジェクトに読み込む。
// カットした後のかけらでも、カット前のクリップでもよい。文字起こしは判定のときの保存分を使う。
async function makeCaptions() {
  const ctx = await prepare();
  if (!ctx.project.path) throw new Error("プロジェクトを一度保存してから使ってください（SRT の保存先に使います）。");
  at("テロップの作成（サーバー）");
  const clips = ctx.targets.map((t) => ({
    media_path: t.clip.path,
    source_in: t.portion.sourceIn,
    source_out: t.portion.sourceOut,
    seq_start: t.portion.startTicks / plan.TICKS_PER_SECOND,
  }));
  let name = "テロップ";
  try { name = ctx.sequence.name || name; } catch (err) { /* 名前を読めない版 */ }
  const started = await request("/api/panel/captions", {
    clips,
    fps: plan.TICKS_PER_SECOND / ctx.frameTicks,
    profile: $("profile").value,
    project_path: ctx.project.path,
    name,
  });
  const result = await waitJob(started.job_id, (frac, message) => {
    setProgress(0.9 * frac);
    setStatus(message);
  }, "テロップを作れませんでした。");
  log("テロップ " + result.cues + " 枚: " + result.srt_path);

  at("SRT をプロジェクトに読み込む");
  let imported = false;
  try {
    imported = await ctx.project.importFiles([result.srt_path], true, await ctx.project.getInsertionBin(), false);
  } catch (err) {
    log("自動の読み込みに失敗しました: " + (err.message || err));
  }
  setProgress(1);
  setStatus("テロップを " + result.cues + " 枚作りました。" + (imported
    ? "プロジェクトパネルに読み込んだ SRT を、シーケンスの先頭（0秒）へドラッグしてください。"
    : "「ファイル」→「読み込み」で次の SRT を読み込み、シーケンスの先頭へドラッグしてください: " + result.srt_path));
}

// ---------------------------------------------------------------- テロップの下段をそろえる

// グラフィックにアップグレードしたテロップを、行数を見て並べ直す。
// 1行のテロップの位置を下段とし、n 行のものは (n-1) 行分だけ上へ。テキストやスタイルは触らない。
const SOURCE_TEXT = /ソーステキスト|Source Text/i;
const POSITION = /^(位置|Position)$/i;

function unwrapValue(v) {
  for (let i = 0; i < 3 && v && typeof v === "object" && "value" in v; i++) v = v.value;
  return v;
}

function pointOf(v) {
  v = unwrapValue(v);
  if (v && typeof v.x === "number" && typeof v.y === "number") return { x: v.x, y: v.y };
  if (Array.isArray(v) && v.length >= 2) return { x: Number(v[0]), y: Number(v[1]) };
  return null;
}

// 記録に出すための、値の先頭300文字
function sampleOf(v) {
  let text;
  try { text = typeof v === "string" ? v : JSON.stringify(v); } catch (err) { text = String(v); }
  text = String(text);
  return text.length > 300 ? text.slice(0, 300) + "…（全 " + text.length + " 文字）" : text;
}

function describeValue(v) {
  const kind = v && v.constructor && v.constructor.name ? v.constructor.name : typeof v;
  return kind + " " + sampleOf(v);
}

// 位置の値を作る（版によって作り方が違うので、順に試す）
function makePoint(original, x, y) {
  try { if (typeof ppro.PointF === "function") return new ppro.PointF(x, y); } catch (err) { /* 次へ */ }
  try { if (ppro.PointF && typeof ppro.PointF.create === "function") return ppro.PointF.create(x, y); } catch (err) { /* 次へ */ }
  const raw = unwrapValue(original);
  if (raw && typeof raw === "object" && "x" in raw) {
    try { raw.x = x; raw.y = y; return raw; } catch (err) { /* 書き換えられない */ }
  }
  return { x, y };
}

// クリップの中から「ソーステキスト」と同じ部品にある「位置」を探す。
async function textLayerOf(project, item) {
  const chain = await item.getComponentChain();
  const components = [];
  project.lockedAccess(() => {
    const count = chain.getComponentCount();
    for (let i = 0; i < count; i++) components.push(chain.getComponentAtIndex(i));
  });
  const layers = [];
  const seen = [];
  let motion = null;
  for (const component of components) {
    const name = await component.getDisplayName();
    let text = null, position = null;
    const isMotion = /^(モーション|Motion)$/.test(name);
    const params = [];
    for (let i = 0; i < component.getParamCount(); i++) {
      const param = component.getParam(i);
      params.push(param.displayName);
      if (isMotion && !motion && POSITION.test(param.displayName)) motion = param;
      else if (!text && SOURCE_TEXT.test(param.displayName)) text = param;
      // ソーステキストより後ろの最初の「位置」が、そのテキストの位置
      else if (text && !position && POSITION.test(param.displayName)) position = param;
    }
    seen.push(name + ": " + params.join(" / "));
    if (text && position) layers.push({ component: name, text, position });
  }
  return { layers, seen, motion };
}

async function readCaption(project, item) {
  const startTicks = ticksOf(await item.getStartTime());
  const where = clock(startTicks);
  const { layers, seen, motion } = await textLayerOf(project, item);
  if (layers.length !== 1) {
    return { where, skip: layers.length ? "テキストが " + layers.length + " 個あるグラフィック" : "テキストの無いクリップ", seen };
  }
  const { text, position } = layers[0];
  let rawText = null;
  try { rawText = unwrapValue(await text.getStartValue()); } catch (err) { /* 読めない */ }
  let str = align.textOf(rawText);
  if (str === null) {
    try { str = await item.getName(); } catch (err) { /* 名前も読めない */ }
  }
  let keyframes = 0;
  try { if (position.isTimeVarying()) keyframes = position.getKeyframeListAsTickTimes().length; } catch (err) { /* */ }
  const startValue = await position.getStartValue();
  const point = pointOf(startValue);
  if (!point) return { where, skip: "位置の値を読めない（" + JSON.stringify(unwrapValue(startValue)) + "）", seen };
  if (keyframes > 1) return { where, skip: "位置にキーフレームがある" };
  return {
    where, item, param: position, startValue, x: point.x, y: point.y,
    startTicks, endTicks: ticksOf(await item.getEndTime()), fromName: rawText === null,
    motion: motion ? pointOf(await motion.getStartValue()) : null,
    text: str, lines: align.lineCount(str), fontSize: align.fontSizeOf(rawText),
    rawType: typeof rawText, rawText,
  };
}

// Premiere 26 はソーステキストの値を渡さない（null）。そのときは保存済みのプロジェクトファイルから
// サーバーに行数を読んでもらい、シーケンス上の開始・終了の位置で突き合わせる。
async function linesFromProject(project, captions) {
  const missing = captions.filter((c) => c.fromName);
  if (!missing.length) return;
  at("プロジェクトファイルから行数を読む");
  setStatus("プロジェクトファイルからテロップの行数を読んでいます…");
  const data = await request("/api/panel/caption_lines", { project_path: project.path || "" });
  const byRange = new Map();
  for (const c of data.captions) byRange.set(c.start + "-" + c.end, c);
  const unmatched = [];
  for (const c of missing) {
    const found = byRange.get(Number(c.startTicks) + "-" + Number(c.endTicks));
    if (found) { c.text = found.text; c.lines = found.lines; }
    else unmatched.push(c);
  }
  log("プロジェクトファイルから行数を読みました（" + (missing.length - unmatched.length) + " / " + missing.length + " 枚）。");
  if (unmatched.length * 2 > missing.length) {
    throw new Error("プロジェクトファイルに無いテロップが " + unmatched.length + " 枚あります。" +
                    "保存（Ctrl+S）してからもう一度押してください。");
  }
  // 少しだけ見つからないものは、行数が分からないので動かさない
  for (const c of unmatched) {
    log("飛ばしました（" + c.where + "）: プロジェクトファイルで行数を読めない（保存前に足したもの？）");
    captions.splice(captions.indexOf(c), 1);
  }
}

// クリップの「モーション」の位置（既定は画面の中央）が動いていると、テキストの位置が同じでも
// 見た目の高さが変わる。そのずれをテキストの位置と同じ単位で dx, dy に入れる。
function addMotionOffsets(captions, frameWidth, frameHeight) {
  const normalized = align.isNormalized(captions);
  let shifted = 0;
  for (const c of captions) {
    if (!c.motion) { c.dx = 0; c.dy = 0; continue; }
    const m = c.motion;
    // モーションの値も割合（0〜1）か px か。テキストの単位に合わせる
    const motionNormalized = Math.abs(m.x) <= 2 && Math.abs(m.y) <= 2;
    let dx = motionNormalized ? m.x - 0.5 : (m.x - frameWidth / 2) / frameWidth;
    let dy = motionNormalized ? m.y - 0.5 : (m.y - frameHeight / 2) / frameHeight;
    if (!normalized) { dx *= frameWidth; dy *= frameHeight; }
    c.dx = dx; c.dy = dy;
    if (Math.abs(dx) > 1e-4 || Math.abs(dy) > 1e-4) shifted++;
  }
  if (shifted) log("クリップのモーションで位置をずらしているテロップ: " + shifted + " 枚（見た目の位置でそろえます）。");
}

async function frameWidthOf(sequence) {
  try {
    const size = await sequence.getFrameSize();
    const w = size && (size.width || (size.right !== undefined ? size.right - size.left : 0));
    if (w > 0) return w;
  } catch (err) { /* 読めない版 */ }
  return 1920;
}

async function frameHeightOf(sequence) {
  try {
    const size = await sequence.getFrameSize();
    const h = size && (size.height || (size.bottom !== undefined ? size.bottom - size.top : 0));
    if (h > 0) return h;
  } catch (err) { /* 読めない版 */ }
  return 1080;
}

async function alignCaptions() {
  at("プロジェクトとシーケンスの読み取り");
  const project = await ppro.Project.getActiveProject();
  const sequence = project && await project.getActiveSequence();
  if (!sequence) throw new Error("シーケンスを開いてください。");
  at("選んだテロップの読み取り");
  const selection = await sequence.getSelection();
  const items = (await selection.getTrackItems()).filter((item) => item instanceof ppro.VideoClipTrackItem);
  if (!items.length) throw new Error("タイムラインでテロップ（グラフィック）を選んでください。");

  const captions = [];
  let shownSeen = false;
  for (let k = 0; k < items.length; k++) {
    const c = await readCaption(project, items[k]);
    if (c.skip) {
      log("飛ばしました（" + c.where + "）: " + c.skip);
      if (c.seen && !shownSeen) { log("  このクリップの中身: " + c.seen.join(" | ")); shownSeen = true; }
      continue;
    }
    captions.push(c);
    if (k % 20 === 0) { setProgress(0.6 * k / items.length); setStatus("読み取り中… " + k + " / " + items.length); }
  }
  if (!captions.length) throw new Error("そろえられるテロップがありませんでした（理由は下の記録）。");
  await linesFromProject(project, captions);
  const unknown = captions.filter((c) => c.text === null);
  if (unknown.length) {
    throw new Error("テロップの文字を読めず、行数が分かりませんでした（" + unknown.length + " 枚）。" +
                    "記録をコピーして送ってください。");
  }

  at("1行の高さを決める");
  const frameHeight = await frameHeightOf(sequence);
  let lineHeight = parseFloat($("lineHeight").value);
  if (!(lineHeight > 0)) {
    const size = captions.map((c) => c.fontSize).find((s) => s > 0);
    if (!size) {
      // フォントサイズのありかを調べられるよう、ソーステキストの生の値を記録に出す
      log("ソーステキストの値（" + captions[0].where + "、型: " + captions[0].rawType + "）: " + sampleOf(captions[0].rawText));
      throw new Error("1行の高さ（px）を入れてください。この版の Premiere はテロップの文字の情報を渡さないため、" +
                      "フォントサイズを自動で読めません。プロパティパネルでフォントサイズを見て、その1.2倍くらいを入れてください。");
    }
    lineHeight = Math.round(size * 1.2);
    $("lineHeight").value = lineHeight;
    log("1行の高さをフォントサイズ " + size + " から " + lineHeight + " px と見積もりました。");
  }
  addMotionOffsets(captions, await frameWidthOf(sequence), frameHeight);
  const result = align.planAlign(captions, lineHeight, frameHeight,
                                {keepX: $("keepX").checked, baseYPx: parseFloat($("baseY").value)});
  const moving = result.moves.filter((m) => m.changed);
  const byLines = {};
  for (const c of captions) byLines[c.lines] = (byLines[c.lines] || 0) + 1;
  const fmt = (p) => result.unit === "px" ? "(" + p.x.toFixed(0) + ", " + p.y.toFixed(0) + ")"
                                           : "(" + p.x.toFixed(4) + ", " + p.y.toFixed(4) + ")";
  log("テロップ " + captions.length + " 枚（" + Object.keys(byLines).map((n) => n + "行: " + byLines[n] + "枚").join("、") +
      "）。下段の位置 " + fmt(result.base) + "（単位: " + result.unit + "、画面の高さ " + frameHeight + "）");
  for (const m of moving.slice(0, 15)) {
    const c = captions[m.index];
    log("  " + c.where + " " + m.lines + "行「" + c.text.replace(/\s+/g, " ").slice(0, 20) + "」 " + fmt(m.from) + " → " + fmt(m.to));
  }
  if (moving.length > 15) log("  …ほか " + (moving.length - 15) + " 枚");

  if ($("alignDryRun").checked) {
    setProgress(1);
    setStatus("確認だけしました: 動かすテロップ " + moving.length + " 枚（変更はしていません）。" +
              "記録を見て問題なければ「確認だけ」を外してもう一度押してください。");
    return;
  }
  if (!moving.length) { setProgress(1); setStatus("すでにそろっています。"); return; }

  at("テロップの位置を書き換える");
  let ok = false;
  project.lockedAccess(() => {
    ok = project.executeTransaction((compound) => {
      for (const m of moving) {
        const c = captions[m.index];
        const point = makePoint(c.startValue, m.to.x, m.to.y);
        if (m === moving[0]) log("位置に書き込む値: " + describeValue(point));
        const keyframe = c.param.createKeyframe(point);
        compound.addAction(c.param.createSetValueAction(keyframe, true));
      }
    }, "テロップの下段をそろえる");
  });
  if (!ok) throw new Error("位置を書き換えられませんでした（タイムラインは変更していません）。");
  setProgress(1);
  setStatus("テロップ " + moving.length + " 枚の位置をそろえました。戻すときは Ctrl+Z 1回。");
}

// 選んだクリップに付いているものを記録に書き出す（判定の食い違いを調べるとき用）
async function diagnose() {
  const ctx = await prepare();
  for (const t of ctx.targets) {
    const info = await inspectClip(ctx.project, ctx.sequence, t.clip);
    const found = plan.cutLosses(info);
    log("V" + (t.clip.trackIndex + 1) + " " + clock(t.clip.startTicks) + "\n" +
        JSON.stringify(info, null, 1) + "\n止める理由: " + (found.blocks.join("、") || "なし"));
  }
  setStatus("選んだクリップの情報を下の記録に書き出しました。");
}

function wire(id, action) {
  $(id).addEventListener("click", async () => {
    if (busy) return;
    setBusy(true);
    setProgress(0);
    stage = "";
    try {
      await action();
    } catch (err) {
      const where = stage ? "「" + stage + "」で失敗: " : "";
      setStatus(where + (err.message || String(err)), true);
      log("エラー（" + (stage || "不明") + "）: " + (err.stack || err));
    } finally {
      setBusy(false);
      refreshStrategy();
    }
  });
}

$("profile").addEventListener("change", () => applyProfile($("profile").value));
$("useRazor").addEventListener("change", refreshStrategy);
wire("previewBtn", preview);
wire("cutBtn", cut);
wire("clearBtn", clearMarkers);
wire("captionBtn", makeCaptions);
wire("alignBtn", alignCaptions);
wire("diagnoseBtn", diagnose);
loadProfiles();
refreshStrategy();
