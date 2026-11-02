// オートカットパネルの時刻の計算。Premiere の API には触らない（node でテストできるように）。
//
// Premiere の時刻は「ティック」（1秒 = 254016000000）。クリップは
//   startTicks / endTicks: シーケンス上の位置
//   inTicks:               素材の中の開始位置（ソースのイン）
// を持つ。サーバーとは秒でやりとりし、ここでフレーム単位のティックに直す。

const TICKS_PER_SECOND = 254016000000;

function frames(seconds, frameTicks) {
  return Math.round((seconds * TICKS_PER_SECOND) / frameTicks);
}

// クリップのうち、判定にかける部分（シーケンス上）。range（イン〜アウト）が無ければクリップ全体。
function analysisPortion(clip, range) {
  const startTicks = range ? Math.max(clip.startTicks, range.inTicks) : clip.startTicks;
  const endTicks = range ? Math.min(clip.endTicks, range.outTicks) : clip.endTicks;
  if (endTicks <= startTicks) return null;
  const srcStartTicks = clip.inTicks + (startTicks - clip.startTicks);
  return {
    startTicks, endTicks, srcStartTicks,
    sourceIn: srcStartTicks / TICKS_PER_SECOND,
    sourceOut: (srcStartTicks + (endTicks - startTicks)) / TICKS_PER_SECOND,
  };
}

// 判定にかけた部分の中の素材の秒を、素材のティックにする（フレームに揃える）。
function sourceTicks(portion, seconds, frameTicks) {
  return portion.srcStartTicks + frames(seconds - portion.sourceIn, frameTicks) * frameTicks;
}

// クリップ全体で残す部分 [{inTicks, outTicks}]（素材のティック、並び順）。
// 判定にかけなかった頭と尻（イン〜アウトの外）はそのまま残す。隣り合う部分はつなぐ。
function keepPieces(clip, portion, keepSeconds, frameTicks) {
  const clipOutTicks = clip.inTicks + (clip.endTicks - clip.startTicks);
  const portionOutTicks = portion.srcStartTicks + (portion.endTicks - portion.startTicks);
  const raw = [];
  if (portion.srcStartTicks > clip.inTicks) raw.push([clip.inTicks, portion.srcStartTicks]);
  for (const [lo, hi] of keepSeconds) {
    const a = Math.max(portion.srcStartTicks, sourceTicks(portion, lo, frameTicks));
    const b = Math.min(portionOutTicks, sourceTicks(portion, hi, frameTicks));
    if (b > a) raw.push([a, b]);
  }
  if (clipOutTicks > portionOutTicks) raw.push([portionOutTicks, clipOutTicks]);
  const pieces = [];
  for (const [a, b] of raw) {
    const last = pieces[pieces.length - 1];
    if (last && a <= last.outTicks) last.outTicks = Math.max(last.outTicks, b);
    else pieces.push({ inTicks: a, outTicks: b });
  }
  return pieces;
}

// 何も消さないか（カットしても同じになるか）。
function unchanged(clip, pieces) {
  return pieces.length === 1 && pieces[0].inTicks === clip.inTicks &&
    pieces[0].outTicks === clip.inTicks + (clip.endTicks - clip.startTicks);
}

// 消す所を、シーケンス上の [{startTicks, durationTicks, reasons}] にする（プレビューのマーカー用）。
function removedOnTimeline(portion, removed, frameTicks) {
  const out = [];
  for (const item of removed) {
    const a = sourceTicks(portion, item.start, frameTicks) - portion.srcStartTicks;
    const b = sourceTicks(portion, item.end, frameTicks) - portion.srcStartTicks;
    if (b > a) {
      out.push({ startTicks: portion.startTicks + a, durationTicks: b - a,
                 reasons: (item.reasons || []).join("、") });
    }
  }
  return out;
}

// レーザー方式（Premiere 27.1 以降）用。残す部分から、シーケンス上の
//   boundaries: 切る位置（ティック）
//   removed:    消す範囲 [[開始, 終了]]（ティック）
// を求める。クリップの両端は切らない（もともと切れている）。
function timelineCuts(clip, pieces) {
  const toTimeline = (src) => clip.startTicks + (src - clip.inTicks);
  const kept = pieces.map((p) => [toTimeline(p.inTicks), toTimeline(p.outTicks)]);
  const removed = [];
  let cursor = clip.startTicks;
  for (const [a, b] of kept) {
    if (a > cursor) removed.push([cursor, a]);
    cursor = Math.max(cursor, b);
  }
  if (clip.endTicks > cursor) removed.push([cursor, clip.endTicks]);
  const boundaries = [];
  for (const [a, b] of removed) {
    for (const t of [a, b]) {
      if (t > clip.startTicks && t < clip.endTicks && !boundaries.includes(t)) boundaries.push(t);
    }
  }
  boundaries.sort((x, y) => x - y);
  return { boundaries, removed };
}

// ---------------------------------------------------------------- 入れ直し方式で失うもの
//
// 入れ直し方式は素材からクリップを作り直すので、元のクリップに付けたものは引き継げない。
// clip は main.js の inspectClip が集めた情報:
//   { speed, reversed, disabled, isSequence, name, projectName, audioTracks: [番号],
//     transitions: 端にあるトランジションの数,
//     video: [{matchName, displayName, params: [{name, keyframes, value}]}],
//     audio: [同じ形（リンクした音声クリップのもの）] }
// 戻り値 { blocks: 止める理由, notes: 止めないが伝えること }。

// 全てのクリップが最初から持っている効果（固定エフェクト）。これ以外が付いていればエフェクトとみなす
const VIDEO_INTRINSIC = /motion|opacity|time ?remap/i;
const AUDIO_INTRINSIC = /volume|panner|balance/i;
// 固定エフェクトの値の既定。数で読めたものだけ比べる（読み方が分からない値は見ない）
const STATIC_DEFAULTS = [
  { component: /motion/i, param: 1, value: 100, label: "スケール" },
  { component: /motion/i, param: 4, value: 0, label: "回転" },
  { component: /opacity/i, param: 0, value: 100, label: "不透明度" },
];

function componentLosses(components, intrinsic, kind) {
  const found = [];
  for (const c of components || []) {
    const name = c.displayName || c.matchName;
    if (!intrinsic.test(c.matchName || "")) {
      found.push(kind + "エフェクト「" + name + "」");
      continue;
    }
    (c.params || []).forEach((param, index) => {
      if (param.keyframes > 0) {
        found.push(kind + "キーフレーム（" + name + " の " + (param.name || index) + "）");
      }
      for (const d of STATIC_DEFAULTS) {
        if (d.component.test(c.matchName) && d.param === index && typeof param.value === "number" &&
            Math.abs(param.value - d.value) > 1e-3) {
          found.push(d.label + "の変更（" + Math.round(param.value * 100) / 100 + "）");
        }
      }
    });
  }
  return found;
}

function cutLosses(clip) {
  const blocks = [];
  const notes = [];
  if (Math.abs((clip.speed || 1) - 1) > 1e-6) blocks.push("速度の変更（" + Math.round(clip.speed * 1000) / 10 + "%）");
  if (clip.reversed) blocks.push("逆再生");
  if (clip.disabled) blocks.push("無効にしたクリップ");
  if (clip.isSequence) blocks.push("ネストしたシーケンス");
  const tracks = (clip.audioTracks || []).slice().sort((a, b) => a - b);
  if (!tracks.length) {
    blocks.push("音声がリンクしていない（入れ直すと音声が付いてしまう）");
  } else if (tracks[tracks.length - 1] - tracks[0] !== tracks.length - 1) {
    blocks.push("リンクした音声が離れたトラックにある（入れ直すと並びが変わる）");
  }
  if (clip.transitions > 0) blocks.push("端のトランジション");
  blocks.push(...componentLosses(clip.video, VIDEO_INTRINSIC, ""));
  blocks.push(...componentLosses(clip.audio, AUDIO_INTRINSIC, "音声の"));
  if (clip.name && clip.projectName && clip.name !== clip.projectName) {
    notes.push("クリップ名「" + clip.name + "」は素材の名前に戻ります");
  }
  return { blocks: Array.from(new Set(blocks)), notes };
}

// 残す部分の合計（ティック）。
function keptTicks(pieces) {
  return pieces.reduce((sum, p) => sum + (p.outTicks - p.inTicks), 0);
}

module.exports = {
  TICKS_PER_SECOND, frames, analysisPortion, keepPieces, unchanged, removedOnTimeline,
  timelineCuts, cutLosses, keptTicks,
};
