// テロップ（グラフィックにアップグレードしたキャプション）の下段をそろえる計算。
// Premiere の API には触らない（node でテストできるように）。
//
// 考え方: 1行のテロップの位置を「下段の位置」とし、n 行のテロップは (n-1) 行分だけ上へ動かす。
// こうすると、行数が違っても一番下の行が同じ高さに来る（2行目が画面の下にはみ出さない）。

// ソーステキストの値から文字列を取り出す。Premiere は文字列のほか、JSON（前に余計なバイトが
// 付くこともある）で返すことがあるので、中の文字のキーを探す。
function textOf(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "object") return findText(value);
  if (typeof value !== "string") return null;
  const brace = value.indexOf("{");
  if (brace >= 0) {
    try {
      const found = findText(JSON.parse(value.slice(brace)));
      if (found !== null) return found;
    } catch (err) { /* JSON ではない文字列 */ }
  }
  return value;
}

const TEXT_KEYS = /^(mText|text|textEditValue|mTextEditValue|string)$/i;

function findText(obj, depth) {
  depth = depth || 0;
  if (!obj || typeof obj !== "object" || depth > 8) return null;
  for (const key of Object.keys(obj)) {
    if (TEXT_KEYS.test(key) && typeof obj[key] === "string") return obj[key];
  }
  for (const key of Object.keys(obj)) {
    const found = findText(obj[key], depth + 1);
    if (found !== null) return found;
  }
  return null;
}

// フォントサイズらしい数値（行の高さの見積もりに使う）。無ければ null。
function fontSizeOf(value) {
  let obj = value;
  if (typeof value === "string") {
    const brace = value.indexOf("{");
    if (brace < 0) return null;
    try { obj = JSON.parse(value.slice(brace)); } catch (err) { return null; }
  }
  return findNumber(obj, /^(m?fontSize|m?FontSize|fontsize)$/, 0);
}

function findNumber(obj, pattern, depth) {
  if (!obj || typeof obj !== "object" || depth > 8) return null;
  for (const key of Object.keys(obj)) {
    if (pattern.test(key) && typeof obj[key] === "number" && obj[key] > 0) return obj[key];
  }
  for (const key of Object.keys(obj)) {
    const found = findNumber(obj[key], pattern, depth + 1);
    if (found !== null) return found;
  }
  return null;
}

// 改行（Premiere のテキストは \r を使う。段落区切りなども数える）
const LINE_BREAKS = new RegExp("\\r\\n|\\r|\\n|" + String.fromCharCode(0x2028, 0x7c, 0x2029, 0x7c, 0x3));

// 行数（空の行は数えない。最低1）。
function lineCount(text) {
  if (typeof text !== "string") return 0;
  return Math.max(1, text.split(LINE_BREAKS).filter((s) => s.trim() !== "").length);
}

// 位置の値が 0〜1 の割合（Premiere の UXP はこちらで返すことがある）か、ピクセルか。
function isNormalized(points) {
  return points.every((p) => Math.abs(p.x) <= 2 && Math.abs(p.y) <= 2);
}

// 最も多い値（少しの誤差はまとめる）。
function mode(values, step) {
  const counts = new Map();
  for (const v of values) {
    const key = Math.round(v / step);
    counts.set(key, (counts.get(key) || []).concat(v));
  }
  let best = null;
  for (const list of counts.values()) if (!best || list.length > best.length) best = list;
  return best ? best[0] : null;
}

// items: [{ lines, x, y, dx, dy }]（今の位置）。lineHeightPx: 1行の高さ（ピクセル）。frameHeight: 画面の高さ。
// options: { keepX, baseYPx }。baseYPx は下段（1行のテロップ）の位置を、画面の上からのピクセルで決めるとき。
// dx, dy はクリップ側の「モーション」で動かしている量（位置と同じ単位。無ければ 0）。
// 見た目の位置（テキストの位置 + モーションのずれ）でそろえ、書き込むのはテキストの位置。
// 戻り値: { base: {x, y}, unit, moves: [{ index, lines, from: {x,y}, to: {x,y}, changed }] }
function planAlign(items, lineHeightPx, frameHeight, options) {
  options = options || {};
  const singles = items.filter((it) => it.lines === 1);
  const fixedY = options.baseYPx > 0;
  if (!singles.length && !fixedY) {
    throw new Error("選んだテロップに1行のものがありません。1行のテロップを1枚以上（正しい位置のもの）含めて選んでください。");
  }
  if (!(lineHeightPx > 0)) throw new Error("1行の高さ（px）を入れてください。");
  const normalized = isNormalized(items);
  const lineHeight = normalized ? lineHeightPx / frameHeight : lineHeightPx;
  const step = normalized ? 0.5 / frameHeight : 0.5;
  const dx = (it) => it.dx || 0;
  const dy = (it) => it.dy || 0;
  // baseYPx があれば、下段の位置はそれ（画面の上からのピクセル）。無ければ1行のテロップで一番多い位置。
  const reference = singles.length ? singles : items;
  const base = {
    x: mode(reference.map((s) => s.x + dx(s)), step),
    y: fixedY ? (normalized ? options.baseYPx / frameHeight : options.baseYPx)
              : mode(singles.map((s) => s.y + dy(s)), step),
  };
  const moves = items.map((it, index) => {
    const to = {
      x: options.keepX ? it.x : base.x - dx(it),
      y: base.y - (it.lines - 1) * lineHeight - dy(it),
    };
    const changed = Math.abs(to.x - it.x) > step || Math.abs(to.y - it.y) > step;
    return { index, lines: it.lines, from: { x: it.x, y: it.y }, to, changed };
  });
  return { base, unit: normalized ? "割合" : "px", lineHeight, moves };
}

module.exports = { textOf, fontSizeOf, lineCount, isNormalized, planAlign };
