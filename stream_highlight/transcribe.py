# -*- coding: utf-8 -*-
"""書き出す区間の音声を文字起こしし、単語ごとの時刻を返す。

自動カットとテロップの土台。配信の音声はマイクとゲーム音が1本に混ざって
いるため、音量では「喋っていない間」を判定できない（README の音声解析の
経緯を参照）。文字起こしが返す単語の時刻を、発話区間として使う。

3時間の配信全体は処理しない。書き出す区間だけを ffmpeg で切り出す。

エンジンは差し替えられるようにしてある。どれも同じ形
（Transcript = テキスト＋単語ごとの開始・終了秒）を返す。

  qwen     Qwen3-ASR-1.7B ＋ Qwen3-ForcedAligner-0.6B（既定）
  whisper  faster-whisper large-v3-turbo（予備）

重いライブラリ（torch など）はここでは読み込まない。使う時点で読み込み、
無ければ何が足りないかを日本語で伝えて、この機能だけを止める。
"""

import gc
import hashlib
import importlib.util
import os
import shutil
import subprocess
import threading
import time
from dataclasses import asdict, dataclass, field

from . import cache, paths

SAMPLE_RATE = 16000
TRANSCRIPT_VERSION = 2      # エンジンの出力の形を変えたら上げる（古いキャッシュを使わない）

# ForcedAligner は1回の処理が最大5分。ただし qwen-asr の公式実装は、時刻を付けるとき
# 180秒ごとに区切っている（MAX_FORCE_ALIGN_INPUT_SECONDS）。こちらも180秒で区切り、
# ライブラリの中でさらに分け直されないようにする（分割点を自分で選べるように）。
QWEN_MAX_CHUNK_SEC = 180.0
# 1回の処理で、音声1秒あたりに追加で使う VRAM（RTX 4060 Ti 8GB・float16 で実測）。
# モデル本体（約5.9GB）の上にこれが乗る。8GB で180秒を一度に処理すると VRAM を超えて
# PC 本体のメモリにはみ出し、処理が2倍近く遅くなった（150秒で8.2GB、180秒で8.9GB）。
QWEN_VRAM_GB_PER_SEC = 0.0165
QWEN_MIN_CHUNK_SEC = 45.0
VRAM_MARGIN_GB = 0.3


def chunk_sec_for_vram(free_gb):
    """読み込み後の VRAM の空きから、1回で処理する長さを決める。

    Premiere や OBS が VRAM を使っていれば、そのぶん短くなる。60秒と120秒で
    1秒あたりの処理時間はほぼ同じだったので、短くしても遅くはならない。
    """
    usable = max(0.0, free_gb - VRAM_MARGIN_GB)
    return max(QWEN_MIN_CHUNK_SEC, min(QWEN_MAX_CHUNK_SEC, usable / QWEN_VRAM_GB_PER_SEC))
# 分割点を探す幅。上限の手前この秒数の中で、いちばん静かな所で切る。
SPLIT_SEARCH_SEC = 60.0
SPLIT_FRAME_SEC = 0.05
SPLIT_SMOOTH_SEC = 0.4      # 静かさをならす幅（語の中の一瞬の途切れを選ばないため）

QWEN_ASR_MODEL = "Qwen/Qwen3-ASR-1.7B"
QWEN_ALIGNER_MODEL = "Qwen/Qwen3-ForcedAligner-0.6B"
WHISPER_MODEL = "large-v3-turbo"

# 画面の選択肢。並び順がそのまま表示順になる。
ENGINE_LABELS = {
    "qwen": "Qwen3-ASR（既定・高精度）",
    "whisper": "faster-whisper large-v3-turbo（予備・導入が簡単）",
}
DEFAULT_ENGINE = "qwen"


class TranscribeError(Exception):
    """利用者に見せる日本語メッセージを持つ失敗。"""


@dataclass
class Word:
    text: str
    start: float
    end: float

    def to_list(self):
        return [self.text, round(self.start, 3), round(self.end, 3)]

    @classmethod
    def from_list(cls, row):
        return cls(str(row[0]), float(row[1]), float(row[2]))


@dataclass
class Transcript:
    """区間ひとつ分の文字起こし。時刻は元の配信の秒。"""
    engine: str
    words: list = field(default_factory=list)
    elapsed: float = 0.0          # 文字起こしにかかった秒数（比較用）
    device: str = ""

    @property
    def text(self):
        return join_words(w.text for w in self.words)

    def as_dict(self):
        data = asdict(self)
        data["words"] = [w.to_list() for w in self.words]
        return data

    @classmethod
    def from_dict(cls, data):
        return cls(engine=data["engine"],
                   words=[Word.from_list(row) for row in data.get("words") or []],
                   elapsed=float(data.get("elapsed") or 0.0),
                   device=data.get("device") or "")


def join_words(parts):
    """単語をつなぐ。日本語は詰め、英数字どうしの間だけ空白を入れる。"""
    out = ""
    for part in parts:
        part = (part or "").strip()
        if not part:
            continue
        if out and _is_ascii_word(out[-1]) and _is_ascii_word(part[0]):
            out += " "
        out += part
    return out


def _is_ascii_word(ch):
    return ch.isascii() and ch.isalnum()


# ---------------------------------------------------------------- 後処理

def remove_loops(words, min_repeat=3, max_unit=12, min_chars=2):
    """同じ語句が3回以上続く「ループ」を1回分だけ残して取り除く。

    ゲーム音やBGMが混ざると、どちらのエンジンでも同じ語句を延々と
    繰り返す失敗が起きる。繰り返しの単位は1〜max_unit語で探す
    （語は文節に近い単位なので、長い決まり文句も入る）。

    1文字の単位は対象外にする。「ははは」「www」のような笑いは本物の発話で、
    消してしまうとテロップが不自然になるため。
    """
    words = list(words)
    out = []
    i = 0
    while i < len(words):
        best = None            # (単位の語数, 繰り返し回数)
        for unit in range(1, max_unit + 1):
            if i + unit * min_repeat > len(words):
                break
            key = _unit_key(words, i, unit)
            if len(key) < min_chars:
                continue
            count = 1
            while (i + unit * (count + 1) <= len(words)
                   and _unit_key(words, i + unit * count, unit) == key):
                count += 1
            if count >= min_repeat and (best is None or unit * count > best[0] * best[1]):
                best = (unit, count)
        if best:
            unit, count = best
            out.extend(words[i:i + unit])
            i += unit * count
        else:
            out.append(words[i])
            i += 1
    return out


def _unit_key(words, start, unit):
    # 句読点や空白の揺れで同じ語句を見逃さないよう、比べるときだけ除く
    text = "".join(w.text for w in words[start:start + unit])
    return "".join(ch for ch in text if ch.isalnum())


def clean_words(words):
    """エンジンの出力を整える。空の語や時刻の逆転を直し、ループを除く。"""
    tidy = []
    for w in words:
        text = (w.text or "").strip()
        if not text:
            continue
        start, end = float(w.start), float(w.end)
        if end < start:
            start, end = end, start
        tidy.append(Word(text, start, end))
    tidy.sort(key=lambda w: (w.start, w.end))
    return remove_loops(tidy)


def attach_punctuation(text, words):
    """時刻の付いた語に、全文にある句読点を戻す。

    ForcedAligner の日本語の出力は1文字ずつで、句読点が落ちている。
    テロップを句点で区切るために、全文を先頭から突き合わせ、語と語の間に
    挟まっていた文字（句読点など）を直前の語の後ろに付け直す。
    全文と突き合わせられない語があれば、そこから先はそのまま返す。
    """
    out = [Word(w.text, w.start, w.end) for w in words]
    pos = 0
    for index, w in enumerate(out):
        token = w.text.strip()
        found = text.find(token, pos) if token else -1
        if found < 0:
            break
        skipped = text[pos:found].strip()
        if skipped and index > 0:
            out[index - 1].text += skipped
        pos = found + len(token)
    else:
        tail = text[pos:].strip()
        if tail and out:
            out[-1].text += tail
    return out


def _script(ch):
    if "぀" <= ch <= "ゟ":
        return "hira"
    if "゠" <= ch <= "ヿ" or ch == "ー":
        return "kata"
    if "一" <= ch <= "鿿" or ch in "々〆":
        return "kanji"
    if ch.isascii() and ch.isalnum():
        return "ascii"
    return "other"


JOIN_GAP_SEC = 0.3     # 文字と文字の間がこれ以上空いていたら、別の語にする
PARTICLES = "はがにでもへと"


def _ends_with_particle(word):
    """語が助詞で終わっていて、次はもう新しい文節か。

    「を」は話し言葉ではほぼ助詞なので、いつでも区切る（ゲームを／やって）。
    それ以外は、漢字・カタカナの語の直後に1文字だけ付いたときに限る
    （部屋に／入る）。「ちょっと」の「と」のような、ひらがなの語の一部は区切らない。
    """
    if not word:
        return False
    if word[-1] == "を":
        return True
    return (word[-1] in PARTICLES and len(word) >= 2
            and _script(word[-2]) in ("kanji", "kata"))


def group_japanese(chars, join_gap=JOIN_GAP_SEC):
    """1文字ずつの語を、文節に近い単位にまとめる。

    「漢字やカタカナの並び＋後ろに続くひらがな」を1つにする簡単な規則
    （今日は / ホラーゲームを / やっていきます）。形態素解析ほど正確ではないが、
    依存を増やさずにテロップの折り返しとループ除去に使える単位になる。

    間が空いた所は必ず分ける。分けないと、語の中に沈黙が隠れて詰められなくなる。
    1文字ずつでない語（英単語など）は、そのまま1語として扱う。
    """
    out = []
    for w in chars:
        text = w.text
        if not out:
            out.append(Word(text, w.start, w.end))
            continue
        prev = out[-1]
        tail = _script(prev.text[-1]) if prev.text else "other"
        # 長音「ー」は直前の文字の続き（「えーと」を「え／ーと」に割らない）
        head = tail if text[:1] == "ー" else (_script(text[0]) if text else "other")
        ends_sentence = prev.text[-1:] in "。！？!?、，,"
        same_word = (
            w.start - prev.end < join_gap
            and not ends_sentence
            and head != "other"
            # ひらがなの後に漢字・カタカナ・英数字が来たら、新しい文節が始まったとみなす
            and not (tail == "hira" and head != "hira")
            and not _ends_with_particle(prev.text)
            # 英数字と日本語の境目でも分ける
            and (head == "ascii") == (tail == "ascii"))
        if same_word:
            prev.text += text
            prev.end = max(prev.end, w.end)
        else:
            out.append(Word(text, w.start, w.end))
    return out


# ---------------------------------------------------------------- 分割

def choose_split_points(energies, frame_sec, total_sec,
                        max_chunk=QWEN_MAX_CHUNK_SEC, search=SPLIT_SEARCH_SEC,
                        smooth_sec=SPLIT_SMOOTH_SEC):
    """長い音声を max_chunk 秒以内に分ける位置（秒）を決める。

    分割点はできるだけ発話の切れ目にしたい。まだ文字起こし前なので、
    上限の手前 search 秒の中で一番静かな所を選ぶ。ゲーム音で
    完全な無音は無いことが多いが、相対的に静かな所は喋りの合間であることが多い。

    静かさは smooth_sec 秒の幅でならしてから比べる。1フレーム（50ms）だけで見ると、
    語の中の一瞬の途切れ（「か」の子音の前など）を選んでしまい、
    「これかな」が「これかだ」に化けた（実機で確認）。
    探す幅は塊の長さの半分までにする。塊が短いときに、先頭近くで切って
    細切れにならないように。

    energies はフレームごとの音の大きさ。戻り値は昇順の分割秒（両端は含まない）。
    """
    half = max(0, int(smooth_sec / frame_sec / 2))
    prefix = [0.0]
    for value in energies:
        prefix.append(prefix[-1] + value)

    def smoothed(k):
        a, b = max(0, k - half), min(len(energies), k + half + 1)
        return (prefix[b] - prefix[a]) / max(1, b - a)

    search = min(search, max_chunk / 2.0)
    points = []
    start = 0.0
    while total_sec - start > max_chunk:
        hi = start + max_chunk
        lo = max(start + max_chunk - search, start + 1.0)
        first = int(lo / frame_sec)
        last = min(int(hi / frame_sec), len(energies))
        if first < last:
            best = min(range(first, last), key=lambda k: (smoothed(k), -k))
            cut = best * frame_sec
        else:
            cut = hi
        cut = min(max(cut, start + 1.0), hi)
        points.append(cut)
        start = cut
    return points


def merge_chunks(chunks):
    """分割して処理した結果を、元の時刻に戻してつなぐ。

    chunks は [(区間の先頭からのずれ秒, 長さ秒, [Word（各断片の先頭が0）])]。
    断片の外にはみ出した時刻は、その断片の範囲に収める。
    """
    out = []
    for offset, length, words in chunks:
        for w in words:
            start = min(max(w.start, 0.0), length) + offset
            end = min(max(w.end, 0.0), length) + offset
            out.append(Word(w.text, start, max(start, end)))
    out.sort(key=lambda w: (w.start, w.end))
    return out


def frame_energies(samples, frame_sec=SPLIT_FRAME_SEC, rate=SAMPLE_RATE):
    """分割点探し用に、フレームごとの平均振幅を出す。"""
    import numpy as np

    size = max(1, int(frame_sec * rate))
    count = len(samples) // size
    if count == 0:
        return []
    frames = np.abs(samples[:count * size]).reshape(count, size)
    return frames.mean(axis=1).tolist()


# ---------------------------------------------------------------- 依存と音声

def find_ffmpeg():
    """ffmpeg の場所。PATH にあればそれを、無ければ imageio-ffmpeg 同梱のものを使う。"""
    found = shutil.which("ffmpeg")
    if found:
        return found
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:          # noqa: BLE001（入っていない・壊れている）
        return ""


def _has(module):
    try:
        return importlib.util.find_spec(module) is not None
    except (ImportError, ValueError):
        return False


ENGINE_MODULES = {
    "qwen": ["numpy", "torch", "qwen_asr"],
    "whisper": ["numpy", "faster_whisper"],
}


def missing(engine):
    """エンジンを使うのに足りないもの（日本語の説明）の一覧。空なら使える。"""
    problems = []
    if paths.is_frozen():
        return ["exe版には文字起こし用のライブラリを同梱していません。"
                "Python から起動（start_hidden.cmd）してください。"]
    if engine not in ENGINE_MODULES:
        return ["文字起こしエンジン「%s」はありません。" % engine]
    if not find_ffmpeg():
        problems.append("ffmpeg が見つかりません（pip install imageio-ffmpeg で入ります）。")
    names = {"numpy": "numpy", "torch": "torch", "qwen_asr": "qwen-asr",
             "faster_whisper": "faster-whisper"}
    lacking = [names[m] for m in ENGINE_MODULES[engine] if not _has(m)]
    if lacking:
        problems.append("%s が入っていません（stream_highlight/requirements-transcribe.txt"
                        " を参照）。" % "、".join(lacking))
    return problems


def capabilities():
    """画面に出す、各エンジンが使えるかどうか。重いライブラリは読み込まない。"""
    return {
        "engines": [{"id": key, "label": label, "missing": missing(key)}
                    for key, label in ENGINE_LABELS.items()],
        "default_engine": DEFAULT_ENGINE,
        "ffmpeg": find_ffmpeg(),
    }


def _no_window():
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def extract_audio(media_path, start, end):
    """動画の start〜end 秒を 16kHz モノラルの float32 配列で取り出す。"""
    import numpy as np

    ffmpeg = find_ffmpeg()
    if not ffmpeg:
        raise TranscribeError("ffmpeg が見つかりません。")
    if not os.path.isfile(media_path):
        raise TranscribeError("動画ファイルが見つかりません: %s" % media_path)
    # -ss を -i より前に置くと、3時間の配信でも先頭から読まずに飛べる
    args = [ffmpeg, "-nostdin", "-v", "error", "-ss", "%.3f" % start,
            "-t", "%.3f" % max(0.0, end - start), "-i", media_path,
            "-vn", "-ac", "1", "-ar", str(SAMPLE_RATE), "-f", "s16le", "-"]
    done = subprocess.run(args, capture_output=True, creationflags=_no_window())
    if done.returncode != 0:
        detail = done.stderr.decode("utf-8", "replace").strip().splitlines()
        raise TranscribeError("音声を切り出せませんでした: %s"
                              % (detail[-1] if detail else "ffmpeg が失敗しました"))
    return np.frombuffer(done.stdout, dtype=np.int16).astype(np.float32) / 32768.0


# ---------------------------------------------------------------- エンジン

def _is_oom(exc):
    """GPUのメモリ不足かどうか。torch の版によって例外の型が違うので文言でも見る。"""
    name = type(exc).__name__
    return "OutOfMemory" in name or "out of memory" in str(exc).lower()


def _cpu_fallback_reason(exc):
    """CPU に切り替えれば続けられる失敗なら、その理由（画面に出す）。違えば空。"""
    if _is_oom(exc):
        return "GPUのメモリが足りない"
    if _is_gpu_library_error(exc):
        return "GPU用のライブラリ（cuBLAS/cuDNN）を読み込めない"
    return ""


def _free_gpu():
    gc.collect()
    try:
        import torch
        if torch.cuda.is_available():
            torch.cuda.empty_cache()
    except Exception:          # noqa: BLE001
        pass


def model_cached(repo_id):
    """モデルがダウンロード済みか（初回ダウンロード中の表示に使う）。"""
    try:
        from huggingface_hub import constants
        folder = os.path.join(constants.HF_HUB_CACHE,
                              "models--" + repo_id.replace("/", "--"))
        return os.path.isdir(os.path.join(folder, "snapshots"))
    except Exception:          # noqa: BLE001
        return False


class Engine:
    """文字起こしエンジンの共通の形。"""
    name = ""
    max_chunk_sec = None       # 1回で処理できる長さ。None なら分けない
    repos = ()                 # 初回ダウンロードの判定に使う

    def __init__(self):
        self.device = ""

    def needs_download(self):
        return not all(model_cached(r) for r in self.repos)

    def load(self, prefer_gpu=True):
        raise NotImplementedError

    def unload(self):
        raise NotImplementedError

    def run(self, samples):
        """先頭を0秒とした Word の一覧を返す。"""
        raise NotImplementedError


class QwenEngine(Engine):
    name = "qwen"
    max_chunk_sec = QWEN_MAX_CHUNK_SEC
    repos = (QWEN_ASR_MODEL, QWEN_ALIGNER_MODEL)

    def __init__(self):
        super().__init__()
        self.model = None

    def load(self, prefer_gpu=True):
        import torch
        from qwen_asr import Qwen3ASRModel

        use_gpu = prefer_gpu and torch.cuda.is_available()
        dtype = torch.float16 if use_gpu else torch.float32
        device = "cuda:0" if use_gpu else "cpu"
        self.model = Qwen3ASRModel.from_pretrained(
            QWEN_ASR_MODEL, dtype=dtype, device_map=device,
            # 180秒の早口でも途中で切れない長さ。ループしてもここで止まる
            max_inference_batch_size=1, max_new_tokens=2048,
            forced_aligner=QWEN_ALIGNER_MODEL,
            forced_aligner_kwargs=dict(dtype=dtype, device_map=device),
        )
        self.device = "GPU" if use_gpu else "CPU"
        self.max_chunk_sec = QWEN_MAX_CHUNK_SEC
        if use_gpu:
            free, _ = torch.cuda.mem_get_info()
            self.max_chunk_sec = chunk_sec_for_vram(free / 1e9)

    def unload(self):
        self.model = None
        _free_gpu()

    def run(self, samples):
        results = self.model.transcribe(audio=[(samples, SAMPLE_RATE)],
                                        language=["Japanese"], return_time_stamps=True)
        if not results or results[0].time_stamps is None:
            return []
        # time_stamps は ForcedAlignResult（.items に1文字ずつ）。版による違いに備えて平らにする
        chars = [Word(str(item.text), float(item.start_time), float(item.end_time))
                 for item in _flatten([results[0].time_stamps])]
        # 日本語は1文字ずつ・句読点なしで返るので、句読点を戻し、文節に近い単位にまとめる
        return group_japanese(attach_punctuation(results[0].text or "", chars))


def _flatten(stamps):
    """time_stamps は版によって入れ子の深さが違うので、平らにして読む。"""
    for item in stamps:
        if hasattr(item, "start_time"):
            yield item
        elif hasattr(item, "items"):
            yield from _flatten(item.items)
        elif isinstance(item, (list, tuple)):
            yield from _flatten(item)


class WhisperEngine(Engine):
    name = "whisper"
    repos = ("mobiuslabsgmbh/faster-whisper-large-v3-turbo",)

    def __init__(self):
        super().__init__()
        self.model = None

    def load(self, prefer_gpu=True):
        _use_torch_cuda_dlls()
        _load_torch_cudnn_first()
        from faster_whisper import WhisperModel

        use_gpu = prefer_gpu and _whisper_gpu()
        self.model = WhisperModel(WHISPER_MODEL, device="cuda" if use_gpu else "cpu",
                                  compute_type="float16" if use_gpu else "int8")
        self.device = "GPU" if use_gpu else "CPU"

    def unload(self):
        self.model = None
        _free_gpu()

    def run(self, samples):
        segments, _ = self.model.transcribe(
            samples, language="ja", word_timestamps=True, vad_filter=True,
            # 前の文脈を引き継ぐと、一度ループすると抜け出せなくなる
            condition_on_previous_text=False)
        words = []
        for seg in segments:
            for w in seg.words or []:
                text = (w.word or "").strip()
                if text:
                    words.append(Word(text, float(w.start), float(w.end)))
        # 日本語は空白が無いので、数文字ずつに割れて返る。Qwen と同じ単位にまとめる
        return extend_word_ends(group_japanese(words), WHISPER_END_LAG_SEC)


# whisper は語の終わりを実際より早く返す（合成音声で実測すると文末で0.1〜0.2秒）。
# そのままだと、余白0.15秒を足してもカット点が語尾ぎりぎりになり、語尾を削ることがあった。
WHISPER_END_LAG_SEC = 0.15


def extend_word_ends(words, lag):
    """語の終わりを lag 秒だけ後ろに延ばす。次の語の始まりは越えない。"""
    out = []
    for index, w in enumerate(words):
        limit = words[index + 1].start if index + 1 < len(words) else w.end + lag
        out.append(Word(w.text, w.start, max(w.end, min(w.end + lag, limit))))
    return out


def _whisper_gpu():
    try:
        import ctranslate2
        return ctranslate2.get_cuda_device_count() > 0
    except Exception:          # noqa: BLE001
        return False


def _use_torch_cuda_dlls():
    """faster-whisper（ctranslate2）が GPU で使う cuBLAS / cuDNN を、torch の同梱品から読ませる。

    ctranslate2 は CUDA のライブラリを同梱しておらず、別に入れていないと
    「cublas64_12.dll が見つからない」で GPU 処理が失敗する。CUDA 版の torch は
    同じ版の DLL を torch/lib に持っているので、そこを探索先に加える。
    """
    if os.name != "nt":
        return
    try:
        spec = importlib.util.find_spec("torch")
    except (ImportError, ValueError):
        return
    if not spec or not spec.submodule_search_locations:
        return
    lib = os.path.join(list(spec.submodule_search_locations)[0], "lib")
    if not os.path.isfile(os.path.join(lib, "cublas64_12.dll")):
        return
    # ctranslate2 は必要になった時点で DLL を探すので、PATH にも足しておく
    if lib not in os.environ.get("PATH", ""):
        os.environ["PATH"] = lib + os.pathsep + os.environ.get("PATH", "")
    try:
        os.add_dll_directory(lib)
    except (AttributeError, OSError):
        pass


def _load_torch_cudnn_first():
    """whisper より先に、torch の cuDNN を読み込んでおく。

    ctranslate2 は自前の cudnn64_9.dll を同梱している。whisper が先にそれを読むと、
    同じプロセスで後から Qwen（torch）を読み込んだときに cuDNN の版が食い違い、
    「Could not load symbol cudnnGetLibConfig」でプロセスごと落ちた（実機で確認）。
    同じ名前の DLL は1つしか読み込まれないので、torch の cuDNN を先に読ませれば
    両方がそれを使う。エンジンの切り替えは、どちらの順でも動くことを確かめた。
    """
    if not _has("torch"):
        return
    try:
        import torch
        if torch.cuda.is_available():
            torch.backends.cudnn.version()
    except Exception:          # noqa: BLE001（torch が壊れていても whisper は CPU で動かせる）
        pass


def _is_gpu_library_error(exc):
    """GPU 用のライブラリが読めない失敗か（この場合も CPU に切り替えれば動く）。"""
    text = str(exc).lower()
    return any(key in text for key in ("cublas", "cudnn", "cuda driver", "libcudart",
                                       "cannot be loaded"))


ENGINES = {"qwen": QwenEngine, "whisper": WhisperEngine}


class EngineHolder:
    """読み込んだモデルを使い回す。ただし同時に持つのは1つだけ。

    VRAM 8GB では2つのエンジンを同時に載せられない。別のエンジンに
    切り替えるときは、先に今のものを解放してから読み込む。
    """

    def __init__(self):
        self._engine = None
        self._lock = threading.Lock()

    @property
    def lock(self):
        return self._lock

    def get(self, name, report=None):
        if self._engine is not None and self._engine.name == name:
            return self._engine
        self.release()
        if name not in ENGINES:
            raise TranscribeError("文字起こしエンジン「%s」はありません。" % name)
        problems = missing(name)
        if problems:
            raise TranscribeError(" ".join(problems))
        engine = ENGINES[name]()
        if report:
            report("文字起こしモデルを初回ダウンロード中（数GBあるため、数分〜数十分かかります）…"
                   if engine.needs_download() else "文字起こしモデルを読み込み中…")
        try:
            engine.load(prefer_gpu=True)
        except Exception as exc:              # noqa: BLE001
            reason = _cpu_fallback_reason(exc)
            if not reason:
                raise
            engine.unload()
            if report:
                report("%sため、CPUで読み込み直しています（遅くなります）…" % reason)
            engine.load(prefer_gpu=False)
        self._engine = engine
        return engine

    def release(self):
        if self._engine is not None:
            self._engine.unload()
            self._engine = None


holder = EngineHolder()


def run_engine(engine, samples, report=None):
    """分割が必要なら分けて処理し、時刻をつないで返す。メモリ不足ならCPUでやり直す。"""
    total = len(samples) / float(SAMPLE_RATE)
    points = []
    if engine.max_chunk_sec and total > engine.max_chunk_sec:
        points = choose_split_points(frame_energies(samples), SPLIT_FRAME_SEC, total,
                                     max_chunk=engine.max_chunk_sec)
    bounds = [0.0] + points + [total]
    chunks = []
    for k in range(len(bounds) - 1):
        lo, hi = bounds[k], bounds[k + 1]
        piece = samples[int(lo * SAMPLE_RATE):int(hi * SAMPLE_RATE)]
        if report and len(bounds) > 2:
            report("（%d/%d 分割）" % (k + 1, len(bounds) - 1))
        try:
            words = engine.run(piece)
        except Exception as exc:              # noqa: BLE001
            reason = _cpu_fallback_reason(exc)
            if not reason or engine.device == "CPU":
                raise
            engine.unload()
            _free_gpu()
            if report:
                report("%sため、CPUに切り替えます（遅くなります）…" % reason)
            engine.load(prefer_gpu=False)
            words = engine.run(piece)
        chunks.append((lo, hi - lo, words))
    return merge_chunks(chunks)


# ---------------------------------------------------------------- キャッシュ

def _media_stamp(media_path):
    """同じ名前でも中身が変われば別物として扱う（大きさと更新時刻で見る）。"""
    try:
        st = os.stat(media_path)
        return "%d-%d" % (st.st_size, int(st.st_mtime))
    except OSError:
        return "missing"


def cache_path(media_path, start, end, engine):
    raw = "|".join([os.path.abspath(media_path), _media_stamp(media_path),
                    "%.3f" % start, "%.3f" % end, engine, str(TRANSCRIPT_VERSION)])
    digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:20]
    return os.path.join(cache.CACHE_DIR, "transcripts", "%s_%s.json.gz" % (engine, digest))


def load_cached(media_path, start, end, engine):
    payload = cache._read_gz(cache_path(media_path, start, end, engine))
    if not payload or payload.get("version") != TRANSCRIPT_VERSION:
        return None
    try:
        return Transcript.from_dict(payload["transcript"])
    except (KeyError, TypeError, ValueError, IndexError):
        return None


def save_cached(media_path, start, end, transcript):
    path = cache_path(media_path, start, end, transcript.engine)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    payload = {"version": TRANSCRIPT_VERSION, "saved_at": time.time(),
               "media": os.path.abspath(media_path), "start": start, "end": end,
               "transcript": transcript.as_dict()}
    # cache._write_gz は CACHE_DIR 直下を作るだけなので、サブフォルダは上で作っておく
    cache._write_gz(path, payload)


def transcribe_segment(media_path, start, end, engine_name, report=None, use_cache=True):
    """区間 start〜end を文字起こしする。時刻は元の配信の秒で返す。

    キャッシュに残すのはエンジンの生の出力（後処理前）。ループ除去などの
    後処理を直しても、文字起こしをやり直さずに済むように。
    """
    if use_cache:
        found = load_cached(media_path, start, end, engine_name)
        if found is not None:
            found.words = clean_words(found.words)
            return found

    samples = extract_audio(media_path, start, end)
    with holder.lock:
        engine = holder.get(engine_name, report)
        began = time.time()
        words = run_engine(engine, samples, report)
        elapsed = time.time() - began
        device = engine.device
    raw = Transcript(engine_name, [Word(w.text, w.start + start, w.end + start) for w in words],
                     elapsed=elapsed, device=device)
    save_cached(media_path, start, end, raw)
    return Transcript(engine_name, clean_words(raw.words), elapsed=elapsed, device=device)
