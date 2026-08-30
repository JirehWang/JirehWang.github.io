#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""台語有聲聖經時間軸校正工具。

流程：下載 FHL 經文與音檔、執行 ASR、用頭/尾三字錨點找出每節位置，
最後以「下一節開始時間」作為本節結束時間並驗證輸出。

ASR 預設使用規格指定的 faster-whisper；保留 Breeze-ASR-26 作為可選後端：

    python breeze_aligner.py --bid 19 --chap 23
    python breeze_aligner.py --backend breeze --bid 19 --chap 23
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Sequence


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


FHL_BOOK_NAMES = {
    1: "Gen", 2: "Ex", 3: "Lev", 4: "Num", 5: "Deut", 6: "Josh", 7: "Judg", 8: "Ruth",
    9: "1Sam", 10: "2Sam", 11: "1Kin", 12: "2Kin", 13: "1Chr", 14: "2Chr", 15: "Ezra",
    16: "Neh", 17: "Esth", 18: "Job", 19: "Ps", 20: "Prov", 21: "Eccl", 22: "Song",
    23: "Isa", 24: "Jer", 25: "Lam", 26: "Ezek", 27: "Dan", 28: "Hos", 29: "Joel",
    30: "Amos", 31: "Obad", 32: "Jonah", 33: "Mic", 34: "Nah", 35: "Hab", 36: "Zeph",
    37: "Hag", 38: "Zech", 39: "Mal", 40: "Matt", 41: "Mark", 42: "Luke", 43: "John",
    44: "Acts", 45: "Rom", 46: "1Cor", 47: "2Cor", 48: "Gal", 49: "Eph", 50: "Phil",
    51: "Col", 52: "1Thess", 53: "2Thess", 54: "1Tim", 55: "2Tim", 56: "Titus",
    57: "Phlm", 58: "Heb", 59: "Jas", 60: "1Pet", 61: "2Pet", 62: "1John", 63: "2John",
    64: "3John", 65: "Jude", 66: "Rev",
}

BOOK_TITLES = {
    1: "創世記", 2: "出埃及記", 3: "利未記", 4: "民數記", 5: "申命記", 6: "約書亞記",
    7: "士師記", 8: "路得記", 9: "撒母耳記上", 10: "撒母耳記下", 11: "列王紀上",
    12: "列王紀下", 13: "歷代志上", 14: "歷代志下", 15: "以斯拉記", 16: "尼希米記",
    17: "以斯帖記", 18: "約伯記", 19: "詩篇", 20: "箴言", 21: "傳道書", 22: "雅歌",
    23: "以賽亞書", 24: "耶利米書", 25: "耶利米哀歌", 26: "以西結書", 27: "但以理書",
    28: "何西阿書", 29: "約珥書", 30: "阿摩司書", 31: "俄巴底亞書", 32: "約拿書",
    33: "彌迦書", 34: "那鴻書", 35: "哈巴谷書", 36: "西番雅書", 37: "哈該書",
    38: "撒迦利亞書", 39: "瑪拉基書", 40: "馬太福音", 41: "馬可福音", 42: "路加福音",
    43: "約翰福音", 44: "使徒行傳", 45: "羅馬書", 46: "哥林多前書", 47: "哥林多後書",
    48: "加拉太書", 49: "以弗所書", 50: "腓立比書", 51: "歌羅西書", 52: "帖撒羅尼迦前書",
    53: "帖撒羅尼迦後書", 54: "提摩太前書", 55: "提摩太後書", 56: "提多書",
    57: "腓利門書", 58: "希伯來書", 59: "雅各書", 60: "彼得前書", 61: "彼得後書",
    62: "約翰一書", 63: "約翰二書", 64: "約翰三書", 65: "猶大書", 66: "啟示錄",
}

CHINESE_NUMERALS = "〇零一二三四五六七八九十百千幾"
VERSE_MARKER_RE = re.compile(rf"^第[{CHINESE_NUMERALS}0-9]+節")
SPOKEN_MARKER_ONLY_RE = re.compile(rf"^(?:第[{CHINESE_NUMERALS}0-9]+(?:節|集|句))+$")
HTML_TAG_RE = re.compile(r"<[^>]*>")
BRACKET_NOTE_RE = re.compile(r"〔[^〕]*〕|【[^】]*】|\[[^\]]*\]")
PAREN_NOTE_RE = re.compile(r"\([^)]*\)|（[^）]*）")
CJK_CHAR_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")

# 台語朗讀中常見的詩篇字母分段口播。單字變體也列入，但只會在
# 每個 ASR span 的開頭剝除，避免誤刪經文中同名的普通字。
HEBREW_SECTION_MARKERS = (
    "阿勒弗", "伯特", "貝特", "基默", "吉默", "達勒特", "達勒", "赫特",
    "瓦烏", "瓦夫", "撒音", "提特", "約德", "卡夫", "拉麥德", "麥姆",
    "努恩", "索梅克", "阿因", "察代", "柯弗", "雷什", "赫", "派", "辛", "陶",
)
SORTED_HEBREW_MARKERS = tuple(sorted(HEBREW_SECTION_MARKERS, key=len, reverse=True))


class AlignmentError(RuntimeError):
    """無法可靠產生時間軸時拋出的錯誤。"""


class TimestampValidationError(AlignmentError):
    """輸出時間軸違反格式或時長規則。"""


@dataclass(frozen=True)
class TranscriptSpan:
    text: str
    start: float
    end: float


@dataclass(frozen=True)
class VerseSpec:
    sec: int
    display_text: str
    match_text: str
    head: str
    tail: str


@dataclass(frozen=True)
class AnchorMatch:
    index: int
    start: float
    end: float
    score: float
    exact: bool


@dataclass
class AlignmentResult:
    verses: list[dict[str, Any]]
    warnings: list[str]
    head_matches: list[AnchorMatch]
    tail_matches: list[Optional[AnchorMatch]]


def _remove_annotations(text: str) -> str:
    """移除 HTML、括號拼音及經文題註；保留原始經文另作顯示。"""
    value = str(text or "")
    value = HTML_TAG_RE.sub("", value)
    value = BRACKET_NOTE_RE.sub("", value)
    value = PAREN_NOTE_RE.sub("", value)
    return value


def clean_text(raw_text: str) -> str:
    """清除 HTML 並整理顯示文字，但保留拼音與題註給前端查看。"""
    value = HTML_TAG_RE.sub("", str(raw_text or ""))
    return re.sub(r"\s+", " ", value).strip()


def normalize_for_matching(text: str) -> str:
    """經文端正規化：只留下可發音的漢字。"""
    return "".join(CJK_CHAR_RE.findall(_remove_annotations(text)))


def _strip_leading_spoken_markers(text: str) -> str:
    value = VERSE_MARKER_RE.sub("", text)
    value = value.strip()
    for _ in range(len(SORTED_HEBREW_MARKERS) + 2):
        changed = False
        for marker in SORTED_HEBREW_MARKERS:
            if value.startswith(marker):
                value = value[len(marker):].lstrip()
                changed = True
                break
        if not changed:
            break
    return value


def _normalize_audio_with_prefix(text: str) -> tuple[str, int, int]:
    """回傳正規化文字、開頭被剝除的漢字數、原始漢字數。

    開頭口播詞被移除時，用剝除字數估算其在 ASR span 中占用的時間，
    讓第一節不再無條件落在 0 秒。
    """
    without_annotations = _remove_annotations(text)
    all_han = "".join(CJK_CHAR_RE.findall(without_annotations))
    if SPOKEN_MARKER_ONLY_RE.fullmatch(all_han):
        spoken = ""
    else:
        spoken = _strip_leading_spoken_markers(all_han)
    prefix_chars = max(0, len(all_han) - len(spoken))
    return spoken, prefix_chars, len(all_han)


def normalize_audio_text(text: str) -> str:
    """ASR 端正規化並移除節數、希伯來字母分段口播。"""
    return _normalize_audio_with_prefix(text)[0]


def head_tail(text: str, size: int = 3) -> tuple[str, str]:
    if size <= 0:
        raise ValueError("錨點長度必須大於 0")
    if len(text) < size:
        raise AlignmentError(f"經文可比對漢字不足 {size} 字：{text!r}")
    return text[:size], text[-size:]


def prepare_verse_specs(records: Iterable[dict[str, Any]]) -> list[VerseSpec]:
    specs: list[VerseSpec] = []
    previous_sec = 0
    for record in records:
        try:
            sec = int(record.get("sec"))
        except (TypeError, ValueError) as exc:
            raise AlignmentError(f"經節編號無效：{record!r}") from exc
        if sec <= previous_sec:
            raise AlignmentError(f"經節編號未嚴格遞增：{previous_sec} -> {sec}")
        raw_text = record.get("bible_text", record.get("text", ""))
        display = clean_text(raw_text)
        match_text = normalize_for_matching(raw_text)
        if not match_text:
            raise AlignmentError(f"第 {sec} 節沒有可比對的漢字")
        head, tail = head_tail(match_text)
        specs.append(VerseSpec(sec, display, match_text, head, tail))
        previous_sec = sec
    if not specs:
        raise AlignmentError("API 沒有回傳任何經節")
    return specs


class TranscriptTimeline:
    """將 ASR spans 串成可由字元位置反推秒數的時間軸。"""

    def __init__(self, text: str, starts: Sequence[float], ends: Sequence[float]):
        self.text = text
        self.starts = list(starts)
        self.ends = list(ends)

    def _make_match(self, index: int, needle_length: int, score: float, exact: bool) -> AnchorMatch:
        last = min(index + needle_length - 1, len(self.starts) - 1)
        return AnchorMatch(index, self.starts[index], self.ends[last], score, exact)

    def find_anchor(
        self,
        needle: str,
        start: int = 0,
        end: Optional[int] = None,
        fuzzy_threshold: float = 2 / 3,
    ) -> Optional[AnchorMatch]:
        if not needle or not self.text:
            return None
        lower = max(0, start)
        upper = len(self.text) if end is None else min(len(self.text), end)
        if upper - lower < len(needle):
            return None

        index = self.text.find(needle, lower, upper)
        if index >= 0 and index + len(needle) <= upper:
            return self._make_match(index, len(needle), 1.0, True)

        best: Optional[AnchorMatch] = None
        last_start = upper - len(needle)
        for candidate_index in range(lower, last_start + 1):
            candidate = self.text[candidate_index:candidate_index + len(needle)]
            score = SequenceMatcher(None, needle, candidate).ratio()
            if score < fuzzy_threshold:
                continue
            match = self._make_match(candidate_index, len(needle), score, False)
            if best is None or match.score > best.score:
                best = match
        return best


def build_transcript_timeline(spans: Iterable[TranscriptSpan]) -> TranscriptTimeline:
    ordered = sorted(spans, key=lambda span: (float(span.start), float(span.end)))
    chars: list[str] = []
    starts: list[float] = []
    ends: list[float] = []

    for span in ordered:
        try:
            span_start = float(span.start)
            span_end = float(span.end)
        except (TypeError, ValueError) as exc:
            raise AlignmentError(f"ASR 時間戳無效：{span!r}") from exc
        if not math.isfinite(span_start) or not math.isfinite(span_end):
            continue
        span_start = max(0.0, span_start)
        span_end = max(span_start, span_end)
        text, prefix_chars, original_han_count = _normalize_audio_with_prefix(span.text)
        if not text:
            continue

        duration = span_end - span_start
        prefix_ratio = prefix_chars / original_han_count if original_han_count else 0.0
        content_start = span_start + duration * prefix_ratio
        chars.append(text)
        text_length = len(text)
        for index in range(text_length):
            starts.append(content_start + duration * index / text_length)
            ends.append(content_start + duration * (index + 1) / text_length)

    if not chars:
        raise AlignmentError("ASR 沒有可用的漢字與時間戳")
    return TranscriptTimeline("".join(chars), starts, ends)


def validate_timestamps(
    verses: Sequence[dict[str, Any]],
    total_duration: float,
    min_verse_seconds: float = 1.5,
    max_verse_seconds: float = 35.0,
) -> None:
    """驗證輸出時間軸的單調性、連續性、範圍與每節合理時長。"""
    try:
        duration = float(total_duration)
    except (TypeError, ValueError) as exc:
        raise TimestampValidationError("音訊總長度不是有效數字") from exc
    if not math.isfinite(duration) or duration <= 0:
        raise TimestampValidationError(f"音訊總長度無效：{total_duration!r}")
    if not verses:
        raise TimestampValidationError("時間軸沒有經節")

    previous_start = -math.inf
    tolerance = 0.011
    for index, verse in enumerate(verses):
        try:
            start = float(verse["start"])
            end = float(verse["end"])
        except (KeyError, TypeError, ValueError) as exc:
            raise TimestampValidationError(f"第 {index + 1} 筆缺少有效 start/end") from exc
        if not math.isfinite(start) or not math.isfinite(end):
            raise TimestampValidationError(f"第 {index + 1} 筆含有非有限時間")
        if start < -tolerance or end < -tolerance or end > duration + tolerance:
            raise TimestampValidationError(
                f"第 {index + 1} 節超出音訊範圍：{start:.2f} -> {end:.2f} / {duration:.2f}"
            )
        if start <= previous_start + tolerance / 10:
            raise TimestampValidationError(f"第 {index + 1} 節開始時間沒有嚴格遞增")
        if end <= start + tolerance / 10:
            raise TimestampValidationError(f"第 {index + 1} 節時長為零或負數")
        verse_duration = end - start
        if verse_duration < min_verse_seconds - tolerance:
            raise TimestampValidationError(
                f"第 {index + 1} 節過短：{verse_duration:.2f}s，最低 {min_verse_seconds:.2f}s"
            )
        if verse_duration > max_verse_seconds + tolerance:
            raise TimestampValidationError(
                f"第 {index + 1} 節過長：{verse_duration:.2f}s，最高 {max_verse_seconds:.2f}s"
            )
        if index > 0:
            previous_end = float(verses[index - 1]["end"])
            if abs(previous_end - start) > tolerance:
                raise TimestampValidationError(f"第 {index} 節與第 {index + 1} 節之間有空隙或重疊")
        previous_start = start

    last_end = float(verses[-1]["end"])
    if abs(last_end - duration) > tolerance:
        raise TimestampValidationError(f"最後一節沒有結束於音訊尾端：{last_end:.2f} != {duration:.2f}")


def align_verse_records(
    records: Iterable[dict[str, Any]],
    transcript_spans: Iterable[TranscriptSpan],
    duration: float,
    *,
    strict_tail: bool = False,
    min_verse_seconds: float = 1.5,
    max_verse_seconds: float = 35.0,
) -> AlignmentResult:
    """用頭三字找起點、用尾三字驗證內容，再以相鄰起點建立連續區間。"""
    specs = prepare_verse_specs(records)
    timeline = build_transcript_timeline(transcript_spans)

    head_matches: list[AnchorMatch] = []
    cursor = 0
    for spec in specs:
        match = timeline.find_anchor(spec.head, cursor)
        if match is None:
            raise AlignmentError(f"找不到第 {spec.sec} 節頭部錨點：{spec.head}")
        if head_matches and match.start <= head_matches[-1].start:
            raise AlignmentError(f"第 {spec.sec} 節頭部錨點未向前：{match.start:.2f}s")
        head_matches.append(match)
        cursor = match.index + len(spec.head)

    warnings: list[str] = []
    tail_matches: list[Optional[AnchorMatch]] = []
    for index, spec in enumerate(specs):
        search_start = head_matches[index].index + len(spec.head)
        search_end = head_matches[index + 1].index if index + 1 < len(head_matches) else None
        tail = timeline.find_anchor(spec.tail, search_start, search_end)
        tail_matches.append(tail)
        if tail is None:
            warning = f"第 {spec.sec} 節找不到尾部錨點：{spec.tail}"
            if strict_tail:
                raise AlignmentError(warning)
            warnings.append(warning)

    total_duration = round(float(duration), 2)
    starts = [round(match.start, 2) for match in head_matches]
    ends = starts[1:] + [total_duration]
    verses = [
        {
            "sec": spec.sec,
            "start": starts[index],
            "end": ends[index],
            "text": spec.display_text,
        }
        for index, spec in enumerate(specs)
    ]
    validate_timestamps(
        verses,
        total_duration,
        min_verse_seconds=min_verse_seconds,
        max_verse_seconds=max_verse_seconds,
    )
    return AlignmentResult(verses, warnings, head_matches, tail_matches)


def _timestamp_pair(value: Any) -> Optional[tuple[float, float]]:
    if not isinstance(value, (tuple, list)) or len(value) < 1:
        return None
    try:
        start = float(value[0]) if value[0] is not None else None
        end = float(value[1]) if len(value) > 1 and value[1] is not None else None
    except (TypeError, ValueError):
        return None
    if start is None or not math.isfinite(start):
        return None
    if end is None or not math.isfinite(end) or end < start:
        end = start + 0.25
    return start, end


def _breeze_spans(result: dict[str, Any]) -> list[TranscriptSpan]:
    spans: list[TranscriptSpan] = []
    for chunk in result.get("chunks", []) if isinstance(result, dict) else []:
        if not isinstance(chunk, dict):
            continue
        pair = _timestamp_pair(chunk.get("timestamp"))
        text = str(chunk.get("text", ""))
        if pair and text.strip():
            spans.append(TranscriptSpan(text, pair[0], pair[1]))
    return spans


def make_faster_whisper_transcriber(
    model_name: str = "small",
    device: str = "cpu",
    compute_type: str = "int8",
    language: str = "zh",
) -> Callable[[Any], list[TranscriptSpan]]:
    """建立 faster-whisper 轉錄器；匯入延後到真正選用此後端時。"""
    try:
        from faster_whisper import WhisperModel
    except ImportError as exc:
        raise AlignmentError(
            "缺少 faster-whisper，請先執行：pip install faster-whisper miniaudio numpy scipy"
        ) from exc

    model = WhisperModel(model_name, device=device, compute_type=compute_type)

    def transcribe(audio: Any) -> list[TranscriptSpan]:
        segments, _ = model.transcribe(
            audio,
            language=language,
            word_timestamps=True,
            vad_filter=True,
            condition_on_previous_text=False,
        )
        spans: list[TranscriptSpan] = []
        for segment in segments:
            words = getattr(segment, "words", None)
            if words:
                for word in words:
                    start = getattr(word, "start", None)
                    end = getattr(word, "end", None)
                    if start is None:
                        start = getattr(segment, "start", None)
                    if end is None:
                        end = getattr(segment, "end", None)
                    pair = _timestamp_pair((start, end))
                    if pair and getattr(word, "word", "").strip():
                        spans.append(TranscriptSpan(word.word, pair[0], pair[1]))
            else:
                pair = _timestamp_pair((getattr(segment, "start", None), getattr(segment, "end", None)))
                if pair and getattr(segment, "text", "").strip():
                    spans.append(TranscriptSpan(segment.text, pair[0], pair[1]))
        return spans

    return transcribe


def make_breeze_transcriber(
    model_name: str = "MediaTek-Research/Breeze-ASR-26",
) -> Callable[[Any], list[TranscriptSpan]]:
    """建立既有 Breeze-ASR-26 後端，供沒有 faster-whisper 的環境選用。"""
    try:
        from transformers import pipeline
    except ImportError as exc:
        raise AlignmentError(
            "缺少 transformers，請先執行：pip install transformers torch"
        ) from exc

    pipe = pipeline(
        "automatic-speech-recognition",
        model=model_name,
        chunk_length_s=30,
        return_timestamps=True,
    )

    def transcribe(audio: Any) -> list[TranscriptSpan]:
        result = pipe(audio, chunk_length_s=30, return_timestamps=True)
        return _breeze_spans(result)

    return transcribe


def decode_audio(audio_path: Path) -> tuple[Any, float]:
    """用 miniaudio 解碼，再以 scipy 轉成 16kHz mono float32。"""
    try:
        import miniaudio
        import numpy as np
        import scipy.signal
    except ImportError as exc:
        raise AlignmentError("缺少音訊依賴，請先執行：pip install miniaudio numpy scipy") from exc

    decoded = miniaudio.decode_file(str(audio_path))
    if decoded.sample_rate <= 0 or decoded.nchannels <= 0:
        raise AlignmentError(f"音訊格式無效：{audio_path}")
    duration = decoded.num_frames / decoded.sample_rate
    samples = np.frombuffer(decoded.samples, dtype=np.int16).astype(np.float32) / 32768.0
    if decoded.nchannels > 1:
        samples = samples.reshape(-1, decoded.nchannels).mean(axis=1)
    target_len = max(1, round(len(samples) * 16000 / decoded.sample_rate))
    audio_16k = scipy.signal.resample(samples, target_len).astype(np.float32)
    return audio_16k, float(duration)


def _read_url(url: str, timeout: float) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def get_chapter_text(
    bid: int,
    chap: int,
    version: str = "tghg",
    timeout: float = 30.0,
) -> list[dict[str, Any]]:
    if bid not in FHL_BOOK_NAMES:
        raise AlignmentError(f"書卷編號必須介於 1 到 66：{bid}")
    if chap <= 0:
        raise AlignmentError(f"章數必須大於 0：{chap}")
    query = urllib.parse.quote(f"{FHL_BOOK_NAMES[bid]} {chap}")
    version_query = urllib.parse.quote(version)
    url = f"https://bible.fhl.net/json/qsb.php?qstr={query}&version={version_query}"
    try:
        data = json.loads(_read_url(url, timeout).decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise AlignmentError(f"無法取得經文 API：{url}") from exc
    records = data.get("record", []) if isinstance(data, dict) else []
    if not isinstance(records, list):
        raise AlignmentError("經文 API 回傳格式不是 record 陣列")
    return records


def download_audio(bid: int, chap: int, out_path: Path, timeout: float = 60.0) -> None:
    """下載 FHL 音檔；官方檔名優先使用補零章號，並相容舊檔名。"""
    base = f"https://media.fhl.net/Taiwanese/{bid}"
    padded = f"{base}/{bid}_{chap:03d}.mp3"
    legacy = f"{base}/{bid}_{chap}.mp3"
    urls = [padded] if padded == legacy else [padded, legacy]

    last_error: Optional[Exception] = None
    for url in urls:
        try:
            payload = _read_url(url, timeout)
            if not payload:
                raise AlignmentError(f"下載到空音檔：{url}")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_bytes(payload)
            return
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, OSError, AlignmentError) as exc:
            last_error = exc
            if isinstance(exc, urllib.error.HTTPError) and exc.code != 404:
                break
    raise AlignmentError(f"無法下載書卷 {bid} 第 {chap} 章音檔") from last_error

def _chapter_title(bid: int, chap: int) -> str:
    title = BOOK_TITLES.get(bid, FHL_BOOK_NAMES.get(bid, str(bid)))
    suffix = "篇" if bid == 19 else "章"
    return f"{title}第{chap}{suffix}"


def _write_json_atomic(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json.tmp",
            prefix=f".{path.stem}.",
            dir=path.parent,
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        temp_path.replace(path)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def align_chapter(
    bid: int,
    chap: int,
    transcriber: Callable[[Any], list[TranscriptSpan]],
    out_dir: Path | str = Path(__file__).resolve().parents[1] / "timestamps",
    *,
    audio_path: Optional[Path | str] = None,
    audio_version: str = "1",
    strict_tail: bool = False,
    timeout: float = 60.0,
    min_verse_seconds: float = 1.5,
    max_verse_seconds: float = 35.0,
) -> Path:
    print(f"\n[開始處理] 書卷 {bid} 第 {chap} 章")
    records = get_chapter_text(bid, chap, timeout=timeout)
    prepare_verse_specs(records)

    output_path = Path(out_dir) / f"{bid}_{chap}.json"
    temporary_audio: Optional[Path] = None
    source_audio = Path(audio_path) if audio_path else None
    try:
        if source_audio is None:
            with tempfile.NamedTemporaryFile(
                suffix=".mp3",
                prefix=f"taiwanese_{bid}_{chap}_",
                delete=False,
            ) as handle:
                temporary_audio = Path(handle.name)
            source_audio = temporary_audio
            print("-> 下載台語音檔中...")
            download_audio(bid, chap, source_audio, timeout=timeout)
        elif not source_audio.exists():
            raise AlignmentError(f"指定的音檔不存在：{source_audio}")

        print("-> 解碼音訊並執行 ASR...")
        audio, duration = decode_audio(source_audio)
        spans = transcriber(audio)
        print(f"-> ASR 完成，共取得 {len(spans)} 個時間片段，音訊 {duration:.2f}s")
        result = align_verse_records(
            records,
            spans,
            duration,
            strict_tail=strict_tail,
            min_verse_seconds=min_verse_seconds,
            max_verse_seconds=max_verse_seconds,
        )
        for warning in result.warnings:
            print(f"[警告] {warning}")

        output = {
            "bid": bid,
            "chap": chap,
            "title": _chapter_title(bid, chap),
            "audio_version": audio_version,
            "total_duration": round(duration, 2),
            "verses": result.verses,
        }
        _write_json_atomic(output_path, output)
        print(f"[完成] 已生成 {output_path}（共 {len(result.verses)} 節）")
        return output_path
    finally:
        if temporary_audio and temporary_audio.exists():
            temporary_audio.unlink()


def build_parser() -> argparse.ArgumentParser:
    default_output = Path(__file__).resolve().parents[1] / "timestamps"
    parser = argparse.ArgumentParser(description="台語有聲聖經時間軸自動校正工具")
    parser.add_argument("--bid", type=int, default=19, help="書卷編號 1-66，預設 19（詩篇）")
    parser.add_argument("--chap", "--chapter", dest="chapters", type=int, nargs="+", default=[23], help="一個或多個章數")
    parser.add_argument("--backend", choices=("faster-whisper", "breeze"), default="faster-whisper")
    parser.add_argument("--model", default=None, help="ASR 模型名稱；未指定時依後端使用 small 或 Breeze-ASR-26")
    parser.add_argument("--device", default="cpu", help="faster-whisper 裝置，預設 cpu")
    parser.add_argument("--compute-type", default="int8", help="faster-whisper 計算型態，預設 int8")
    parser.add_argument("--language", default="zh", help="ASR 語言提示，預設 zh")
    parser.add_argument("--output-dir", type=Path, default=default_output)
    parser.add_argument("--audio", type=Path, default=None, help="使用本機音檔；批次處理時只能指定一章")
    parser.add_argument("--audio-version", default="1")
    parser.add_argument("--strict-tail", action="store_true", help="找不到尾部三字時直接停止")
    parser.add_argument("--timeout", type=float, default=60.0)
    parser.add_argument("--min-verse-seconds", type=float, default=1.5)
    parser.add_argument("--max-verse-seconds", type=float, default=35.0)
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.audio and len(args.chapters) != 1:
        print("[錯誤] --audio 只能搭配單一章數", file=sys.stderr)
        return 2

    model_name = args.model or (
        "small" if args.backend == "faster-whisper" else "MediaTek-Research/Breeze-ASR-26"
    )
    if args.backend == "faster-whisper":
        transcriber = make_faster_whisper_transcriber(
            model_name=model_name,
            device=args.device,
            compute_type=args.compute_type,
            language=args.language,
        )
    else:
        transcriber = make_breeze_transcriber(model_name=model_name)

    failures = 0
    for chap in args.chapters:
        try:
            align_chapter(
                args.bid,
                chap,
                transcriber,
                args.output_dir,
                audio_path=args.audio,
                audio_version=args.audio_version,
                strict_tail=args.strict_tail,
                timeout=args.timeout,
                min_verse_seconds=args.min_verse_seconds,
                max_verse_seconds=args.max_verse_seconds,
            )
        except (AlignmentError, OSError, urllib.error.URLError) as exc:
            failures += 1
            print(f"[錯誤] 第 {chap} 章處理失敗：{exc}", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
