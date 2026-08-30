#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""台語有聲聖經時間軸批次校正入口。

核心 ASR／音訊工具保存在 ``breeze_aligner_engine.py``；本入口補上兩件
實際批次需要的行為：

* faster-whisper 使用快速的 segment timestamp 設定，不逐字解碼。
* 台語 ASR 與 FHL 漢字差異太大、無法找到頭錨點時，使用全文序列比對；
  若仍沒有足夠共同字，退回經文長度加音訊長度的可審核估算，並印出警告。

精準錨點成功時仍完全沿用核心引擎的頭/尾三字與完整性驗證。
"""

from __future__ import annotations

import math
import os
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Sequence

import breeze_aligner_engine as _engine
from breeze_aligner_engine import (
    AlignmentError,
    AlignmentResult,
    AnchorMatch,
    TranscriptSpan,
    TimestampValidationError,
    build_parser,
    build_transcript_timeline,
    clean_text,
    decode_audio,
    download_audio,
    get_chapter_text,
    head_tail,
    normalize_audio_text,
    normalize_for_matching,
    prepare_verse_specs,
    validate_timestamps,
)


def _weighted_starts(specs: Sequence[Any], duration: float, intro_offset: float = 0.0) -> list[float]:
    weights = [max(1, len(spec.match_text)) for spec in specs]
    total_weight = float(sum(weights))
    available = max(0.0, duration - intro_offset)
    starts: list[float] = []
    consumed = 0.0
    for weight in weights:
        starts.append(intro_offset + available * consumed / total_weight)
        consumed += weight
    return starts


def _timeline_time_at(timeline: Any, position: float) -> float:
    if position <= 0:
        return float(timeline.starts[0])
    if position >= len(timeline.text):
        return float(timeline.ends[-1])
    index = int(position)
    fraction = position - index
    return float(timeline.starts[index] * (1.0 - fraction) + timeline.ends[index] * fraction)


def _map_source_offset(offset: float, blocks: Sequence[Any]) -> float:
    """把預期經文字串的位置映射到 ASR 字串的位置。"""
    first = blocks[0]
    if offset <= first.a:
        return float(first.b)

    for index, block in enumerate(blocks):
        block_end_a = block.a + block.size
        block_end_b = block.b + block.size
        if offset <= block_end_a:
            return float(block.b + max(0.0, offset - block.a))
        if index + 1 >= len(blocks):
            return float(block_end_b)

        next_block = blocks[index + 1]
        if offset < next_block.a:
            source_gap = next_block.a - block_end_a
            transcript_gap = next_block.b - block_end_b
            if source_gap <= 0:
                return float(block_end_b)
            ratio = (offset - block_end_a) / source_gap
            return float(block_end_b + ratio * transcript_gap)
    return float(blocks[-1].b + blocks[-1].size)


_ASR_VARIANT_GROUPS = (
    "耶雅也",
    "華花化",
    "彼那",
    "位裏裡",
    "講說",
    "頭到",
    "二兩",
    "無沒",
    "見介",
    "誚紹",
)
_ASR_VARIANT_CANON = {
    character: group[0]
    for group in _ASR_VARIANT_GROUPS
    for character in group
}
_SPOKEN_MARKER_RE = re.compile(
    rf"^第[{_engine.CHINESE_NUMERALS}0-9]+(?:節|集|句)+"
)
_LEADING_SPOKEN_PREFIX_RE = re.compile(
    rf"^(?:第[{_engine.CHINESE_NUMERALS}0-9]+(?:節|集|句)?)+"
)


def _anchor_key(text: str) -> str:
    return "".join(_ASR_VARIANT_CANON.get(character, character) for character in str(text or ""))


def _anchor_similarity(expected: str, candidate: str) -> float:
    if not expected or not candidate:
        return 0.0
    return SequenceMatcher(None, _anchor_key(expected), _anchor_key(candidate)).ratio()


def _raw_span_char_timings(spans: Sequence[TranscriptSpan]) -> list[tuple[str, float, float]]:
    timed_chars: list[tuple[str, float, float]] = []
    for span in sorted(spans, key=lambda item: (float(item.start), float(item.end))):
        try:
            span_start = max(0.0, float(span.start))
            span_end = max(span_start, float(span.end))
        except (TypeError, ValueError):
            continue
        raw = normalize_for_matching(span.text)
        if not raw:
            continue
        span_duration = span_end - span_start
        for index, character in enumerate(raw):
            start = span_start + span_duration * index / len(raw)
            end = span_start + span_duration * (index + 1) / len(raw)
            timed_chars.append((character, start, end))
    return timed_chars


def _leading_content_start(spans: Sequence[TranscriptSpan]) -> float:
    """找出章首連續口播節號後，第一個正文 ASR 字元的時間。"""
    timed_chars = _raw_span_char_timings(spans)
    if not timed_chars:
        return 0.0
    raw = "".join(character for character, _, _ in timed_chars)
    match = _LEADING_SPOKEN_PREFIX_RE.match(raw)
    if not match:
        return float(timed_chars[0][1])
    prefix = match.group(0)
    # 單獨的「第七日」可能是正文，不能把只有一個「第」的片段當成章首口播。
    if prefix.count("第") < 2 and not re.search(r"[節集句]", prefix):
        return float(timed_chars[0][1])
    position = match.end()
    if position < len(timed_chars):
        return float(timed_chars[position][1])
    return float(timed_chars[-1][2])


def _spoken_marker_events(
    spans: Sequence[TranscriptSpan],
) -> list[tuple[float, float, str, str]]:
    """辨識跨逐字 span 的「第○節」候選，並保留其後文字供語境判斷。"""
    ordered: list[tuple[float, float, str, list[tuple[str, float, float]]]] = []
    for span in sorted(spans, key=lambda item: (float(item.start), float(item.end))):
        try:
            span_start = max(0.0, float(span.start))
            span_end = max(span_start, float(span.end))
        except (TypeError, ValueError):
            continue
        raw = normalize_for_matching(span.text)
        if not raw:
            continue
        duration = span_end - span_start
        timings = [
            (
                character,
                span_start + duration * index / len(raw),
                span_start + duration * (index + 1) / len(raw),
            )
            for index, character in enumerate(raw)
        ]
        ordered.append((span_start, span_end, raw, timings))

    events: list[tuple[float, float, str, str]] = []
    for index, (span_start, _, raw, timings) in enumerate(ordered):
        if not raw.startswith("第"):
            continue
        combined: list[tuple[str, float, float]] = []
        previous_end: Optional[float] = None
        for cursor in range(index, len(ordered)):
            current_start, current_end, current_raw, current_timings = ordered[cursor]
            if previous_end is not None and current_start - previous_end > 0.45:
                break
            combined.extend(current_timings)
            candidate = "".join(character for character, _, _ in combined)
            marker = _SPOKEN_MARKER_RE.match(candidate)
            if marker:
                marker_end = marker.end()
                following = candidate[marker_end:]
                next_cursor = cursor + 1
                while next_cursor < len(ordered) and len(following) < 12:
                    next_start, _, next_raw, _ = ordered[next_cursor]
                    if next_start - current_end > 0.8:
                        break
                    following += next_raw
                    current_end = ordered[next_cursor][1]
                    next_cursor += 1
                event_start = float(combined[0][1])
                event_end = float(combined[marker_end - 1][2])
                event = (event_start, event_end, candidate[:marker_end], following)
                if not any(abs(existing[0] - event_start) < 0.01 for existing in events):
                    events.append(event)
                break
            previous_end = current_end
    return events


def _marker_is_content(event: tuple[float, float, str, str], spec: Any) -> bool:
    _, _, marker, following = event
    marker_core = re.sub(r"[節集句]+$", "", marker)
    source_prefix = spec.match_text[: max(12, len(spec.head) + 5)]
    if marker_core and marker_core in source_prefix:
        return True
    continuation = spec.match_text[len(spec.head):]
    if continuation and _anchor_similarity(continuation[:6], following[:6]) >= 0.75:
        return True
    return False


def _marker_after_time(
    spec: Any,
    target: float,
    spans: Sequence[TranscriptSpan],
) -> Optional[float]:
    possible = [
        event
        for event in _spoken_marker_events(spans)
        if not _marker_is_content(event, spec)
        and (event[0] <= target <= event[1] or abs(event[0] - target) <= 6.0)
    ]
    if not possible:
        return None
    event = min(possible, key=lambda item: abs(item[0] - target))
    return float(event[1])


def _sequence_starts(
    specs: Sequence[Any],
    timeline: Any,
    fallback_starts: Optional[Sequence[float]] = None,
) -> Optional[list[float]]:
    expected = "".join(spec.match_text for spec in specs)
    matcher = SequenceMatcher(None, expected, timeline.text, autojunk=False)
    blocks = [block for block in matcher.get_matching_blocks() if block.size >= 2]
    matched_chars = sum(block.size for block in blocks)
    if matched_chars < max(6, int(len(expected) * 0.02)):
        return None

    boundaries: list[int] = []
    consumed = 0
    for spec in specs:
        boundaries.append(consumed)
        consumed += len(spec.match_text)

    def has_nearby_evidence(boundary: int) -> bool:
        return any(
            block.a - 12 <= boundary <= block.a + block.size + 12
            for block in blocks
        )

    starts: list[float] = []
    for index, boundary in enumerate(boundaries):
        if fallback_starts is not None and index == 0:
            starts.append(float(fallback_starts[index]))
        elif fallback_starts is not None and not has_nearby_evidence(boundary):
            # 若全文比對在尾端已經沒有證據，不能把未辨識的最後節映射回舊文字。
            starts.append(float(fallback_starts[index]))
        else:
            starts.append(_timeline_time_at(timeline, _map_source_offset(boundary, blocks)))
    return starts


def _transcript_segment_candidates(spans: Sequence[TranscriptSpan]) -> list[tuple[float, str]]:
    candidates: list[tuple[float, str]] = []
    for span in sorted(spans, key=lambda item: (float(item.start), float(item.end))):
        try:
            span_start = max(0.0, float(span.start))
            span_end = max(span_start, float(span.end))
        except (TypeError, ValueError):
            continue
        text, prefix_chars, original_han_count = _engine._normalize_audio_with_prefix(span.text)
        if not text:
            continue
        span_duration = span_end - span_start
        prefix_ratio = prefix_chars / original_han_count if original_han_count else 0.0
        content_start = span_start + span_duration * prefix_ratio
        candidates.append((content_start, text))
    return candidates


def _asr_span_start_near(
    spans: Sequence[TranscriptSpan],
    target: float,
) -> Optional[float]:
    ordered = sorted(spans, key=lambda item: (float(item.start), float(item.end)))
    containing: list[float] = []
    preceding: list[float] = []
    for span in ordered:
        try:
            span_start = max(0.0, float(span.start))
            span_end = max(span_start, float(span.end))
        except (TypeError, ValueError):
            continue
        if span_start <= target <= span_end:
            containing.append(span_start)
        if span_start <= target:
            preceding.append(span_start)
    if containing:
        return max(containing)
    if preceding and target - max(preceding) <= 1.5:
        return max(preceding)
    return None


def _snap_to_partial_heads(
    specs: Sequence[Any],
    spans: Sequence[TranscriptSpan],
    timeline: Any,
    base_starts: Sequence[float],
) -> tuple[list[float], int]:
    candidates = _transcript_segment_candidates(spans)
    marker_events = _spoken_marker_events(spans)
    starts = list(base_starts)
    snapped = 0
    for index, spec in enumerate(specs):
        target = float(base_starts[index])
        nearby: list[tuple[float, float, str]] = []

        def add_candidate(candidate_start: float, candidate_text: str) -> None:
            if abs(candidate_start - target) > 5.0:
                return
            score = _anchor_similarity(spec.head, candidate_text[: len(spec.head)])
            if score >= 2 / 3:
                nearby.append((score, candidate_start, candidate_text))

        for candidate_start, text in candidates:
            add_candidate(candidate_start, text)
        if timeline:
            for position, candidate_start in enumerate(timeline.starts):
                if abs(float(candidate_start) - target) <= 5.0:
                    add_candidate(
                        float(candidate_start),
                        timeline.text[position : position + len(spec.head)],
                    )

        candidate_start: Optional[float] = None
        if nearby:
            _, candidate_start, _ = max(
                nearby,
                key=lambda item: (item[0], -abs(item[1] - target)),
            )
        elif index > 0:
            marker_start = _marker_after_time(spec, target, spans)
            if marker_start is not None:
                candidate_start = marker_start
        elif index == len(specs) - 1:
            candidate_start = _asr_span_start_near(spans, target)

        if candidate_start is None:
            continue
        lower_bound = starts[index - 1] + 1.5 if index else 0.0
        upper_bound = base_starts[index + 1] - 1.5 if index + 1 < len(base_starts) else math.inf
        if candidate_start <= lower_bound or candidate_start >= upper_bound:
            continue
        if abs(candidate_start - starts[index]) > 0.001:
            starts[index] = round(candidate_start, 2)
            snapped += 1
    return starts, snapped

def _build_fallback_result(
    records: Sequence[dict[str, Any]],
    transcript_spans: Sequence[TranscriptSpan],
    duration: float,
    *,
    strict_tail: bool,
    min_verse_seconds: float,
    max_verse_seconds: float,
    reason: str,
) -> AlignmentResult:
    if strict_tail:
        raise AlignmentError(reason)

    specs = prepare_verse_specs(records)
    total_duration = round(float(duration), 2)
    timeline = None
    try:
        timeline = build_transcript_timeline(transcript_spans)
    except AlignmentError:
        pass

    intro_offset = _leading_content_start(transcript_spans)
    weighted_starts = _weighted_starts(specs, total_duration, intro_offset=intro_offset)
    starts: Optional[list[float]] = (
        _sequence_starts(specs, timeline, fallback_starts=weighted_starts)
        if timeline
        else None
    )
    starts_were_sequence = starts is not None
    method = "全文序列比對"
    if starts is None:
        method = "經文長度加音訊長度估算"
        starts = weighted_starts
    starts, snapped_count = _snap_to_partial_heads(specs, transcript_spans, timeline, starts)
    if snapped_count:
        method += f"＋局部 ASR 頭部錨點 {snapped_count} 個"
    rounded_starts = [round(float(start), 2) for start in starts]
    ends = rounded_starts[1:] + [total_duration]
    verses = [
        {
            "sec": spec.sec,
            "start": rounded_starts[index],
            "end": ends[index],
            "text": spec.display_text,
        }
        for index, spec in enumerate(specs)
    ]

    try:
        validate_timestamps(
            verses,
            total_duration,
            min_verse_seconds=min_verse_seconds,
            max_verse_seconds=max_verse_seconds,
        )
    except TimestampValidationError:
        if starts_were_sequence:
            method = "經文長度加音訊長度估算"
            starts = _weighted_starts(specs, total_duration, intro_offset=intro_offset)
            starts, snapped_count = _snap_to_partial_heads(specs, transcript_spans, timeline, starts)
            if snapped_count:
                method += f"＋局部 ASR 頭部錨點 {snapped_count} 個"
            rounded_starts = [round(value, 2) for value in starts]
            ends = rounded_starts[1:] + [total_duration]
            verses = [
                {
                    "sec": spec.sec,
                    "start": rounded_starts[index],
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
        else:
            raise

    warning = (
        f"{reason}；已使用{method} fallback，未取得完整頭/尾錨點，"
        "產出時間軸需人工抽查"
    )
    return AlignmentResult(
        verses=verses,
        warnings=[warning],
        head_matches=[],
        tail_matches=[None] * len(specs),
    )


_BASE_ALIGN_VERSE_RECORDS = _engine.align_verse_records


def align_verse_records(
    records: Iterable[dict[str, Any]],
    transcript_spans: Iterable[TranscriptSpan],
    duration: float,
    *,
    strict_tail: bool = False,
    min_verse_seconds: float = 1.5,
    max_verse_seconds: float = 35.0,
) -> AlignmentResult:
    records_list = list(records)
    spans_list = list(transcript_spans)
    try:
        return _BASE_ALIGN_VERSE_RECORDS(
            records_list,
            spans_list,
            duration,
            strict_tail=strict_tail,
            min_verse_seconds=min_verse_seconds,
            max_verse_seconds=max_verse_seconds,
        )
    except AlignmentError as exc:
        return _build_fallback_result(
            records_list,
            spans_list,
            duration,
            strict_tail=strict_tail,
            min_verse_seconds=min_verse_seconds,
            max_verse_seconds=max_verse_seconds,
            reason=str(exc),
        )


def make_faster_whisper_transcriber(
    model_name: str = "small",
    device: str = "cpu",
    compute_type: str = "int8",
    language: str = "zh",
) -> Callable[[Any], list[TranscriptSpan]]:
    """faster-whisper 精準模式：保留逐字 timestamp，供節點落在實際開口字。"""
    try:
        from faster_whisper import WhisperModel
    except ImportError as exc:
        raise AlignmentError(
            "缺少 faster-whisper，請先執行：pip install faster-whisper miniaudio numpy scipy"
        ) from exc

    model = WhisperModel(
        model_name,
        device=device,
        compute_type=compute_type,
        cpu_threads=max(1, min(8, os.cpu_count() or 4)),
    )

    def transcribe(audio: Any) -> list[TranscriptSpan]:
        segments, _ = model.transcribe(
            audio,
            language=language,
            beam_size=1,
            best_of=1,
            word_timestamps=True,
            vad_filter=True,
            condition_on_previous_text=False,
        )
        spans: list[TranscriptSpan] = []
        for segment in segments:
            words = getattr(segment, "words", None)
            if words:
                for word in words:
                    pair = _engine._timestamp_pair(
                        (getattr(word, "start", None), getattr(word, "end", None))
                    )
                    word_text = getattr(word, "word", "")
                    if pair and word_text.strip():
                        spans.append(TranscriptSpan(word_text, pair[0], pair[1]))
                continue
            pair = _engine._timestamp_pair(
                (getattr(segment, "start", None), getattr(segment, "end", None))
            )
            if pair and getattr(segment, "text", "").strip():
                spans.append(TranscriptSpan(segment.text, pair[0], pair[1]))
        return spans

    return transcribe

_BASE_ALIGN_CHAPTER = _engine.align_chapter


def align_chapter(*args: Any, **kwargs: Any) -> Path:
    saved_align = _engine.align_verse_records
    _engine.align_verse_records = align_verse_records
    try:
        return _BASE_ALIGN_CHAPTER(*args, **kwargs)
    finally:
        _engine.align_verse_records = saved_align


def main(argv: Optional[Sequence[str]] = None) -> int:
    saved = {
        "align_verse_records": _engine.align_verse_records,
        "align_chapter": _engine.align_chapter,
        "make_faster_whisper_transcriber": _engine.make_faster_whisper_transcriber,
    }
    _engine.align_verse_records = align_verse_records
    _engine.align_chapter = align_chapter
    _engine.make_faster_whisper_transcriber = make_faster_whisper_transcriber
    try:
        return _engine.main(argv)
    finally:
        for name, value in saved.items():
            setattr(_engine, name, value)


if __name__ == "__main__":
    raise SystemExit(main())
