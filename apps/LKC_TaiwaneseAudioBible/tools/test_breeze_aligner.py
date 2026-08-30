import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import (  # noqa: E402
    TranscriptSpan,
    TimestampValidationError,
    align_verse_records,
    normalize_audio_text,
    normalize_for_matching,
    validate_timestamps,
)


class BreezeAlignerTests(unittest.TestCase):
    def test_normalization_removes_annotations_and_keeps_han_characters(self):
        raw = "<b>耶和華是我的牧者</b>(kheh)，我無欠缺。〔大衛的詩〕"

        self.assertEqual(
            normalize_for_matching(raw),
            "耶和華是我的牧者我無欠缺",
        )

    def test_audio_normalization_removes_verse_number_and_hebrew_marker(self):
        raw = "第一節 阿勒弗，耶和華是我的牧者。"

        self.assertEqual(normalize_audio_text(raw), "耶和華是我的牧者")

    def test_alignment_uses_head_and_tail_anchors_and_next_start_boundaries(self):
        records = [
            {"sec": 1, "bible_text": "耶和華是我的牧者，我無欠缺。"},
            {"sec": 2, "bible_text": "伊互我倒佇青翠的草埔。"},
        ]
        transcript = [
            TranscriptSpan("第一節阿勒弗耶和華是我的牧者我無欠缺", 0.0, 5.0),
            TranscriptSpan("第二節伯特伊互我倒佇青翠的草埔", 5.0, 10.0),
        ]

        result = align_verse_records(records, transcript, duration=10.0)

        self.assertEqual(result.warnings, [])
        self.assertGreater(result.verses[0]["start"], 0.0)
        self.assertEqual(result.verses[0]["end"], result.verses[1]["start"])
        self.assertEqual(result.verses[-1]["end"], 10.0)

    def test_alignment_reports_unmatched_tail_without_corrupting_boundaries(self):
        records = [
            {"sec": 1, "bible_text": "耶和華是我的牧者，我無欠缺。"},
            {"sec": 2, "bible_text": "伊互我倒佇青翠的草埔。"},
        ]
        transcript = [
            TranscriptSpan("耶和華是我的牧者", 0.0, 4.0),
            TranscriptSpan("伊互我倒佇青翠的草埔", 4.0, 10.0),
        ]

        result = align_verse_records(records, transcript, duration=10.0)

        self.assertTrue(any("尾部" in warning for warning in result.warnings))
        self.assertEqual(result.verses[0]["end"], result.verses[1]["start"])

    def test_timestamp_validation_rejects_zero_length_or_overlong_verses(self):
        verses = [
            {"sec": 1, "start": 0.0, "end": 1.0, "text": "短"},
            {"sec": 2, "start": 1.0, "end": 40.0, "text": "長"},
        ]

        with self.assertRaises(TimestampValidationError):
            validate_timestamps(verses, total_duration=40.0)


if __name__ == "__main__":
    unittest.main()
