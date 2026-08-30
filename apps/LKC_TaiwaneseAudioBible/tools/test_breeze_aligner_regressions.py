import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import TranscriptSpan, align_verse_records  # noqa: E402


class BreezeAlignerRegressionTests(unittest.TestCase):
    def test_fallback_uses_contiguous_asr_text_for_a_near_head_variant(self):
        records = [
            {"sec": 1, "bible_text": "甲甲甲乙乙乙丙丙丙丁丁丁。"},
            {"sec": 2, "bible_text": "彼位的金是真好亦出真珠碧玉。"},
        ]
        transcript = [
            TranscriptSpan("甲甲甲乙乙乙丙丙丙丁丁丁", 0.0, 12.0),
            TranscriptSpan("那", 12.0, 12.4),
            TranscriptSpan("裏", 12.4, 12.8),
            TranscriptSpan("的", 12.8, 13.2),
            TranscriptSpan("金是真好也出真珠碧玉", 13.2, 20.0),
        ]

        result = align_verse_records(records, transcript, duration=20.0)

        self.assertEqual(result.verses[1]["start"], 12.0)

    def test_fallback_does_not_map_an_unmatched_last_verse_back_into_previous_text(self):
        records = [
            {"sec": 1, "bible_text": "甲甲甲乙乙乙丙丙丙丁丁丁。"},
            {"sec": 2, "bible_text": "戊戊戊己己己庚庚庚辛辛辛。"},
            {"sec": 3, "bible_text": "翁某二人平平褪腹裼也無見誚。"},
        ]
        transcript = [
            TranscriptSpan("甲甲甲乙乙乙丙丙丙丁丁丁", 0.0, 10.0),
            TranscriptSpan("戊戊戊己己己庚庚庚辛辛辛", 10.0, 20.0),
        ]

        result = align_verse_records(records, transcript, duration=30.0)

        self.assertGreaterEqual(result.verses[2]["start"], 20.0)


if __name__ == "__main__":
    unittest.main()
