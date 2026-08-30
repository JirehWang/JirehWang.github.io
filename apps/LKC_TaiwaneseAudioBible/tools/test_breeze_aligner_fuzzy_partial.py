import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import TranscriptSpan, align_verse_records  # noqa: E402


class BreezeAlignerFuzzyPartialTests(unittest.TestCase):
    def test_fallback_accepts_a_nearby_head_with_one_asr_character_error(self):
        records = [
            {"sec": 1, "bible_text": "甲甲甲乙乙乙。"},
            {"sec": 2, "bible_text": "總是有霧對地起。"},
        ]
        transcript = [
            TranscriptSpan("第一集", 0.0, 1.0),
            TranscriptSpan("戊戊戊己己己", 1.0, 10.0),
            TranscriptSpan("總是無辜對待去", 10.0, 20.0),
        ]

        result = align_verse_records(records, transcript, duration=20.0)

        self.assertEqual(result.verses[1]["start"], 10.0)


if __name__ == "__main__":
    unittest.main()
