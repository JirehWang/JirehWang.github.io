import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import TranscriptSpan, align_verse_records  # noqa: E402


class BreezeAlignerPartialAnchorTests(unittest.TestCase):
    def test_fallback_uses_a_later_exact_head_anchor_after_an_early_mismatch(self):
        records = [
            {"sec": 1, "bible_text": "甲甲甲乙乙乙。"},
            {"sec": 2, "bible_text": "總是有霧對地起。"},
        ]
        transcript = [
            TranscriptSpan("第一集", 0.0, 1.0),
            TranscriptSpan("戊戊戊己己己", 1.0, 10.0),
            TranscriptSpan("總是有部對著去", 10.0, 20.0),
        ]

        result = align_verse_records(records, transcript, duration=20.0)

        self.assertEqual(result.verses[1]["start"], 10.0)
        self.assertEqual(result.verses[0]["end"], result.verses[1]["start"])


if __name__ == "__main__":
    unittest.main()
