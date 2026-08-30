import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import TranscriptSpan, align_verse_records  # noqa: E402


class BreezeAlignerIntroTests(unittest.TestCase):
    def test_fallback_preserves_leading_spoken_intro_before_first_verse(self):
        records = [
            {"sec": 1, "bible_text": "甲甲甲乙乙乙。"},
            {"sec": 2, "bible_text": "丙丙丙丁丁丁。"},
        ]
        transcript = [
            TranscriptSpan("第一集", 0.0, 2.0),
            TranscriptSpan("戊戊戊己己己", 2.0, 20.0),
        ]

        result = align_verse_records(records, transcript, duration=20.0)

        self.assertEqual(result.verses[0]["start"], 2.0)
        self.assertEqual(result.verses[-1]["end"], 20.0)


if __name__ == "__main__":
    unittest.main()
