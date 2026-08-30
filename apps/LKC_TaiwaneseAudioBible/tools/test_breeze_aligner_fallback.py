import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import TranscriptSpan, align_verse_records  # noqa: E402


class BreezeAlignerFallbackTests(unittest.TestCase):
    def test_unmatched_asr_uses_auditable_weighted_fallback(self):
        records = [
            {"sec": 1, "bible_text": "甲甲甲乙乙乙。"},
            {"sec": 2, "bible_text": "丙丙丙丁丁丁。"},
        ]
        transcript = [TranscriptSpan("戊戊戊己己己", 0.0, 20.0)]

        result = align_verse_records(records, transcript, duration=20.0)

        self.assertEqual(len(result.verses), 2)
        self.assertEqual(result.verses[0]["end"], result.verses[1]["start"])
        self.assertEqual(result.verses[-1]["end"], 20.0)
        self.assertTrue(any("fallback" in warning for warning in result.warnings))


if __name__ == "__main__":
    unittest.main()
