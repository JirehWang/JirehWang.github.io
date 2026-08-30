import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import TranscriptSpan, align_verse_records  # noqa: E402


class BreezeAlignerNearHeadRegressionTests(unittest.TestCase):
    def test_variant_head_is_anchored_at_the_first_asr_word(self):
        records = [
            {"sec": 1, "bible_text": "甲" * 30 + "。"},
            {"sec": 2, "bible_text": "彼位的金是真好亦出真珠碧玉。"},
        ]
        transcript = [
            TranscriptSpan("亂" * 30, 0.0, 12.0),
            TranscriptSpan("那", 12.0, 12.4),
            TranscriptSpan("裏", 12.4, 12.8),
            TranscriptSpan("的", 12.8, 13.2),
            TranscriptSpan("金是真好也出真珠碧玉", 13.2, 20.0),
        ]

        result = align_verse_records(records, transcript, duration=20.0)

        self.assertEqual(result.verses[1]["start"], 12.0)


if __name__ == "__main__":
    unittest.main()
