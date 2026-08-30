import sys
import unittest
from pathlib import Path


TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))

from breeze_aligner import normalize_audio_text  # noqa: E402


class BreezeAlignerMarkerTests(unittest.TestCase):
    def test_content_starting_with第七日_is_not_deleted_as_a_verse_marker(self):
        raw = "第七集 兄弟萬一所造的工"

        self.assertEqual(normalize_audio_text(raw), "第七集兄弟萬一所造的工")


if __name__ == "__main__":
    unittest.main()
