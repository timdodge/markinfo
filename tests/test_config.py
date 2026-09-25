import tempfile
import unittest
from pathlib import Path

from markinfo.config import (
    AppConfig,
    load_config_file,
    normalize_ui_scale,
    save_config,
    step_ui_scale,
)


class UiScaleTests(unittest.TestCase):
    def test_normalize_clamps(self) -> None:
        self.assertEqual(normalize_ui_scale(1.5), 1.5)
        self.assertEqual(normalize_ui_scale(0.1), 1.0)
        self.assertEqual(normalize_ui_scale(9), 2.5)
        self.assertEqual(normalize_ui_scale("nope"), 1.0)

    def test_step_up_and_down(self) -> None:
        self.assertEqual(step_ui_scale(1.0, 1), 1.25)
        self.assertEqual(step_ui_scale(1.25, 1), 1.5)
        self.assertEqual(step_ui_scale(2.5, 1), 2.5)
        self.assertEqual(step_ui_scale(1.0, -1), 1.0)
        self.assertEqual(step_ui_scale(1.5, -1), 1.25)
        self.assertEqual(step_ui_scale(1.3, 1), 1.5)
        self.assertEqual(step_ui_scale(1.3, -1), 1.25)

    def test_roundtrip_in_config_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "markinfo.cfg"
            path.write_text("[database]\ntype = sqlite\npath = keep.db\n", encoding="utf-8")
            cfg = AppConfig(ui_scale=1.5, source=path)
            save_config(cfg)
            loaded = load_config_file(path)
            self.assertAlmostEqual(loaded.ui_scale, 1.5)
            self.assertEqual(loaded.database.path, "keep.db")


if __name__ == "__main__":
    unittest.main()
