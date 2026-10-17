import unittest

import tkinter as tk
from tkinter import ttk

from markinfo.gui import OLD_FG, ROW_BASE, ROW_STRIPE, STRIPE_TAG, restripe_tree


class RowBandTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self.tree = ttk.Treeview(self.root, columns=("name", "level"), show="headings")
        self.tree.tag_configure(STRIPE_TAG, background=ROW_STRIPE)
        self.tree.tag_configure("old", foreground=OLD_FG)

    def tearDown(self) -> None:
        self.root.destroy()

    def test_bands_follow_visual_order_and_keep_other_tags(self) -> None:
        self.tree.insert("", "end", iid="a", values=("ADA", "1"), tags=("old",))
        self.tree.insert("", "end", iid="b", values=("BO", "2"))
        self.tree.insert("", "end", iid="c", values=("CY", "3"), tags=("old",))
        self.tree.insert("", "end", iid="d", values=("DEE", "4"), tags=(STRIPE_TAG,))

        restripe_tree(self.tree)
        self.assertEqual(_tags(self.tree, "a"), ("old",))
        self.assertEqual(_tags(self.tree, "b"), (STRIPE_TAG,))
        self.assertEqual(_tags(self.tree, "c"), ("old",))
        self.assertEqual(_tags(self.tree, "d"), (STRIPE_TAG,))

        self.tree.move("d", "", 0)
        self.tree.move("c", "", 1)
        restripe_tree(self.tree)
        self.assertEqual(_tags(self.tree, "d"), ())
        self.assertEqual(_tags(self.tree, "c"), ("old", STRIPE_TAG))
        self.assertEqual(_tags(self.tree, "a"), ("old",))
        self.assertEqual(_tags(self.tree, "b"), (STRIPE_TAG,))

    def test_stripe_tag_is_a_full_row_background(self) -> None:
        painted = self.tree.tk.call(self.tree._w, "tag", "configure", STRIPE_TAG, "-background")
        self.assertEqual(str(painted).lower(), ROW_STRIPE)
        self.assertNotEqual(ROW_STRIPE, ROW_BASE)

    def test_previous_season_text_is_dark_enough_for_both_bands(self) -> None:
        for background in (ROW_BASE, ROW_STRIPE):
            self.assertGreater(_contrast(OLD_FG, background), 4.5)


def _tags(tree: ttk.Treeview, item: str) -> tuple[str, ...]:
    raw = tree.item(item, "tags")
    if not raw:
        return ()
    if isinstance(raw, str):
        return tuple(raw.split())
    return tuple(str(tag) for tag in raw)


def _channel(value: int) -> float:
    channel = value / 255
    if channel <= 0.04045:
        return channel / 12.92
    return ((channel + 0.055) / 1.055) ** 2.4


def _luminance(hex_color: str) -> float:
    red = int(hex_color[1:3], 16)
    green = int(hex_color[3:5], 16)
    blue = int(hex_color[5:7], 16)
    return 0.2126 * _channel(red) + 0.7152 * _channel(green) + 0.0722 * _channel(blue)


def _contrast(foreground: str, background: str) -> float:
    lighter = max(_luminance(foreground), _luminance(background))
    darker = min(_luminance(foreground), _luminance(background))
    return (lighter + 0.05) / (darker + 0.05)
