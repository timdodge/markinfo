import tempfile
import unittest
import zipfile
from pathlib import Path

from markinfo.database import connect_sqlite
from markinfo.models import Player, notes_indicator
from markinfo.parser import process_input_file

FIXTURES = Path(__file__).resolve().parent / "fixtures"


class DatabaseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.db = connect_sqlite(Path(self.tmp.name) / "test.db")
        self.db.session = 5

    def tearDown(self) -> None:
        self.db.close()
        self.tmp.cleanup()

    def _player(self, **kwargs) -> Player:
        data = dict(
            league="PREMIER",
            team="ARSENAL",
            name="SMITH",
            age=24,
            level=8,
            cat="S",
            pos="DF",
            type="---",
            sess=0,
            notes="",
        )
        data.update(kwargs)
        return Player(**data)

    def test_schema_and_lookups(self) -> None:
        self.assertTrue(self.db.table_exists("Players"))
        self.assertIn("GK", self.db.get_posns())
        self.assertIn("P/S", self.db.get_cats())
        self.assertIn("STAR", self.db.get_types())
        self.assertEqual(self.db.session, 5)

    def test_league_team_player_crud(self) -> None:
        self.db.add_league("premier")
        self.assertEqual(self.db.get_leagues(), ["PREMIER"])
        self.db.add_team("arsenal", "PREMIER")
        self.assertEqual(self.db.get_teams("PREMIER"), ["ARSENAL"])
        key, created = self.db.add_player(self._player())
        self.assertTrue(created)
        self.assertGreater(key, 0)
        squad = self.db.get_squad("ARSENAL", "PREMIER")
        self.assertEqual(len(squad), 1)
        self.assertEqual(squad[0].name, "SMITH")
        self.assertEqual(squad[0].sess, 5)

        self.db.rename_team("ARSENAL", "THE GUNNERS", "PREMIER")
        self.assertEqual(self.db.get_teams("PREMIER"), ["THE GUNNERS"])
        self.db.rename_league("PREMIER", "ENGLAND")
        self.assertEqual(self.db.get_leagues(), ["ENGLAND"])

        self.db.del_player("SMITH", "THE GUNNERS", "ENGLAND")
        self.assertEqual(len(self.db.get_squad("THE GUNNERS", "ENGLAND")), 0)
        self.db.del_team("THE GUNNERS", "ENGLAND")
        self.assertEqual(self.db.get_teams("ENGLAND"), [])
        self.db.del_league("ENGLAND")
        self.assertEqual(self.db.get_leagues(), [])

    def test_add_player_updates_newer_session_only(self) -> None:
        self.db.add_player(self._player(level=8, notes="old"))
        key, created = self.db.add_player(self._player(level=9, notes="new"))
        self.assertFalse(created)
        player = self.db.get_squad("ARSENAL", "PREMIER")[0]
        self.assertEqual(player.level, 9)
        self.assertEqual(player.notes, "new")

        self.db.session = 4
        self.db.add_player(self._player(level=1, notes="stale"))
        player = self.db.get_squad("ARSENAL", "PREMIER")[0]
        self.assertEqual(player.level, 9)
        self.assertEqual(player.notes, "new")

    def test_unknown_cat_and_pos_preserved(self) -> None:
        self.db.add_player(self._player(cat="S", pos="DF"))
        self.db.add_player(self._player(cat="?", pos="?"))
        player = self.db.get_squad("ARSENAL", "PREMIER")[0]
        self.assertEqual(player.cat, "S")
        self.assertEqual(player.pos, "DF")

    def test_keeper_forces_no_cat(self) -> None:
        self.db.add_player(self._player(pos="GK", cat="S"))
        player = self.db.get_squad("ARSENAL", "PREMIER")[0]
        self.assertEqual(player.cat, "---")

    def test_move_player(self) -> None:
        self.db.add_player(self._player())
        moved = self._player(team="CHELSEA")
        self.db.move_player(moved, "ARSENAL")
        self.assertEqual(len(self.db.get_squad("ARSENAL", "PREMIER")), 0)
        self.assertEqual(len(self.db.get_squad("CHELSEA", "PREMIER")), 1)

    def test_new_season_prunes_and_ages(self) -> None:
        self.db.add_player(self._player(name="CURRENT"))
        self.db.add_player(self._player(name="OLD", sess=-1, team="ARSENAL"))
        self.db.add_player(self._player(name="ANCIENT", sess=-2, team="ARSENAL"))
        self.db.new_season("PREMIER", session_limit=2)
        names = {p.name: p.sess for p in self.db.get_squad("ARSENAL", "PREMIER")}
        self.assertEqual(names["CURRENT"], -1)
        self.assertEqual(names["OLD"], -2)
        self.assertNotIn("ANCIENT", names)

    def test_search(self) -> None:
        self.db.add_player(self._player(name="SMITH", pos="DF", age=24, level=8))
        self.db.add_player(self._player(name="JONES", team="CHELSEA", pos="GK", age=30, level=10))
        hits = self.db.search_players(name="smi")
        self.assertEqual([p.name for p in hits], ["SMITH"])
        hits = self.db.search_players(pos="GK")
        self.assertEqual([p.name for p in hits], ["JONES"])
        hits = self.db.search_players(min_age=25)
        self.assertEqual([p.name for p in hits], ["JONES"])

    def test_import_zip(self) -> None:
        zip_path = Path(self.tmp.name) / "sho.zip"
        with zipfile.ZipFile(zip_path, "w") as archive:
            archive.write(FIXTURES / "squad.txt", "squad.txt")
            archive.write(FIXTURES / "auction.txt", "nested/auction.txt")
        added, updated = process_input_file(
            zip_path, "PREMIER", self.db.add_player, self.db.move_player
        )
        self.assertGreater(added, 0)
        self.assertEqual(updated, 0)
        teams = set(self.db.get_teams("PREMIER"))
        self.assertIn("ARSENAL", teams)
        self.assertIn("MANCHESTER UNITED", teams)
        self.assertIn("CHELSEA", teams)

    def test_import_deals_moves_players(self) -> None:
        self.db.add_player(self._player(name="JOHN SMITH", team="ARSENAL"))
        added, updated = process_input_file(
            FIXTURES / "deals.txt",
            "PREMIER",
            self.db.add_player,
            self.db.move_player,
        )
        self.assertGreater(added + updated, 0)
        arsenal = [p.name for p in self.db.get_squad("ARSENAL", "PREMIER")]
        chelsea = [p.name for p in self.db.get_squad("CHELSEA", "PREMIER")]
        self.assertNotIn("JOHN SMITH", arsenal)
        self.assertIn("JOHN SMITH", chelsea)


class NotesIndicatorTests(unittest.TestCase):
    def test_short_numeric_kept(self) -> None:
        self.assertEqual(notes_indicator("12 injured"), "12")
        self.assertEqual(notes_indicator("i3 note"), "i3")
        self.assertEqual(notes_indicator("s2"), "s2")

    def test_long_notes_tick(self) -> None:
        self.assertEqual(notes_indicator("long note here"), "✓")
        self.assertEqual(notes_indicator(""), "")


if __name__ == "__main__":
    unittest.main()
