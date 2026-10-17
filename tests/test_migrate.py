import tempfile
import unittest
from pathlib import Path

from markinfo.app import main
from markinfo.database import connect_sqlite
from markinfo.migrate import MigrationError, migrate_jet_to_sqlite

MDB = Path(__file__).resolve().parents[1] / "vb6" / "MarkInfo21.mdb"


@unittest.skipUnless(MDB.is_file(), "VB6 Jet database is not in the tree")
class MigrateJetTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dest = Path(self.tmp.name) / "markinfo.db"

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_copies_keys_session_and_player(self) -> None:
        result = migrate_jet_to_sqlite(MDB, self.dest)
        self.assertEqual(result.leagues, 1)
        self.assertEqual(result.teams, 1)
        self.assertEqual(result.players, 1)
        self.assertEqual(result.session, 2)

        db = connect_sqlite(self.dest)
        try:
            self.assertEqual(db.session, 2)
            self.assertEqual(db.get_leagues(), ["WELLY"])
            self.assertEqual(db.get_teams("WELLY"), ["TEAMY"])
            squad = db.get_squad("TEAMY", "WELLY")
            self.assertEqual(len(squad), 1)
            player = squad[0]
            self.assertEqual(player.name, "DAVE")
            self.assertEqual(player.age, 17)
            self.assertEqual(player.level, 0)
            self.assertEqual(player.cat, "P")
            self.assertEqual(player.pos, "FWS")
            self.assertEqual(player.type, "---")
            self.assertEqual(player.sess, 1)
            self.assertEqual(player.notes, "test")
            self.assertIn("GK", db.get_posns())
            self.assertIn("P/S", db.get_cats())
            self.assertIn("STAR", db.get_types())

            league_key = db._fetchone("SELECT LeagueKey FROM Leagues")
            team_key = db._fetchone("SELECT TeamKey, LeagueKey FROM Teams")
            player_key = db._fetchone("SELECT PlayerKey, TeamKey FROM Players")
            self.assertEqual(int(league_key["LeagueKey"]), 1)
            self.assertEqual(int(team_key["TeamKey"]), 1)
            self.assertEqual(int(team_key["LeagueKey"]), 1)
            self.assertEqual(int(player_key["PlayerKey"]), 1)
            self.assertEqual(int(player_key["TeamKey"]), 1)
            self.assertEqual(db.add_league("OTHER"), 2)
        finally:
            db.close()

    def test_refuses_to_replace_an_existing_file(self) -> None:
        self.dest.write_text("keep", encoding="utf-8")
        with self.assertRaises(MigrationError):
            migrate_jet_to_sqlite(MDB, self.dest)
        self.assertEqual(self.dest.read_text(encoding="utf-8"), "keep")

    def test_cli_writes_the_sqlite_file(self) -> None:
        code = main(["--migrate-vb6", str(MDB), "--output", str(self.dest)])
        self.assertEqual(code, 0)
        self.assertTrue(self.dest.is_file())
        db = connect_sqlite(self.dest)
        try:
            self.assertEqual(db.get_squad("TEAMY", "WELLY")[0].name, "DAVE")
        finally:
            db.close()

    def test_cli_requires_both_paths(self) -> None:
        self.assertEqual(main(["--migrate-vb6", str(MDB)]), 2)
        self.assertEqual(main(["--output", str(self.dest)]), 2)


class MigrateMissingFileTests(unittest.TestCase):
    def test_missing_jet_file(self) -> None:
        with self.assertRaises(MigrationError):
            migrate_jet_to_sqlite(
                "/tmp/does-not-exist-markinfo.mdb",
                "/tmp/does-not-matter-markinfo.db",
            )
