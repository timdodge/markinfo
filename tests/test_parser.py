import unittest
from pathlib import Path

from markinfo.models import AddOp, MoveOp
from markinfo.parser import (
    classify_file,
    parse_auction_line,
    parse_deals_line,
    parse_shomatch_text,
    parse_squad_line,
    tokenize_line,
)

FIXTURES = Path(__file__).resolve().parent / "fixtures"


class TokenizeTests(unittest.TestCase):
    def test_strips_format_tags_and_uppercases(self) -> None:
        words = tokenize_line("  <$BOLD>Smith   df 24-8 s arsenal  ")
        self.assertEqual(words, ["SMITH", "DF", "24-8", "S", "ARSENAL"])

    def test_classify(self) -> None:
        self.assertEqual(classify_file("International Squad"), "squad")
        self.assertEqual(classify_file("REPRESENTATIVE SQUAD"), "squad")
        self.assertEqual(classify_file("Auction Results"), "auction")
        self.assertEqual(classify_file("Private Deals"), "deals")
        self.assertIsNone(classify_file("something else"))


class SquadParseTests(unittest.TestCase):
    def test_outfield_player(self) -> None:
        player = parse_squad_line(
            ["SMITH", "DF", "24-8", "S", "ARSENAL"], "PREMIER"
        )
        assert player is not None
        self.assertEqual(player.name, "SMITH")
        self.assertEqual(player.pos, "DF")
        self.assertEqual(player.age, 24)
        self.assertEqual(player.level, 8)
        self.assertEqual(player.cat, "S")
        self.assertEqual(player.team, "ARSENAL")
        self.assertEqual(player.type, "---")
        self.assertEqual(player.league, "PREMIER")

    def test_keeper_skips_nationality_token(self) -> None:
        player = parse_squad_line(
            ["JONES", "(E)", "GK", "28-10", "MANCHESTER", "UNITED"], "PREMIER"
        )
        assert player is not None
        self.assertEqual(player.name, "JONES")
        self.assertEqual(player.pos, "GK")
        self.assertEqual(player.cat, "---")
        self.assertEqual(player.team, "MANCHESTER UNITED")
        self.assertEqual(player.age, 28)
        self.assertEqual(player.level, 10)

    def test_star_and_schoolboy_types(self) -> None:
        star = parse_squad_line(
            ["VAN", "BASTEN", "FW", "30-14", "P/S", "AC", "MILAN"], "SERIE A"
        )
        sby = parse_squad_line(
            ["YOUNGSTER", "MF", "17-4", "S", "YOUTH", "FC"], "PREMIER"
        )
        assert star and sby
        self.assertEqual(star.type, "STAR")
        self.assertEqual(sby.type, "SBY")

    def test_fixture_file(self) -> None:
        ops = parse_shomatch_text((FIXTURES / "squad.txt").read_text(), "PREMIER")
        names = [op.player.name for op in ops if isinstance(op, AddOp)]
        self.assertEqual(names, ["SMITH", "JONES", "VAN BASTEN", "YOUNGSTER"])


class AuctionParseTests(unittest.TestCase):
    def test_bought_player(self) -> None:
        player = parse_auction_line(
            ["LOT1)", "DF", "24-8", "JOHN", "SMITH", "(ARSENAL)", "150K"],
            "PREMIER",
        )
        assert player is not None
        self.assertEqual(player.name, "JOHN SMITH")
        self.assertEqual(player.team, "ARSENAL")
        self.assertEqual(player.pos, "DF")
        self.assertEqual(player.age, 24)
        self.assertEqual(player.level, 8)
        self.assertEqual(player.cat, "?")

    def test_keeper_cat(self) -> None:
        player = parse_auction_line(
            ["LOT2)", "GK", "19-6", "BILLY", "KEEPER", "(CHELSEA)", "80K"],
            "PREMIER",
        )
        assert player is not None
        self.assertEqual(player.cat, "---")
        self.assertEqual(player.pos, "GK")

    def test_not_sold_skipped(self) -> None:
        player = parse_auction_line(
            ["LOT3)", "MF", "22-7", "NOT", "SOLD.", "(EVERTON)", "0K"],
            "PREMIER",
        )
        self.assertIsNone(player)

    def test_fixture_file(self) -> None:
        ops = parse_shomatch_text((FIXTURES / "auction.txt").read_text(), "PREMIER")
        self.assertEqual(len(ops), 2)
        self.assertEqual(ops[0].player.team, "ARSENAL")
        self.assertEqual(ops[1].player.team, "CHELSEA")


class DealsParseTests(unittest.TestCase):
    def test_sale(self) -> None:
        ops = parse_deals_line(
            [
                "ARSENAL", "SOLD", "JOHN", "SMITH", "(24-8", "DF)",
                "TO", "CHELSEA", "FOR", "500K",
            ],
            "PREMIER",
        )
        self.assertEqual(len(ops), 1)
        self.assertIsInstance(ops[0], MoveOp)
        self.assertEqual(ops[0].old_team, "ARSENAL")
        self.assertEqual(ops[0].player.team, "CHELSEA")
        self.assertEqual(ops[0].player.name, "JOHN SMITH")
        self.assertEqual(ops[0].player.pos, "DF")
        self.assertEqual(ops[0].player.age, 24)
        self.assertEqual(ops[0].player.level, 8)

    def test_swap(self) -> None:
        ops = parse_deals_line(
            [
                "LIVERPOOL", "SWAPPED", "JANE", "DOE", "(22-7", "MF)",
                "FOR", "BOB", "ROE", "(25-9", "FW)", "FROM", "EVERTON.",
            ],
            "PREMIER",
        )
        self.assertEqual(len(ops), 2)
        outgoing, incoming = ops
        self.assertEqual(outgoing.player.name, "JANE DOE")
        self.assertEqual(outgoing.old_team, "LIVERPOOL")
        self.assertEqual(outgoing.player.team, "EVERTON")
        self.assertEqual(incoming.player.name, "BOB ROE")
        self.assertEqual(incoming.old_team, "EVERTON")
        self.assertEqual(incoming.player.team, "LIVERPOOL")

    def test_fixture_file(self) -> None:
        ops = parse_shomatch_text((FIXTURES / "deals.txt").read_text(), "PREMIER")
        self.assertEqual(len(ops), 3)


if __name__ == "__main__":
    unittest.main()
