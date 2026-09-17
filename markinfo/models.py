"""Domain objects for MarkInfo."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Iterator, List, Optional


NO_CAT = "---"
UNKNOWN = "?"
KEEPER = "GK"
STAR = "STAR"
SBY = "SBY"
APP = "APP"
FUT = "FUT"
POWER = "P"
SKILL = "S"
POWER_SKILL = "P/S"

CATEGORIES = (UNKNOWN, NO_CAT, POWER, SKILL, POWER_SKILL)
POSITIONS = (
    UNKNOWN,
    KEEPER,
    "SW",
    "DF",
    "DF/A",
    "MF",
    "MF/D",
    "MF/G",
    "MF/A",
    "MF/A/D",
    "FW",
    "FWT",
    "FWS",
    "WG",
    "UT",
)
TYPES = (NO_CAT, APP, FUT, SBY, STAR)

NAME_MAX = 20
NOTES_MAX = 255
SESSION_MIN = 1
SESSION_MAX = 16
AGE_MIN = 17
AGE_MAX = 99
LEVEL_MIN = 0
LEVEL_MAX = 99


def infer_type(age: int, level: int) -> str:
    """Infer player type the same way the original shomatch import did."""
    if level > 12:
        return STAR
    if age == 17:
        return SBY
    return NO_CAT


def notes_indicator(notes: Optional[str]) -> str:
    """Short grid marker for notes, matching the VB6 flexgrid logic."""
    if not notes:
        return ""
    first = notes.split()[0] if notes.split() else notes
    if len(first) <= 3 and (
        first.isdigit()
        or (first[:1].lower() == "i" and first[1:].isdigit())
        or (first[:1].lower() == "s" and first[1:].isdigit())
    ):
        return first
    return "✓"


@dataclass
class Player:
    league: str = ""
    team: str = ""
    name: str = ""
    age: int = AGE_MIN
    level: int = LEVEL_MIN
    cat: str = UNKNOWN
    pos: str = UNKNOWN
    type: str = NO_CAT
    sess: int = 0
    notes: str = ""

    def normalised_name(self) -> str:
        return (self.name or "").strip().upper()[:NAME_MAX]


@dataclass
class Squad:
    players: List[Player] = field(default_factory=list)

    def add(self, player: Player) -> None:
        self.players.append(player)

    def __iter__(self) -> Iterator[Player]:
        return iter(self.players)

    def __len__(self) -> int:
        return len(self.players)

    def __getitem__(self, index: int) -> Player:
        return self.players[index]


@dataclass(frozen=True)
class AddOp:
    player: Player


@dataclass(frozen=True)
class MoveOp:
    player: Player
    old_team: str


Op = AddOp | MoveOp
