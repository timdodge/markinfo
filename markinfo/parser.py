"""Parse Kickabout shomatch zip/text files into database operations."""

from __future__ import annotations

import os
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Iterable, Optional

from markinfo.models import (
    AddOp,
    KEEPER,
    MoveOp,
    NO_CAT,
    Op,
    Player,
    POWER,
    POWER_SKILL,
    SKILL,
    UNKNOWN,
    infer_type,
)

FORMAT_RE = re.compile(r"<\$[^>]*>")
INT_RE = re.compile(r"-?\d+")

INT_SQUAD = "INTERNATIONAL SQUAD"
REP_SQUAD = "REPRESENTATIVE SQUAD"
AUCTION = "AUCTION RESULTS"
DEALS = "PRIVATE DEALS"


class ParseError(RuntimeError):
    pass


def parse_int(value: str) -> int:
    """VB6 CInt-style parse: leading signed digits, trailing junk ignored."""
    match = INT_RE.search(value)
    if not match:
        raise ValueError(f"not an integer: {value!r}")
    return int(match.group(0))


def tokenize_line(line: str) -> list[str]:
    line = FORMAT_RE.sub("", line.upper())
    return line.split()


def classify_file(first_line: str) -> Optional[str]:
    text = first_line.upper()
    if INT_SQUAD in text or REP_SQUAD in text:
        return "squad"
    if AUCTION in text:
        return "auction"
    if DEALS in text:
        return "deals"
    return None


def _age_level_word(word: str) -> bool:
    return len(word) >= 4 and word[2] == "-"


def parse_squad_line(words: list[str], league: str) -> Optional[Player]:
    last = len(words) - 1
    if last < 4:
        return None
    found = None
    for i in range(3, last + 1):
        if not _age_level_word(words[i]):
            continue
        nxt = words[i + 1] if i + 1 <= last else ""
        prev = words[i - 1]
        if nxt in (SKILL, POWER, POWER_SKILL) or prev == KEEPER:
            found = i
            break
    if found is None or found >= last:
        return None
    i = found
    if words[i - 1] == KEEPER and words[i + 1] not in (POWER, SKILL, POWER_SKILL):
        team = " ".join(words[i + 1 :])
    else:
        team = " ".join(words[i + 2 :])
    name_parts = []
    for token in words[: i - 1]:
        if not (token.endswith(")") and len(token) <= 3):
            name_parts.append(token)
    name = " ".join(name_parts)
    age = parse_int(words[i][:2])
    level = parse_int(words[i][3:])
    pos = words[i - 1]
    cat = NO_CAT if words[i - 1] == KEEPER else words[i + 1]
    return Player(
        league=league,
        team=team,
        name=name,
        age=age,
        level=level,
        cat=cat,
        pos=pos,
        type=infer_type(age, level),
        sess=0,
        notes="",
    )


def parse_auction_line(words: list[str], league: str) -> Optional[Player]:
    last = len(words) - 1
    if last < 4:
        return None
    found = None
    for i in range(2, last + 1):
        if _age_level_word(words[i]) and words[i - 2].endswith(")"):
            found = i
            break
    if found is None or found >= last:
        return None
    if words[3] == "NOT" and words[4] == "SOLD.":
        return None
    idx = last - 1
    team = words[idx][:-1]
    while not team.startswith("("):
        idx -= 1
        if idx < 0:
            return None
        team = words[idx] + " " + team
    team = team[1:]
    name_parts: list[str] = []
    while idx > 3:
        idx -= 1
        name_parts.insert(0, words[idx])
    name = " ".join(name_parts)
    age = parse_int(words[2][:2])
    level = parse_int(words[2][3:])
    pos = words[1]
    cat = NO_CAT if pos == KEEPER else UNKNOWN
    return Player(
        league=league,
        team=team,
        name=name,
        age=age,
        level=level,
        cat=cat,
        pos=pos,
        type=infer_type(age, level),
        sess=0,
        notes="",
    )


def _is_age_token(word: str) -> bool:
    return (
        len(word) >= 4
        and word.startswith("(")
        and word[1:3].isdigit()
        and word[3] == "-"
    )


def parse_deals_line(words: list[str], league: str) -> list[MoveOp]:
    last = len(words) - 1
    if last < 4:
        return []
    i = 0
    old_parts: list[str] = []
    while i <= last and words[i] not in ("SWAPPED", "SOLD"):
        old_parts.append(words[i])
        i += 1
    if i > last:
        return []
    old_team = " ".join(old_parts)
    switch = False
    pending: list[Player] = []
    try:
        while i <= last and words[i] in ("SWAPPED", "SOLD", "AND", "FOR"):
            i += 1
            name_parts: list[str] = []
            while i <= last and not _is_age_token(words[i]):
                name_parts.append(words[i])
                i += 1
            name = " ".join(name_parts)
            age = parse_int(words[i][1:3])
            dash = words[i].find("-")
            level = parse_int(words[i][dash + 1 :])
            ptype = infer_type(age, level)
            i += 1
            pos_word = words[i]
            bracket = pos_word.find("[")
            if bracket == -1:
                pos = pos_word[:-1]
            else:
                pos = pos_word[:bracket]
                i += 1
            i += 1
            cat = UNKNOWN
            if i <= last and words[i] == KEEPER:
                cat = NO_CAT
            if i <= last and words[i] == "{P/S}":
                cat = POWER_SKILL
                i += 1
            if i <= last and words[i] == "{SBY}":
                ptype = "SBY"
                i += 1
            elif i <= last and words[i] == "{FUT}":
                ptype = "FUT"
                i += 1
            elif i <= last and words[i] == "{APP}":
                ptype = "APP"
                i += 1
            elif i <= last and words[i] == "{STAR}":
                ptype = "STAR"
                i += 1
            pending.append(
                Player(
                    league=league,
                    team=old_team if switch else "",
                    name=name,
                    age=age,
                    level=level,
                    cat=cat,
                    pos=pos,
                    type=ptype,
                    sess=0,
                    notes="",
                )
            )
            if i <= last and words[i] == "FOR":
                switch = True
            if i + 1 <= last:
                nxt = words[i + 1]
                if len(nxt) >= 2 and nxt[-1] == "K" and nxt[:-1].isdigit():
                    i += 2
        if i > last:
            return []
        if words[i] != "TO":
            while i <= last and words[i] != "FROM":
                i += 1
            i += 1
            new_team = " ".join(words[i:]).rstrip(" .")
        else:
            i += 1
            new_parts: list[str] = []
            while i <= last and words[i] != "FOR":
                new_parts.append(words[i])
                i += 1
            new_team = " ".join(new_parts)
    except (IndexError, ValueError):
        return []
    ops: list[MoveOp] = []
    for player in pending:
        if player.team == "":
            player.team = new_team
            ops.append(MoveOp(player=player, old_team=old_team))
        else:
            ops.append(MoveOp(player=player, old_team=new_team))
    return ops


def parse_shomatch_text(text: str, league: str) -> list[Op]:
    lines = text.splitlines()
    if not lines:
        return []
    kind = None
    ops: list[Op] = []
    for raw in lines:
        if kind is None:
            kind = classify_file(raw)
            if kind is None:
                return []
            continue
        words = tokenize_line(raw)
        if len(words) < 5:
            continue
        if kind == "squad":
            player = parse_squad_line(words, league)
            if player:
                ops.append(AddOp(player))
        elif kind == "auction":
            player = parse_auction_line(words, league)
            if player:
                ops.append(AddOp(player))
        elif kind == "deals":
            ops.extend(parse_deals_line(words, league))
    return ops


def read_text_file(path: str | Path) -> str:
    data = Path(path).read_bytes()
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def iter_input_texts(path: str | Path) -> Iterable[tuple[str, str]]:
    """Yield (source_name, text) for a zip archive or a single text file."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".zip":
        with tempfile.TemporaryDirectory(prefix="markinfo") as tmp:
            with zipfile.ZipFile(path) as archive:
                for info in archive.infolist():
                    if info.is_dir():
                        continue
                    name = os.path.basename(info.filename)
                    if not name.lower().endswith(".txt"):
                        continue
                    target = Path(tmp) / name
                    with archive.open(info) as src, open(target, "wb") as dst:
                        dst.write(src.read())
                    yield name, read_text_file(target)
    elif suffix == ".txt":
        yield path.name, read_text_file(path)
    else:
        raise ParseError(f"Unsupported file type: {path}")


def process_input_file(
    path: str | Path,
    league: str,
    add_player: Callable[[Player], tuple[int, bool]],
    move_player: Callable[[Player, str], tuple[int, bool]],
) -> tuple[int, int]:
    """Apply shomatch ops through the provided database callbacks."""
    added = 0
    updated = 0
    for _name, text in iter_input_texts(path):
        for op in parse_shomatch_text(text, league):
            if isinstance(op, MoveOp):
                _key, created = move_player(op.player, op.old_team)
            else:
                _key, created = add_player(op.player)
            if created:
                added += 1
            else:
                updated += 1
    return added, updated
