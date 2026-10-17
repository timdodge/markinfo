"""Copy a VB6 Jet (Access) database into a new SQLite file."""

from __future__ import annotations

import csv
import io
import os
import shutil
import sqlite3
import subprocess
from dataclasses import dataclass
from pathlib import Path

from markinfo.database import connect_sqlite

# mdb-export writes this in place of NULL. It is not a character that
# appears in MarkInfo text, so blank notes stay distinct from missing ones.
_NULL = "\x1e"

_LOOKUP_TABLES = (
    ("Categories", "PlayerCat"),
    ("Positions", "PlayerPos"),
    ("Types", "PlayerType"),
)


class MigrationError(Exception):
    """The Jet file could not be copied into SQLite."""


@dataclass(frozen=True)
class MigrationResult:
    path: Path
    leagues: int
    teams: int
    players: int
    session: int


def migrate_jet_to_sqlite(
    mdb_path: str | Path,
    sqlite_path: str | Path,
    *,
    overwrite: bool = False,
) -> MigrationResult:
    """Read a VB6 MarkInfo Jet database and write a new SQLite database.

    League, team, and player keys are kept so the relationships in the
    Access file survive. The destination is replaced only when ``overwrite``
    is set, and only after the new file has been written successfully.
    """
    source = Path(mdb_path)
    destination = Path(sqlite_path)
    if shutil.which("mdb-export") is None:
        raise MigrationError(
            "mdbtools is required to read a VB6 database. "
            "Install the mdbtools package and try again."
        )
    if not source.is_file():
        raise MigrationError(f"Jet database not found: {source}")
    if destination.exists() and not overwrite:
        raise MigrationError(f"SQLite file already exists: {destination}")
    if destination.exists() and destination.is_dir():
        raise MigrationError(f"SQLite path is a directory: {destination}")

    temporary = destination.with_name(destination.name + ".migrating")
    if temporary.exists():
        temporary.unlink()
    try:
        result = _write_database(source, temporary)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    destination.parent.mkdir(parents=True, exist_ok=True)
    os.replace(temporary, destination)
    return MigrationResult(
        path=destination,
        leagues=result.leagues,
        teams=result.teams,
        players=result.players,
        session=result.session,
    )


def _write_database(source: Path, destination: Path) -> MigrationResult:
    db = connect_sqlite(destination)
    try:
        for table, column in _LOOKUP_TABLES:
            for row in _export_rows(source, table):
                value = _required_text(row, table, column)
                db._insert_ignore(
                    f"INSERT INTO {table} ({column}) VALUES (?)",
                    (value,),
                )
        session = _session(_export_rows(source, "Admin"))
        db.session = session

        leagues = _export_rows(source, "Leagues")
        teams = _export_rows(source, "Teams")
        players = _export_rows(source, "Players")
        try:
            for row in leagues:
                db._execute(
                    "INSERT INTO Leagues (LeagueKey, LeagueName) VALUES (?, ?)",
                    (
                        _required_int(row, "Leagues", "LeagueKey"),
                        _required_text(row, "Leagues", "LeagueName"),
                    ),
                ).close()
            for row in teams:
                name = _optional_text(row, "TeamName") or ""
                db._execute(
                    "INSERT INTO Teams (TeamKey, LeagueKey, TeamName) VALUES (?, ?, ?)",
                    (
                        _required_int(row, "Teams", "TeamKey"),
                        _required_int(row, "Teams", "LeagueKey"),
                        name,
                    ),
                ).close()
            for row in players:
                db._execute(
                    """
                    INSERT INTO Players (
                        PlayerKey, TeamKey, PlayerName, PlayerAge, PlayerLevel,
                        PlayerCat, PlayerPos, PlayerType, PlayerSess, PlayerNotes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        _required_int(row, "Players", "PlayerKey"),
                        _required_int(row, "Players", "TeamKey"),
                        _required_text(row, "Players", "PlayerName"),
                        _required_int(row, "Players", "PlayerAge"),
                        _required_int(row, "Players", "PlayerLevel"),
                        _required_text(row, "Players", "PlayerCat"),
                        _required_text(row, "Players", "PlayerPos"),
                        _required_text(row, "Players", "PlayerType"),
                        _required_int(row, "Players", "PlayerSess"),
                        _optional_text(row, "PlayerNotes"),
                    ),
                ).close()
        except sqlite3.IntegrityError as exc:
            raise MigrationError(f"Jet data does not fit the SQLite schema: {exc}") from exc
    except OSError as exc:
        raise MigrationError(str(exc)) from exc
    finally:
        db.close()
    return MigrationResult(
        path=destination,
        leagues=len(leagues),
        teams=len(teams),
        players=len(players),
        session=session,
    )


def _export_rows(mdb: Path, table: str) -> list[dict[str, str | None]]:
    try:
        proc = subprocess.run(
            ["mdb-export", "-0", _NULL, str(mdb), table],
            check=False,
            capture_output=True,
        )
    except OSError as exc:
        raise MigrationError(f"Could not run mdb-export: {exc}") from exc
    if proc.returncode != 0:
        detail = proc.stderr.decode("utf-8", "replace").strip()
        raise MigrationError(detail or f"mdb-export failed for {table}")
    text = proc.stdout.decode("utf-8", "replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        raise MigrationError(f"{table} export had no header row")
    rows: list[dict[str, str | None]] = []
    for raw in reader:
        rows.append(
            {
                key: None if value == _NULL else value
                for key, value in raw.items()
                if key is not None
            }
        )
    return rows


def _session(rows: list[dict[str, str | None]]) -> int:
    if not rows:
        return 1
    values = [_required_int(row, "Admin", "CurrSession") for row in rows]
    return max(values)


def _required_text(row: dict[str, str | None], table: str, column: str) -> str:
    value = row.get(column)
    if value is None or not value.strip():
        raise MigrationError(f"{table} row is missing {column}")
    return value.strip()


def _optional_text(row: dict[str, str | None], column: str) -> str | None:
    value = row.get(column)
    if value is None:
        return None
    value = value.strip()
    return value or None


def _required_int(row: dict[str, str | None], table: str, column: str) -> int:
    value = row.get(column)
    if value is None or not str(value).strip():
        raise MigrationError(f"{table} row is missing {column}")
    try:
        return int(str(value).strip())
    except ValueError as exc:
        raise MigrationError(f"{table}.{column} is not a number: {value}") from exc
