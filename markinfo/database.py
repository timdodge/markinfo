"""Database access layer. SQLite by default, MySQL optional."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, Optional, Sequence

from markinfo.config import DatabaseConfig
from markinfo.models import (
    CATEGORIES,
    NAME_MAX,
    NO_CAT,
    NOTES_MAX,
    POSITIONS,
    Player,
    Squad,
    TYPES,
    UNKNOWN,
)


class DatabaseError(RuntimeError):
    pass


def _row_value(row: Any, key: str, default: Any = None) -> Any:
    if row is None:
        return default
    try:
        value = row[key]
    except (KeyError, IndexError, TypeError):
        return default
    return default if value is None else value


class DBManager:
    def __init__(self, conn: Any, dialect: str) -> None:
        if dialect not in ("sqlite", "mysql"):
            raise DatabaseError(f"Unsupported database type: {dialect}")
        self.conn = conn
        self.dialect = dialect
        self._session: Optional[int] = None
        self._ensure_schema()

    # -- connection helpers -------------------------------------------------

    def _sql(self, sql: str) -> str:
        if self.dialect == "mysql":
            return sql.replace("?", "%s")
        return sql

    def _execute(self, sql: str, params: Sequence[Any] = ()) -> Any:
        cur = self.conn.cursor()
        cur.execute(self._sql(sql), tuple(params))
        if self.dialect == "sqlite":
            self.conn.commit()
        return cur

    def _fetchone(self, sql: str, params: Sequence[Any] = ()) -> Any:
        cur = self._execute(sql, params)
        try:
            return cur.fetchone()
        finally:
            cur.close()

    def _fetchall(self, sql: str, params: Sequence[Any] = ()) -> list:
        cur = self._execute(sql, params)
        try:
            return list(cur.fetchall())
        finally:
            cur.close()

    def _insert_id(self, cur: Any) -> int:
        return int(cur.lastrowid or 0)

    def _insert_ignore(self, sql: str, params: Sequence[Any] = ()) -> None:
        if self.dialect == "sqlite":
            sql = sql.replace("INSERT INTO", "INSERT OR IGNORE INTO", 1)
        else:
            sql = sql.replace("INSERT INTO", "INSERT IGNORE INTO", 1)
        cur = self._execute(sql, params)
        cur.close()

    def close(self) -> None:
        try:
            self.conn.close()
        except Exception:
            pass

    def table_exists(self, name: str) -> bool:
        try:
            cur = self._execute(f"SELECT COUNT(*) AS n FROM {name}")
            cur.close()
            return True
        except Exception:
            if self.dialect == "sqlite":
                self.conn.rollback()
            return False

    # -- schema -------------------------------------------------------------

    def _ensure_schema(self) -> None:
        if self.dialect == "sqlite":
            self._create_sqlite()
        else:
            self._create_mysql()
        self._seed()

    def _create_sqlite(self) -> None:
        statements = [
            """
            CREATE TABLE IF NOT EXISTS Admin (
                CurrSession INTEGER NOT NULL PRIMARY KEY
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS Categories (
                PlayerCat VARCHAR(3) NOT NULL PRIMARY KEY
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS Positions (
                PlayerPos VARCHAR(6) NOT NULL PRIMARY KEY
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS Types (
                PlayerType VARCHAR(4) NOT NULL PRIMARY KEY
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS Leagues (
                LeagueKey INTEGER PRIMARY KEY AUTOINCREMENT,
                LeagueName VARCHAR(20) NOT NULL UNIQUE
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS Teams (
                TeamKey INTEGER PRIMARY KEY AUTOINCREMENT,
                LeagueKey INTEGER NOT NULL,
                TeamName VARCHAR(20) NOT NULL,
                UNIQUE (LeagueKey, TeamName),
                FOREIGN KEY (LeagueKey) REFERENCES Leagues(LeagueKey)
                    ON DELETE CASCADE
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS Players (
                PlayerKey INTEGER PRIMARY KEY AUTOINCREMENT,
                TeamKey INTEGER NOT NULL,
                PlayerName VARCHAR(20) NOT NULL,
                PlayerAge INTEGER NOT NULL,
                PlayerLevel INTEGER NOT NULL,
                PlayerCat VARCHAR(3) NOT NULL,
                PlayerPos VARCHAR(6) NOT NULL,
                PlayerType VARCHAR(4) NOT NULL,
                PlayerSess INTEGER NOT NULL,
                PlayerNotes VARCHAR(255),
                UNIQUE (TeamKey, PlayerName),
                FOREIGN KEY (TeamKey) REFERENCES Teams(TeamKey)
                    ON DELETE CASCADE
            )
            """,
        ]
        for sql in statements:
            cur = self._execute(sql)
            cur.close()

    def _create_mysql(self) -> None:
        statements = [
            """
            CREATE TABLE IF NOT EXISTS Admin (
                CurrSession TINYINT NOT NULL PRIMARY KEY
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS Categories (
                PlayerCat VARCHAR(3) NOT NULL PRIMARY KEY
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS Positions (
                PlayerPos VARCHAR(6) NOT NULL PRIMARY KEY
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS Types (
                PlayerType VARCHAR(4) NOT NULL PRIMARY KEY
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS Leagues (
                LeagueKey INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                LeagueName VARCHAR(20) NOT NULL,
                UNIQUE INDEX IX_Leagues_LeagueName (LeagueName)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS Teams (
                TeamKey INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                LeagueKey INT NOT NULL,
                TeamName VARCHAR(20) NOT NULL,
                UNIQUE INDEX IX_Teams_TeamName (LeagueKey, TeamName),
                CONSTRAINT fk_teams_league
                    FOREIGN KEY (LeagueKey) REFERENCES Leagues(LeagueKey)
                    ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
            """
            CREATE TABLE IF NOT EXISTS Players (
                PlayerKey INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                TeamKey INT NOT NULL,
                PlayerName VARCHAR(20) NOT NULL,
                PlayerAge TINYINT NOT NULL,
                PlayerLevel TINYINT NOT NULL,
                PlayerCat VARCHAR(3) NOT NULL,
                PlayerPos VARCHAR(6) NOT NULL,
                PlayerType VARCHAR(4) NOT NULL,
                PlayerSess SMALLINT NOT NULL,
                PlayerNotes VARCHAR(255) NULL,
                UNIQUE INDEX IX_Players_PlayerName (TeamKey, PlayerName),
                CONSTRAINT fk_players_team
                    FOREIGN KEY (TeamKey) REFERENCES Teams(TeamKey)
                    ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """,
        ]
        for sql in statements:
            cur = self._execute(sql)
            cur.close()

    def _seed(self) -> None:
        for cat in CATEGORIES:
            self._insert_ignore("INSERT INTO Categories (PlayerCat) VALUES (?)", (cat,))
        for pos in POSITIONS:
            self._insert_ignore("INSERT INTO Positions (PlayerPos) VALUES (?)", (pos,))
        for ptype in TYPES:
            self._insert_ignore("INSERT INTO Types (PlayerType) VALUES (?)", (ptype,))
        if self._fetchone("SELECT CurrSession FROM Admin") is None:
            cur = self._execute("INSERT INTO Admin (CurrSession) VALUES (?)", (1,))
            cur.close()

    # -- lookups ------------------------------------------------------------

    def get_cats(self) -> list[str]:
        rows = self._fetchall("SELECT PlayerCat FROM Categories ORDER BY PlayerCat")
        return [str(_row_value(r, "PlayerCat")) for r in rows]

    def get_posns(self) -> list[str]:
        rows = self._fetchall("SELECT PlayerPos FROM Positions ORDER BY PlayerPos")
        return [str(_row_value(r, "PlayerPos")) for r in rows]

    def get_types(self) -> list[str]:
        rows = self._fetchall("SELECT PlayerType FROM Types ORDER BY PlayerType")
        return [str(_row_value(r, "PlayerType")) for r in rows]

    def _valid_pos(self, pos: str) -> str:
        row = self._fetchone(
            "SELECT PlayerPos FROM Positions WHERE PlayerPos = ?", (pos,)
        )
        return str(_row_value(row, "PlayerPos")).strip() if row else UNKNOWN

    def _valid_cat(self, cat: str) -> str:
        row = self._fetchone(
            "SELECT PlayerCat FROM Categories WHERE PlayerCat = ?", (cat,)
        )
        return str(_row_value(row, "PlayerCat")).strip() if row else UNKNOWN

    def _valid_type(self, ptype: str) -> str:
        row = self._fetchone(
            "SELECT PlayerType FROM Types WHERE PlayerType = ?", (ptype,)
        )
        return str(_row_value(row, "PlayerType")).strip() if row else NO_CAT

    # -- session ------------------------------------------------------------

    @property
    def session(self) -> int:
        if self._session is None:
            self._session = self._get_session()
        return self._session

    @session.setter
    def session(self, value: int) -> None:
        value = int(value)
        if self._session != value:
            self._session = value
            self._let_session(value)

    def _get_session(self) -> int:
        row = self._fetchone("SELECT CurrSession FROM Admin")
        if row is None:
            return 1
        return int(_row_value(row, "CurrSession", 1))

    def _let_session(self, value: int) -> None:
        row = self._fetchone("SELECT CurrSession FROM Admin")
        if row is None:
            cur = self._execute(
                "INSERT INTO Admin (CurrSession) VALUES (?)", (value,)
            )
        else:
            cur = self._execute("UPDATE Admin SET CurrSession = ?", (value,))
        cur.close()

    def get_max_player_key(self) -> int:
        row = self._fetchone("SELECT MAX(PlayerKey) AS MaxPlayerKey FROM Players")
        value = _row_value(row, "MaxPlayerKey")
        return int(value) if value is not None else 0

    # -- leagues / teams ----------------------------------------------------

    def add_league(self, league: str) -> int:
        league = league.strip().upper()[:NAME_MAX]
        row = self._fetchone(
            "SELECT LeagueKey FROM Leagues WHERE LeagueName = ?", (league,)
        )
        if row:
            return int(_row_value(row, "LeagueKey"))
        cur = self._execute(
            "INSERT INTO Leagues (LeagueName) VALUES (?)", (league,)
        )
        key = self._insert_id(cur)
        cur.close()
        return key

    def add_team(self, team: str, league: str) -> int:
        team = team.strip().upper()[:NAME_MAX]
        league = league.strip().upper()[:NAME_MAX]
        row = self._fetchone(
            """
            SELECT t.TeamKey
            FROM Teams t
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE t.TeamName = ? AND l.LeagueName = ?
            """,
            (team, league),
        )
        if row:
            return int(_row_value(row, "TeamKey"))
        league_key = self.add_league(league)
        cur = self._execute(
            "INSERT INTO Teams (LeagueKey, TeamName) VALUES (?, ?)",
            (league_key, team),
        )
        key = self._insert_id(cur)
        cur.close()
        return key

    def get_leagues(self) -> list[str]:
        rows = self._fetchall("SELECT LeagueName FROM Leagues ORDER BY LeagueName")
        return [str(_row_value(r, "LeagueName")) for r in rows]

    def get_teams(self, league: str) -> list[str]:
        rows = self._fetchall(
            """
            SELECT t.TeamName
            FROM Teams t
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE l.LeagueName = ?
            ORDER BY t.TeamName
            """,
            (league,),
        )
        return [str(_row_value(r, "TeamName")) for r in rows]

    def rename_league(self, old_name: str, new_name: str) -> None:
        new_name = new_name.strip().upper()[:NAME_MAX]
        cur = self._execute(
            "UPDATE Leagues SET LeagueName = ? WHERE LeagueName = ?",
            (new_name, old_name),
        )
        cur.close()

    def rename_team(self, old_name: str, new_name: str, league: str) -> None:
        new_name = new_name.strip().upper()[:NAME_MAX]
        row = self._fetchone(
            "SELECT LeagueKey FROM Leagues WHERE LeagueName = ?", (league,)
        )
        if not row:
            raise DatabaseError("League not found")
        league_key = int(_row_value(row, "LeagueKey"))
        cur = self._execute(
            """
            UPDATE Teams SET TeamName = ?
            WHERE TeamName = ? AND LeagueKey = ?
            """,
            (new_name, old_name, league_key),
        )
        cur.close()

    def del_league(self, league: str) -> None:
        teams = self._fetchall(
            """
            SELECT t.TeamKey
            FROM Teams t
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE l.LeagueName = ?
            """,
            (league,),
        )
        for row in teams:
            team_key = int(_row_value(row, "TeamKey"))
            cur = self._execute("DELETE FROM Players WHERE TeamKey = ?", (team_key,))
            cur.close()
            cur = self._execute("DELETE FROM Teams WHERE TeamKey = ?", (team_key,))
            cur.close()
        cur = self._execute("DELETE FROM Leagues WHERE LeagueName = ?", (league,))
        cur.close()

    def del_team(self, team: str, league: str) -> None:
        row = self._fetchone(
            """
            SELECT t.TeamKey
            FROM Teams t
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE t.TeamName = ? AND l.LeagueName = ?
            """,
            (team, league),
        )
        if not row:
            return
        team_key = int(_row_value(row, "TeamKey"))
        cur = self._execute("DELETE FROM Players WHERE TeamKey = ?", (team_key,))
        cur.close()
        cur = self._execute("DELETE FROM Teams WHERE TeamKey = ?", (team_key,))
        cur.close()

    # -- players ------------------------------------------------------------

    def add_player(self, player: Player) -> tuple[int, bool]:
        """Insert or update a player. Returns (player_key, created)."""
        name = player.normalised_name()
        team = player.team.strip().upper()[:NAME_MAX]
        league = player.league.strip().upper()[:NAME_MAX]
        pos = self._valid_pos(player.pos)
        if pos == "GK":
            cat = NO_CAT
        else:
            cat = self._valid_cat(player.cat)
        ptype = self._valid_type(player.type)
        sess = self.session if player.sess == 0 else int(player.sess)
        notes = (player.notes or "").strip()[:NOTES_MAX] or None

        row = self._fetchone(
            """
            SELECT p.PlayerKey, p.PlayerSess, p.PlayerCat, p.PlayerPos
            FROM Players p
            JOIN Teams t ON p.TeamKey = t.TeamKey
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE p.PlayerName = ? AND t.TeamName = ? AND l.LeagueName = ?
            """,
            (name, team, league),
        )
        if row:
            player_key = int(_row_value(row, "PlayerKey"))
            existing_sess = int(_row_value(row, "PlayerSess", 0))
            if sess >= existing_sess:
                existing_cat = str(_row_value(row, "PlayerCat", "")).strip()
                existing_pos = str(_row_value(row, "PlayerPos", "")).strip()
                if cat == UNKNOWN and existing_cat != UNKNOWN:
                    cat = existing_cat
                if pos == UNKNOWN and existing_pos != UNKNOWN:
                    pos = existing_pos
                cur = self._execute(
                    """
                    UPDATE Players
                    SET PlayerAge = ?, PlayerLevel = ?, PlayerCat = ?,
                        PlayerPos = ?, PlayerType = ?, PlayerSess = ?,
                        PlayerNotes = ?
                    WHERE PlayerKey = ?
                    """,
                    (
                        int(player.age),
                        int(player.level),
                        cat,
                        pos,
                        ptype,
                        sess,
                        notes,
                        player_key,
                    ),
                )
                cur.close()
            return player_key, False

        team_key = self.add_team(team, league)
        cur = self._execute(
            """
            INSERT INTO Players (
                TeamKey, PlayerName, PlayerAge, PlayerLevel, PlayerCat,
                PlayerPos, PlayerType, PlayerSess, PlayerNotes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                team_key,
                name,
                int(player.age),
                int(player.level),
                cat,
                pos,
                ptype,
                sess,
                notes,
            ),
        )
        player_key = self._insert_id(cur)
        cur.close()
        return player_key, True

    def del_player(self, name: str, team: str, league: str) -> None:
        row = self._fetchone(
            """
            SELECT p.PlayerKey
            FROM Players p
            JOIN Teams t ON p.TeamKey = t.TeamKey
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE p.PlayerName = ? AND t.TeamName = ? AND l.LeagueName = ?
            """,
            (name, team, league),
        )
        if not row:
            return
        cur = self._execute(
            "DELETE FROM Players WHERE PlayerKey = ?",
            (int(_row_value(row, "PlayerKey")),),
        )
        cur.close()

    def move_player(self, player: Player, old_team: str) -> tuple[int, bool]:
        self.del_player(player.normalised_name(), old_team, player.league)
        return self.add_player(player)

    def get_squad(self, team: str, league: str) -> Squad:
        rows = self._fetchall(
            """
            SELECT p.PlayerName, p.PlayerAge, p.PlayerLevel, p.PlayerCat,
                   p.PlayerPos, p.PlayerType, p.PlayerSess, p.PlayerNotes
            FROM Players p
            JOIN Teams t ON p.TeamKey = t.TeamKey
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE t.TeamName = ? AND l.LeagueName = ?
            ORDER BY p.PlayerName
            """,
            (team, league),
        )
        squad = Squad()
        for row in rows:
            notes = _row_value(row, "PlayerNotes") or ""
            squad.add(
                Player(
                    league=league,
                    team=team,
                    name=str(_row_value(row, "PlayerName") or ""),
                    age=int(_row_value(row, "PlayerAge") or 0),
                    level=int(_row_value(row, "PlayerLevel") or 0),
                    cat=str(_row_value(row, "PlayerCat") or "").strip(),
                    pos=str(_row_value(row, "PlayerPos") or "").strip(),
                    type=str(_row_value(row, "PlayerType") or "").strip(),
                    sess=int(_row_value(row, "PlayerSess") or 0),
                    notes=str(notes),
                )
            )
        return squad

    def search_players(
        self,
        name: str = "",
        league: str = "",
        team: str = "",
        pos: str = "",
        cat: str = "",
        ptype: str = "",
        min_age: Optional[int] = None,
        max_age: Optional[int] = None,
        min_level: Optional[int] = None,
        max_level: Optional[int] = None,
    ) -> list[Player]:
        clauses = ["1=1"]
        params: list[Any] = []
        if name:
            clauses.append("p.PlayerName LIKE ?")
            params.append(f"%{name.strip().upper()}%")
        if league:
            clauses.append("l.LeagueName = ?")
            params.append(league)
        if team:
            clauses.append("t.TeamName = ?")
            params.append(team)
        if pos:
            clauses.append("p.PlayerPos = ?")
            params.append(pos)
        if cat:
            clauses.append("p.PlayerCat = ?")
            params.append(cat)
        if ptype:
            clauses.append("p.PlayerType = ?")
            params.append(ptype)
        if min_age is not None:
            clauses.append("p.PlayerAge >= ?")
            params.append(min_age)
        if max_age is not None:
            clauses.append("p.PlayerAge <= ?")
            params.append(max_age)
        if min_level is not None:
            clauses.append("p.PlayerLevel >= ?")
            params.append(min_level)
        if max_level is not None:
            clauses.append("p.PlayerLevel <= ?")
            params.append(max_level)
        where = " AND ".join(clauses)
        rows = self._fetchall(
            f"""
            SELECT l.LeagueName, t.TeamName, p.PlayerName, p.PlayerAge,
                   p.PlayerLevel, p.PlayerCat, p.PlayerPos, p.PlayerType,
                   p.PlayerSess, p.PlayerNotes
            FROM Players p
            JOIN Teams t ON p.TeamKey = t.TeamKey
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE {where}
            ORDER BY l.LeagueName, t.TeamName, p.PlayerName
            """,
            params,
        )
        results: list[Player] = []
        for row in rows:
            notes = _row_value(row, "PlayerNotes") or ""
            results.append(
                Player(
                    league=str(_row_value(row, "LeagueName") or ""),
                    team=str(_row_value(row, "TeamName") or ""),
                    name=str(_row_value(row, "PlayerName") or ""),
                    age=int(_row_value(row, "PlayerAge") or 0),
                    level=int(_row_value(row, "PlayerLevel") or 0),
                    cat=str(_row_value(row, "PlayerCat") or "").strip(),
                    pos=str(_row_value(row, "PlayerPos") or "").strip(),
                    type=str(_row_value(row, "PlayerType") or "").strip(),
                    sess=int(_row_value(row, "PlayerSess") or 0),
                    notes=str(notes),
                )
            )
        return results

    def new_season(self, league: str, session_limit: int = 2) -> None:
        """Age session numbers and drop players older than the keep window."""
        rows = self._fetchall(
            """
            SELECT p.PlayerKey, p.PlayerSess
            FROM Players p
            JOIN Teams t ON p.TeamKey = t.TeamKey
            JOIN Leagues l ON t.LeagueKey = l.LeagueKey
            WHERE l.LeagueName = ?
            """,
            (league,),
        )
        cutoff = -1 * (session_limit - 1)
        for row in rows:
            key = int(_row_value(row, "PlayerKey"))
            sess = int(_row_value(row, "PlayerSess"))
            if sess < cutoff:
                cur = self._execute(
                    "DELETE FROM Players WHERE PlayerKey = ?", (key,)
                )
            elif sess < 0:
                cur = self._execute(
                    "UPDATE Players SET PlayerSess = PlayerSess - 1 "
                    "WHERE PlayerKey = ?",
                    (key,),
                )
            else:
                cur = self._execute(
                    "UPDATE Players SET PlayerSess = -1 WHERE PlayerKey = ?",
                    (key,),
                )
            cur.close()


def connect_sqlite(path: str | Path) -> DBManager:
    db_path = Path(path)
    if db_path.parent and str(db_path.parent) not in ("", "."):
        db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return DBManager(conn, "sqlite")


def connect_mysql(
    host: str,
    port: int,
    user: str,
    password: str,
    database: str,
) -> DBManager:
    try:
        import pymysql
        from pymysql.cursors import DictCursor
    except ImportError as exc:
        raise DatabaseError(
            "MySQL support requires pymysql. Install with: pip install pymysql"
        ) from exc
    try:
        conn = pymysql.connect(
            host=host,
            port=int(port),
            user=user,
            password=password,
            database=database,
            charset="utf8mb4",
            autocommit=True,
            cursorclass=DictCursor,
        )
    except Exception as exc:
        raise DatabaseError(f"Could not connect to MySQL: {exc}") from exc
    return DBManager(conn, "mysql")


def connect(cfg: DatabaseConfig) -> DBManager:
    if cfg.kind == "mysql":
        return connect_mysql(cfg.host, cfg.port, cfg.user, cfg.password, cfg.name)
    return connect_sqlite(cfg.path)
