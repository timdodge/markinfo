"""Configuration and command-line options."""

from __future__ import annotations

import argparse
import configparser
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


DEFAULT_SQLITE_NAME = "markinfo.db"
CONFIG_NAMES = ("markinfo.cfg", "markinfo.ini")
UI_SCALE_STEPS = (1.0, 1.25, 1.5, 1.75, 2.0, 2.5)
UI_SCALE_MIN = UI_SCALE_STEPS[0]
UI_SCALE_MAX = UI_SCALE_STEPS[-1]


@dataclass
class DatabaseConfig:
    kind: str = "sqlite"  # sqlite | mysql
    path: str = DEFAULT_SQLITE_NAME
    host: str = "localhost"
    port: int = 3306
    user: str = "markinfo"
    password: str = ""
    name: str = "markinfo"


@dataclass
class AppConfig:
    database: DatabaseConfig = field(default_factory=DatabaseConfig)
    ui_scale: float = 1.0
    source: Optional[Path] = None


def default_config_paths() -> list[Path]:
    paths = [Path.cwd() / name for name in CONFIG_NAMES]
    home = Path.home() / ".markinfo"
    paths.extend(home / name for name in CONFIG_NAMES)
    return paths


def default_save_path() -> Path:
    return Path.home() / ".markinfo" / "markinfo.cfg"


def normalize_ui_scale(value: float) -> float:
    try:
        scale = float(value)
    except (TypeError, ValueError):
        return 1.0
    return min(UI_SCALE_MAX, max(UI_SCALE_MIN, scale))


def step_ui_scale(current: float, delta: int) -> float:
    """Move one preset step up (delta > 0) or down (delta < 0)."""
    current = normalize_ui_scale(current)
    if delta > 0:
        for step in UI_SCALE_STEPS:
            if step > current + 1e-9:
                return step
        return UI_SCALE_MAX
    if delta < 0:
        for step in reversed(UI_SCALE_STEPS):
            if step < current - 1e-9:
                return step
        return UI_SCALE_MIN
    return current


def load_config_file(path: Optional[Path] = None) -> AppConfig:
    cfg = AppConfig()
    parser = configparser.ConfigParser()
    if path is not None:
        cfg.source = path
        files = [path] if path.is_file() else []
    else:
        files = [p for p in default_config_paths() if p.is_file()]
    if not files:
        return cfg
    used = files[0]
    cfg.source = used
    parser.read(used, encoding="utf-8")
    db = cfg.database
    if parser.has_section("database"):
        section = parser["database"]
        db.kind = section.get("type", db.kind).strip().lower()
        db.path = section.get("path", db.path)
    if parser.has_section("mysql"):
        section = parser["mysql"]
        db.host = section.get("host", db.host)
        db.port = section.getint("port", db.port)
        db.user = section.get("user", db.user)
        db.password = section.get("password", db.password)
        db.name = section.get("database", db.name)
    if parser.has_section("ui"):
        cfg.ui_scale = normalize_ui_scale(parser["ui"].getfloat("scale", 1.0))
    return cfg


def save_config(cfg: AppConfig) -> Path:
    path = cfg.source if cfg.source is not None else default_save_path()
    parser = configparser.ConfigParser()
    if path.is_file():
        parser.read(path, encoding="utf-8")
    if not parser.has_section("ui"):
        parser.add_section("ui")
    parser.set("ui", "scale", f"{cfg.ui_scale:g}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        parser.write(handle)
    cfg.source = path
    return path


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="markinfo",
        description="MarkInfo - Kickabout player database manager",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="Path to markinfo.cfg / markinfo.ini",
    )
    parser.add_argument(
        "--database",
        "-d",
        help="SQLite database file (default: ./markinfo.db)",
    )
    parser.add_argument(
        "--mysql",
        action="store_true",
        help="Use MySQL instead of SQLite",
    )
    parser.add_argument("--host", help="MySQL host")
    parser.add_argument("--port", type=int, help="MySQL port")
    parser.add_argument("--user", help="MySQL user")
    parser.add_argument("--password", help="MySQL password")
    parser.add_argument("--dbname", help="MySQL database name")
    parser.add_argument(
        "--ui-scale",
        type=float,
        dest="ui_scale",
        help="UI size multiplier (1.0 = 100%%, 1.5 = 150%%, max 2.5)",
    )
    return parser.parse_args(argv)


def config_from_args(args: argparse.Namespace) -> AppConfig:
    cfg = load_config_file(args.config)
    db = cfg.database
    if args.mysql:
        db.kind = "mysql"
    if args.database:
        db.kind = "sqlite"
        db.path = args.database
    if args.host:
        db.host = args.host
    if args.port:
        db.port = args.port
    if args.user:
        db.user = args.user
    if args.password is not None:
        db.password = args.password
    if args.dbname:
        db.name = args.dbname
    if args.ui_scale is not None:
        cfg.ui_scale = normalize_ui_scale(args.ui_scale)
    env_kind = os.environ.get("MARKINFO_DB_TYPE")
    if env_kind:
        db.kind = env_kind.strip().lower()
    env_path = os.environ.get("MARKINFO_DB")
    if env_path and db.kind == "sqlite" and not args.database:
        db.path = env_path
    return cfg
