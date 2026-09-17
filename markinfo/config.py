"""Configuration and command-line options."""

from __future__ import annotations

import argparse
import configparser
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


DEFAULT_SQLITE_NAME = "markinfo.db"
CONFIG_NAMES = ("markinfo.cfg", "markinfo.ini")


@dataclass
class DatabaseConfig:
    kind: str = "sqlite"  # sqlite | mysql
    path: str = DEFAULT_SQLITE_NAME
    host: str = "localhost"
    port: int = 3306
    user: str = "markinfo"
    password: str = ""
    name: str = "markinfo"


def default_config_paths() -> list[Path]:
    paths = [Path.cwd() / name for name in CONFIG_NAMES]
    home = Path.home() / ".markinfo"
    paths.extend(home / name for name in CONFIG_NAMES)
    return paths


def load_config_file(path: Optional[Path] = None) -> DatabaseConfig:
    cfg = DatabaseConfig()
    parser = configparser.ConfigParser()
    if path is not None:
        files = [path]
    else:
        files = [p for p in default_config_paths() if p.is_file()]
    if not files:
        return cfg
    parser.read(files[0], encoding="utf-8")
    if parser.has_section("database"):
        section = parser["database"]
        cfg.kind = section.get("type", cfg.kind).strip().lower()
        cfg.path = section.get("path", cfg.path)
    if parser.has_section("mysql"):
        section = parser["mysql"]
        cfg.host = section.get("host", cfg.host)
        cfg.port = section.getint("port", cfg.port)
        cfg.user = section.get("user", cfg.user)
        cfg.password = section.get("password", cfg.password)
        cfg.name = section.get("database", cfg.name)
    return cfg


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
    return parser.parse_args(argv)


def config_from_args(args: argparse.Namespace) -> DatabaseConfig:
    cfg = load_config_file(args.config)
    if args.mysql:
        cfg.kind = "mysql"
    if args.database:
        cfg.kind = "sqlite"
        cfg.path = args.database
    if args.host:
        cfg.host = args.host
    if args.port:
        cfg.port = args.port
    if args.user:
        cfg.user = args.user
    if args.password is not None:
        cfg.password = args.password
    if args.dbname:
        cfg.name = args.dbname
    env_kind = os.environ.get("MARKINFO_DB_TYPE")
    if env_kind:
        cfg.kind = env_kind.strip().lower()
    env_path = os.environ.get("MARKINFO_DB")
    if env_path and cfg.kind == "sqlite" and not args.database:
        cfg.path = env_path
    return cfg
