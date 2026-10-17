"""Application entry point."""

from __future__ import annotations

import sys
from typing import Optional

from markinfo.config import config_from_args, parse_args
from markinfo.database import DatabaseError, connect
from markinfo.gui import MarkInfoApp
from markinfo.migrate import MigrationError, migrate_jet_to_sqlite


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    if args.migrate_vb6 or args.output:
        return _migrate_cli(args)
    cfg = config_from_args(args)
    try:
        db = connect(cfg.database)
    except DatabaseError as exc:
        print(f"markinfo: {exc}", file=sys.stderr)
        return 1
    MarkInfoApp(db, cfg).run()
    return 0


def _migrate_cli(args) -> int:
    if not args.migrate_vb6 or not args.output:
        print(
            "markinfo: --migrate-vb6 and --output are both required",
            file=sys.stderr,
        )
        return 2
    try:
        result = migrate_jet_to_sqlite(args.migrate_vb6, args.output)
    except MigrationError as exc:
        print(f"markinfo: {exc}", file=sys.stderr)
        return 1
    print(
        f"Wrote {result.path}: {result.leagues} league(s), "
        f"{result.teams} team(s), {result.players} player(s), "
        f"session {result.session}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
