"""Application entry point."""

from __future__ import annotations

import sys
from typing import Optional

from markinfo.config import config_from_args, parse_args
from markinfo.database import DatabaseError, connect
from markinfo.gui import MarkInfoApp


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    cfg = config_from_args(args)
    try:
        db = connect(cfg.database)
    except DatabaseError as exc:
        print(f"markinfo: {exc}", file=sys.stderr)
        return 1
    MarkInfoApp(db, cfg).run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
