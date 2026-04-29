"""Console entrypoint for the original HWP parser CLI.

This wrapper intentionally delegates to ``hwp_full_parser.core.main`` so the
original trial-and-error parser logic, patches, CLI semantics, and web UI remain
intact.
"""

from __future__ import annotations

from .core import main as _core_main


def main() -> int:
    return _core_main()


if __name__ == "__main__":
    raise SystemExit(main())
