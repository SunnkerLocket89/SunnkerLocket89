from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable


SANDBOX_PREFIX = "sandbox:"


def normalize_path(path_str: str) -> tuple[Path, str]:
    """Return a normalized ``Path`` along with a human readable label."""

    cleaned = path_str.strip()
    raw = Path(cleaned).expanduser()
    label = cleaned

    if cleaned.startswith(SANDBOX_PREFIX):
        # remove the ``sandbox:`` prefix which is not valid locally
        stripped = cleaned[len(SANDBOX_PREFIX) :]
        raw = Path(stripped.strip()).expanduser()

    return raw, label


def _suggest_missing_path_hints(path: Path) -> list[str]:
    """Return contextual hints to help fix a missing dataset path."""

    hints: list[str] = []

    missing_parent = _first_missing_parent(path)
    if missing_parent is not None:
        parent_container = missing_parent.parent
        if parent_container == missing_parent:
            hints.append(
                f"Hint: The base location '{missing_parent}' is unavailable; ensure the "
                "dataset drive is mounted or adjust the path."
            )
        else:
            if str(parent_container) == ".":
                container_label = "the current working directory"
            else:
                container_label = f"'{parent_container}'"
            hints.append(
                "Hint: The directory "
                f"'{missing_parent}' does not exist under {container_label}. "
                "Mount the dataset or correct the path before retrying."
            )

    # Detect an accidental ``/file/`` segment that often appears in copied paths.
    anchor = path.anchor
    parts = list(path.parts)
    parts_without_anchor = parts[1:] if anchor else parts

    seen_candidates: set[Path] = set()

    for idx, part in enumerate(parts_without_anchor):
        if part != "file":
            continue

        candidate_parts = parts_without_anchor[:idx] + parts_without_anchor[idx + 1 :]
        candidate = Path(anchor, *candidate_parts) if anchor else Path(*candidate_parts)

        if candidate == path or candidate in seen_candidates:
            continue

        seen_candidates.add(candidate)
        hints.append(
            "Hint: The segment 'file' looks extraneous; try removing it, e.g. "
            f"'{candidate}'."
        )

    return hints


def _suggest_alternative_locations(path: Path, *, limit: int = 3) -> list[str]:
    """Return hints for similarly named files found in common search roots."""

    name = path.name
    if not name:
        return []

    search_roots: list[Path] = []

    def _append_if_exists(candidate: Path) -> None:
        try:
            if candidate.exists() and candidate not in search_roots:
                search_roots.append(candidate)
        except OSError:
            pass

    _append_if_exists(Path.cwd())
    _append_if_exists(Path.home())
    _append_if_exists(Path("/workspace"))

    hints: list[str] = []
    seen: set[Path] = set()

    for root in search_roots:
        matches = _iter_matches(root, name, limit=limit - len(hints))
        for match in matches:
            if match in seen:
                continue
            seen.add(match)
            hints.append(f"Hint: Found '{name}' at '{match}'.")
            if len(hints) >= limit:
                return hints

    return hints


def _iter_matches(root: Path, name: str, *, limit: int) -> Iterable[Path]:
    """Yield up to ``limit`` matches for ``name`` within ``root``."""

    if limit <= 0:
        return []

    found: list[Path] = []

    try:
        for candidate in root.rglob(name):
            found.append(candidate)
            if len(found) >= limit:
                break
    except OSError:
        return []

    return found


def _first_missing_parent(path: Path) -> Path | None:
    """Return the first missing parent directory for ``path`` if any."""

    for parent in path.parents:
        if not parent.exists():
            return parent
    return None


def main(path_str: str) -> None:
    if not path_str.strip():
        print("Error: path cannot be empty")
        sys.exit(1)

    path, label = normalize_path(path_str)

    if path.exists():
        print(f"Found: {path}")
        return

    print(f"Missing: {path}")

    if label != str(path):
        print(
            f"Hint: drop the '{SANDBOX_PREFIX}' prefix and use the local path "
            f"'{path}' instead."
        )

    if not path.is_absolute():
        print("Hint: Provide an absolute path, e.g. /mnt/data/file.xlsx")

    for hint in _suggest_missing_path_hints(path):
        print(hint)

    for hint in _suggest_alternative_locations(path):
        print(hint)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python check_dataset_path.py <path>")
        sys.exit(1)
    main(sys.argv[1])
