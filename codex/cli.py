"""Command line interface for managing Codex API credentials."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

DEFAULT_CONFIG_PATH = Path("~/.config/codex/config.json").expanduser()


class CLIError(RuntimeError):
    """Raised when CLI usage results in an error condition."""


class APIKeyWriter:
    """Persist API keys to a configuration file."""

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = path or DEFAULT_CONFIG_PATH

    def write(self, key: str, *, overwrite: bool = False) -> Path:
        """Write *key* to the configuration file."""
        key = key.strip()
        if not key:
            raise CLIError("No API key provided")

        config_dir = self.path.parent
        config_dir.mkdir(parents=True, exist_ok=True)

        if self.path.exists() and not overwrite:
            raise CLIError(
                f"Refusing to overwrite existing configuration at {self.path}. "
                "Use --overwrite to replace it."
            )

        data = {"api_key": key}
        with self.path.open("w", encoding="utf-8") as handle:
            json.dump(data, handle)
            handle.write("\n")
        return self.path


class LoginCommand:
    """Handle the ``codex login`` command."""

    def __init__(self, writer: Optional[APIKeyWriter] = None) -> None:
        self.writer = writer or APIKeyWriter()

    def run(self, with_api_key: bool, *, overwrite: bool = False) -> str:
        """Execute the command and return a status message."""
        if not with_api_key:
            raise CLIError(
                "No API key source provided. Use --with-api-key and pipe the key via STDIN."
            )

        if sys.stdin.isatty():
            raise CLIError(
                "--with-api-key requires the key to be provided through STDIN, "
                "for example: codex login --with-api-key < my_key.txt"
            )

        api_key = sys.stdin.read()
        path = self.writer.write(api_key, overwrite=overwrite)
        return f"API key saved to {path}"


class CodexCLI:
    """Main CLI entry point."""

    def __init__(self) -> None:
        self.parser = argparse.ArgumentParser(
            prog="codex",
            description="Manage Codex API credentials.",
        )
        subparsers = self.parser.add_subparsers(dest="command", required=True)

        login_parser = subparsers.add_parser(
            "login", help="Store your Codex API key in the local configuration file."
        )
        login_parser.add_argument(
            "--with-api-key",
            action="store_true",
            help="Read the API key from STDIN."
        )
        login_parser.add_argument(
            "--overwrite",
            action="store_true",
            help="Overwrite an existing configuration file if present.",
        )
        login_parser.add_argument(
            "--config-path",
            type=str,
            help=(
                "Write the API key to a custom configuration file instead of the "
                "default location."
            ),
        )

    def run(self, argv: Optional[list[str]] = None) -> str:
        """Parse *argv* and dispatch to the appropriate sub-command."""
        args = self.parser.parse_args(argv)

        if args.command == "login":
            config_path = Path(args.config_path).expanduser() if args.config_path else None
            command = LoginCommand(writer=APIKeyWriter(config_path))
            return command.run(with_api_key=args.with_api_key, overwrite=args.overwrite)

        raise CLIError(f"Unknown command: {args.command}")


def main(argv: Optional[list[str]] = None) -> int:
    """Entry point for the ``codex`` console script."""
    cli = CodexCLI()
    try:
        message = cli.run(argv)
    except CLIError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    else:
        print(message)
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
