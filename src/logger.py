"""
===========================================================
Cisco Partner Intelligence Automation Platform

Module:
logger.py

Purpose:
Shared logging infrastructure used by every script in the
platform. Writes timestamped, level-tagged messages to the
console and, when enabled, appends them to the run log.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

from datetime import datetime
from pathlib import Path

from config import (
    ENABLE_RUN_LOG,
    RUN_LOG_FILE,
    TIMESTAMP_FORMAT,
)


def _format_message(message: str, level: str) -> str:
    """Compose a timestamped, level-tagged log line."""
    timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)
    return f"[{timestamp}] [{level}] {message}"


def _write_to_run_log(text: str) -> None:
    """Append a single log line to the run log file."""
    Path(RUN_LOG_FILE).parent.mkdir(parents=True, exist_ok=True)
    with open(RUN_LOG_FILE, "a", encoding="utf-8") as file:
        file.write(text + "\n")


def log(message: str, level: str = "INFO") -> None:
    """Log a message to the console and, if enabled, the run log.

    Parameters
    ----------
    message:
        The message to log.
    level:
        Severity label, typically INFO, WARNING, or ERROR.
    """
    formatted_message = _format_message(message, level)
    print(formatted_message)
    if ENABLE_RUN_LOG:
        _write_to_run_log(formatted_message)
