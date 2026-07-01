from datetime import datetime
from pathlib import Path

from config import (
    RUN_LOG_FILE,
    ENABLE_RUN_LOG
)


def log(message, level="INFO"):
    """
    Logs a message to both the console and the run log.

    Parameters
    ----------
    message : str
        Message to log.

    level : str
        INFO | WARNING | ERROR
    """

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    formatted_message = (
        f"[{timestamp}] "
        f"[{level}] "
        f"{message}"
    )

    print(formatted_message)

    if ENABLE_RUN_LOG:

        Path(RUN_LOG_FILE).parent.mkdir(
            parents=True,
            exist_ok=True
        )

        with open(
            RUN_LOG_FILE,
            "a",
            encoding="utf-8"
        ) as file:

            file.write(
                formatted_message + "\n"
            )