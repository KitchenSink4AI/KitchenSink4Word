"""One line, once, asking for a star.

The products are free and community-supported, and the single most useful
thing a happy user can do costs them one click. This says so exactly once per
installation and then never again.

WHAT IT WILL NOT DO. It never touches stdout, which carries the protocol. It
never asks the network anything. It never blocks, retries, or waits. It runs
after startup has already succeeded, so a failure here cannot stop the server,
and every error path is silence.

THE MARKER IS WRITTEN BEFORE THE LINE IS PRINTED, deliberately. If the state
directory cannot be written, the line is not shown at all. "Shown once" turning
into "shown every single start" is the one failure mode worth designing
against, and a nudge nobody ever sees is a smaller loss than a nag.

Set KS4W_STAR_NUDGE=off to silence it without waiting for the marker.
"""

import os
import sys
from pathlib import Path

__all__ = ["MESSAGE", "MARKER_NAME", "marker_path", "announce_once"]

MARKER_NAME = "star-nudge"

ENV_TOGGLE = "KS4W_STAR_NUDGE"

_OFF = {"off", "0", "false", "no", "none"}

MESSAGE = (
    "KitchenSink4Word is community-supported. If it helps you, star "
    "github.com/KitchenSink4AI/KitchenSink4Word. Shown once."
)


def marker_path() -> Path:
    """The marker, in the state directory this product already writes to."""
    from .update_check import cache_path

    return cache_path().parent / MARKER_NAME


def announce_once(stream=None) -> bool:
    """Show the line if it has never been shown. True when it was shown."""
    if os.environ.get(ENV_TOGGLE, "").strip().lower() in _OFF:
        return False
    try:
        marker = marker_path()
        marker.parent.mkdir(parents=True, exist_ok=True)
        # O_EXCL so two servers starting at the same moment cannot both claim
        # the first run. The loser gets FileExistsError and stays quiet.
        os.close(os.open(str(marker), os.O_CREAT | os.O_EXCL | os.O_WRONLY))
    except Exception:
        return False
    try:
        (stream or sys.stderr).write(MESSAGE + "\n")
    except Exception:
        return False
    return True
