"""xlgrep: grep for the cells of Excel workbooks."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("xlgrep")
except PackageNotFoundError:  # running from a source tree without installation
    __version__ = "0.0.0"
