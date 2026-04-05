"""
Filesystem MCP Server
Provides file system tools via FastMCP.
All operations work relative to the current working directory.
"""

import os
import shutil
import logging
from pathlib import Path
from fastmcp import FastMCP

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="[FS-MCP] %(asctime)s %(levelname)s │ %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("filesystem_mcp")

# ── Server ─────────────────────────────────────────────────────────────────────
mcp = FastMCP(
    name="filesystem-mcp",
)


def _abs(name: str) -> Path:
    """Resolve a filename relative to CWD."""
    return Path(os.getcwd()) / name


# ── Tools ──────────────────────────────────────────────────────────────────────

@mcp.tool()
def create_file(filename: str) -> str:
    """
    Create an empty file at *filename* (relative to CWD).
    Parent directories are created automatically.
    """
    path = _abs(filename)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.touch(exist_ok=True)
    log.info("create_file → %s", path)
    return f"Created file: {path}"


@mcp.tool()
def write_file(filename: str, content: str) -> str:
    """
    Write *content* to *filename*, overwriting any existing content.
    Parent directories are created automatically.
    """
    path = _abs(filename)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    log.info("write_file → %s (%d bytes)", path, len(content))
    return f"Written {len(content)} bytes to: {path}"


@mcp.tool()
def append_file(filename: str, content: str) -> str:
    """
    Append *content* to *filename*.
    Creates the file if it does not exist.
    """
    path = _abs(filename)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(content)
    log.info("append_file → %s (+%d bytes)", path, len(content))
    return f"Appended {len(content)} bytes to: {path}"


@mcp.tool()
def modify_file(filename: str, old: str, new: str) -> str:
    """
    Replace the first occurrence of *old* with *new* inside *filename*.
    Returns an error string if the file is not found or *old* is not present.
    """
    path = _abs(filename)
    if not path.exists():
        msg = f"ERROR: file not found: {path}"
        log.warning(msg)
        return msg
    text = path.read_text(encoding="utf-8")
    if old not in text:
        msg = f"ERROR: substring not found in {path}"
        log.warning(msg)
        return msg
    path.write_text(text.replace(old, new, 1), encoding="utf-8")
    log.info("modify_file → %s", path)
    return f"Modified file: {path}"


@mcp.tool()
def delete_file(filename: str) -> str:
    """
    Delete a file at *filename*.
    Returns an error string if the file does not exist.
    """
    path = _abs(filename)
    if not path.exists():
        msg = f"ERROR: file not found: {path}"
        log.warning(msg)
        return msg
    path.unlink()
    log.info("delete_file → %s", path)
    return f"Deleted file: {path}"


@mcp.tool()
def delete_directory(dirname: str) -> str:
    """
    Recursively delete a directory at *dirname*.
    Returns an error string if the directory does not exist.
    """
    path = _abs(dirname)
    if not path.exists():
        msg = f"ERROR: directory not found: {path}"
        log.warning(msg)
        return msg
    shutil.rmtree(path)
    log.info("delete_directory → %s", path)
    return f"Deleted directory: {path}"


@mcp.tool()
def create_file_with_content(filename: str, content: str) -> str:
    """
    Create a file at *filename* and immediately write *content* to it.
    Combines create_file + write_file in a single atomic call.
    """
    path = _abs(filename)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    log.info("create_file_with_content → %s (%d bytes)", path, len(content))
    return f"Created file with content: {path} ({len(content)} bytes)"


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    log.info("Starting Filesystem MCP Server …")
    mcp.run(transport="stdio")
