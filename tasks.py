"""Automation tasks for the Lesson Scribe project."""

from __future__ import annotations

import os
import shlex
import subprocess
import sys
from pathlib import Path

from invoke import Context, task

ROOT_DIR = Path(__file__).parent.resolve()
SRC_DIR = ROOT_DIR / "src"
VENV_DIR = ROOT_DIR / ".venv"
VENV_BIN = "Scripts" if os.name == "nt" else "bin"
VENV_PYTHON = VENV_DIR / VENV_BIN / ("python.exe" if os.name == "nt" else "python")


def _quote(path: Path | str) -> str:
    """Return a safely shell-escaped representation of *path*."""

    value = str(path)
    if os.name == "nt":
        return subprocess.list2cmdline([value])
    return shlex.quote(value)


def _base_env() -> dict[str, str]:
    """Return a baseline environment with the repository on ``PYTHONPATH``."""

    env = os.environ.copy()
    env["PYTHONPATH"] = str(SRC_DIR)
    return env


def _venv_env() -> dict[str, str]:
    """Environment variables adjusted so commands run inside the virtualenv."""

    env = _base_env()
    env.update(
        {
            "VIRTUAL_ENV": str(VENV_DIR),
            "PATH": os.pathsep.join(
                [str(VENV_DIR / VENV_BIN), env.get("PATH", "")]
            ),
        }
    )
    return env


def _run(context: Context, command: str, *, use_venv: bool = False) -> None:
    """Execute *command* from the repository root directory."""

    env = _venv_env() if use_venv else _base_env()
    with context.cd(str(ROOT_DIR)):
        context.run(command, pty=True, echo=True, env=env)


def _ensure_virtualenv(context: Context) -> None:
    """Create the local virtual environment when it does not exist."""

    if VENV_PYTHON.exists():
        return

    python_executable = sys.executable or "python3"
    with context.cd(str(ROOT_DIR)):
        context.run(
            f"{_quote(python_executable)} -m venv {_quote(VENV_DIR)}",
            pty=True,
            echo=True,
        )


def _ensure_installation(context: Context) -> None:
    """Install project dependencies inside the managed virtual environment."""

    _ensure_virtualenv(context)

    # Upgrade pip to benefit from the latest resolver and wheels.
    _run(
        context,
        f"{_quote(VENV_PYTHON)} -m pip install --upgrade pip",
        use_venv=True,
    )

    requirements = ROOT_DIR / "requirements.txt"
    if requirements.exists() and requirements.stat().st_size > 0:
        _run(
            context,
            f"{_quote(VENV_PYTHON)} -m pip install -r {_quote(requirements)}",
            use_venv=True,
        )

    # Install the project in editable mode inside the environment.
    _run(
        context,
        f"{_quote(VENV_PYTHON)} -m pip install -e .",
        use_venv=True,
    )


def _is_project_installed(context: Context) -> bool:
    """Return True if the package is installed in the managed environment."""

    if not VENV_PYTHON.exists():
        return False

    result = context.run(
        f"{_quote(VENV_PYTHON)} -m pip show lesson-scribe",
        env=_venv_env(),
        hide=True,
        warn=True,
    )
    return result.ok


@task
def install(context: Context) -> None:
    """Set up the local virtual environment and install the project."""

    _ensure_installation(context)


@task
def run(context: Context) -> None:
    """Launch the Lesson Scribe desktop application from the virtualenv."""

    if not _is_project_installed(context):
        _ensure_installation(context)

    _run(context, f"{_quote(VENV_PYTHON)} -m lesson_scribe", use_venv=True)


@task
def lint(context: Context) -> None:
    """Run Ruff inside the managed virtual environment when available."""

    if not _is_project_installed(context):
        _ensure_installation(context)

    ruff = "ruff"
    result = context.run(
        f"{ruff} --version",
        env=_venv_env(),
        hide=True,
        warn=True,
    )
    if result.ok:
        _run(context, f"{ruff} check src", use_venv=True)
    else:
        print("Ruff is not installed. Install it with `pip install ruff` inside the venv.")
