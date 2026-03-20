import nox

nox.options.sessions = ["lint", "test"]

SOURCES = [
    "france_regions",
]


MAX_LINE_LEN = 120


@nox.session()
def lint(session):
    session.install("black", "flake8", "isort")
    session.run("black", "--check", "-l", str(MAX_LINE_LEN), *SOURCES)
    session.run("flake8", "--max-line-length", str(MAX_LINE_LEN), "--per-file-ignores", "ositah/apps/*/callbacks.py:F403,F405 ositah/apps/*/main.py:F403,F405", *SOURCES)
    session.run("isort", "--check", "--profile", "black", "-l", str(MAX_LINE_LEN), *SOURCES)


@nox.session(name="format")
def format_(session):
    session.install("black", "isort")
    session.run("black", "-l", str(MAX_LINE_LEN), *SOURCES)
    session.run("isort", "--profile", "black", "-l", str(MAX_LINE_LEN), *SOURCES)


@nox.session(python=["2.7", "3.5", "3.6", "3.7", "3.8", "3.9"])
def test(session):
    session.install("pytest")
    session.install("dash[testing]")
    session.install(".")
    session.run("pytest", "--headless", "tests")


@nox.session()
def doctest(session):
    session.install("pytest")
    session.install("dash[testing]")
    session.install("-r", "docs/requirements.txt")
    session.install(".")
    session.run(
        "pytest",
        "--headless",
        "-v",
        "docs/components_page/components/__tests__",
    )
