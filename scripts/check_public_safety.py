from __future__ import annotations

import re
import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

TOKEN_PATTERNS = [
    ("private key", re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----")),
    ("OpenAI-style API key", re.compile(r"\bsk-[A-Za-z0-9_-]{20,}\b")),
    ("Google API key", re.compile(r"\bAIza[0-9A-Za-z_-]{30,}\b")),
    (
        "GitHub token",
        re.compile(r"\b(?:ghp|gho|ghu|ghs|ghr)_[A-Za-z0-9]{20,}\b|\bgithub_pat_[A-Za-z0-9_]{20,}"),
    ),
    ("AWS access key", re.compile(r"\bAKIA[0-9A-Z]{16}\b")),
    ("Slack webhook", re.compile(r"https://hooks\.slack\.com/services/[A-Za-z0-9/_-]{20,}", re.I)),
    ("DeepL Free API key", re.compile(r"\b[0-9a-f]{8}-[0-9a-f-]{20,}:fx\b", re.I)),
]

FORBIDDEN_SUFFIXES = {".pem", ".key", ".p12", ".pfx", ".xls", ".xlsx", ".xlsm", ".xlsb"}


def tracked_files() -> list[Path]:
    output = subprocess.check_output(["git", "ls-files", "-z"], cwd=ROOT)
    return [ROOT / item.decode("utf-8") for item in output.split(b"\0") if item]


def is_placeholder(value: str) -> bool:
    clean = value.strip().strip("\"'")
    if not clean:
        return True
    return bool(re.match(r"^(?:replace|example|your[_-]|ここに|<|\.\.\.)", clean, re.I))


def path_issue(path: Path) -> str | None:
    rel = path.relative_to(ROOT)
    name = path.name.lower()

    if name == ".env" or name.startswith(".env."):
        return "local environment file must not be tracked"

    if any(part.lower() == "credentials" for part in rel.parts):
        return "credential directory must not be tracked"

    if path.suffix.lower() in FORBIDDEN_SUFFIXES:
        return "credential or Excel data file must not be tracked"

    return None


def scan_text(path: Path, text: str) -> list[str]:
    rel = path.relative_to(ROOT).as_posix()
    findings: list[str] = []

    for label, pattern in TOKEN_PATTERNS:
        if pattern.search(text):
            findings.append(f"{rel}: possible {label}")

    fallback = re.search(
        r'DEEPL_AUTH_KEY_FALLBACK\s+As\s+String\s*=\s*"([^"]*)"',
        text,
        re.I,
    )
    if fallback and not is_placeholder(fallback.group(1)):
        findings.append(f"{rel}: DEEPL_AUTH_KEY_FALLBACK contains a non-placeholder value")

    legacy = re.search(
        r'DEEPL_AUTH_KEY\s+As\s+String\s*=\s*"([^"]*)"',
        text,
        re.I,
    )
    if legacy and not is_placeholder(legacy.group(1)):
        findings.append(f"{rel}: DEEPL_AUTH_KEY contains a non-placeholder value")

    return findings


def main() -> int:
    findings: list[str] = []
    tracked = tracked_files()

    for path in tracked:
        issue = path_issue(path)
        if issue:
            findings.append(f"{path.relative_to(ROOT).as_posix()}: {issue}")
            continue

        try:
            if path.stat().st_size > 2 * 1024 * 1024:
                continue
            raw = path.read_bytes()
        except OSError as exc:
            findings.append(f"{path.relative_to(ROOT).as_posix()}: cannot inspect file: {exc}")
            continue

        if b"\0" in raw:
            continue

        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            continue

        findings.extend(scan_text(path, text))

    if findings:
        print("Public repository safety check failed:")
        for finding in findings:
            print(f"- {finding}")
        return 1

    print(f"Public repository safety check passed ({len(tracked)} tracked files scanned).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
