#!/usr/bin/env python3
"""Validate the standalone public development-potential-assessment Skill."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote


PUBLIC_NAME = "hehe-development-potential-assessment"
TEXT_SUFFIXES = {".md", ".json", ".py", ".ps1", ".txt", ".html", ".yml", ".yaml"}
FORBIDDEN_PARTS = {"_本地维护", "90_历史备份", "__pycache__"}
PRIVATE_PATTERNS = {
    "local Windows user path": re.compile(r"[A-Za-z]:\\Users\\", re.I),
    "local Git workspace path": re.compile(r"D:\\LenovoSoftstore\\", re.I),
    "private tanzi-pro capability": re.compile(r"(?<![A-Za-z0-9-])tanzi-pro(?![A-Za-z0-9-])", re.I),
    "private AgentKey capability": re.compile(r"(?<![A-Za-z0-9-])agentkey(?![A-Za-z0-9-])", re.I),
}
LINK_RE = re.compile(r"!?\[[^\]]*\]\(([^)]+)\)")
NAME_RE = re.compile(r"(?m)^name:\s*([^\s]+)\s*$")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def check_links(skill_dir: Path, errors: list[str]) -> None:
    root = skill_dir.resolve()
    for md in skill_dir.rglob("*.md"):
        for raw in LINK_RE.findall(read_text(md)):
            target = raw.strip().strip("<>").split(maxsplit=1)[0]
            if not target or target.startswith("#") or re.match(r"^[a-z][a-z0-9+.-]*:", target, re.I):
                continue
            target = unquote(target.split("#", 1)[0])
            resolved = (md.parent / target).resolve()
            try:
                resolved.relative_to(root)
            except ValueError:
                errors.append(f"cross-package link: {md.relative_to(skill_dir)} -> {raw}")
                continue
            if not resolved.exists():
                errors.append(f"broken link: {md.relative_to(skill_dir)} -> {raw}")


def main() -> int:
    repo = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path(__file__).resolve().parents[1]
    errors: list[str] = []
    required_root = [repo / "README.md", repo / "LICENSE", repo / "release-manifest.json", repo / "skills" / PUBLIC_NAME]
    for path in required_root:
        if not path.exists():
            errors.append(f"missing required path: {path.relative_to(repo)}")
    if errors:
        print("\n".join(f"ERROR: {item}" for item in errors))
        return 1

    release = json.loads(read_text(repo / "release-manifest.json"))
    if release.get("entry_skill") != PUBLIC_NAME or release.get("install_policy") != "standalone":
        errors.append("release manifest must declare the public Skill as standalone entry")
    skills = release.get("skills", [])
    if skills != [{"name": PUBLIC_NAME, "role": "standalone-specialized-industry-research"}]:
        errors.append("release manifest must contain exactly one standalone specialized industry Skill")

    skill_dir = repo / "skills" / PUBLIC_NAME
    skill_file = skill_dir / "SKILL.md"
    match = NAME_RE.search(read_text(skill_file)) if skill_file.is_file() else None
    if not match or match.group(1) != PUBLIC_NAME:
        errors.append("public SKILL.md frontmatter name is incorrect")

    package = json.loads(read_text(skill_dir / "package-manifest.json"))
    if package.get("entry_skill") != PUBLIC_NAME or package.get("install_policy") != "standalone":
        errors.append("package manifest must keep the public Skill standalone")
    if package.get("required_specialty_skills"):
        errors.append("standalone Skill must not require industry-researcher specialties")
    for relative in package.get("required_local_files", []):
        if not (skill_dir / relative).is_file():
            errors.append(f"missing package file: {relative}")
    check_links(skill_dir, errors)

    for path in repo.rglob("*"):
        relative = path.relative_to(repo)
        if any(part in FORBIDDEN_PARTS for part in relative.parts):
            errors.append(f"forbidden path: {relative}")
        if path.is_file() and (path.suffix.lower() == ".pyc" or "_CO" in path.name):
            errors.append(f"forbidden generated file: {relative}")
        if path.is_file() and path.suffix.lower() in TEXT_SUFFIXES and path.resolve() != Path(__file__).resolve():
            text = read_text(path)
            for label, pattern in PRIVATE_PATTERNS.items():
                if pattern.search(text):
                    errors.append(f"{label}: {relative}")
            if re.search(r"\$development-potential-assessment\b", text):
                errors.append(f"unprefixed public invocation: {relative}")

    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        print(f"Validation failed with {len(errors)} issue(s).")
        return 1
    print(f"Validated standalone public Skill: {PUBLIC_NAME}")
    print("Name, package contract, required files, links, and public-package hygiene passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
