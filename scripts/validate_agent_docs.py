"""エージェント文書の必須ファイルが存在するか検証するスクリプト（Cursor 版）。"""

from __future__ import annotations

import sys
from pathlib import Path

REQUIRED_FILES = [
    "AGENTS.md",
    ".cursor/doc/memory_stream.md",
    ".cursor/rules/agent_core.mdc",
    ".cursor/rules/agent_implement_entry.mdc",
    ".cursor/rules/code_comments.mdc",
    ".cursor/rules/documentation_wording.mdc",
    ".cursor/rules/engineer_signals_core.mdc",
    ".cursor/rules/external_api.mdc",
    ".cursor/rules/gdrive_mcp_core.mdc",
    ".cursor/rules/git_workflow.mdc",
    ".cursor/rules/god_class_watch.mdc",
    ".cursor/rules/japanese_tech_writing.mdc",
    ".cursor/rules/memory_logger.mdc",
    ".cursor/rules/naming_conventions.mdc",
    ".cursor/rules/project_identity.mdc",
    ".cursor/rules/safe_operations_core.mdc",
    ".cursor/rules/src_readme_policy.mdc",
    ".cursor/rules/template_sync.mdc",
    ".cursor/rules/testing_rules.mdc",
    ".cursor/rules/workspace_boundary.mdc",
    ".cursor/skills/agent-session-record/SKILL.md",
    ".cursor/skills/audit-documentation/SKILL.md",
    ".cursor/skills/audit-implementation/SKILL.md",
    ".cursor/skills/audit-operations/SKILL.md",
    ".cursor/skills/audit-post-prototype/SKILL.md",
    ".cursor/skills/audit-refactor-full/SKILL.md",
    ".cursor/skills/audit-refactoring/SKILL.md",
    ".cursor/skills/audit-security/SKILL.md",
    ".cursor/skills/changelog-entry/SKILL.md",
    ".cursor/skills/d-format-code-guide/SKILL.md",
    ".cursor/skills/design-decision-record/SKILL.md",
    ".cursor/skills/gdrive-mcp/SKILL.md",
    ".cursor/skills/implementation-phase1/SKILL.md",
    ".cursor/skills/implementation-phase2/SKILL.md",
    ".cursor/skills/japanese-tech-writing/SKILL.md",
    ".cursor/skills/known-error-entry/SKILL.md",
    ".cursor/skills/phase3-doc-updates/SKILL.md",
    ".cursor/skills/records-split/SKILL.md",
    ".cursor/skills/refactoring-report/SKILL.md",
    ".cursor/skills/safe-operations-detail/SKILL.md",
    ".cursor/skills/task-value-first/SKILL.md",
    "doc/ai/README.md",
    "doc/specs/00_プロジェクト概要.md",
    "doc/specs/00_開発日誌.md",
    "doc/specs/01_ABC見送り・ギャップ台帳.md",
    "doc/specs/02_要件定義.md",
    "doc/specs/03_システム設計.md",
    "doc/specs/04_機能一覧.md",
    "doc/specs/05_ディレクトリ構成.md",
    "doc/specs/06_ROADMAP.md",
    "doc/specs/07_CHANGELOG.md",
    "doc/初心者ガイド.md",
]


def main() -> None:
    """メイン処理。"""
    repo_root = Path(".")
    missing: list[str] = []

    for filepath in REQUIRED_FILES:
        if not (repo_root / filepath).exists():
            missing.append(filepath)

    if missing:
        print("ERROR: The following required agent documentation files are missing:")
        for item in missing:
            print(f"  - {item}")
        sys.exit(1)

    print(f"OK: All {len(REQUIRED_FILES)} required agent documentation files exist.")


if __name__ == "__main__":
    main()
