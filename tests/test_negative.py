from __future__ import annotations

import re
import subprocess
from pathlib import Path

import pytest


def run_cmd(args: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        args,
        cwd=str(cwd),
        check=True,
        text=True,
        capture_output=True,
    )


def find_cloudflare_masking_markers(html_text: str) -> list[str]:
    """Return matched masking signatures that indicate likely CF/obfuscation corruption."""
    patterns = {
        "cf_email_protection_path": r"/cdn-cgi/l/email-protection",
        "cf_email_data_attr": r"data-cfemail=",
        "cf_email_token": r"__cf_email__",
        "cf_email_bracketed": r"\[email\s*protected\]",
            "cf_decode_script": r'/cdn-cgi/scripts/[^"]*email-decode\.min\.js',
        # URL obfuscation variants often seen after copy/upload sanitization.
        "obfuscated_scheme_hxxp": r"\bhxxps?://",
        "obfuscated_dot_bracket": r"\[\.\]",
    }
    hits: list[str] = []
    for name, pattern in patterns.items():
        if re.search(pattern, html_text, flags=re.IGNORECASE):
            hits.append(name)
    return hits


@pytest.mark.negative
def test_generated_master_has_no_cloudflare_masking(repo_root: Path, python_exe: str) -> None:
    """Guardrail: generated master.html should not include obfuscated/masked artifacts."""
    run_cmd(
        [python_exe, "code/run_tests.py", "run", "--scenario", "hiv_teams"],
        cwd=repo_root,
    )

    master = repo_root / "tests" / "hiv_teams" / "out" / "pages" / "master.html"
    assert master.exists(), f"Missing generated file: {master}"

    text = master.read_text(encoding="utf-8")
    hits = find_cloudflare_masking_markers(text)
    assert not hits, f"Detected Cloudflare/obfuscation artifacts in master.html: {hits}"


@pytest.mark.negative
def test_cloudflare_masking_detector_fires_on_corrupted_master(repo_root: Path, python_exe: str, tmp_path: Path) -> None:
    """Negative control: if masking is injected, detector must flag it."""
    run_cmd(
        [python_exe, "code/run_tests.py", "run", "--scenario", "hiv_teams"],
        cwd=repo_root,
    )

    source_master = repo_root / "tests" / "hiv_teams" / "out" / "pages" / "master.html"
    assert source_master.exists(), f"Missing generated file: {source_master}"

    corrupted = tmp_path / "master_corrupted.html"
    injected = source_master.read_text(encoding="utf-8") + "\n" + "\n".join(
        [
            "<a href=\"/cdn-cgi/l/email-protection#4f2a223e2a3f232a0f2a372e223f232a612c2022\">[email protected]</a>",
            "<script src=\"/cdn-cgi/scripts/5c5dd728/cloudflare-static/email-decode.min.js\"></script>",
            "Visit hxxps://example[.]org/path for details",
        ]
    )
    corrupted.write_text(injected, encoding="utf-8")

    hits = find_cloudflare_masking_markers(corrupted.read_text(encoding="utf-8"))
    # This is the key negative assertion: detector must flag the injected corruption.
    assert hits, "Expected masking detector to flag injected Cloudflare/obfuscation markers"
