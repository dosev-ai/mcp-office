from __future__ import annotations

import importlib.util
import io
import pathlib
import tarfile
import tempfile
import unittest
import zipfile

SCRIPT_PATH = pathlib.Path(__file__).with_name("verify_distribution.py")
SPEC = importlib.util.spec_from_file_location("verify_distribution", SCRIPT_PATH)
assert SPEC and SPEC.loader
verify_distribution = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(verify_distribution)

PROJECT = '''[project]
name = "mcp-office"
version = "1.2.3"
[project.scripts]
excelmcp = "excelmcp.server:main"
[tool.setuptools]
packages = ["excelmcp", "wordmcp._com"]
[tool.setuptools.package-dir]
excelmcp = "excelmcp/src/excelmcp"
wordmcp = "wordmcp/src/wordmcp"
'''


class VerifyDistributionTests(unittest.TestCase):
    def make_fixture(
        self,
        *,
        wheel_package: bool = True,
        sdist_package: bool = True,
        entry_target: str = "excelmcp.server:main",
    ):
        td = tempfile.TemporaryDirectory()
        root = pathlib.Path(td.name)
        project = root / "pyproject.toml"
        project.write_text(PROJECT, encoding="utf-8")
        dist = root / "dist"
        dist.mkdir()

        wheel = dist / "mcp_office-1.2.3-py3-none-any.whl"
        with zipfile.ZipFile(wheel, "w") as archive:
            archive.writestr("excelmcp/__init__.py", "")
            if wheel_package:
                archive.writestr("wordmcp/_com/__init__.py", "")
            archive.writestr(
                "mcp_office-1.2.3.dist-info/METADATA",
                "Name: mcp-office\nVersion: 1.2.3\n",
            )
            archive.writestr(
                "mcp_office-1.2.3.dist-info/entry_points.txt",
                f"[console_scripts]\nexcelmcp = {entry_target}\n",
            )

        sdist = dist / "mcp_office-1.2.3.tar.gz"
        with tarfile.open(sdist, "w:gz") as archive:
            entries = [
                ("mcp_office-1.2.3/pyproject.toml", PROJECT.encode()),
                ("mcp_office-1.2.3/README.md", b"readme"),
                ("mcp_office-1.2.3/excelmcp/src/excelmcp/__init__.py", b""),
            ]
            if sdist_package:
                entries.append(
                    ("mcp_office-1.2.3/wordmcp/src/wordmcp/_com/__init__.py", b"")
                )
            for name, data in entries:
                info = tarfile.TarInfo(name)
                info.size = len(data)
                archive.addfile(info, io.BytesIO(data))

        return td, project, dist

    def test_happy_path(self):
        td, project, dist = self.make_fixture()
        with td:
            name, version, _, _ = verify_distribution.verify(project, dist, "v1.2.3")
            self.assertEqual((name, version), ("mcp-office", "1.2.3"))

    def test_rejects_tag_version_mismatch(self):
        td, project, dist = self.make_fixture()
        with td, self.assertRaisesRegex(SystemExit, "does not match project version"):
            verify_distribution.verify(project, dist, "v9.9.9")

    def test_rejects_missing_wheel_package(self):
        td, project, dist = self.make_fixture(wheel_package=False)
        with td, self.assertRaisesRegex(
            SystemExit,
            "wheel is missing configured package 'wordmcp._com'",
        ):
            verify_distribution.verify(project, dist)

    def test_rejects_wrong_console_script(self):
        td, project, dist = self.make_fixture(entry_target="excelmcp.wrong:main")
        with td, self.assertRaisesRegex(SystemExit, "console script 'excelmcp' maps"):
            verify_distribution.verify(project, dist)

    def test_rejects_missing_sdist_package(self):
        td, project, dist = self.make_fixture(sdist_package=False)
        with td, self.assertRaisesRegex(
            SystemExit,
            "sdist is missing configured package 'wordmcp._com'",
        ):
            verify_distribution.verify(project, dist)


if __name__ == "__main__":
    unittest.main()
