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

PACKAGES = [
    "excelmcp",
    "pptmcp",
    "wordmcp",
    "wordmcp._com",
    "wordmcp._docx",
    "mcpshared",
]
SCRIPTS = {
    "excelmcp": "excelmcp.server:main",
    "pptmcp": "pptmcp.server:main",
    "wordmcp": "wordmcp.server:main",
}
PACKAGE_DIRS = {
    "excelmcp": "excelmcp/src/excelmcp",
    "pptmcp": "pptmcp/src/pptmcp",
    "wordmcp": "wordmcp/src/wordmcp",
    "mcpshared": "shared/src/mcpshared",
}


def _project_text(
    packages: list[str] | None = None,
    scripts: dict[str, str] | None = None,
) -> str:
    packages = PACKAGES if packages is None else packages
    scripts = SCRIPTS if scripts is None else scripts
    text = '[project]\nname = "mcp-office"\nversion = "1.2.3"\n[project.scripts]\n'
    text += "".join(f'{name} = "{target}"\n' for name, target in scripts.items())
    text += "[tool.setuptools]\npackages = ["
    text += ", ".join(f'"{package}"' for package in packages)
    text += "]\n[tool.setuptools.package-dir]\n"
    text += "".join(
        f'{name} = "{source}"\n' for name, source in PACKAGE_DIRS.items()
    )
    return text


def _source_path(package: str) -> str:
    root = max(
        (
            name
            for name in PACKAGE_DIRS
            if package == name or package.startswith(f"{name}.")
        ),
        key=len,
    )
    suffix = package[len(root):].lstrip(".").replace(".", "/")
    return PACKAGE_DIRS[root] + (f"/{suffix}" if suffix else "")


class VerifyDistributionTests(unittest.TestCase):
    def make_fixture(
        self,
        *,
        project_text: str | None = None,
        wheel_omit: str | None = None,
        sdist_omit: str | None = None,
        wrong_entry: str | None = None,
        sdist_version: str = "1.2.3",
    ):
        td = tempfile.TemporaryDirectory()
        root = pathlib.Path(td.name)
        project = root / "pyproject.toml"
        project_text = project_text or _project_text()
        project.write_text(project_text, encoding="utf-8")
        dist = root / "dist"
        dist.mkdir()

        wheel = dist / "mcp_office-1.2.3-py3-none-any.whl"
        with zipfile.ZipFile(wheel, "w") as archive:
            for package in PACKAGES:
                path = package.replace(".", "/") + "/__init__.py"
                if path != wheel_omit:
                    archive.writestr(path, "")
            if "excelmcp/extra.py" != wheel_omit:
                archive.writestr("excelmcp/extra.py", "")
            archive.writestr(
                "mcp_office-1.2.3.dist-info/METADATA",
                "Name: mcp-office\nVersion: 1.2.3\n",
            )
            entry_points = dict(SCRIPTS)
            if wrong_entry is not None:
                entry_points["excelmcp"] = wrong_entry
            archive.writestr(
                "mcp_office-1.2.3.dist-info/entry_points.txt",
                "[console_scripts]\n"
                + "".join(
                    f"{name} = {target}\n"
                    for name, target in entry_points.items()
                ),
            )

        sdist = dist / "mcp_office-1.2.3.tar.gz"
        with tarfile.open(sdist, "w:gz") as archive:
            root_pkg_info = f"Name: mcp-office\nVersion: {sdist_version}\n".encode()
            entries = [
                ("mcp_office-1.2.3/pyproject.toml", project_text.encode()),
                ("mcp_office-1.2.3/README.md", b"readme"),
                ("mcp_office-1.2.3/PKG-INFO", root_pkg_info),
                (
                    "mcp_office-1.2.3/mcp_office.egg-info/PKG-INFO",
                    root_pkg_info,
                ),
            ]
            entries.extend(
                (
                    f"mcp_office-1.2.3/{_source_path(package)}/__init__.py",
                    b"",
                )
                for package in PACKAGES
            )
            entries.append(
                ("mcp_office-1.2.3/excelmcp/src/excelmcp/extra.py", b"")
            )
            for name, data in entries:
                if name == sdist_omit:
                    continue
                info = tarfile.TarInfo(name)
                info.size = len(data)
                archive.addfile(info, io.BytesIO(data))

        return td, project, dist

    def test_happy_path_allows_egg_info_pkg_info(self):
        td, project, dist = self.make_fixture()
        with td:
            name, version, _, _ = verify_distribution.verify(
                project, dist, "v1.2.3"
            )
            self.assertEqual((name, version), ("mcp-office", "1.2.3"))

    def test_rejects_tag_version_mismatch(self):
        td, project, dist = self.make_fixture()
        with td, self.assertRaisesRegex(SystemExit, "does not match project version"):
            verify_distribution.verify(project, dist, "v9.9.9")

    def test_rejects_required_suite_removed_from_build_config(self):
        packages = [package for package in PACKAGES if package != "pptmcp"]
        scripts = {
            name: target for name, target in SCRIPTS.items() if name != "pptmcp"
        }
        td, project, dist = self.make_fixture(
            project_text=_project_text(packages=packages, scripts=scripts)
        )
        with td, self.assertRaisesRegex(SystemExit, "missing required suite packages"):
            verify_distribution.verify(project, dist)

    def test_rejects_missing_wheel_package(self):
        td, project, dist = self.make_fixture(
            wheel_omit="wordmcp/_com/__init__.py"
        )
        with td, self.assertRaisesRegex(
            SystemExit,
            "wheel is missing configured package 'wordmcp._com'",
        ):
            verify_distribution.verify(project, dist)

    def test_rejects_wrong_console_script(self):
        td, project, dist = self.make_fixture(
            wrong_entry="excelmcp.wrong:main"
        )
        with td, self.assertRaisesRegex(SystemExit, "console script 'excelmcp' maps"):
            verify_distribution.verify(project, dist)

    def test_rejects_missing_root_sdist_pkg_info_even_with_egg_info_copy(self):
        td, project, dist = self.make_fixture(
            sdist_omit="mcp_office-1.2.3/PKG-INFO"
        )
        with td, self.assertRaisesRegex(SystemExit, "expected root sdist PKG-INFO"):
            verify_distribution.verify(project, dist)

    def test_rejects_missing_sdist_package_content(self):
        td, project, dist = self.make_fixture(
            sdist_omit=(
                "mcp_office-1.2.3/wordmcp/src/wordmcp/_com/__init__.py"
            )
        )
        with td, self.assertRaisesRegex(
            SystemExit,
            "sdist is missing wheel package contents",
        ):
            verify_distribution.verify(project, dist)

    def test_rejects_wheel_module_missing_from_sdist_parity(self):
        td, project, dist = self.make_fixture(
            sdist_omit="mcp_office-1.2.3/excelmcp/src/excelmcp/extra.py"
        )
        with td, self.assertRaisesRegex(
            SystemExit,
            "sdist is missing wheel package contents",
        ):
            verify_distribution.verify(project, dist)

    def test_rejects_sdist_module_missing_from_wheel_parity(self):
        td, project, dist = self.make_fixture(
            wheel_omit="excelmcp/extra.py"
        )
        with td, self.assertRaisesRegex(
            SystemExit,
            "wheel is missing sdist package contents",
        ):
            verify_distribution.verify(project, dist)

    def test_rejects_sdist_identity_mismatch(self):
        td, project, dist = self.make_fixture(sdist_version="9.9.9")
        with td, self.assertRaisesRegex(SystemExit, "sdist metadata Version"):
            verify_distribution.verify(project, dist)


if __name__ == "__main__":
    unittest.main()
