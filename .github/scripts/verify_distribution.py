from __future__ import annotations

import argparse
import configparser
import email.parser
import pathlib
import tarfile
import tomllib
import zipfile

REQUIRED_PACKAGES = frozenset({
    "excelmcp",
    "pptmcp",
    "wordmcp",
    "wordmcp._com",
    "wordmcp._docx",
    "mcpshared",
})
REQUIRED_SCRIPTS = {
    "excelmcp": "excelmcp.server:main",
    "pptmcp": "pptmcp.server:main",
    "wordmcp": "wordmcp.server:main",
}


def fail(message: str) -> None:
    raise SystemExit(f"distribution verification failed: {message}")


def _source_package_path(package: str, package_dirs: dict[str, str]) -> str:
    candidates = [
        key for key in package_dirs
        if package == key or package.startswith(f"{key}.")
    ]
    if not candidates:
        return package.replace(".", "/")
    root = max(candidates, key=len)
    base = package_dirs[root].rstrip("/")
    suffix = package[len(root):].lstrip(".").replace(".", "/")
    return f"{base}/{suffix}" if suffix else base


def _parse_metadata(raw: bytes):
    return email.parser.Parser().parsestr(raw.decode("utf-8"))


def _verify_identity(metadata, name: str, version: str, label: str) -> None:
    if metadata.get("Name") != name:
        fail(f"{label} metadata Name={metadata.get('Name')!r}, expected {name!r}")
    if metadata.get("Version") != version:
        fail(f"{label} metadata Version={metadata.get('Version')!r}, expected {version!r}")


def verify(
    project_path: pathlib.Path,
    dist_dir: pathlib.Path,
    release_tag: str | None = None,
) -> tuple[str, str, pathlib.Path, pathlib.Path]:
    if not project_path.is_file():
        fail(f"missing project file: {project_path}")
    if not dist_dir.is_dir():
        fail(f"missing distribution directory: {dist_dir}")

    with project_path.open("rb") as handle:
        config = tomllib.load(handle)

    project = config.get("project", {})
    name = project.get("name")
    version = project.get("version")
    if name != "mcp-office":
        fail(f"expected project name 'mcp-office', got {name!r}")
    if not isinstance(version, str) or not version:
        fail("project.version must be a non-empty string")

    if release_tag:
        tag_version = release_tag[1:] if release_tag.startswith("v") else release_tag
        if tag_version != version:
            fail(f"release tag {release_tag!r} does not match project version {version!r}")

    expected_packages = config.get("tool", {}).get("setuptools", {}).get("packages", [])
    configured_package_set = set(expected_packages)
    missing_required_packages = sorted(REQUIRED_PACKAGES - configured_package_set)
    if missing_required_packages:
        fail(
            "configured distribution is missing required suite packages: "
            f"{missing_required_packages}"
        )

    expected_scripts = project.get("scripts", {})
    for script_name, target in REQUIRED_SCRIPTS.items():
        if expected_scripts.get(script_name) != target:
            fail(f"required console script {script_name!r} must map to {target!r}")

    package_dirs = config.get("tool", {}).get("setuptools", {}).get("package-dir", {})

    wheels = sorted(dist_dir.glob("*.whl"))
    sdists = sorted(dist_dir.glob("*.tar.gz"))
    if len(wheels) != 1:
        fail(f"expected exactly one wheel, found {[p.name for p in wheels]}")
    if len(sdists) != 1:
        fail(f"expected exactly one sdist, found {[p.name for p in sdists]}")

    wheel = wheels[0]
    sdist = sdists[0]
    normalized_version = version.replace("-", "_")
    if not wheel.name.startswith(f"mcp_office-{normalized_version}-"):
        fail(f"unexpected wheel filename: {wheel.name}")
    if not sdist.name.startswith(f"mcp_office-{version}"):
        fail(f"unexpected sdist filename: {sdist.name}")

    wheel_files_by_package: dict[str, set[str]] = {}
    with zipfile.ZipFile(wheel) as archive:
        names = archive.namelist()
        metadata_paths = [p for p in names if p.endswith(".dist-info/METADATA")]
        entry_point_paths = [p for p in names if p.endswith(".dist-info/entry_points.txt")]
        if len(metadata_paths) != 1:
            fail(f"expected one METADATA file, found {metadata_paths}")
        if len(entry_point_paths) != 1:
            fail(f"expected one entry_points.txt, found {entry_point_paths}")

        _verify_identity(
            _parse_metadata(archive.read(metadata_paths[0])), name, version, "wheel"
        )

        for package in expected_packages:
            prefix = package.replace(".", "/") + "/"
            package_files = {
                member[len(prefix):]
                for member in names
                if member.startswith(prefix) and not member.endswith("/")
            }
            if not package_files:
                fail(f"wheel is missing configured package {package!r}")
            wheel_files_by_package[package] = package_files

        parser_cfg = configparser.ConfigParser()
        parser_cfg.read_string(archive.read(entry_point_paths[0]).decode("utf-8"))
        installed_scripts = (
            dict(parser_cfg["console_scripts"])
            if parser_cfg.has_section("console_scripts")
            else {}
        )
        for script_name, target in expected_scripts.items():
            if installed_scripts.get(script_name) != target:
                fail(
                    f"console script {script_name!r} maps to "
                    f"{installed_scripts.get(script_name)!r}, expected {target!r}"
                )

    with tarfile.open(sdist, "r:gz") as archive:
        members = archive.getmembers()
        sdist_names = [member.name for member in members]
        if not any(member.endswith("/pyproject.toml") for member in sdist_names):
            fail("sdist is missing pyproject.toml")
        if not any(member.endswith("/README.md") for member in sdist_names):
            fail("sdist is missing README.md")

        pkg_infos = [
            member
            for member in members
            if member.name.endswith("/PKG-INFO") and member.isfile()
        ]
        if len(pkg_infos) != 1:
            fail(f"expected one sdist PKG-INFO, found {[m.name for m in pkg_infos]}")
        pkg_info_handle = archive.extractfile(pkg_infos[0])
        if pkg_info_handle is None:
            fail("cannot read sdist PKG-INFO")
        _verify_identity(_parse_metadata(pkg_info_handle.read()), name, version, "sdist")

        for package in expected_packages:
            source_path = _source_package_path(package, package_dirs).strip("/")
            marker = f"/{source_path}/"
            sdist_package_files: set[str] = set()
            for member in members:
                if not member.isfile():
                    continue
                padded = f"/{member.name.strip('/')}"
                marker_index = padded.find(marker)
                if marker_index >= 0:
                    sdist_package_files.add(padded[marker_index + len(marker):])
            if not sdist_package_files:
                fail(
                    f"sdist is missing configured package {package!r} "
                    f"at {source_path!r}"
                )
            missing_from_sdist = sorted(
                wheel_files_by_package[package] - sdist_package_files
            )
            if missing_from_sdist:
                fail(
                    f"sdist is missing wheel package contents for {package!r}: "
                    f"{missing_from_sdist}"
                )

    return name, version, wheel, sdist


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify the mcp-office release artifacts.")
    parser.add_argument("--dist", default="dist", help="Distribution directory")
    parser.add_argument("--project", default="pyproject.toml", help="Root pyproject.toml")
    parser.add_argument("--release-tag", default=None, help="Optional release tag, e.g. v0.7.0")
    args = parser.parse_args()

    name, version, wheel, sdist = verify(
        pathlib.Path(args.project), pathlib.Path(args.dist), args.release_tag
    )
    print(
        f"distribution verification PASS: name={name} version={version} "
        f"wheel={wheel.name} sdist={sdist.name}"
    )


if __name__ == "__main__":
    main()
