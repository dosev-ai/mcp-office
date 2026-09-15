from __future__ import annotations

import argparse
import configparser
import email.parser
import pathlib
import tarfile
import tomllib
import zipfile


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

    expected_packages = config.get("tool", {}).get("setuptools", {}).get("packages", [])
    if not expected_packages:
        fail("tool.setuptools.packages is empty")
    package_dirs = config.get("tool", {}).get("setuptools", {}).get("package-dir", {})

    with zipfile.ZipFile(wheel) as archive:
        names = archive.namelist()
        metadata_paths = [p for p in names if p.endswith(".dist-info/METADATA")]
        entry_point_paths = [p for p in names if p.endswith(".dist-info/entry_points.txt")]
        if len(metadata_paths) != 1:
            fail(f"expected one METADATA file, found {metadata_paths}")
        if len(entry_point_paths) != 1:
            fail(f"expected one entry_points.txt, found {entry_point_paths}")

        metadata = email.parser.Parser().parsestr(
            archive.read(metadata_paths[0]).decode("utf-8")
        )
        if metadata.get("Name") != name:
            fail(f"wheel metadata Name={metadata.get('Name')!r}, expected {name!r}")
        if metadata.get("Version") != version:
            fail(f"wheel metadata Version={metadata.get('Version')!r}, expected {version!r}")

        for package in expected_packages:
            prefix = package.replace(".", "/") + "/"
            if not any(member.startswith(prefix) for member in names):
                fail(f"wheel is missing configured package {package!r}")

        expected_scripts = project.get("scripts", {})
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
        sdist_names = archive.getnames()
        if not any(member.endswith("/pyproject.toml") for member in sdist_names):
            fail("sdist is missing pyproject.toml")
        if not any(member.endswith("/README.md") for member in sdist_names):
            fail("sdist is missing README.md")
        normalized_members = [f"/{member.strip('/')}/" for member in sdist_names]
        for package in expected_packages:
            source_path = _source_package_path(package, package_dirs).strip("/")
            marker = f"/{source_path}/"
            if not any(marker in member for member in normalized_members):
                fail(
                    f"sdist is missing configured package {package!r} "
                    f"at {source_path!r}"
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
