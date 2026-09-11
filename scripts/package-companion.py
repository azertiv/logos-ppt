"""Build the Windows portable archive on macOS/Linux/Windows; no compilation.
Official runtime archives must be present in .tmp/downloads. Hashes are pinned.
"""
import hashlib
import json
from pathlib import Path
import shutil
import zipfile

ROOT = Path(__file__).resolve().parent.parent
LOCK = json.loads((ROOT / "bridge/runtime-lock.json").read_text())
OUT = ROOT / "dist-companion"
NAME = "Atelier-Pictos-Codex-Windows-x64"
PACKAGE = OUT / NAME


def sha(path):
    with path.open("rb") as f:
        return hashlib.file_digest(f, "sha256").hexdigest()


def build():
    launcher = ROOT / "bridge/windows/bin/Release/net48/Neurow.Pictos.exe"
    if not launcher.exists():
        raise SystemExit("Compilez d’abord le lanceur : dotnet build bridge/windows/Neurow.Pictos.csproj -c Release")
    for info in (LOCK["node"], LOCK["codex"], *LOCK["licenses"]):
        archive = ROOT / ".tmp/downloads" / info["archive"]
        if not archive.exists():
            raise SystemExit(f"Archive officielle manquante : {archive}\nSource : {info['url']}")
        if sha(archive) != info["sha256"]:
            raise SystemExit(f"Somme de contrôle incorrecte : {archive}")
    if PACKAGE.exists():
        shutil.rmtree(PACKAGE)
    (PACKAGE / "runtime").mkdir(parents=True)
    (PACKAGE / "public").mkdir()
    shutil.copytree(ROOT / "bridge", PACKAGE / "bridge", ignore=shutil.ignore_patterns("__pycache__", "windows"))
    shutil.copy2(launcher, PACKAGE / "Neurow.Pictos.exe")
    shutil.copy2(launcher.with_suffix(".exe.config"), PACKAGE / "Neurow.Pictos.exe.config")
    for name in ["ai-tasks.js", "ai-providers.js"]:
        shutil.copy2(ROOT / "public" / name, PACKAGE / "public" / name)
    with zipfile.ZipFile(ROOT / ".tmp/downloads" / LOCK["node"]["archive"]) as archive:
        prefix = f"node-v{LOCK['node']['version']}-win-x64/"
        for source, dest in [("node.exe", "node.exe"), ("LICENSE", "NODE-LICENSE.txt")]:
            (PACKAGE / "runtime" / dest).write_bytes(archive.read(prefix + source))
    with zipfile.ZipFile(ROOT / ".tmp/downloads" / LOCK["codex"]["archive"]) as archive:
        for source in ["codex-x86_64-pc-windows-msvc.exe", "codex-command-runner.exe", "codex-windows-sandbox-setup.exe"]:
            target = "codex.exe" if source.startswith("codex-x86_64") else source
            (PACKAGE / "runtime" / target).write_bytes(archive.read(source))
    shutil.copy2(ROOT / "docs/COMPAGNON_WINDOWS.md", PACKAGE / "LIRE-MOI.txt")
    shutil.copy2(ROOT / "manifest.xml", PACKAGE / "manifest.xml")
    for name in ["CODEX-LICENSE.txt", "CODEX-NOTICE.txt"]:
        shutil.copy2(ROOT / ".tmp/downloads" / name, PACKAGE / "runtime" / name)
    hashes = {str(p.relative_to(PACKAGE)).replace("\\", "/"): sha(p) for p in sorted(PACKAGE.rglob("*")) if p.is_file()}
    (PACKAGE / "CONTENU-SHA256.json").write_text(json.dumps(hashes, indent=2) + "\n")
    archive_path = OUT / (NAME + ".zip")
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for p in sorted(PACKAGE.rglob("*")):
            if p.is_file():
                archive.write(p, p.relative_to(OUT))
    with zipfile.ZipFile(archive_path) as archive:
        if archive.testzip() is not None:
            raise SystemExit("Archive de sortie endommagée.")
    (OUT / (NAME + ".sha256")).write_text(f"{sha(archive_path)}  {archive_path.name}\n")
    print(f"Archive portable : {archive_path}\nTaille : {archive_path.stat().st_size / 1024 / 1024:.1f} Mio")


if __name__ == "__main__":
    build()
