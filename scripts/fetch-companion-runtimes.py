"""Fetch only the pinned, hash-verified runtime files needed for the Windows ZIP."""
import hashlib
import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent.parent
LOCK = json.loads((ROOT / 'bridge/runtime-lock.json').read_text())
DOWNLOADS = ROOT / '.tmp/downloads'


def digest(path):
    with path.open('rb') as file:
        return hashlib.file_digest(file, 'sha256').hexdigest()


def main():
    DOWNLOADS.mkdir(parents=True, exist_ok=True)
    for item in [LOCK['node'], LOCK['codex'], *LOCK['licenses']]:
        target = DOWNLOADS / item['archive']
        if target.exists() and digest(target) == item['sha256']:
            continue
        temporary = target.with_suffix(target.suffix + '.part')
        subprocess.run(['curl', '--fail', '--silent', '--show-error', '--location', '--retry', '3', item['url'], '-o', str(temporary)], check=True)
        if digest(temporary) != item['sha256']:
            temporary.unlink()
            raise SystemExit('Empreinte incorrecte : ' + item['archive'])
        temporary.replace(target)
        print('Vérifié : ' + item['archive'])


if __name__ == '__main__':
    main()
