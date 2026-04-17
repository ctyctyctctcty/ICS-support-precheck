from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Dict, Any

ROOT = Path(__file__).resolve().parents[1]


def load_env() -> Dict[str, str]:
    env_path = ROOT / 'config' / '.env'
    values: Dict[str, str] = {}
    if env_path.exists():
        for raw in env_path.read_text(encoding='utf-8').splitlines():
            line = raw.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            key, value = line.split('=', 1)
            values[key.strip()] = value.strip().strip('"')
    for key, value in values.items():
        os.environ.setdefault(key, value)
    return values


def load_settings() -> Dict[str, Any]:
    settings_path = ROOT / 'config' / 'settings.json'
    with settings_path.open('r', encoding='utf-8-sig') as fp:
        settings = json.load(fp)
    return settings


def resolve_path(path_value: str) -> Path:
    path = Path(path_value)
    if path.is_absolute():
        return path
    return ROOT / path


def data_dirs(settings: Dict[str, Any]) -> Dict[str, Path]:
    result = {name: resolve_path(value) for name, value in settings['data_dirs'].items()}
    for path in result.values():
        path.mkdir(parents=True, exist_ok=True)
    return result


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 1
    while True:
        candidate = parent / f'{stem}_{index}{suffix}'
        if not candidate.exists():
            return candidate
        index += 1