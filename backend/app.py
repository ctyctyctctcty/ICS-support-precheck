from __future__ import annotations

import shutil
import sys
from pathlib import Path
from threading import Lock
from typing import Dict, List
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'src'
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from config import data_dirs, load_env, load_settings, unique_path  # noqa: E402
from main import move_source, process_file, write_error_report  # noqa: E402

RESULT_BUCKETS = {'network_ready', 'needs_confirmation', 'error'}
PROCESS_LOCK = Lock()

app = FastAPI(title='ICS Support Precheck')

app.add_middleware(
    CORSMiddleware,
    allow_origins=['http://localhost:5173', 'http://127.0.0.1:5173'],
    allow_credentials=True,
    allow_methods=['*'],
    allow_headers=['*'],
)


def _runtime() -> tuple[Dict, Dict[str, str], Dict[str, Path]]:
    env = load_env()
    settings = load_settings()
    dirs = data_dirs(settings)
    return settings, env, dirs


def _folder_snapshot(path: Path) -> set[str]:
    if not path.exists():
        return set()
    return {item.name for item in path.iterdir() if item.is_file() and item.name != '.gitkeep'}


def _result_files(bucket: str, names: List[str]) -> List[Dict[str, str]]:
    return [
        {
            'name': name,
            'bucket': bucket,
            'download_url': f'/api/files/{bucket}/{quote(name)}',
        }
        for name in sorted(names)
    ]


def _save_upload(upload: UploadFile, source_dir: Path) -> Path:
    original_name = Path(upload.filename or '').name
    if not original_name:
        raise HTTPException(status_code=400, detail='File name is missing.')
    if original_name.startswith('~$') or Path(original_name).suffix.lower() != '.xlsx':
        raise HTTPException(status_code=400, detail=f'Only normal .xlsx files are supported: {original_name}')

    target = unique_path(source_dir / original_name)
    with target.open('wb') as fp:
        shutil.copyfileobj(upload.file, fp)
    return target


@app.get('/api/health')
def health() -> Dict[str, object]:
    _, _, dirs = _runtime()
    return {
        'ok': True,
        'folders': {key: str(value) for key, value in dirs.items()},
    }


@app.post('/api/process')
def process_uploads(files: List[UploadFile] = File(...)) -> Dict[str, object]:
    if not files:
        raise HTTPException(status_code=400, detail='No files were uploaded.')

    results = []
    with PROCESS_LOCK:
        settings, env, dirs = _runtime()
        for upload in files:
            before = {bucket: _folder_snapshot(dirs[bucket]) for bucket in RESULT_BUCKETS}
            source_path = _save_upload(upload, dirs['source'])
            try:
                status = process_file(source_path, settings, env, dirs)
            except Exception as exc:
                if source_path.exists():
                    write_error_report(source_path, dirs['error'], [f'Unexpected error occurred while processing: {exc}'], env)
                    move_source(source_path, dirs['error'])
                status = 'error'

            after = _folder_snapshot(dirs[status])
            new_names = sorted(after - before[status])
            results.append({
                'original_name': upload.filename,
                'status': status,
                'folder': str(dirs[status]),
                'files': _result_files(status, new_names),
            })

    return {'results': results}


@app.get('/api/files/{bucket}/{file_name}')
def download_file(bucket: str, file_name: str) -> FileResponse:
    if bucket not in RESULT_BUCKETS:
        raise HTTPException(status_code=404, detail='Unknown output folder.')

    _, _, dirs = _runtime()
    file_path = (dirs[bucket] / Path(file_name).name).resolve()
    bucket_path = dirs[bucket].resolve()
    if bucket_path not in file_path.parents or not file_path.is_file():
        raise HTTPException(status_code=404, detail='File not found.')
    return FileResponse(file_path, filename=file_path.name)


dist_dir = ROOT / 'frontend' / 'dist'
if dist_dir.exists():
    app.mount('/', StaticFiles(directory=dist_dir, html=True), name='frontend')
