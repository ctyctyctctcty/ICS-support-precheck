from __future__ import annotations

from pathlib import Path
from typing import Iterable


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8')


def applicant_error_message(file_name: str, blockers: Iterable[str]) -> str:
    lines = [
        '申請書の内容を確認したところ、以下の不備がありました。',
        'お手数ですが、修正のうえ再提出をお願いいたします。',
        '',
        f'対象ファイル: {file_name}',
        '',
        '不備内容:',
    ]
    lines.extend(f'- {item}' for item in blockers)
    lines.append('')
    return '\n'.join(lines)


def confirmation_message(file_name: str, confirmations: Iterable[str]) -> str:
    lines = [
        'この申請書は標準フォーマットへの変換が完了していますが、以下の確認事項があります。',
        '申請者へ確認後、問題なければ生成された standard xlsx を network team へ連携してください。',
        '',
        f'対象ファイル: {file_name}',
        '',
        '確認事項:',
    ]
    lines.extend(f'- {item}' for item in confirmations)
    lines.append('')
    return '\n'.join(lines)
