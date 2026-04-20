from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Dict, List, Optional, Sequence, Tuple

from ai_client import AIError, format_ai_errors, load_ai_config, model_candidates, chat_json
from checks import StandardRow
from xlsx_io import SheetData, workbook_as_text

HEADER_ALIASES = {
    'account': ['対象アカウント', 'アカウント', 'ユーザーid', 'ユーザid', 'userid', 'user id', 'userID', 'アカウントid'],
    'name': ['対象者氏名', '氏名', '利用者名', '対象者名', '名前', 'ユーザー名', 'name', 'full name', 'user name', 'target name'],
    'company': ['対象者会社名', '会社名', '所属会社', '会社', 'ベンダー', '協力会社', 'company', 'vendor'],
    'email': ['対象者メールアドレス', 'メールアドレス', 'メール', 'email', 'e-mail'],
    'hostname': ['接続先サーバー名', 'サーバー名', '接続先server', 'hostname', 'host name', 'ホスト名', '接続先ホスト名'],
    'ip': ['ipアドレス', 'ip address', 'ip', '接続先ip', '接続先ipアドレス'],
    'admin_name': ['管理者氏名', '管理者名', '管理者'],
}

SAME_AS_ABOVE = {'同上', '同じ', '同', '〃', '上記同様', '前行同様'}

AI_ROWS_SCHEMA = {
    'type': 'object',
    'properties': {
        'rows': {
            'type': 'array',
            'items': {
                'type': 'object',
                'properties': {
                    'account': {'type': 'string', 'description': 'Target account or user ID.'},
                    'name': {'type': 'string', 'description': 'Applicant or target user full name.'},
                    'company': {'type': 'string', 'description': 'Company name.'},
                    'email': {'type': 'string', 'description': 'Email address.'},
                    'hostname': {'type': 'string', 'description': 'Destination server or host name.'},
                    'ip': {'type': 'string', 'description': 'Destination IPv4, CIDR, or Internet Access.'},
                    'admin_name': {'type': 'string', 'description': 'Manager or administrator name, if present.'},
                },
                'required': ['account', 'name', 'company', 'email', 'hostname', 'ip', 'admin_name'],
                'additionalProperties': False,
            },
        },
    },
    'required': ['rows'],
    'additionalProperties': False,
}


@dataclass
class ParseResult:
    rows: List[StandardRow]
    blockers: List[str]
    used_ai: bool = False
    ai_model: str = ''


def normalize_header(value: str) -> str:
    text = str(value or '').lower()
    text = text.replace('\n', '').replace('\r', '').replace(' ', '').replace('　', '')
    text = re.sub(r'[（）()\[\]【】「」『』:：\\._\-]', '', text)
    return text


def _alias_lookup() -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    for key, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            lookup[normalize_header(alias)] = key
    return lookup


def _cell_key(value: str, lookup: Dict[str, str]) -> Optional[str]:
    norm = normalize_header(value)
    if norm in lookup:
        return lookup[norm]
    for alias, key in lookup.items():
        if alias and alias in norm:
            return key
    return None


def _is_example_row(row: Sequence[str]) -> bool:
    return any('例' in str(cell) or 'サンプル' in str(cell).lower() for cell in row)


def _same_as_above(value: str) -> bool:
    return str(value or '').strip() in SAME_AS_ABOVE


def parse_with_rules(sheets: Sequence[SheetData]) -> ParseResult:
    lookup = _alias_lookup()
    best: Optional[Tuple[SheetData, int, Dict[str, int]]] = None
    best_score = 0

    for sheet in sheets:
        for row_idx, row in enumerate(sheet.rows):
            mapping: Dict[str, int] = {}
            for col_idx, cell in enumerate(row):
                key = _cell_key(cell, lookup)
                if key and key not in mapping:
                    mapping[key] = col_idx
            score = len(mapping)
            if score > best_score:
                best = (sheet, row_idx, mapping)
                best_score = score

    if best is None or best_score < 3:
        return ParseResult([], ['申請書のヘッダー行を認識できませんでした。'])

    sheet, header_idx, mapping = best
    if not {'account', 'name', 'ip'}.issubset(mapping):
        missing = {'account': '対象アカウント', 'name': '対象者氏名', 'ip': 'IPアドレス'}
        not_found = [label for key, label in missing.items() if key not in mapping]
        return ParseResult([], [f'必須ヘッダーを認識できませんでした: {", ".join(not_found)}'])

    parsed: List[StandardRow] = []
    previous: Dict[str, str] = {}
    for row in sheet.rows[header_idx + 1:]:
        if _is_example_row(row):
            continue
        if not any(str(cell).strip() for cell in row):
            continue

        values: Dict[str, str] = {}
        for key, col_idx in mapping.items():
            raw = row[col_idx].strip() if col_idx < len(row) else ''
            if _same_as_above(raw):
                raw = previous.get(key, '')
            values[key] = raw

        if not any(values.get(key, '') for key in ('account', 'name', 'company', 'email', 'hostname', 'ip')):
            continue

        previous.update({key: value for key, value in values.items() if value})
        parsed.append(StandardRow(
            userID=values.get('account', ''),
            name=values.get('name', ''),
            company=values.get('company', ''),
            email=values.get('email', ''),
            hostname=values.get('hostname', ''),
            IP=values.get('ip', ''),
        ))

    if not parsed:
        return ParseResult([], ['申請データ行を取得できませんでした。'])
    return ParseResult(parsed, [])


def _ai_messages(workbook_text: str) -> List[Dict[str, str]]:
    return [
        {
            'role': 'system',
            'content': (
                'You are a careful data extraction engine for Japanese VPN application Excel files. '
                'Return only fields that are present or directly implied by same-as-above notation. '
                'Do not invent accounts, names, hostnames, emails, companies, or IP addresses.'
            ),
        },
        {
            'role': 'user',
            'content': f'''
Extract VPN application rows from this Japanese Excel text.

Rules:
- Put Japanese Internet Access-like values in ip as exactly "Internet Access".
- Carry forward "同上", "〃", or similar same-as-above values from the previous data row.
- Ignore example/sample rows.
- If a value is missing, use an empty string.
- Return JSON that matches the schema.

Excel text:
{workbook_text}
'''.strip(),
        },
    ]


def parse_with_ai(path: Path, env: Dict[str, str]) -> ParseResult:
    try:
        config = load_ai_config(env)
    except AIError as exc:
        return ParseResult([], [f'AI設定が正しくありません: {exc}'])

    if not config.enabled:
        return ParseResult([], ['AI解析用のAPIキーが未設定です。'])

    workbook_text = workbook_as_text(path)
    errors: List[str] = []
    for model in model_candidates(config.parse_model, config.parse_fallback_model):
        try:
            data = chat_json(config, model, _ai_messages(workbook_text), AI_ROWS_SCHEMA, 'vpn_application_rows')
        except AIError as exc:
            errors.append(f'{model}: {exc}')
            continue

        items = data.get('rows', [])
        if not isinstance(items, list):
            errors.append(f'{model}: rows が配列ではありません。')
            continue

        rows: List[StandardRow] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            rows.append(StandardRow(
                userID=str(item.get('account', '') or ''),
                name=str(item.get('name', '') or ''),
                company=str(item.get('company', '') or ''),
                email=str(item.get('email', '') or ''),
                hostname=str(item.get('hostname', '') or ''),
                IP=str(item.get('ip', '') or ''),
            ))
        if rows:
            return ParseResult(rows, [], used_ai=True, ai_model=model)
        errors.append(f'{model}: 申請データ行を取得できませんでした。')

    return ParseResult([], [f'AI解析に失敗しました: {format_ai_errors(errors)}'])


def parse_application(path: Path, sheets: Sequence[SheetData], env: Dict[str, str]) -> ParseResult:
    rule_result = parse_with_rules(sheets)
    if not rule_result.blockers:
        return rule_result
    ai_result = parse_with_ai(path, env)
    if not ai_result.blockers:
        return ai_result
    return ParseResult([], rule_result.blockers + ai_result.blockers)



