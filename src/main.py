from __future__ import annotations

import os
import shutil
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

from ai_client import draft_issue_reply
from checks import (
    CheckResult,
    StandardRow,
    ad_user_check,
    dhcp_check,
    normalize_ip_value,
    reverse_dns_check,
    validate_row,
)
from config import data_dirs, load_env, load_settings, unique_path
from parser import parse_application, parse_with_ai
from reports import applicant_error_message, confirmation_message, write_text
from xlsx_io import read_workbook, write_standard_workbook


def standard_name(source_path: Path) -> str:
    return f'{source_path.stem}_standard.xlsx'


def report_name(source_path: Path, suffix: str) -> str:
    return f'{source_path.stem}_{suffix}.txt'


def move_source(source_path: Path, target_dir: Path) -> Path:
    target = unique_path(target_dir / source_path.name)
    shutil.move(str(source_path), str(target))
    return target


def validate_rows(rows: List[StandardRow], internet_aliases: List[str]) -> Tuple[List[StandardRow], List[str], List[str]]:
    valid_rows: List[StandardRow] = []
    blockers: List[str] = []
    notes: List[str] = []
    for index, row in enumerate(rows, start=1):
        valid, row_blockers = validate_row(row, internet_aliases)
        if row_blockers:
            blockers.extend(f'行{index}: {item}' for item in row_blockers)
            continue
        if valid is not None:
            valid_rows.append(valid)
    if not valid_rows and not blockers:
        blockers.append('申請データ行を取得できませんでした。')
    return valid_rows, blockers, notes


def _env_value(env: Dict[str, str], key: str, default: str = '') -> str:
    return os.getenv(key, env.get(key, default)).strip()


def run_checks(rows: List[StandardRow], settings: Dict, env: Dict[str, str]) -> CheckResult:
    result = CheckResult(rows=rows)
    internet_aliases = settings.get('internet_aliases', [])
    required_group = _env_value(env, 'REQUIRED_SECURITY_GROUP', settings.get('required_security_group', ''))
    dhcp_server = _env_value(env, 'DHCP_SERVER')
    dhcp_username = _env_value(env, 'DHCP_QUERY_USERNAME')
    dhcp_password = _env_value(env, 'DHCP_QUERY_PASSWORD')
    dhcp_domain = _env_value(env, 'DHCP_QUERY_DOMAIN')

    for row in rows:
        normalized_ip, ip_kind, ip_error = normalize_ip_value(row.IP, internet_aliases)
        if ip_error or not normalized_ip or not ip_kind:
            result.blockers.append(f'{row.userID}: {ip_error or "IPアドレスを確認できません。"}')
            continue
        row.IP = normalized_ip

        ad_blockers, ad_confirmations = ad_user_check(row.userID, required_group)
        result.blockers.extend(ad_blockers)
        result.confirmations.extend(ad_confirmations)
        if ad_blockers:
            continue

        result.confirmations.extend(reverse_dns_check(row, ip_kind))
        result.confirmations.extend(dhcp_check(row, ip_kind, dhcp_server, dhcp_username, dhcp_password, dhcp_domain))

    return result


def write_success_outputs(source_path: Path, target_dir: Path, settings: Dict, rows: List[StandardRow]) -> Path:
    headers = settings['standard_columns']
    workbook_path = unique_path(target_dir / standard_name(source_path))
    write_standard_workbook(workbook_path, headers, [row.as_dict() for row in rows])
    return workbook_path


def _report_text(source_path: Path, issues: Iterable[str], report_kind: str, env: Dict[str, str]) -> str:
    issue_list = list(issues)
    ai_text = draft_issue_reply(source_path.name, issue_list, report_kind, env)
    if report_kind == 'error':
        fallback = applicant_error_message(source_path.name, issue_list)
        detail_title = '検出された不備:'
    else:
        fallback = confirmation_message(source_path.name, issue_list)
        detail_title = '検出された確認事項:'

    if not ai_text:
        return fallback

    detail = '\n'.join(['', '---', detail_title, *[f'- {item}' for item in issue_list], ''])
    return ai_text + detail


def write_error_report(source_path: Path, target_dir: Path, issues: Iterable[str], env: Dict[str, str]) -> None:
    write_text(unique_path(target_dir / report_name(source_path, 'error')), _report_text(source_path, issues, 'error', env))


def write_confirmation_report(source_path: Path, target_dir: Path, issues: Iterable[str], env: Dict[str, str]) -> None:
    write_text(unique_path(target_dir / report_name(source_path, 'check')), _report_text(source_path, issues, 'confirmation', env))


def process_file(source_path: Path, settings: Dict, env: Dict[str, str], dirs: Dict[str, Path]) -> str:
    sheets = read_workbook(source_path)
    parse_result = parse_application(source_path, sheets, env)
    if parse_result.blockers:
        target_dir = dirs['error']
        write_error_report(source_path, target_dir, parse_result.blockers, env)
        move_source(source_path, target_dir)
        return 'error'

    rows, validation_blockers, _ = validate_rows(parse_result.rows, settings.get('internet_aliases', []))
    if validation_blockers and (_env_value(env, 'OPENROUTER_API_KEY') or _env_value(env, 'OPENAI_API_KEY')) and not parse_result.used_ai:
        ai_result = parse_with_ai(source_path, env)
        if not ai_result.blockers:
            rows, validation_blockers, _ = validate_rows(ai_result.rows, settings.get('internet_aliases', []))

    if validation_blockers:
        target_dir = dirs['error']
        write_error_report(source_path, target_dir, validation_blockers, env)
        move_source(source_path, target_dir)
        return 'error'

    check_result = run_checks(rows, settings, env)
    if check_result.blockers:
        target_dir = dirs['error']
        write_error_report(source_path, target_dir, check_result.blockers, env)
        move_source(source_path, target_dir)
        return 'error'

    if check_result.confirmations:
        target_dir = dirs['needs_confirmation']
        write_success_outputs(source_path, target_dir, settings, rows)
        write_confirmation_report(source_path, target_dir, check_result.confirmations, env)
        move_source(source_path, target_dir)
        return 'needs_confirmation'

    target_dir = dirs['network_ready']
    write_success_outputs(source_path, target_dir, settings, rows)
    move_source(source_path, target_dir)
    return 'network_ready'


def main() -> None:
    env = load_env()
    settings = load_settings()
    dirs = data_dirs(settings)
    source_dir = dirs['source']
    files = sorted(path for path in source_dir.glob('*.xlsx') if path.is_file() and not path.name.startswith('~$'))
    if not files:
        print(f'No .xlsx files found in {source_dir}')
        return

    stats = Counter()
    for source_path in files:
        print(f'Processing {source_path.name}...')
        try:
            result = process_file(source_path, settings, env, dirs)
        except Exception as exc:
            target_dir = dirs['error']
            write_error_report(source_path, target_dir, [f'処理中に予期しないエラーが発生しました: {exc}'], env)
            move_source(source_path, target_dir)
            result = 'error'
        stats[result] += 1
        print(f'  -> {result}')

    print('Summary:', dict(stats))


if __name__ == '__main__':
    main()


