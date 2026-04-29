from __future__ import annotations

import csv
import ipaddress
import json
import re
import socket
import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
ACCOUNT_RE = re.compile(r'^[A-Za-z0-9._-]{1,64}$')
HOSTNAME_RE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9.-]{0,253}$')
DHCP_STATUS_FILE = 'dhcp_export_status.json'


@dataclass
class StandardRow:
    userID: str
    name: str
    company: str = ''
    email: str = ''
    hostname: str = ''
    IP: str = ''

    def as_dict(self) -> Dict[str, str]:
        return {
            'userID': self.userID,
            'name': self.name,
            'company': self.company,
            'email': self.email,
            'hostname': self.hostname,
            'IP': self.IP,
        }


@dataclass
class CheckResult:
    rows: List[StandardRow] = field(default_factory=list)
    blockers: List[str] = field(default_factory=list)
    confirmations: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    @property
    def has_blockers(self) -> bool:
        return bool(self.blockers)

    @property
    def has_confirmations(self) -> bool:
        return bool(self.confirmations)


@dataclass(frozen=True)
class DhcpRange:
    start: ipaddress.IPv4Address
    end: ipaddress.IPv4Address
    source: str = ''
    name: str = ''
    range_type: str = 'scope'

    def contains(self, ip: ipaddress.IPv4Address) -> bool:
        return int(self.start) <= int(ip) <= int(self.end)

    def label(self) -> str:
        label_parts = []
        if self.name:
            label_parts.append(self.name)
        label_parts.append(f'{self.start}-{self.end}')
        if self.range_type and self.range_type != 'scope':
            label_parts.append(self.range_type)
        if self.source:
            label_parts.append(f'source: {self.source}')
        return ' / '.join(label_parts)


@dataclass
class DhcpReferenceState:
    ranges: List[DhcpRange] = field(default_factory=list)
    exclusions: List[DhcpRange] = field(default_factory=list)
    issues: List[str] = field(default_factory=list)
    latest_export_label: str = ''


def compact_text(value: str) -> str:
    return re.sub(r'\s+', '', str(value or '')).lower()


def normalize_internet(value: str, aliases: Sequence[str]) -> Optional[str]:
    text = compact_text(value)
    alias_set = {compact_text(alias) for alias in aliases}
    if text in alias_set:
        return 'Internet Access'
    return None


def normalize_ip_value(value: str, internet_aliases: Sequence[str]) -> tuple[Optional[str], Optional[str], Optional[str]]:
    raw = str(value or '').strip()
    if not raw:
        return None, None, 'IP address is missing.'

    internet = normalize_internet(raw, internet_aliases)
    if internet:
        return internet, 'internet', None

    try:
        if '/' in raw:
            network = ipaddress.ip_network(raw, strict=False)
            if network.version != 4:
                return None, None, 'Only IPv4 CIDR ranges are supported.'
            if network.prefixlen == 32:
                return str(network.network_address), 'ip', None
            return str(network), 'cidr', None
        ip = ipaddress.ip_address(raw)
        if ip.version != 4:
            return None, None, 'Only IPv4 addresses are supported.'
        return str(ip), 'ip', None
    except ValueError:
        return None, None, f'Invalid IP address format: {raw}'


def validate_row(row: StandardRow, internet_aliases: Sequence[str]) -> tuple[Optional[StandardRow], List[str]]:
    blockers: List[str] = []
    user_id = row.userID.strip()
    name = row.name.strip()
    ip_value, _, ip_error = normalize_ip_value(row.IP, internet_aliases)

    if not user_id:
        blockers.append('Target account is missing.')
    elif not ACCOUNT_RE.fullmatch(user_id):
        blockers.append(f'Invalid target account format: {user_id}')

    if not name:
        blockers.append('Target user name is missing.')

    if ip_error:
        blockers.append(ip_error)

    email = row.email.strip()
    if email and not EMAIL_RE.fullmatch(email):
        blockers.append(f'Invalid target user email address format: {email}')

    hostname = row.hostname.strip()
    if hostname and not HOSTNAME_RE.fullmatch(hostname):
        blockers.append(f'Invalid destination hostname format: {hostname}')

    if blockers:
        return None, blockers

    return StandardRow(
        userID=user_id,
        name=name,
        company=row.company.strip(),
        email=email,
        hostname=hostname,
        IP=ip_value or row.IP.strip(),
    ), []


def short_hostname(value: str) -> str:
    return str(value or '').strip().rstrip('.').split('.')[0].lower()


def jp_item(user_id: str, summary: str, detail: str = '') -> str:
    prefix = f'{user_id}: ' if user_id else ''
    if detail:
        return f'{prefix}{summary} ({detail})'
    return f'{prefix}{summary}'


_NSLOOKUP_NAME_PATTERNS = (
    re.compile(r'^\s*name\s*=\s*(.+?)\s*$', re.IGNORECASE),
    re.compile(r'^\s*name\s*:\s*(.+?)\s*$', re.IGNORECASE),
    re.compile(r'^\s*名前\s*:\s*(.+?)\s*$'),
)


def _extract_nslookup_names(output: str) -> List[str]:
    names: List[str] = []
    seen = set()
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        for pattern in _NSLOOKUP_NAME_PATTERNS:
            match = pattern.match(line)
            if not match:
                continue
            candidate = match.group(1).strip().rstrip('.')
            if candidate and candidate.lower() not in seen:
                names.append(candidate)
                seen.add(candidate.lower())
            break
    return names


def _reverse_lookup_names(ip: str) -> tuple[List[str], Optional[str]]:
    names: List[str] = []
    seen = set()

    try:
        primary_name, aliases, _ = socket.gethostbyaddr(ip)
        for candidate in [primary_name, *aliases]:
            normalized = str(candidate or '').strip().rstrip('.')
            if normalized and normalized.lower() not in seen:
                names.append(normalized)
                seen.add(normalized.lower())
    except Exception:
        pass

    try:
        completed = subprocess.run(
            ['nslookup', ip],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except Exception as exc:
        return names, str(exc) if not names else None

    output = (completed.stdout or '') + '\n' + (completed.stderr or '')
    parsed_names = _extract_nslookup_names(output)
    for candidate in parsed_names:
        if candidate.lower() not in seen:
            names.append(candidate)
            seen.add(candidate.lower())

    if names:
        return names, None

    if completed.returncode != 0:
        return [], f'nslookup rc={completed.returncode}'
    return [], None


def reverse_dns_check(row: StandardRow, ip_kind: str) -> List[str]:
    if ip_kind != 'ip' or not row.hostname.strip():
        return []

    names, lookup_error = _reverse_lookup_names(row.IP)
    if not names:
        detail = f'IP={row.IP}, 申請ホスト名={row.hostname}, 手動確認'
        if lookup_error:
            detail = f'{detail}, detail={lookup_error}'
        return [jp_item(row.userID, '逆引きホスト名なし', detail)]

    requested = short_hostname(row.hostname)
    if any(short_hostname(name) == requested for name in names):
        return []
    return [jp_item(row.userID, 'ホスト名一致しない', f'申請={row.hostname}, 逆引き={", ".join(names)}')]


def _first_value(record: Dict[str, Any], keys: Iterable[str]) -> str:
    lower_map = {str(key).lower().replace('_', '').replace('-', ''): value for key, value in record.items()}
    for key in keys:
        normalized = key.lower().replace('_', '').replace('-', '')
        value = lower_map.get(normalized)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ''


def _parse_ipv4(value: str) -> Optional[ipaddress.IPv4Address]:
    try:
        ip = ipaddress.ip_address(str(value).strip())
    except ValueError:
        return None
    if ip.version != 4:
        return None
    return ip


def _range_type_from_record(record: Dict[str, Any], default: str = 'scope') -> str:
    value = _first_value(record, ['RangeType', 'Type', 'Kind', 'RecordType'])
    text = compact_text(value)
    if text in {'exclusion', 'excluded', 'exclude', '除外'}:
        return 'exclusion'
    if text in {'scope', 'range', 'dhcp'}:
        return 'scope'
    return default


def _range_from_record(record: Dict[str, Any], source: str, default_type: str = 'scope') -> Optional[DhcpRange]:
    start_text = _first_value(record, ['StartRange', 'Start', 'RangeStart', 'StartAddress', 'From'])
    end_text = _first_value(record, ['EndRange', 'End', 'RangeEnd', 'EndAddress', 'To'])
    name = _first_value(record, ['Name', 'ScopeName', 'ScopeId', 'Subnet', 'Network'])
    range_type = _range_type_from_record(record, default_type)

    if not start_text or not end_text:
        cidr_text = _first_value(record, ['CIDR', 'Cidr', 'NetworkCidr', 'Prefix'])
        if cidr_text:
            try:
                network = ipaddress.ip_network(cidr_text, strict=False)
            except ValueError:
                return None
            if network.version != 4:
                return None
            return DhcpRange(network.network_address, network.broadcast_address, source=source, name=name or str(network), range_type=range_type)
        return None

    start = _parse_ipv4(start_text)
    end = _parse_ipv4(end_text)
    if start is None or end is None:
        return None
    if int(start) > int(end):
        start, end = end, start
    return DhcpRange(start, end, source=source, name=name, range_type=range_type)


def _json_records(data: Any) -> List[Dict[str, Any]]:
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if isinstance(data, dict):
        for key in ('scopes', 'ranges', 'dhcp_ranges', 'DhcpRanges', 'Scopes'):
            value = data.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, dict)]
        return [data]
    return []


def _json_records_by_type(data: Any) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    if isinstance(data, dict):
        scope_records: List[Dict[str, Any]] = []
        exclusion_records: List[Dict[str, Any]] = []
        for key in ('scopes', 'ranges', 'dhcp_ranges', 'DhcpRanges', 'Scopes'):
            value = data.get(key)
            if isinstance(value, list):
                scope_records.extend(item for item in value if isinstance(item, dict))
        for key in ('exclusions', 'excluded_ranges', 'Exclusions', 'ExcludedRanges'):
            value = data.get(key)
            if isinstance(value, list):
                exclusion_records.extend(item for item in value if isinstance(item, dict))
        if scope_records or exclusion_records:
            return scope_records, exclusion_records
    return _json_records(data), []


def _candidate_dhcp_files(path: Path) -> List[Path]:
    if path.is_dir():
        return sorted([
            file_path
            for file_path in [*path.glob('*.csv'), *path.glob('*.json')]
            if file_path.name.lower() != DHCP_STATUS_FILE
        ])
    return [path]


def _file_is_stale(path: Path, max_age_days: int) -> bool:
    try:
        age_seconds = max(0.0, time.time() - path.stat().st_mtime)
    except OSError:
        return True
    return age_seconds > max_age_days * 86400


def _validate_status_file(path: Path, data_files: List[Path], max_age_days: int) -> tuple[List[str], str]:
    issues: List[str] = []
    latest_export_label = ''
    if not path.exists():
        return issues, latest_export_label

    try:
        status = json.loads(path.read_text(encoding='utf-8-sig'))
    except Exception as exc:
        return [f'DHCP status file could not be read: {path.name} ({exc})'], latest_export_label

    exported_at = str(status.get('exported_at', '') or '').strip()
    scope_count = int(status.get('scope_count', 0) or 0)
    success = bool(status.get('success'))
    latest_export_label = exported_at or path.name

    if not success:
        error_text = str(status.get('error', '') or 'unknown export error')
        issues.append(f'DHCP export status indicates failure: {error_text}')
    if scope_count <= 0:
        issues.append('DHCP export status indicates zero scopes.')

    file_names = [str(name).strip() for name in status.get('data_files', []) if str(name).strip()]
    if file_names:
        existing_names = {file_path.name for file_path in data_files}
        missing = [name for name in file_names if name not in existing_names]
        if missing:
            issues.append(f'DHCP export status references missing files: {", ".join(missing)}')

    if _file_is_stale(path, max_age_days):
        issues.append(f'DHCP export status is older than {max_age_days} days.')
    return issues, latest_export_label


def _load_dhcp_ranges_from_file(path: Path) -> List[DhcpRange]:
    suffix = path.suffix.lower()
    ranges: List[DhcpRange] = []
    if suffix == '.json':
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        scope_records, exclusion_records = _json_records_by_type(data)
        for record in scope_records:
            parsed = _range_from_record(record, path.name, 'scope')
            if parsed:
                ranges.append(parsed)
        for record in exclusion_records:
            parsed = _range_from_record(record, path.name, 'exclusion')
            if parsed:
                ranges.append(parsed)
    elif suffix == '.csv':
        with path.open('r', encoding='utf-8-sig', newline='') as fp:
            reader = csv.DictReader(fp)
            for record in reader:
                parsed = _range_from_record(record, path.name)
                if parsed:
                    ranges.append(parsed)
    return ranges


def load_dhcp_ranges(reference_path: str) -> List[DhcpRange]:
    raw = str(reference_path or '').strip().strip('"')
    if not raw:
        return []
    path = Path(raw).expanduser()
    if not path.exists():
        return []

    files = _candidate_dhcp_files(path)
    ranges: List[DhcpRange] = []
    for file_path in files:
        try:
            ranges.extend(item for item in _load_dhcp_ranges_from_file(file_path) if item.range_type != 'exclusion')
        except Exception:
            continue
    return ranges


def load_dhcp_reference_ranges(reference_path: str) -> tuple[List[DhcpRange], List[DhcpRange]]:
    raw = str(reference_path or '').strip().strip('"')
    if not raw:
        return [], []
    path = Path(raw).expanduser()
    if not path.exists():
        return [], []

    scopes: List[DhcpRange] = []
    exclusions: List[DhcpRange] = []
    for file_path in _candidate_dhcp_files(path):
        try:
            for item in _load_dhcp_ranges_from_file(file_path):
                if item.range_type == 'exclusion':
                    exclusions.append(item)
                else:
                    scopes.append(item)
        except Exception:
            continue
    return scopes, exclusions


def validate_dhcp_reference(reference_path: str, max_age_days: int = 35) -> DhcpReferenceState:
    raw = str(reference_path or '').strip().strip('"')
    if not raw:
        return DhcpReferenceState(issues=['DHCP参照未設定 (DHCP範囲確認は手動対応)'])

    path = Path(raw).expanduser()
    if not path.exists():
        return DhcpReferenceState(issues=[f'DHCP参照ファイルなし (path={reference_path})'])

    files = _candidate_dhcp_files(path)
    if not files:
        return DhcpReferenceState(issues=['DHCP参照ファイルなし (CSV/JSONが見つかりません)'])

    issues: List[str] = []
    if all(_file_is_stale(file_path, max_age_days) for file_path in files):
        issues.append(f'DHCP参照データ期限切れ ({max_age_days}日超)')

    status_path = path / DHCP_STATUS_FILE if path.is_dir() else path.with_name(DHCP_STATUS_FILE)
    status_issues, latest_export_label = _validate_status_file(status_path, files, max_age_days)
    issues.extend(status_issues)

    ranges, exclusions = load_dhcp_reference_ranges(reference_path)
    if not ranges:
        issues.append('DHCP範囲読み取り不可 (有効なStartRange/EndRangeがありません)')
    elif len(ranges) < len(files):
        issues.append('DHCP参照データ一部読取失敗 (一部ファイルを確認してください)')

    return DhcpReferenceState(ranges=ranges, exclusions=exclusions, issues=issues, latest_export_label=latest_export_label)


def dhcp_reference_issues(reference_path: str, max_age_days: int = 35) -> List[str]:
    state = validate_dhcp_reference(reference_path, max_age_days=max_age_days)
    issues: List[str] = []
    for item in state.issues:
        detail = item if not state.latest_export_label else f'{item}, latest={state.latest_export_label}'
        issues.append(jp_item('', 'DHCP参照異常', detail))
    return issues


def dhcp_check(
    row: StandardRow,
    ip_kind: str,
    reference_path: str = '',
    ranges: Optional[Sequence[DhcpRange]] = None,
    exclusions: Optional[Sequence[DhcpRange]] = None,
    username: str = '',
    password: str = '',
    domain: str = '',
) -> List[str]:
    del username, password, domain
    if ip_kind != 'ip':
        return []

    ip = _parse_ipv4(row.IP)
    if ip is None:
        return []

    effective_ranges = list(ranges or [])
    effective_exclusions = list(exclusions or [])
    if not effective_ranges and reference_path:
        effective_ranges, effective_exclusions = load_dhcp_reference_ranges(reference_path)

    for excluded_range in effective_exclusions:
        if excluded_range.contains(ip):
            return []

    for dhcp_range in effective_ranges:
        if dhcp_range.contains(ip):
            return [jp_item(row.userID, 'DHCP範囲内', f'IP={row.IP}, range={dhcp_range.label()}, 固定IPか申請者へ確認')]
    return []


def ad_user_check(user_id: str, required_group: str) -> tuple[List[str], List[str]]:
    blockers: List[str] = []
    confirmations: List[str] = []
    try:
        completed = subprocess.run(
            ['net', 'user', user_id, '/domain'],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except Exception as exc:
        blockers.append(f'{user_id}: Could not run net user /domain. Detail: {exc}')
        return blockers, confirmations

    output = (completed.stdout or '') + '\n' + (completed.stderr or '')
    if completed.returncode != 0:
        blockers.append(f'{user_id}: Target account does not exist or could not be verified. Please confirm the account with the applicant.')
        return blockers, confirmations

    if required_group:
        normalized_output = compact_text(output)
        normalized_group = compact_text(required_group)
        if normalized_group and normalized_group not in normalized_output:
            confirmations.append(jp_item(user_id, '権限グループ未所属', f'group={required_group}, 申請者または管理者へ確認'))
    return blockers, confirmations
