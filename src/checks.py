from __future__ import annotations

import csv
import ipaddress
import json
import re
import socket
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
ACCOUNT_RE = re.compile(r'^[A-Za-z0-9._-]{1,64}$')
HOSTNAME_RE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9.-]{0,253}$')


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

    def contains(self, ip: ipaddress.IPv4Address) -> bool:
        return int(self.start) <= int(ip) <= int(self.end)

    def label(self) -> str:
        label_parts = []
        if self.name:
            label_parts.append(self.name)
        label_parts.append(f'{self.start}-{self.end}')
        if self.source:
            label_parts.append(f'source: {self.source}')
        return ' / '.join(label_parts)


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


def _range_from_record(record: Dict[str, Any], source: str) -> Optional[DhcpRange]:
    start_text = _first_value(record, ['StartRange', 'Start', 'RangeStart', 'StartAddress', 'From'])
    end_text = _first_value(record, ['EndRange', 'End', 'RangeEnd', 'EndAddress', 'To'])
    name = _first_value(record, ['Name', 'ScopeName', 'ScopeId', 'Subnet', 'Network'])

    if not start_text or not end_text:
        cidr_text = _first_value(record, ['CIDR', 'Cidr', 'NetworkCidr', 'Prefix'])
        if cidr_text:
            try:
                network = ipaddress.ip_network(cidr_text, strict=False)
            except ValueError:
                return None
            if network.version != 4:
                return None
            return DhcpRange(network.network_address, network.broadcast_address, source=source, name=name or str(network))
        return None

    start = _parse_ipv4(start_text)
    end = _parse_ipv4(end_text)
    if start is None or end is None:
        return None
    if int(start) > int(end):
        start, end = end, start
    return DhcpRange(start, end, source=source, name=name)


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


def _load_dhcp_ranges_from_file(path: Path) -> List[DhcpRange]:
    suffix = path.suffix.lower()
    ranges: List[DhcpRange] = []
    if suffix == '.json':
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        for record in _json_records(data):
            parsed = _range_from_record(record, path.name)
            if parsed:
                ranges.append(parsed)
    elif suffix == '.csv':
        with path.open('r', encoding='utf-8-sig', newline='') as fp:
            for record in csv.DictReader(fp):
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

    files: List[Path]
    if path.is_dir():
        files = sorted([*path.glob('*.csv'), *path.glob('*.json')])
    else:
        files = [path]

    ranges: List[DhcpRange] = []
    for file_path in files:
        try:
            ranges.extend(_load_dhcp_ranges_from_file(file_path))
        except Exception:
            continue
    return ranges


def dhcp_check(
    row: StandardRow,
    ip_kind: str,
    reference_path: str,
    username: str = '',
    password: str = '',
    domain: str = '',
) -> List[str]:
    del username, password, domain
    if ip_kind != 'ip':
        return []
    if not reference_path:
        return [jp_item(row.userID, 'DHCP参照未設定', f'IP={row.IP}, DHCP範囲内か手動確認')]

    ip = _parse_ipv4(row.IP)
    if ip is None:
        return []

    reference = Path(str(reference_path).strip().strip('"')).expanduser()
    if not reference.exists():
        return [jp_item(row.userID, 'DHCP参照ファイルなし', f'path={reference_path}, IP={row.IP}, DHCP範囲内か手動確認')]

    ranges = load_dhcp_ranges(reference_path)
    if not ranges:
        return [jp_item(row.userID, 'DHCP範囲読み取り不可', f'path={reference_path}, IP={row.IP}, DHCP範囲内か手動確認')]

    for dhcp_range in ranges:
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
