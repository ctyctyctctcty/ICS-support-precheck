from __future__ import annotations

import ipaddress
import re
import subprocess
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

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
        return None, None, 'IPアドレスが未記入です。'

    internet = normalize_internet(raw, internet_aliases)
    if internet:
        return internet, 'internet', None

    try:
        if '/' in raw:
            network = ipaddress.ip_network(raw, strict=False)
            if network.version != 4:
                return None, None, 'IPv4以外のCIDRは対応していません。'
            if network.prefixlen == 32:
                return str(network.network_address), 'ip', None
            return str(network), 'cidr', None
        ip = ipaddress.ip_address(raw)
        if ip.version != 4:
            return None, None, 'IPv4以外のIPアドレスは対応していません。'
        return str(ip), 'ip', None
    except ValueError:
        return None, None, f'IPアドレスの形式が正しくありません: {raw}'


def validate_row(row: StandardRow, internet_aliases: Sequence[str]) -> tuple[Optional[StandardRow], List[str]]:
    blockers: List[str] = []
    user_id = row.userID.strip()
    name = row.name.strip()
    ip_value, _, ip_error = normalize_ip_value(row.IP, internet_aliases)

    if not user_id:
        blockers.append('対象アカウントが未記入です。')
    elif not ACCOUNT_RE.fullmatch(user_id):
        blockers.append(f'対象アカウントの形式が正しくありません: {user_id}')

    if not name:
        blockers.append('対象者氏名が未記入です。')

    if ip_error:
        blockers.append(ip_error)

    email = row.email.strip()
    if email and not EMAIL_RE.fullmatch(email):
        blockers.append(f'対象者メールアドレスの形式が正しくありません: {email}')

    hostname = row.hostname.strip()
    if hostname and not HOSTNAME_RE.fullmatch(hostname):
        blockers.append(f'接続先サーバー名の形式が正しくありません: {hostname}')

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


def reverse_dns_check(row: StandardRow, ip_kind: str) -> List[str]:
    if ip_kind != 'ip' or not row.hostname.strip():
        return []

    try:
        completed = subprocess.run(
            ['nslookup', row.IP],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except Exception as exc:
        return [f'{row.userID}: nslookup {row.IP} を実行できませんでした。手動確認してください。詳細: {exc}']

    output = (completed.stdout or '') + '\n' + (completed.stderr or '')
    if completed.returncode != 0:
        return [f'{row.userID}: nslookup {row.IP} が失敗しました。接続先サーバー名 {row.hostname} を手動確認してください。']

    names = []
    for line in output.splitlines():
        if 'name =' in line.lower():
            names.append(line.split('=', 1)[1].strip().rstrip('.'))
        elif line.strip().lower().startswith('name:'):
            names.append(line.split(':', 1)[1].strip().rstrip('.'))

    if not names:
        return [f'{row.userID}: nslookup {row.IP} でホスト名を確認できませんでした。申請値 {row.hostname} を手動確認してください。']

    requested = short_hostname(row.hostname)
    if any(short_hostname(name) == requested for name in names):
        return []
    return [f'{row.userID}: DNS逆引き結果 ({", ".join(names)}) と申請された接続先サーバー名 ({row.hostname}) が一致しません。']


def _ps_single_quote(value: str) -> str:
    return str(value).replace("'", "''")


def _dhcp_command(ip: str, dhcp_server: str, username: str = '', password: str = '', domain: str = '') -> List[str]:
    server = _ps_single_quote(dhcp_server)
    address = _ps_single_quote(ip)
    if username and password:
        qualified_user = f'{domain}\\{username}' if domain else username
        user = _ps_single_quote(qualified_user)
        secret = _ps_single_quote(password)
        script = (
            f"$secure = ConvertTo-SecureString '{secret}' -AsPlainText -Force; "
            f"$cred = [System.Management.Automation.PSCredential]::new('{user}', $secure); "
            f"$session = New-CimSession -ComputerName '{server}' -Credential $cred; "
            "try { "
            f"Get-DhcpServerv4Lease -CimSession $session -IPAddress '{address}' -ErrorAction SilentlyContinue | "
            "Select-Object -First 1 -ExpandProperty IPAddress "
            "} finally { if ($session) { Remove-CimSession $session } }"
        )
    else:
        script = (
            f"Get-DhcpServerv4Lease -ComputerName '{server}' -IPAddress '{address}' "
            "-ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty IPAddress"
        )
    return ['powershell', '-NoProfile', '-Command', script]


def dhcp_check(
    row: StandardRow,
    ip_kind: str,
    dhcp_server: str,
    username: str = '',
    password: str = '',
    domain: str = '',
) -> List[str]:
    if ip_kind != 'ip':
        return []
    if not dhcp_server:
        return [f'{row.userID}: DHCPサーバー設定が未設定のため、DHCP対象か手動確認してください。']

    try:
        completed = subprocess.run(
            _dhcp_command(row.IP, dhcp_server, username, password, domain),
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except Exception as exc:
        return [f'{row.userID}: DHCP確認を実行できませんでした。手動確認してください。詳細: {exc}']

    if completed.returncode != 0:
        detail = (completed.stderr or '').strip()
        return [f'{row.userID}: DHCP確認に失敗しました。手動確認してください。詳細: {detail}']
    if (completed.stdout or '').strip():
        return [f'{row.userID}: 対象IP {row.IP} はDHCPリース対象の可能性があります。申請者へ固定IPか確認してください。']
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
        blockers.append(f'{user_id}: net user /domain を実行できませんでした。詳細: {exc}')
        return blockers, confirmations

    output = (completed.stdout or '') + '\n' + (completed.stderr or '')
    if completed.returncode != 0:
        blockers.append(f'{user_id}: 対象アカウントが存在しない、または確認できません。申請者へアカウントを確認してください。')
        return blockers, confirmations

    if required_group and required_group not in output:
        confirmations.append(f'{user_id}: 対象アカウントが {required_group} に所属していません。申請者または管理者へ確認してください。')
    return blockers, confirmations
