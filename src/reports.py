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


def support_confirmation_text(item: str) -> str:
    if 'DHCP参照異常' in item:
        return (
            'DHCP参照データに問題があります。AD側のDHCP出力ファイルが最新か確認してください。'
            f'必要に応じて、申請IPがDHCP配布範囲内か手動で確認してください。詳細: {item}'
        )
    if 'DHCP範囲内' in item:
        return (
            '申請IPがDHCP配布範囲に含まれています。固定IPとして利用してよいか申請者へ確認してください。'
            f'詳細: {item}'
        )
    if 'ホスト名一致しない' in item:
        return (
            '申請された接続先ホスト名とDNS逆引き結果が一致しません。接続先サーバー名が正しいか確認してください。'
            f'詳細: {item}'
        )
    if '逆引き確認不可' in item or '逆引きホスト名なし' in item:
        return (
            'DNS逆引きで接続先ホスト名を確認できませんでした。申請されたサーバー名が正しいか手動で確認してください。'
            f'詳細: {item}'
        )
    if '権限グループ未所属' in item:
        return (
            '対象アカウントが必要な権限グループに所属していません。申請者または管理者へ利用権限を確認してください。'
            f'詳細: {item}'
        )
    return item


def confirmation_message(file_name: str, confirmations: Iterable[str]) -> str:
    lines = [
        '標準フォーマットへの変換は完了しています。',
        'network team へ連携する前に、以下の内容を確認してください。',
        '',
        f'対象ファイル: {file_name}',
        '',
        '対応が必要な確認事項:',
    ]
    lines.extend(f'- {support_confirmation_text(item)}' for item in confirmations)
    lines.append('')
    lines.append('確認が完了したら、生成された standard xlsx を network team へ連携してください。')
    lines.append('')
    return '\n'.join(lines)
