from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from ai_client import draft_issue_reply, load_ai_config, model_candidates
from checks import StandardRow, _dhcp_command, normalize_ip_value, validate_row
from parser import parse_with_rules
from xlsx_io import SheetData, read_workbook, write_standard_workbook


INTERNET_ALIASES = ['internet', 'internet access', 'インターネット', 'インターネットアクセス']


class CoreFlowTests(unittest.TestCase):
    def test_write_and_read_standard_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'standard.xlsx'
            headers = ['userID', 'name', 'company', 'email', 'hostname', 'IP']
            rows = [{
                'userID': 'user01',
                'name': '山田 太郎',
                'company': 'Example',
                'email': 'user01@example.com',
                'hostname': 'server01',
                'IP': '192.0.2.10',
            }]
            write_standard_workbook(path, headers, rows)
            workbook = read_workbook(path)

        self.assertEqual(workbook[0].rows[0], headers)
        self.assertEqual(workbook[0].rows[1], [rows[0][header] for header in headers])

    def test_parse_japanese_headers_and_same_as_above(self) -> None:
        sheet = SheetData('申請', [
            ['対象アカウント', '対象者氏名', '会社名', 'メールアドレス', '接続先サーバー名', 'IPアドレス'],
            ['user01', '山田 太郎', 'Example', 'user01@example.com', 'server01', '192.0.2.10'],
            ['user02', '佐藤 花子', '同上', 'user02@example.com', '同上', 'インターネット'],
        ])

        result = parse_with_rules([sheet])

        self.assertEqual(result.blockers, [])
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[1].company, 'Example')
        self.assertEqual(result.rows[1].hostname, 'server01')
        self.assertEqual(result.rows[1].IP, 'インターネット')

    def test_validate_row_normalizes_internet_access(self) -> None:
        row = StandardRow(userID='user01', name='山田 太郎', IP='インターネットアクセス')
        valid, blockers = validate_row(row, INTERNET_ALIASES)

        self.assertEqual(blockers, [])
        self.assertIsNotNone(valid)
        self.assertEqual(valid.IP, 'Internet Access')

    def test_normalize_ip_accepts_cidr(self) -> None:
        value, kind, error = normalize_ip_value('192.0.2.0/24', INTERNET_ALIASES)

        self.assertEqual(value, '192.0.2.0/24')
        self.assertEqual(kind, 'cidr')
        self.assertIsNone(error)

    def test_dhcp_command_uses_cim_session_when_credentials_are_present(self) -> None:
        command = _dhcp_command('192.0.2.10', 'dhcp01.example.local', 'reader', 'secret', 'EXAMPLE')
        script = command[-1]

        self.assertIn('New-CimSession', script)
        self.assertIn("EXAMPLE\\reader", script)
        self.assertIn("192.0.2.10", script)

    def test_openrouter_config_defaults_to_confirmed_models(self) -> None:
        config = load_ai_config({
            'AI_PROVIDER': 'openrouter',
            'OPENROUTER_API_KEY': 'test-key',
        })

        self.assertEqual(config.provider, 'openrouter')
        self.assertEqual(config.base_url, 'https://openrouter.ai/api/v1')
        self.assertEqual(config.parse_model, 'google/gemini-2.5-pro')
        self.assertEqual(config.parse_fallback_model, '')
        self.assertEqual(config.reply_model, 'google/gemini-2.5-flash')

    def test_model_candidates_deduplicates_fallback(self) -> None:
        self.assertEqual(model_candidates('a', 'a'), ['a'])
        self.assertEqual(model_candidates('a', 'b'), ['a', 'b'])

    def test_draft_issue_reply_without_key_uses_no_network(self) -> None:
        self.assertIsNone(draft_issue_reply('request.xlsx', ['IPアドレスが未記入です。'], 'error', {'AI_PROVIDER': 'openrouter'}))


if __name__ == '__main__':
    unittest.main()

