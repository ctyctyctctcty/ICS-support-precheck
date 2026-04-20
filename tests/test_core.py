from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from ai_client import draft_issue_reply, load_ai_config, model_candidates
from checks import StandardRow, dhcp_check, load_dhcp_ranges, normalize_ip_value, validate_row
from parser import parse_with_rules
from xlsx_io import SheetData, read_workbook, write_standard_workbook


INTERNET_ALIASES = ['internet', 'internet access']


class CoreFlowTests(unittest.TestCase):
    def test_write_and_read_standard_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'standard.xlsx'
            headers = ['userID', 'name', 'company', 'email', 'hostname', 'IP']
            rows = [{
                'userID': 'user01',
                'name': 'User One',
                'company': 'Example',
                'email': 'user01@example.com',
                'hostname': 'server01',
                'IP': '192.0.2.10',
            }]
            write_standard_workbook(path, headers, rows)
            workbook = read_workbook(path)

        self.assertEqual(workbook[0].rows[0], headers)
        self.assertEqual(workbook[0].rows[1], [rows[0][header] for header in headers])

    def test_parse_english_headers(self) -> None:
        sheet = SheetData('Request', [
            ['userID', 'name', 'company', 'email', 'hostname', 'IP'],
            ['user01', 'User One', 'Example', 'user01@example.com', 'server01', '192.0.2.10'],
            ['user02', 'User Two', 'Example', 'user02@example.com', 'server02', 'internet access'],
        ])

        result = parse_with_rules([sheet])

        self.assertEqual(result.blockers, [])
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[1].company, 'Example')
        self.assertEqual(result.rows[1].hostname, 'server02')
        self.assertEqual(result.rows[1].IP, 'internet access')

    def test_validate_row_normalizes_internet_access(self) -> None:
        row = StandardRow(userID='user01', name='User One', IP='internet access')
        valid, blockers = validate_row(row, INTERNET_ALIASES)

        self.assertEqual(blockers, [])
        self.assertIsNotNone(valid)
        self.assertEqual(valid.IP, 'Internet Access')

    def test_normalize_ip_accepts_cidr(self) -> None:
        value, kind, error = normalize_ip_value('192.0.2.0/24', INTERNET_ALIASES)

        self.assertEqual(value, '192.0.2.0/24')
        self.assertEqual(kind, 'cidr')
        self.assertIsNone(error)

    def test_dhcp_check_uses_local_csv_reference(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / 'dhcp_scopes.csv'
            csv_path.write_text(
                'ScopeId,Name,StartRange,EndRange\n'
                '192.0.2.0,Office LAN,192.0.2.10,192.0.2.50\n',
                encoding='utf-8',
            )
            row = StandardRow(userID='user01', name='User One', IP='192.0.2.20')
            messages = dhcp_check(row, 'ip', str(csv_path))

        self.assertEqual(len(messages), 1)
        self.assertIn('DHCP範囲内', messages[0])
        self.assertIn('固定IPか申請者へ確認', messages[0])
        self.assertIn('Office LAN', messages[0])

    def test_dhcp_check_accepts_json_directory_reference(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            json_path = Path(tmp) / 'scopes.json'
            json_path.write_text(
                '{"scopes": [{"name": "Branch LAN", "start": "198.51.100.10", "end": "198.51.100.20"}]}',
                encoding='utf-8',
            )
            ranges = load_dhcp_ranges(tmp)
            hit = dhcp_check(StandardRow('user01', 'User', IP='198.51.100.15'), 'ip', tmp)
            miss = dhcp_check(StandardRow('user02', 'User', IP='198.51.100.30'), 'ip', tmp)

        self.assertEqual(len(ranges), 1)
        self.assertEqual(len(hit), 1)
        self.assertEqual(miss, [])

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
        self.assertIsNone(draft_issue_reply('request.xlsx', ['IP address is missing.'], 'error', {'AI_PROVIDER': 'openrouter'}))


if __name__ == '__main__':
    unittest.main()
