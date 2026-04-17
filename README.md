# ICS Support Precheck

Support team tool for checking applicant VPN request workbooks and converting them to the standard workbook format used by the network team.

This is a standalone support precheck/conversion tool. Keep it separate from any network-team production system so applicant intake checks and network-side processing do not get mixed together.

## Flow

1. Put applicant `.xlsx` files in `data/source`.
2. Run the tool from this directory.
3. The tool first tries rule-based parsing, then uses AI fallback when parsing fails or validation needs another extraction attempt.
4. The tool validates required fields, normalizes Internet Access / IP values, checks AD / reverse DNS / DHCP where configured, and writes a standard workbook when possible.
5. Processed files are moved out of `data/source` so the next run only handles new files.

## Output folders

- `data/network_ready`: the request can be sent directly to the network team.
- `data/needs_confirmation`: a standard workbook was generated, but support team confirmation is needed before sending it onward.
- `data/error`: a blocker prevents standard workbook generation or onward handling.

Each confirmation or error case also gets a UTF-8 `.txt` report with the reason. When an AI API key is configured, the report starts with a polite human-readable Japanese draft and then keeps the raw detected details for audit.

## Run

```powershell
py src\main.py
```

## Test

The tests are offline and do not call AD, DHCP, DNS, OpenRouter, or OpenAI.

```powershell
py -m unittest discover -s tests
```

## AI provider and models

Copy `config/.env.example` to `config/.env` and set the company API key.

Recommended OpenRouter setup:

```dotenv
AI_PROVIDER=openrouter
OPENROUTER_API_KEY=...
OPENROUTER_BASE_URL=https://openrouter.ai/api/v1
AI_PARSE_MODEL=google/gemini-2.5-pro
AI_PARSE_FALLBACK_MODEL=
AI_REPLY_MODEL=google/gemini-2.5-flash
```

Model roles:

- `AI_PARSE_MODEL`: high-accuracy workbook extraction. Default: `google/gemini-2.5-pro`.
- `AI_PARSE_FALLBACK_MODEL`: optional second extraction attempt when the primary model fails. Leave empty for the first rollout.
- `AI_REPLY_MODEL`: cheaper model for drafting polite Japanese messages from detected issues. Default: `google/gemini-2.5-flash`.

OpenAI direct mode is still supported for compatibility:

```dotenv
AI_PROVIDER=openai
OPENAI_API_KEY=...
AI_PARSE_MODEL=gpt-4.1
AI_REPLY_MODEL=gpt-4.1
```

If no AI API key is configured, the tool uses rule-based parsing and local report templates only.

## Company-specific settings

Do not commit real company settings or credentials. Put them in `config/.env`.

```dotenv
REQUIRED_SECURITY_GROUP=your-vpn-access-group-name
```

If `REQUIRED_SECURITY_GROUP` is empty, AD existence checks still run, but group membership confirmation is skipped.

## DHCP permissions

For DHCP checking, request a read-only account that can query DHCP scope and lease information. It should be equivalent to read-only access for `Get-DhcpServerv4Lease` against the DHCP server and should not need permissions to create, modify, or delete DHCP configuration.

Set these in `config/.env` when a dedicated read-only account is required:

```dotenv
DHCP_SERVER=dhcp01.example.local
DHCP_QUERY_USERNAME=readonly-user
DHCP_QUERY_PASSWORD=change-me
DHCP_QUERY_DOMAIN=EXAMPLE
```

If `DHCP_QUERY_USERNAME` and `DHCP_QUERY_PASSWORD` are empty, the current Windows user context is used for the DHCP query.
