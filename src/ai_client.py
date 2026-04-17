from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional


DEFAULT_OPENROUTER_BASE_URL = 'https://openrouter.ai/api/v1'
DEFAULT_OPENAI_BASE_URL = 'https://api.openai.com/v1'
DEFAULT_PARSE_MODEL = 'google/gemini-2.5-pro'
DEFAULT_PARSE_FALLBACK_MODEL = ''
DEFAULT_REPLY_MODEL = 'google/gemini-2.5-flash'

REPLY_SCHEMA = {
    'type': 'object',
    'properties': {
        'subject': {'type': 'string', 'description': 'Short Japanese subject line.'},
        'body': {'type': 'string', 'description': 'Polite Japanese message body.'},
    },
    'required': ['subject', 'body'],
    'additionalProperties': False,
}


class AIError(RuntimeError):
    pass


@dataclass(frozen=True)
class AIConfig:
    provider: str
    api_key: str
    base_url: str
    parse_model: str
    parse_fallback_model: str
    reply_model: str

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)


def env_value(env: Dict[str, str], key: str, default: str = '') -> str:
    return os.getenv(key, env.get(key, default)).strip()


def load_ai_config(env: Dict[str, str]) -> AIConfig:
    provider = env_value(env, 'AI_PROVIDER', 'auto').lower() or 'auto'
    if provider == 'auto':
        provider = 'openrouter' if env_value(env, 'OPENROUTER_API_KEY') else 'openai'

    if provider == 'openrouter':
        api_key = env_value(env, 'OPENROUTER_API_KEY')
        base_url = env_value(env, 'OPENROUTER_BASE_URL', DEFAULT_OPENROUTER_BASE_URL) or DEFAULT_OPENROUTER_BASE_URL
        parse_model = env_value(env, 'AI_PARSE_MODEL', DEFAULT_PARSE_MODEL) or DEFAULT_PARSE_MODEL
        parse_fallback_model = env_value(env, 'AI_PARSE_FALLBACK_MODEL', DEFAULT_PARSE_FALLBACK_MODEL)
        reply_model = env_value(env, 'AI_REPLY_MODEL', DEFAULT_REPLY_MODEL) or DEFAULT_REPLY_MODEL
    elif provider == 'openai':
        api_key = env_value(env, 'OPENAI_API_KEY')
        base_url = env_value(env, 'OPENAI_BASE_URL', DEFAULT_OPENAI_BASE_URL) or DEFAULT_OPENAI_BASE_URL
        parse_model = env_value(env, 'AI_PARSE_MODEL', env_value(env, 'OPENAI_MODEL', 'gpt-4.1')) or 'gpt-4.1'
        parse_fallback_model = env_value(env, 'AI_PARSE_FALLBACK_MODEL')
        reply_model = env_value(env, 'AI_REPLY_MODEL', env_value(env, 'OPENAI_REPLY_MODEL', parse_model)) or parse_model
    else:
        raise AIError(f'Unsupported AI_PROVIDER: {provider}')

    return AIConfig(
        provider=provider,
        api_key=api_key,
        base_url=base_url.rstrip('/'),
        parse_model=parse_model,
        parse_fallback_model=parse_fallback_model,
        reply_model=reply_model,
    )


def _strip_json_fence(text: str) -> str:
    value = text.strip()
    if value.startswith('```'):
        value = re.sub(r'^```(?:json)?', '', value).strip()
        value = re.sub(r'```$', '', value).strip()
    return value


def _extract_message_content(data: Dict[str, Any]) -> str:
    choices = data.get('choices') or []
    if choices:
        message = choices[0].get('message') or {}
        content = message.get('content', '')
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict):
                    parts.append(str(item.get('text') or item.get('content') or ''))
                else:
                    parts.append(str(item))
            return ''.join(parts)
        return str(content or '')
    return ''


def chat_json(
    config: AIConfig,
    model: str,
    messages: List[Dict[str, str]],
    schema: Dict[str, Any],
    schema_name: str,
    timeout: int = 60,
) -> Dict[str, Any]:
    if not config.enabled:
        raise AIError('AI API key is not configured.')
    if not model:
        raise AIError('AI model is not configured.')

    payload = {
        'model': model,
        'messages': messages,
        'temperature': 0,
        'response_format': {
            'type': 'json_schema',
            'json_schema': {
                'name': schema_name,
                'strict': True,
                'schema': schema,
            },
        },
    }
    request = urllib.request.Request(
        f'{config.base_url}/chat/completions',
        data=json.dumps(payload).encode('utf-8'),
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {config.api_key}',
        },
        method='POST',
    )

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode('utf-8', errors='replace')
        raise AIError(f'HTTP {exc.code}: {body}') from exc
    except Exception as exc:
        raise AIError(str(exc)) from exc

    content = _strip_json_fence(_extract_message_content(data))
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError as exc:
        raise AIError(f'AI response was not valid JSON: {content[:500]}') from exc
    if not isinstance(parsed, dict):
        raise AIError('AI response JSON root was not an object.')
    return parsed


def model_candidates(primary: str, fallback: str = '') -> List[str]:
    models: List[str] = []
    for model in (primary, fallback):
        if model and model not in models:
            models.append(model)
    return models


def format_ai_errors(errors: Iterable[str]) -> str:
    return ' / '.join(str(error) for error in errors if str(error).strip())


def draft_issue_reply(file_name: str, issues: Iterable[str], report_kind: str, env: Dict[str, str]) -> Optional[str]:
    issue_list = [str(item).strip() for item in issues if str(item).strip()]
    if not issue_list:
        return None

    try:
        config = load_ai_config(env)
    except AIError:
        return None
    if not config.enabled or not config.reply_model:
        return None

    audience = '申請者' if report_kind == 'error' else 'Support team'
    purpose = '再提出をお願いする' if report_kind == 'error' else 'network team へ送る前の確認事項を共有する'
    messages = [
        {
            'role': 'system',
            'content': (
                'You write concise, polite Japanese business messages. '
                'Use only the facts supplied by the user. Do not add new checks, causes, promises, deadlines, or technical facts.'
            ),
        },
        {
            'role': 'user',
            'content': f'''
以下の検出結果をもとに、{audience}向けの自然な日本語メッセージを作成してください。
目的: {purpose}
対象ファイル: {file_name}

条件:
- 丁寧だが長すぎない文面にする。
- 箇条書きを使ってよい。
- 検出結果にない情報は追加しない。
- 件名と本文を返す。

検出結果:
{chr(10).join(f'- {item}' for item in issue_list)}
'''.strip(),
        },
    ]
    try:
        data = chat_json(config, config.reply_model, messages, REPLY_SCHEMA, 'issue_reply', timeout=45)
    except AIError:
        return None

    subject = str(data.get('subject', '') or '').strip()
    body = str(data.get('body', '') or '').strip()
    if not subject or not body:
        return None
    return f'件名: {subject}\n\n{body}\n'

