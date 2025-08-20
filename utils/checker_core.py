"""Common utilities for AI grammar checking."""
from typing import List, Dict, Callable, Optional
import time

try:
    import litellm
except ImportError as e:
    raise ImportError(
        "缺少必要的库: litellm. 请运行: pip install litellm"
    ) from e


def create_prompt(text: str, language: str = "中文", check_type: str = "grammar", *, custom_requirement: str = "") -> str:
    """Create prompt for grammar or additional checks."""
    if language == "中文":
        if check_type == "grammar":
            return f"""请检查以下文本的语法错误，只需要指出语法问题并给出简洁的修改建议：

文本：{text}

请用中文回答，格式如下：
- 如果没有语法错误，请仅回答\"语法正确\"
- 如果有语法错误，简洁地指出问题和建议
"""
        else:
            return f"""请对以下文本进行检查：{custom_requirement}

文本：{text}

请用中文给出简洁的评价和建议：
"""
    else:
        if check_type == "grammar":
            return f"""Please check the following text for grammar errors and provide concise suggestions:

Text: {text}

Please respond in English:
- If there are no grammar errors, please only respond \"Grammar is correct\"
- If there are grammar errors, briefly point out the issues and suggestions
"""
        else:
            return f"""Please check the following text for: {custom_requirement}

Text: {text}

Please provide concise evaluation and suggestions in English:
"""


def call_ai_api(prompt: str, provider: str, model: str, api_key: str, *, max_retries: int = 3, retry_delay: float = 1.0) -> str:
    """Call AI API via litellm with retry logic."""
    if not api_key:
        raise ValueError("API密钥不能为空")

    for attempt in range(max_retries):
        try:
            if provider == "openai":
                litellm.openai_key = api_key
            elif provider == "gemini":
                litellm.gemini_key = api_key
            response = litellm.completion(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=500,
                temperature=0.3,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
            else:
                return f"API调用失败: {str(e)}"


def process_paragraphs(
    paragraphs: List[str],
    config: Dict,
    *,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[Dict]:
    """Process paragraphs using AI grammar check and additional checks.

    Args:
        paragraphs: list of paragraph texts.
        config: configuration dict containing at least:
            language, provider, model, api_key, additional_checks,
            session_refresh_interval, max_retries, retry_delay.
        progress_callback: optional callable accepting (index, total, message).

    Returns:
        List of result dictionaries per paragraph.
    """
    results: List[Dict] = []
    total = len(paragraphs)
    additional_checks = config.get("additional_checks", [])
    interval = config.get("session_refresh_interval", 3)
    for i, paragraph in enumerate(paragraphs):
        if progress_callback:
            progress_callback(i, total, f"处理第 {i+1}/{total} 段...")
        if i > 0 and interval and i % interval == 0:
            if progress_callback:
                progress_callback(i, total, f"刷新AI会话... 第 {i+1}/{total} 段")
            time.sleep(1)
        row: Dict = {"原始文本": paragraph}
        prompt = create_prompt(paragraph, config.get("language", "中文"), "grammar")
        row["语法检查"] = call_ai_api(
            prompt,
            config["provider"],
            config["model"],
            config["api_key"],
            max_retries=config.get("max_retries", 3),
            retry_delay=config.get("retry_delay", 1),
        )
        for j, check in enumerate(additional_checks):
            if check.strip():
                add_prompt = create_prompt(
                    paragraph,
                    config.get("language", "中文"),
                    "additional",
                    custom_requirement=check,
                )
                row[f"额外检查_{j+1}"] = call_ai_api(
                    add_prompt,
                    config["provider"],
                    config["model"],
                    config["api_key"],
                    max_retries=config.get("max_retries", 3),
                    retry_delay=config.get("retry_delay", 1),
                )
        results.append(row)
        time.sleep(0.5)
    return results
