"""Simple i18n mapping for desktop application."""

# Mapping uses Chinese text as the key to avoid scattering ids across the
# application. Each entry provides translations for supported languages.

translations = {
    "AI语法检查器": {"zh": "AI语法检查器", "en": "AI Grammar Checker"},
    "⚙️ 配置设置": {"zh": "⚙️ 配置设置", "en": "⚙️ Configuration"},
    "🔑 API设置": {"zh": "🔑 API设置", "en": "🔑 API Settings"},
    "OpenAI API Key:": {"zh": "OpenAI API Key:", "en": "OpenAI API Key:"},
    "Gemini API Key:": {"zh": "Gemini API Key:", "en": "Gemini API Key:"},
    "🤖 模型设置": {"zh": "🤖 模型设置", "en": "🤖 Model Settings"},
    "AI供应商:": {"zh": "AI供应商:", "en": "Provider:"},
    "模型:": {"zh": "模型:", "en": "Model:"},
    "🌐 语言设置": {"zh": "🌐 语言设置", "en": "🌐 Language"},
    "中文": {"zh": "中文", "en": "Chinese"},
    "English": {"zh": "English", "en": "English"},
    "🔧 高级设置": {"zh": "🔧 高级设置", "en": "🔧 Advanced Settings"},
    "最大重试次数:": {"zh": "最大重试次数:", "en": "Max Retries:"},
    "重试延迟(秒):": {"zh": "重试延迟(秒):", "en": "Retry Delay (s):"},
    "会话刷新间隔:": {"zh": "会话刷新间隔:", "en": "Session Refresh Interval:"},
    "💾 配置管理": {"zh": "💾 配置管理", "en": "💾 Config Management"},
    "保存配置": {"zh": "保存配置", "en": "Save Config"},
    "加载配置": {"zh": "加载配置", "en": "Load Config"},
    "📁 文件操作": {"zh": "📁 文件操作", "en": "📁 File"},
    "选择Word文档": {"zh": "选择Word文档", "en": "Select Word Document"},
    "未选择文件": {"zh": "未选择文件", "en": "No file selected"},
    "文档预览将在这里显示...": {
        "zh": "文档预览将在这里显示...",
        "en": "Document preview will appear here...",
    },
    "📊 输出设置": {"zh": "📊 输出设置", "en": "📊 Output Settings"},
    "浏览": {"zh": "浏览", "en": "Browse"},
    "Excel输出路径:": {"zh": "Excel输出路径:", "en": "Excel Output Path:"},
    "✅ 额外检查要求": {"zh": "✅ 额外检查要求", "en": "✅ Additional Checks"},
    "输入新的检查要求...": {
        "zh": "输入新的检查要求...",
        "en": "Enter new check requirement...",
    },
    "添加": {"zh": "添加", "en": "Add"},
    "删除选中项": {"zh": "删除选中项", "en": "Remove Selected"},
    "🚀 开始语法检查": {
        "zh": "🚀 开始语法检查",
        "en": "🚀 Start Grammar Check",
    },
    "文档包含 {} 个段落\n\n": {
        "zh": "文档包含 {} 个段落\n\n",
        "en": "Document contains {} paragraphs\n\n",
    },
    "段落 {}: {}...\n\n": {
        "zh": "段落 {}: {}...\n\n",
        "en": "Paragraph {}: {}...\n\n",
    },
    "错误": {"zh": "错误", "en": "Error"},
    "读取Word文档失败: {}": {
        "zh": "读取Word文档失败: {}",
        "en": "Failed to read Word document: {}",
    },
    "警告": {"zh": "警告", "en": "Warning"},
    "请先选择Word文档": {
        "zh": "请先选择Word文档",
        "en": "Please select a Word document first",
    },
    "请输入{} API密钥": {
        "zh": "请输入{} API密钥",
        "en": "Please enter {} API key",
    },
    "成功": {"zh": "成功", "en": "Success"},
    "配置已保存到 config.json": {
        "zh": "配置已保存到 config.json",
        "en": "Configuration saved to config.json",
    },
    "配置文件加载成功": {
        "zh": "配置文件加载成功",
        "en": "Configuration file loaded successfully",
    },
    "加载配置文件失败: {}": {
        "zh": "加载配置文件失败: {}",
        "en": "Failed to load config file: {}",
    },
    "完成": {"zh": "完成", "en": "Done"},
    "语法检查完成！\n结果已保存到: {}\n共处理 {} 个段落": {
        "zh": "语法检查完成！\n结果已保存到: {}\n共处理 {} 个段落",
        "en": "Grammar check completed!\nResults saved to: {}\nProcessed {} paragraphs",
    },
    "保存Excel文件失败: {}": {
        "zh": "保存Excel文件失败: {}",
        "en": "Failed to save Excel file: {}",
    },
    "处理过程中出现错误: {}": {
        "zh": "处理过程中出现错误: {}",
        "en": "An error occurred during processing: {}",
    },
    "API密钥不能为空": {
        "zh": "API密钥不能为空",
        "en": "API key cannot be empty",
    },
    "保存配置失败: {}": {
        "zh": "保存配置失败: {}",
        "en": "Failed to save config: {}",
    },
    "语法检查结果.xlsx": {
        "zh": "语法检查结果.xlsx",
        "en": "grammar_check_results.xlsx",
    },
    "保存Excel文件": {"zh": "保存Excel文件", "en": "Save Excel File"},
    "选择配置文件": {"zh": "选择配置文件", "en": "Select Config File"},
}


def get_text(text: str, lang: str) -> str:
    """Return translation for given text and language."""
    return translations.get(text, {}).get(lang, text)

