# Word文档AI语法检查器 - 使用说明

## 功能概述

这个脚本可以：
- 读取Word文档并按段落分割
- 使用AI（OpenAI或Gemini）进行语法检查
- 支持自定义的额外检查要求
- 每3段自动重新开始AI会话
- 将结果保存为结构化的Excel文件

## 安装依赖

```bash
pip install python-docx pandas litellm tqdm openpyxl
```

## 配置设置

### 1. 配置文件说明 (config.json)

| 参数 | 说明 | 示例值 |
|------|------|--------|
| `model` | AI模型选择 | `"gpt-3.5-turbo"` 或 `"gemini-pro"` |
| `openai_api_key` | OpenAI API密钥 | `"sk-..."` |
| `gemini_api_key` | Gemini API密钥 | `"AIza..."` |
| `max_retries` | API调用失败时的最大重试次数 | `3` |
| `retry_delay` | 重试间隔（秒） | `1` |
| `session_refresh_interval` | 多少段后重新开始AI会话 | `3` |
| `additional_checks` | 默认的额外检查要求 | 数组格式 |

### 2. API密钥获取

**OpenAI:**
- 访问 https://platform.openai.com/api-keys
- 创建新的API密钥
- 复制密钥到配置文件

**Gemini:**
- 访问 https://makersuite.google.com/app/apikey
- 获取API密钥
- 复制密钥到配置文件

## 使用方法

### 基本用法

```bash
python grammar_checker.py document.docx
```

### 指定输出文件

```bash
python grammar_checker.py document.docx -o 检查结果.xlsx
```

### 添加额外检查要求

```bash
python grammar_checker.py document.docx --additional-checks "检查用词准确性" "检查逻辑连贯性"
```

### 使用自定义配置文件

```bash
python grammar_checker.py document.docx -c my_config.json
```

## 输出文件结构

Excel文件包含以下列：

| 列名 | 内容 |
|------|------|
| 原始文本 | Word文档中的原始段落 |
| 语法检查 | AI的语法检查意见 |
| 额外检查_XXX | 每个额外检查要求的结果（如果有） |

## 高级功能

### 1. 会话管理
- 每处理3个段落后自动重新开始AI会话
- 避免上下文过长影响检查质量
- 可通过 `session_refresh_interval` 参数调整

### 2. 错误处理
- 自动重试API调用失败
- 详细的错误日志
- 优雅的异常处理

### 3. 进度显示
- 实时显示处理进度
- 会话刷新提示
- 处理统计信息

## 常见问题

### Q: API调用失败怎么办？
A: 脚本会自动重试，检查：
- API密钥是否正确
- 网络连接是否正常
- API配额是否充足

### Q: 如何调整检查的详细程度？
A: 修改脚本中的提示词模板，在 `create_grammar_prompt` 和 `create_additional_check_prompt` 方法中

### Q: 支持哪些文档格式？
A: 目前仅支持 .docx 格式的Word文档

### Q: 如何批量处理多个文档？
A: 可以编写批处理脚本或修改主程序支持文件夹输入

## 注意事项

1. **API成本**: 每段文本都会调用AI API，注意控制成本
2. **处理时间**: 大文档需要较长处理时间
3. **网络要求**: 需要稳定的网络连接
4. **文档格式**: 确保Word文档格式正确，避免特殊字符
