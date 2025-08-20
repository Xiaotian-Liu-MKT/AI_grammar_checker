#!/usr/bin/env python3
"""
Word文档AI语法检查脚本
功能：读取Word文档，分段后使用AI进行语法检查，结果保存到Excel
"""

import os
import json
import time
import argparse
from typing import List, Dict, Optional, Tuple
from pathlib import Path


class ConfigFileNotFoundError(Exception):
    """在配置文件缺失时抛出的异常"""

    def __init__(self, message: str, config: Dict):
        super().__init__(message)
        self.config = config

try:
    from docx import Document
    import pandas as pd
    import litellm
    from tqdm import tqdm
except ImportError as e:
    print(f"缺少必要的库: {e}")
    print("请运行: pip install python-docx pandas litellm tqdm openpyxl")
    exit(1)


class GrammarChecker:
    def __init__(self, config_path: str = "config.json"):
        """
        初始化语法检查器
        
        Args:
            config_path: 配置文件路径
        """
        self.config = self.load_config(config_path)
        self.session_count = 0
        self.paragraph_count = 0
        
        # 设置litellm
        if self.config.get("openai_api_key"):
            litellm.openai_key = self.config["openai_api_key"]
        if self.config.get("gemini_api_key"):
            litellm.gemini_key = self.config["gemini_api_key"]
    
    def load_config(self, config_path: str) -> Dict:
        """加载配置文件"""
        if not os.path.exists(config_path):
            # 创建默认配置文件
            default_config = {
                "model": "gpt-3.5-turbo",  # 或 "gemini-pro"
                "openai_api_key": "",
                "gemini_api_key": "",
                "max_retries": 3,
                "retry_delay": 1,
                "session_refresh_interval": 3,
                "additional_checks": []
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            raise ConfigFileNotFoundError(
                f"已创建默认配置文件: {config_path}", default_config
            )

        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def read_word_document(self, file_path: str) -> List[str]:
        """
        读取Word文档并按段落分割
        
        Args:
            file_path: Word文档路径
            
        Returns:
            分段后的文本列表
        """
        try:
            doc = Document(file_path)
            paragraphs = []
            
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:  # 跳过空段落
                    paragraphs.append(text)
            
            print(f"成功读取文档，共 {len(paragraphs)} 个段落")
            return paragraphs
            
        except Exception as e:
            print(f"读取Word文档时出错: {e}")
            return []
    
    def create_grammar_prompt(self, text: str) -> str:
        """创建语法检查的提示词"""
        return f"""请检查以下文本的语法错误，只需要指出语法问题并给出简洁的修改建议：

文本：{text}

请用中文回答，格式如下：
- 如果没有语法错误，回答"语法正确"
- 如果有语法错误，简洁地指出问题和建议
"""
    
    def create_additional_check_prompt(self, text: str, check_requirement: str) -> str:
        """创建额外检查的提示词"""
        return f"""请对以下文本进行检查：{check_requirement}

文本：{text}

请用中文给出简洁的评价和建议：
"""
    
    def call_ai_api(self, prompt: str, max_retries: int = None) -> str:
        """
        调用AI API
        
        Args:
            prompt: 提示词
            max_retries: 最大重试次数
            
        Returns:
            AI的回复
        """
        max_retries = max_retries or self.config.get("max_retries", 3)
        
        for attempt in range(max_retries):
            try:
                response = litellm.completion(
                    model=self.config["model"],
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=500,
                    temperature=0.3
                )
                return response.choices[0].message.content.strip()
                
            except Exception as e:
                print(f"API调用失败 (尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(self.config.get("retry_delay", 1))
                else:
                    return f"API调用失败: {str(e)}"
    
    def should_refresh_session(self) -> bool:
        """判断是否需要刷新会话"""
        return self.paragraph_count % self.config.get("session_refresh_interval", 3) == 0
    
    def process_paragraphs(self, paragraphs: List[str], 
                          additional_checks: List[str] = None) -> List[Dict]:
        """
        处理所有段落
        
        Args:
            paragraphs: 段落列表
            additional_checks: 额外检查要求列表
            
        Returns:
            处理结果列表
        """
        results = []
        additional_checks = additional_checks or self.config.get("additional_checks", [])
        
        print(f"开始处理 {len(paragraphs)} 个段落...")
        if additional_checks:
            print(f"额外检查要求: {additional_checks}")
        
        for i, paragraph in enumerate(tqdm(paragraphs, desc="处理进度")):
            self.paragraph_count += 1
            
            # 检查是否需要刷新会话
            if self.should_refresh_session():
                self.session_count += 1
                print(f"\n刷新AI会话 (第 {self.session_count} 次)")
            
            result = {
                "original_text": paragraph,
                "grammar_check": "",
                "additional_checks": {}
            }
            
            # 语法检查
            grammar_prompt = self.create_grammar_prompt(paragraph)
            grammar_result = self.call_ai_api(grammar_prompt)
            result["grammar_check"] = grammar_result
            
            # 额外检查
            for check_name in additional_checks:
                additional_prompt = self.create_additional_check_prompt(paragraph, check_name)
                additional_result = self.call_ai_api(additional_prompt)
                result["additional_checks"][check_name] = additional_result
            
            results.append(result)
            
            # 简短延迟避免API限流
            time.sleep(0.5)
        
        return results
    
    def save_to_excel(self, results: List[Dict], output_path: str):
        """
        保存结果到Excel文件
        
        Args:
            results: 处理结果
            output_path: 输出文件路径
        """
        try:
            # 准备数据
            data = []
            for result in results:
                row = {
                    "原始文本": result["original_text"],
                    "语法检查": result["grammar_check"]
                }
                
                # 添加额外检查列
                for check_name, check_result in result["additional_checks"].items():
                    row[f"额外检查_{check_name}"] = check_result
                
                data.append(row)
            
            # 创建DataFrame并保存
            df = pd.DataFrame(data)
            df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"结果已保存到: {output_path}")
            
        except Exception as e:
            print(f"保存Excel文件时出错: {e}")
    
    def run(self, word_file: str, output_file: str = None, 
            additional_checks: List[str] = None):
        """
        运行完整的检查流程
        
        Args:
            word_file: Word文档路径
            output_file: 输出Excel文件路径
            additional_checks: 额外检查要求
        """
        # 设置输出文件名
        if not output_file:
            word_path = Path(word_file)
            output_file = word_path.parent / f"{word_path.stem}_语法检查结果.xlsx"
        
        print("=== Word文档AI语法检查器 ===")
        print(f"输入文件: {word_file}")
        print(f"输出文件: {output_file}")
        print(f"使用模型: {self.config['model']}")
        
        # 读取文档
        paragraphs = self.read_word_document(word_file)
        if not paragraphs:
            print("未能读取到有效段落，程序退出")
            return
        
        # 处理段落
        results = self.process_paragraphs(paragraphs, additional_checks)
        
        # 保存结果
        self.save_to_excel(results, str(output_file))
        
        print(f"\n=== 处理完成 ===")
        print(f"总段落数: {len(paragraphs)}")
        print(f"AI会话次数: {self.session_count + 1}")
        print(f"结果文件: {output_file}")


def main():
    parser = argparse.ArgumentParser(description="Word文档AI语法检查器")
    parser.add_argument("word_file", help="Word文档路径")
    parser.add_argument("-o", "--output", help="输出Excel文件路径")
    parser.add_argument("-c", "--config", default="config.json", help="配置文件路径")
    parser.add_argument("--additional-checks", nargs="*", 
                       help="额外检查要求，例如：--additional-checks '检查用词准确性' '检查逻辑连贯性'")
    
    args = parser.parse_args()
    
    # 检查输入文件
    if not os.path.exists(args.word_file):
        print(f"错误：文件不存在 {args.word_file}")
        return
    
    # 创建检查器并运行
    try:
        checker = GrammarChecker(args.config)
    except ConfigFileNotFoundError as e:
        print(e)
        print("使用默认配置继续运行。")
        checker = GrammarChecker(args.config)
    checker.run(
        word_file=args.word_file,
        output_file=args.output,
        additional_checks=args.additional_checks
    )


if __name__ == "__main__":
    main()
