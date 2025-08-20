#!/usr/bin/env python3
"""
Word文档AI语法检查脚本
功能：读取Word文档，分段后使用AI进行语法检查，结果保存到Excel
"""

import os
import json
import argparse
from typing import List, Dict
from pathlib import Path

from utils.checker_core import process_paragraphs as core_process_paragraphs


class ConfigFileNotFoundError(Exception):
    """在配置文件缺失时抛出的异常"""

    def __init__(self, message: str, config: Dict):
        super().__init__(message)
        self.config = config

try:
    from docx import Document
    import pandas as pd
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
        provider = "gemini" if "gemini" in self.config["model"].lower() else "openai"
        api_key = (
            self.config.get("gemini_api_key")
            if provider == "gemini"
            else self.config.get("openai_api_key")
        )
        if not api_key:
            print(f"错误：缺少{provider.upper()} API密钥")
            return
        cfg = {
            "language": "中文",
            "provider": provider,
            "model": self.config["model"],
            "api_key": api_key,
            "additional_checks": additional_checks or self.config.get("additional_checks", []),
            "session_refresh_interval": self.config.get("session_refresh_interval", 3),
            "max_retries": self.config.get("max_retries", 3),
            "retry_delay": self.config.get("retry_delay", 1),
        }

        with tqdm(total=len(paragraphs), desc="处理进度") as pbar:
            def callback(i, total, message):
                pbar.update(1)
            results = core_process_paragraphs(paragraphs, cfg, progress_callback=callback)

        # 保存结果
        self.save_to_excel(results, str(output_file))

        session_interval = cfg["session_refresh_interval"]
        total_sessions = len(paragraphs) // session_interval + 1
        print(f"\n=== 处理完成 ===")
        print(f"总段落数: {len(paragraphs)}")
        print(f"AI会话次数: {total_sessions}")
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
