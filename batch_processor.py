#!/usr/bin/env python3
"""
批量处理Word文档的语法检查
用于一次性处理多个Word文档
"""

import os
import glob
from pathlib import Path
from grammar_checker import GrammarChecker, ConfigFileNotFoundError


def batch_process_documents(input_folder: str, output_folder: str = None, 
                           file_pattern: str = "*.docx",
                           additional_checks: list = None):
    """
    批量处理Word文档
    
    Args:
        input_folder: 输入文件夹路径
        output_folder: 输出文件夹路径（可选，默认为输入文件夹）
        file_pattern: 文件匹配模式
        additional_checks: 额外检查要求
    """
    # 设置路径
    input_path = Path(input_folder)
    if not input_path.exists():
        print(f"输入文件夹不存在: {input_folder}")
        return
    
    if output_folder:
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)
    else:
        output_path = input_path
    
    # 查找所有Word文档
    pattern = str(input_path / file_pattern)
    word_files = glob.glob(pattern)
    
    if not word_files:
        print(f"在 {input_folder} 中未找到匹配的文件: {file_pattern}")
        return
    
    print(f"找到 {len(word_files)} 个文档需要处理")
    
    # 初始化检查器
    try:
        checker = GrammarChecker()
    except ConfigFileNotFoundError as e:
        print(e)
        print("使用默认配置继续运行。")
        checker = GrammarChecker()
    
    # 处理每个文档
    for i, word_file in enumerate(word_files, 1):
        print(f"\n=== 处理文档 {i}/{len(word_files)} ===")
        
        # 生成输出文件名
        word_path = Path(word_file)
        output_file = output_path / f"{word_path.stem}_语法检查结果.xlsx"
        
        try:
            checker.run(
                word_file=word_file,
                output_file=str(output_file),
                additional_checks=additional_checks
            )
            print(f"✓ 完成: {word_path.name}")
            
        except Exception as e:
            print(f"✗ 处理失败 {word_path.name}: {e}")
    
    print(f"\n=== 批量处理完成 ===")
    print(f"输出文件夹: {output_path}")


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="批量处理Word文档语法检查")
    parser.add_argument("input_folder", help="输入文件夹路径")
    parser.add_argument("-o", "--output", help="输出文件夹路径")
    parser.add_argument("-p", "--pattern", default="*.docx", help="文件匹配模式")
    parser.add_argument("--additional-checks", nargs="*", 
                       help="额外检查要求")
    
    args = parser.parse_args()
    
    batch_process_documents(
        input_folder=args.input_folder,
        output_folder=args.output,
        file_pattern=args.pattern,
        additional_checks=args.additional_checks
    )


if __name__ == "__main__":
    main()
