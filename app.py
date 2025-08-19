#!/usr/bin/env python3
"""
Word文档AI语法检查器 - Streamlit Web界面
现代化的图形界面，支持所有自定义功能
"""

import streamlit as st
import json
import os
import time
import tempfile
from pathlib import Path
from typing import List, Dict, Optional
import pandas as pd

try:
    from docx import Document
    import litellm
    from io import BytesIO
except ImportError as e:
    st.error(f"缺少必要的库: {e}")
    st.info("请运行: pip install python-docx litellm streamlit")
    st.stop()

# 页面配置
st.set_page_config(
    page_title="AI语法检查器",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

class StreamlitGrammarChecker:
    def __init__(self):
        self.available_models = {
            "openai": [],
            "gemini": []
        }
        self.session_initialized = False
    
    def initialize_session_state(self):
        """初始化Streamlit会话状态"""
        if not self.session_initialized:
            # 默认配置
            default_config = {
                "openai_api_key": "",
                "gemini_api_key": "",
                "max_retries": 3,
                "retry_delay": 1,
                "session_refresh_interval": 3,
                "additional_checks": []
            }
            
            for key, value in default_config.items():
                if key not in st.session_state:
                    st.session_state[key] = value
            
            # 其他会话状态
            if "language" not in st.session_state:
                st.session_state.language = "中文"
            if "provider" not in st.session_state:
                st.session_state.provider = "openai"
            if "model" not in st.session_state:
                st.session_state.model = "gpt-3.5-turbo"
            if "output_path" not in st.session_state:
                st.session_state.output_path = str(Path.home() / "语法检查结果.xlsx")
            
            self.session_initialized = True
    
    def get_available_models(self, provider: str, api_key: str) -> List[str]:
        """获取可用模型列表"""
        if not api_key:
            return []
        
        try:
            if provider == "openai":
                # OpenAI常用模型（实际使用时可以通过API获取）
                return [
                    "gpt-4o",
                    "gpt-4o-mini",
                    "gpt-4-turbo",
                    "gpt-4",
                    "gpt-3.5-turbo"
                ]
            elif provider == "gemini":
                # Gemini模型
                return [
                    "gemini-pro",
                    "gemini-pro-vision",
                    "gemini-1.5-pro",
                    "gemini-1.5-flash"
                ]
        except Exception as e:
            st.warning(f"获取{provider}模型列表失败: {e}")
        
        return []
    
    def read_word_document(self, uploaded_file) -> List[str]:
        """读取上传的Word文档"""
        try:
            # 保存临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
            
            # 读取文档
            doc = Document(tmp_path)
            paragraphs = []
            
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    paragraphs.append(text)
            
            # 清理临时文件
            os.unlink(tmp_path)
            
            return paragraphs
            
        except Exception as e:
            st.error(f"读取Word文档时出错: {e}")
            return []
    
    def create_prompts(self, text: str, language: str, check_type: str = "grammar", 
                      custom_requirement: str = "") -> str:
        """创建AI提示词"""
        if language == "中文":
            if check_type == "grammar":
                return f"""请检查以下文本的语法错误，只需要指出语法问题并给出简洁的修改建议：

文本：{text}

请用中文回答，格式如下：
- 如果没有语法错误，回答"语法正确"
- 如果有语法错误，简洁地指出问题和建议
"""
            else:
                return f"""请对以下文本进行检查：{custom_requirement}

文本：{text}

请用中文给出简洁的评价和建议：
"""
        else:  # 英文
            if check_type == "grammar":
                return f"""Please check the following text for grammar errors and provide concise suggestions:

Text: {text}

Please respond in English:
- If there are no grammar errors, respond "Grammar is correct"
- If there are grammar errors, briefly point out the issues and suggestions
"""
            else:
                return f"""Please check the following text for: {custom_requirement}

Text: {text}

Please provide concise evaluation and suggestions in English:
"""
    
    def call_ai_api(self, prompt: str, provider: str, model: str, api_key: str) -> str:
        """调用AI API"""
        try:
            # 设置API密钥
            if provider == "openai":
                litellm.openai_key = api_key
            elif provider == "gemini":
                litellm.gemini_key = api_key
            
            response = litellm.completion(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=500,
                temperature=0.3
            )
            return response.choices[0].message.content.strip()
            
        except Exception as e:
            return f"API调用失败: {str(e)}"
    
    def process_paragraphs(self, paragraphs: List[str], config: Dict) -> pd.DataFrame:
        """处理段落并返回结果DataFrame"""
        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        language = config["language"]
        provider = config["provider"]
        model = config["model"]
        api_key = config["api_key"]
        additional_checks = config["additional_checks"]
        session_interval = config["session_refresh_interval"]
        
        for i, paragraph in enumerate(paragraphs):
            # 更新进度
            progress = (i + 1) / len(paragraphs)
            progress_bar.progress(progress)
            status_text.text(f"处理第 {i+1}/{len(paragraphs)} 段...")
            
            # 检查是否需要刷新会话
            if i > 0 and i % session_interval == 0:
                status_text.text(f"刷新AI会话... 第 {i+1}/{len(paragraphs)} 段")
                time.sleep(1)
            
            result_row = {"原始文本": paragraph}
            
            # 语法检查
            grammar_prompt = self.create_prompts(paragraph, language, "grammar")
            grammar_result = self.call_ai_api(grammar_prompt, provider, model, api_key)
            result_row["语法检查"] = grammar_result
            
            # 额外检查
            for j, check_requirement in enumerate(additional_checks):
                if check_requirement.strip():
                    additional_prompt = self.create_prompts(
                        paragraph, language, "additional", check_requirement
                    )
                    additional_result = self.call_ai_api(
                        additional_prompt, provider, model, api_key
                    )
                    result_row[f"额外检查_{j+1}"] = additional_result
            
            results.append(result_row)
            time.sleep(0.5)  # 避免API限流
        
        progress_bar.progress(1.0)
        status_text.text("处理完成！")
        
        return pd.DataFrame(results)


def main():
    # 创建检查器实例
    checker = StreamlitGrammarChecker()
    checker.initialize_session_state()
    
    # 主标题
    st.title("📝 AI语法检查器")
    st.markdown("---")
    
    # 侧边栏：配置设置
    with st.sidebar:
        st.header("⚙️ 配置设置")
        
        # API设置
        st.subheader("🔑 API设置")
        openai_key = st.text_input(
            "OpenAI API Key",
            value=st.session_state.openai_api_key,
            type="password",
            help="请输入您的OpenAI API密钥"
        )
        st.session_state.openai_api_key = openai_key
        
        gemini_key = st.text_input(
            "Gemini API Key", 
            value=st.session_state.gemini_api_key,
            type="password",
            help="请输入您的Gemini API密钥"
        )
        st.session_state.gemini_api_key = gemini_key
        
        # 模型设置
        st.subheader("🤖 模型设置")
        provider = st.selectbox(
            "选择AI供应商",
            ["openai", "gemini"],
            index=0 if st.session_state.provider == "openai" else 1
        )
        st.session_state.provider = provider
        
        # 根据选择的供应商获取API密钥
        current_api_key = openai_key if provider == "openai" else gemini_key
        
        # 获取可用模型
        available_models = checker.get_available_models(provider, current_api_key)
        
        if available_models:
            model_index = 0
            if st.session_state.model in available_models:
                model_index = available_models.index(st.session_state.model)
            
            model = st.selectbox(
                "选择模型",
                available_models,
                index=model_index
            )
            st.session_state.model = model
        else:
            st.warning("请先输入API密钥以获取可用模型")
            model = st.text_input("手动输入模型名称", value=st.session_state.model)
            st.session_state.model = model
        
        # 语言设置
        st.subheader("🌐 语言设置")
        language = st.radio(
            "选择检查语言",
            ["中文", "English"],
            index=0 if st.session_state.language == "中文" else 1
        )
        st.session_state.language = language
        
        # 高级设置
        with st.expander("🔧 高级设置"):
            max_retries = st.number_input(
                "最大重试次数",
                min_value=1,
                max_value=10,
                value=st.session_state.max_retries
            )
            st.session_state.max_retries = max_retries
            
            retry_delay = st.number_input(
                "重试延迟(秒)",
                min_value=0.1,
                max_value=10.0,
                value=float(st.session_state.retry_delay),
                step=0.1
            )
            st.session_state.retry_delay = retry_delay
            
            session_interval = st.number_input(
                "会话刷新间隔(段落数)",
                min_value=1,
                max_value=20,
                value=st.session_state.session_refresh_interval
            )
            st.session_state.session_refresh_interval = session_interval
        
        # 配置文件操作
        st.subheader("💾 配置管理")
        if st.button("保存配置到文件"):
            config_data = {
                "openai_api_key": st.session_state.openai_api_key,
                "gemini_api_key": st.session_state.gemini_api_key,
                "provider": st.session_state.provider,
                "model": st.session_state.model,
                "language": st.session_state.language,
                "max_retries": st.session_state.max_retries,
                "retry_delay": st.session_state.retry_delay,
                "session_refresh_interval": st.session_state.session_refresh_interval,
                "additional_checks": st.session_state.additional_checks
            }
            
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            st.success("配置已保存到 config.json")
        
        uploaded_config = st.file_uploader("上传配置文件", type=["json"])
        if uploaded_config:
            try:
                config_data = json.load(uploaded_config)
                for key, value in config_data.items():
                    if key in st.session_state:
                        st.session_state[key] = value
                st.success("配置文件加载成功！")
                st.experimental_rerun()
            except Exception as e:
                st.error(f"配置文件加载失败: {e}")
    
    # 主界面
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 文件上传")
        uploaded_file = st.file_uploader(
            "选择Word文档",
            type=["docx"],
            help="请上传.docx格式的Word文档"
        )
        
        if uploaded_file:
            st.success(f"已上传: {uploaded_file.name}")
            
            # 读取并显示文档预览
            paragraphs = checker.read_word_document(uploaded_file)
            if paragraphs:
                st.info(f"文档包含 {len(paragraphs)} 个段落")
                
                with st.expander("📄 文档预览（前3段）"):
                    for i, para in enumerate(paragraphs[:3]):
                        st.write(f"**段落 {i+1}:** {para[:100]}...")
    
    with col2:
        st.header("📊 输出设置")
        output_path = st.text_input(
            "Excel输出路径",
            value=st.session_state.output_path,
            help="指定Excel文件的保存路径"
        )
        st.session_state.output_path = output_path
        
        if st.button("📂 选择保存文件夹"):
            st.info("请在文本框中直接输入完整路径")
    
    # 额外检查设置
    st.header("✅ 额外检查要求")
    
    # 显示现有的额外检查
    if "additional_checks" not in st.session_state:
        st.session_state.additional_checks = []
    
    # 添加新的检查要求
    new_check = st.text_input("添加新的检查要求")
    if st.button("➕ 添加") and new_check:
        st.session_state.additional_checks.append(new_check)
        st.experimental_rerun()
    
    # 显示和管理现有检查要求
    if st.session_state.additional_checks:
        st.subheader("当前检查要求:")
        for i, check in enumerate(st.session_state.additional_checks):
            col_check, col_delete = st.columns([4, 1])
            with col_check:
                st.write(f"{i+1}. {check}")
            with col_delete:
                if st.button("🗑️", key=f"delete_{i}"):
                    st.session_state.additional_checks.pop(i)
                    st.experimental_rerun()
    
    # 运行检查
    st.markdown("---")
    if st.button("🚀 开始语法检查", type="primary", use_container_width=True):
        # 验证必要条件
        if not uploaded_file:
            st.error("请先上传Word文档")
            return
        
        api_key = (st.session_state.openai_api_key if st.session_state.provider == "openai" 
                  else st.session_state.gemini_api_key)
        
        if not api_key:
            st.error(f"请先输入{st.session_state.provider.upper()} API密钥")
            return
        
        # 读取文档
        paragraphs = checker.read_word_document(uploaded_file)
        if not paragraphs:
            st.error("无法读取文档内容")
            return
        
        # 配置参数
        config = {
            "language": st.session_state.language,
            "provider": st.session_state.provider,
            "model": st.session_state.model,
            "api_key": api_key,
            "additional_checks": st.session_state.additional_checks,
            "session_refresh_interval": st.session_state.session_refresh_interval
        }
        
        # 处理文档
        st.subheader("🔄 处理中...")
        
        try:
            result_df = checker.process_paragraphs(paragraphs, config)
            
            # 显示结果
            st.subheader("📋 检查结果")
            st.dataframe(result_df, use_container_width=True)
            
            # 保存到Excel
            try:
                result_df.to_excel(st.session_state.output_path, index=False, engine='openpyxl')
                st.success(f"结果已保存到: {st.session_state.output_path}")
            except Exception as e:
                st.error(f"保存Excel文件失败: {e}")
            
            # 提供下载链接
            excel_buffer = BytesIO()
            result_df.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)
            
            st.download_button(
                label="📥 下载Excel文件",
                data=excel_buffer,
                file_name=f"语法检查结果_{time.strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"处理过程中出现错误: {e}")


if __name__ == "__main__":
    main()
