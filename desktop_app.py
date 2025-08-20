#!/usr/bin/env python3
"""
Word文档AI语法检查器 - PyQt6桌面版
现代化的原生桌面界面
"""

import sys
import json
import os
import time
import tempfile
from pathlib import Path
from typing import List, Dict
import pandas as pd

try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit, QFileDialog,
        QProgressBar, QTabWidget, QFormLayout, QSpinBox, QDoubleSpinBox,
        QCheckBox, QListWidget, QListWidgetItem, QMessageBox, QSplitter,
        QGroupBox, QRadioButton, QButtonGroup, QScrollArea
    )
    from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
    from PyQt6.QtGui import QFont, QIcon, QPixmap
    from docx import Document
except ImportError as e:
    print(f"缺少必要的库: {e}")
    print("请运行: pip install PyQt6 python-docx litellm pandas openpyxl")
    sys.exit(1)

from utils.checker_core import process_paragraphs as core_process_paragraphs


class ProcessingThread(QThread):
    """后台处理线程"""
    progress_updated = pyqtSignal(int, str)
    processing_finished = pyqtSignal(pd.DataFrame)
    error_occurred = pyqtSignal(str)

    def __init__(self, paragraphs, config):
        super().__init__()
        self.paragraphs = paragraphs
        self.config = config

    def run(self):
        """运行处理任务"""
        try:
            def callback(i: int, total: int, message: str):
                progress = int((i + 1) / total * 100)
                self.progress_updated.emit(progress, message)

            results = core_process_paragraphs(
                self.paragraphs, self.config, progress_callback=callback
            )
            df = pd.DataFrame(results)
            self.processing_finished.emit(df)
        except Exception as e:
            self.error_occurred.emit(str(e))


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.current_paragraphs = []
        self.processing_thread = None
        self.init_ui()
        self.load_config()
    
    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("AI语法检查器")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QHBoxLayout(central_widget)
        
        # 创建分割器
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        
        # 左侧：配置面板
        left_widget = self.create_config_panel()
        splitter.addWidget(left_widget)
        
        # 右侧：主要操作面板
        right_widget = self.create_main_panel()
        splitter.addWidget(right_widget)
        
        # 设置分割器比例
        splitter.setSizes([400, 800])
    
    def create_config_panel(self) -> QWidget:
        """创建配置面板"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 标题
        title = QLabel("⚙️ 配置设置")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # API设置组
        api_group = QGroupBox("🔑 API设置")
        api_layout = QFormLayout(api_group)
        
        self.openai_key_input = QLineEdit()
        self.openai_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        api_layout.addRow("OpenAI API Key:", self.openai_key_input)
        
        self.gemini_key_input = QLineEdit()
        self.gemini_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        api_layout.addRow("Gemini API Key:", self.gemini_key_input)
        
        scroll_layout.addWidget(api_group)
        
        # 模型设置组
        model_group = QGroupBox("🤖 模型设置")
        model_layout = QFormLayout(model_group)
        
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["openai", "gemini"])
        self.provider_combo.currentTextChanged.connect(self.on_provider_changed)
        model_layout.addRow("AI供应商:", self.provider_combo)
        
        self.model_combo = QComboBox()
        model_layout.addRow("模型:", self.model_combo)
        
        scroll_layout.addWidget(model_group)
        
        # 语言设置组
        language_group = QGroupBox("🌐 语言设置")
        language_layout = QVBoxLayout(language_group)
        
        self.language_group = QButtonGroup()
        self.chinese_radio = QRadioButton("中文")
        self.english_radio = QRadioButton("English")
        self.chinese_radio.setChecked(True)
        
        self.language_group.addButton(self.chinese_radio, 0)
        self.language_group.addButton(self.english_radio, 1)
        
        language_layout.addWidget(self.chinese_radio)
        language_layout.addWidget(self.english_radio)
        scroll_layout.addWidget(language_group)
        
        # 高级设置组
        advanced_group = QGroupBox("🔧 高级设置")
        advanced_layout = QFormLayout(advanced_group)
        
        self.max_retries_spin = QSpinBox()
        self.max_retries_spin.setRange(1, 10)
        self.max_retries_spin.setValue(3)
        advanced_layout.addRow("最大重试次数:", self.max_retries_spin)
        
        self.retry_delay_spin = QDoubleSpinBox()
        self.retry_delay_spin.setRange(0.1, 10.0)
        self.retry_delay_spin.setValue(1.0)
        self.retry_delay_spin.setSingleStep(0.1)
        advanced_layout.addRow("重试延迟(秒):", self.retry_delay_spin)
        
        self.session_interval_spin = QSpinBox()
        self.session_interval_spin.setRange(1, 20)
        self.session_interval_spin.setValue(3)
        advanced_layout.addRow("会话刷新间隔:", self.session_interval_spin)
        
        scroll_layout.addWidget(advanced_group)
        
        # 配置文件操作
        config_group = QGroupBox("💾 配置管理")
        config_layout = QVBoxLayout(config_group)
        
        save_config_btn = QPushButton("保存配置")
        save_config_btn.clicked.connect(self.save_config)
        config_layout.addWidget(save_config_btn)
        
        load_config_btn = QPushButton("加载配置")
        load_config_btn.clicked.connect(self.load_config_file)
        config_layout.addWidget(load_config_btn)
        
        scroll_layout.addWidget(config_group)
        
        scroll_area.setWidget(scroll_content)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)
        
        return widget
    
    def create_main_panel(self) -> QWidget:
        """创建主操作面板"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 文件上传区域
        file_group = QGroupBox("📁 文件操作")
        file_layout = QVBoxLayout(file_group)
        
        file_btn_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("选择Word文档")
        self.select_file_btn.clicked.connect(self.select_word_file)
        file_btn_layout.addWidget(self.select_file_btn)
        
        self.file_label = QLabel("未选择文件")
        file_btn_layout.addWidget(self.file_label)
        file_btn_layout.addStretch()
        
        file_layout.addLayout(file_btn_layout)
        
        # 文档预览
        self.preview_text = QTextEdit()
        self.preview_text.setMaximumHeight(150)
        self.preview_text.setPlaceholderText("文档预览将在这里显示...")
        file_layout.addWidget(self.preview_text)
        
        layout.addWidget(file_group)
        
        # 输出设置
        output_group = QGroupBox("📊 输出设置")
        output_layout = QFormLayout(output_group)
        
        output_btn_layout = QHBoxLayout()
        self.output_path_input = QLineEdit()
        self.output_path_input.setText(str(Path.home() / "语法检查结果.xlsx"))
        output_btn_layout.addWidget(self.output_path_input)
        
        browse_btn = QPushButton("浏览")
        browse_btn.clicked.connect(self.browse_output_path)
        output_btn_layout.addWidget(browse_btn)
        
        output_layout.addRow("Excel输出路径:", output_btn_layout)
        layout.addWidget(output_group)
        
        # 额外检查要求
        checks_group = QGroupBox("✅ 额外检查要求")
        checks_layout = QVBoxLayout(checks_group)
        
        add_check_layout = QHBoxLayout()
        self.new_check_input = QLineEdit()
        self.new_check_input.setPlaceholderText("输入新的检查要求...")
        add_check_layout.addWidget(self.new_check_input)
        
        add_check_btn = QPushButton("添加")
        add_check_btn.clicked.connect(self.add_check_requirement)
        add_check_layout.addWidget(add_check_btn)
        
        checks_layout.addLayout(add_check_layout)
        
        self.checks_list = QListWidget()
        checks_layout.addWidget(self.checks_list)
        
        remove_check_btn = QPushButton("删除选中项")
        remove_check_btn.clicked.connect(self.remove_check_requirement)
        checks_layout.addWidget(remove_check_btn)
        
        layout.addWidget(checks_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel()
        self.status_label.setVisible(False)
        layout.addWidget(self.status_label)
        
        # 开始按钮
        self.start_btn = QPushButton("🚀 开始语法检查")
        self.start_btn.setMinimumHeight(50)
        self.start_btn.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.start_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.start_btn)
        
        return widget
    
    def on_provider_changed(self, provider):
        """供应商改变时更新模型列表"""
        self.model_combo.clear()
        
        if provider == "openai":
            models = ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-4", "gpt-3.5-turbo"]
        elif provider == "gemini":
            models = ["gemini-pro", "gemini-pro-vision", "gemini-1.5-pro", "gemini-1.5-flash"]
        
        self.model_combo.addItems(models)
    
    def select_word_file(self):
        """选择Word文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Word文档", "", "Word文档 (*.docx)"
        )
        
        if file_path:
            self.file_label.setText(Path(file_path).name)
            self.load_word_document(file_path)
    
    def load_word_document(self, file_path):
        """加载Word文档"""
        try:
            doc = Document(file_path)
            paragraphs = []
            
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    paragraphs.append(text)
            
            self.current_paragraphs = paragraphs
            
            # 显示预览
            preview_text = f"文档包含 {len(paragraphs)} 个段落\n\n"
            for i, para in enumerate(paragraphs[:3]):
                preview_text += f"段落 {i+1}: {para[:100]}...\n\n"
            
            self.preview_text.setText(preview_text)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取Word文档失败: {e}")
    
    def browse_output_path(self):
        """浏览输出路径"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel文件", "", "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            self.output_path_input.setText(file_path)
    
    def add_check_requirement(self):
        """添加检查要求"""
        text = self.new_check_input.text().strip()
        if text:
            self.checks_list.addItem(text)
            self.new_check_input.clear()
    
    def remove_check_requirement(self):
        """删除检查要求"""
        current_row = self.checks_list.currentRow()
        if current_row >= 0:
            self.checks_list.takeItem(current_row)
    
    def get_additional_checks(self) -> List[str]:
        """获取额外检查要求列表"""
        checks = []
        for i in range(self.checks_list.count()):
            checks.append(self.checks_list.item(i).text())
        return checks
    
    def save_config(self):
        """保存配置"""
        config = {
            "openai_api_key": self.openai_key_input.text(),
            "gemini_api_key": self.gemini_key_input.text(),
            "provider": self.provider_combo.currentText(),
            "model": self.model_combo.currentText(),
            "language": "中文" if self.chinese_radio.isChecked() else "English",
            "max_retries": self.max_retries_spin.value(),
            "retry_delay": self.retry_delay_spin.value(),
            "session_refresh_interval": self.session_interval_spin.value(),
            "additional_checks": self.get_additional_checks()
        }
        
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, "成功", "配置已保存到 config.json")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存配置失败: {e}")
    
    def load_config(self):
        """加载配置"""
        try:
            if os.path.exists("config.json"):
                with open("config.json", "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                self.openai_key_input.setText(config.get("openai_api_key", ""))
                self.gemini_key_input.setText(config.get("gemini_api_key", ""))
                
                provider = config.get("provider", "openai")
                self.provider_combo.setCurrentText(provider)
                self.on_provider_changed(provider)
                
                self.model_combo.setCurrentText(config.get("model", "gpt-3.5-turbo"))
                
                language = config.get("language", "中文")
                if language == "中文":
                    self.chinese_radio.setChecked(True)
                else:
                    self.english_radio.setChecked(True)
                
                self.max_retries_spin.setValue(config.get("max_retries", 3))
                self.retry_delay_spin.setValue(config.get("retry_delay", 1.0))
                self.session_interval_spin.setValue(config.get("session_refresh_interval", 3))
                
                # 加载额外检查要求
                for check in config.get("additional_checks", []):
                    self.checks_list.addItem(check)
        
        except Exception as e:
            print(f"加载配置失败: {e}")
    
    def load_config_file(self):
        """从文件加载配置"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择配置文件", "", "JSON文件 (*.json)"
        )
        
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                # 这里可以复用load_config的逻辑
                QMessageBox.information(self, "成功", "配置文件加载成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载配置文件失败: {e}")
    
    def start_processing(self):
        """开始处理"""
        # 验证输入
        if not self.current_paragraphs:
            QMessageBox.warning(self, "警告", "请先选择Word文档")
            return
        
        provider = self.provider_combo.currentText()
        api_key = (self.openai_key_input.text() if provider == "openai" 
                  else self.gemini_key_input.text())
        
        if not api_key:
            QMessageBox.warning(self, "警告", f"请输入{provider.upper()} API密钥")
            return
        
        # 准备配置
        config = {
            "provider": provider,
            "model": self.model_combo.currentText(),
            "api_key": api_key,
            "language": "中文" if self.chinese_radio.isChecked() else "English",
            "session_refresh_interval": self.session_interval_spin.value(),
            "additional_checks": self.get_additional_checks()
        }
        
        # 禁用开始按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.status_label.setVisible(True)
        
        # 创建并启动处理线程
        self.processing_thread = ProcessingThread(self.current_paragraphs, config)
        self.processing_thread.progress_updated.connect(self.update_progress)
        self.processing_thread.processing_finished.connect(self.on_processing_finished)
        self.processing_thread.error_occurred.connect(self.on_processing_error)
        self.processing_thread.start()
    
    def update_progress(self, value, message):
        """更新进度"""
        self.progress_bar.setValue(value)
        self.status_label.setText(message)
    
    def on_processing_finished(self, result_df):
        """处理完成"""
        try:
            # 保存Excel文件
            output_path = self.output_path_input.text()
            result_df.to_excel(output_path, index=False, engine='openpyxl')
            
            QMessageBox.information(
                self, "完成", 
                f"语法检查完成！\n结果已保存到: {output_path}\n共处理 {len(result_df)} 个段落"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存Excel文件失败: {e}")
        
        finally:
            # 重置界面
            self.start_btn.setEnabled(True)
            self.progress_bar.setVisible(False)
            self.status_label.setVisible(False)
    
    def on_processing_error(self, error_message):
        """处理错误"""
        QMessageBox.critical(self, "错误", f"处理过程中出现错误: {error_message}")
        
        # 重置界面
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setVisible(False)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用现代化样式
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
