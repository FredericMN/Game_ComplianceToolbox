import os
import sys
import json
from datetime import datetime
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QHBoxLayout, QWidget, QFrame, QPushButton,
    QTextEdit, QTreeWidget, QTreeWidgetItem, QScrollArea
)
from PySide6.QtCore import Qt, QThread, Signal
from.base_interface import BaseInterface
from PySide6.QtGui import QFont, QIcon, QColor, QPalette
from utils.environment_checker import EnvironmentChecker


class WelcomeInterface(BaseInterface):
    """欢迎页界面"""

    environment_check_started = Signal()
    environment_check_finished = Signal(bool, bool)  # (has_errors, is_new_check)
    card_clicked = Signal(str)  # 新增信号，用于发送点击的卡片名称
    is_checking_env = False  # 新增标志，用于记录是否正在进行环境检测

    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = None
        self.environment_checker = None
        self._connections_established = False
        self.env_result_file = self.get_env_result_file_path()
        self.init_ui()

    def get_current_dir(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.dirname(os.path.abspath(__file__))

    def get_env_result_file_path(self):
        CURRENT_DIR = self.get_current_dir()
        return os.path.join(CURRENT_DIR, 'env_check_result.json')

    def check_env_status(self):
        """检查环境状态，增加缓存时间限制"""
        if os.path.exists(self.env_result_file):
            try:
                with open(self.env_result_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    date_str = data.get('date')
                    result = data.get('result')
                    check_time = datetime.strptime(date_str, "%Y-%m-%d")
                    
                    # 如果检测结果超过7天，重新检测
                    if (datetime.now() - check_time).days > 7:
                        self.run_environment_check()
                        return False
                        
                    if date_str and (result is not None):
                        timestamp = datetime.now().strftime("%H:%M:%S")
                        if result:
                            message = f"[{timestamp}] {date_str} 检测环境：通过。可直接使用。如遇问题可再次检测！"
                            # 如果检测通过，启用导航栏，但不触发弹窗
                            self.environment_check_finished.emit(False, False)
                        else:
                            message = f"[{timestamp}] {date_str} 检测环境：不通过。建议再次检测或获取帮助！"
                            # 如果检测不通过，保持导航栏禁用状态
                            self.environment_check_finished.emit(True, False)
                        self.output_text_edit.append(message)
                        return True
            except Exception as e:
                self.output_text_edit.append(f"[{datetime.now().strftime('%H:%M:%S')}] 读取检测结果失败，需要重新检测。错误信息：{str(e)}")

        # 如果没有历史记录或读取失败，则进行新的检测
        self.run_environment_check()
        return False

    def init_ui(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: #f5f6fa;
            }
        """)

        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 顶部欢迎区域
        welcome_widget = QWidget()
        welcome_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 10px;
            }
        """)
        welcome_layout = QVBoxLayout(welcome_widget)
        welcome_layout.setContentsMargins(10, 5, 10, 5)

        welcome_label = QLabel("欢迎使用合规工具箱")
        welcome_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                font-size: 24px;
                font-weight: bold;
                padding: 5px;
            }
        """)
        welcome_label.setAlignment(Qt.AlignCenter)
        welcome_layout.addWidget(welcome_label)
        main_layout.addWidget(welcome_widget)

        # 功能卡片区域 - 调整内边距和间距
        functions_widget = QWidget()
        functions_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 8px; /* 减小内边距 */
            }
        """)
        functions_grid = QHBoxLayout(functions_widget)
        functions_grid.setSpacing(8)
        functions_grid.setContentsMargins(8, 6, 8, 6)  # 减小上下边距

        # 左右两列的容器
        left_column = QVBoxLayout()
        right_column = QVBoxLayout()
        left_column.setSpacing(4)  # 减小行间距
        right_column.setSpacing(4)  # 减小行间距

        functions = [
            {"name": "文档风险词汇批量检测", "description": "检测并标记文档中的风险词汇。", "icon": "🔍"},
            {"name": "新游爬虫", "description": "爬取TapTap上的新游信息并匹配版号。", "icon": "🕷️"},
            {"name": "版号匹配", "description": "匹配游戏的版号信息。", "icon": "📋"},
            {"name": "词表对照", "description": "对照两个词表的差异。", "icon": "📊"},
            {"name": "大模型语义分析", "description": "通过大模型审核文本，标记高风险内容。", "icon": "🤖"},
            {"name": "大模型文案正向优化", "description": "通过大模型输出语句的正向优化。", "icon": "✨"},
            {"name": "设定", "description": "配置工具的相关设置。", "icon": "⚙️"}
        ]

        # 功能卡片样式调整
        for i, func in enumerate(functions):
            card = QWidget()
            card.setObjectName("functionCard")  # 添加对象名，用于样式表识别
            card.setStyleSheet("""
                QWidget#functionCard {
                    background-color: #f8f9fa;
                    border-radius: 8px;
                    padding: 4px 6px;
                    margin: 1px;
                }
                QWidget#functionCard:hover {
                    background-color: #e9ecef;
                }
                QWidget#functionCard[disabled="true"] {
                    background-color: #f8f9fa;
                }
                QLabel {
                    background: transparent;  /* 确保标签背景透明 */
                }
            """)
            # 使用 setCursor 来设置鼠标指针样式
            card.setCursor(Qt.PointingHandCursor)

            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(6, 4, 6, 4)
            card_layout.setSpacing(8)

            icon_label = QLabel(func["icon"])
            icon_label.setStyleSheet("""
                QLabel {
                    font-size: 20px;
                    min-width: 30px;
                    background: transparent;
                }
            """)

            text_widget = QWidget()
            text_widget.setStyleSheet("background: transparent;")  # 确保背景透明
            text_layout = QVBoxLayout(text_widget)
            text_layout.setSpacing(0)
            text_layout.setContentsMargins(0, 0, 0, 0)

            name_label = QLabel(func["name"])
            name_label.setStyleSheet("""
                QLabel {
                    color: #2c3e50;
                    font-size: 14px;
                    font-weight: bold;
                    background: transparent;
                }
            """)

            desc_label = QLabel(func["description"])
            desc_label.setStyleSheet("""
                QLabel {
                    color: #7f8c8d;
                    font-size: 12px;
                    background: transparent;
                }
            """)
            desc_label.setWordWrap(True)

            text_layout.addWidget(name_label)
            text_layout.addWidget(desc_label)

            card_layout.addWidget(icon_label)
            card_layout.addWidget(text_widget, 1)

            # 为卡片添加点击事件
            card.mousePressEvent = lambda event, name=func["name"]: self.on_card_clicked(name)
            
            # 存储卡片引用以便后续控制
            card.setProperty("disabled", False)  # 添加禁用状态属性
            if not hasattr(self, 'function_cards'):
                self.function_cards = []
            self.function_cards.append(card)

            if i % 2 == 0:
                left_column.addWidget(card)
            else:
                right_column.addWidget(card)

        functions_grid.addLayout(left_column)
        functions_grid.addLayout(right_column)
        main_layout.addWidget(functions_widget)

        # 环境检测区域
        env_widget = QWidget()
        env_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 6px;
            }
        """)
        env_layout = QHBoxLayout(env_widget)
        env_layout.setSpacing(10)
        env_layout.setContentsMargins(8, 4, 8, 4)

        # 左侧检测按钮和说明区域
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(4)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        self.check_env_button = QPushButton("检测运行环境")
        self.check_env_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: bold;
                font-size: 14px;
                min-width: 160px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #2573a7;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.check_env_button.clicked.connect(self.run_environment_check)

        description_label = QLabel("每次运行软件时会自动检测运行环境\n需要已安装Edge浏览器")
        description_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 12px;
                margin-top: 2px;
            }
        """)
        description_label.setAlignment(Qt.AlignLeft)
        description_label.setWordWrap(True)

        left_layout.addWidget(self.check_env_button)
        left_layout.addWidget(description_label)
        left_layout.addStretch()  # 添加弹性空间，使按钮和说明文字固定在顶部

        # 右侧输出区域
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 10px;
                background-color: #f8f9fa;
                color: #2c3e50;
                font-size: 12px;
                line-height: 1.4;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
        """)
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("环境检测信息将在此处显示...")
        self.output_text_edit.setMinimumHeight(230)  # 增加最小高度
        self.output_text_edit.setMaximumHeight(230)  # 增加最大高度

        # 设置左右区域的比例为 1:2，让输出区域更大
        env_layout.addWidget(left_container, 1)
        env_layout.addWidget(self.output_text_edit, 2)
        main_layout.addWidget(env_widget)

        # 调整主布局的间距
        main_layout.setSpacing(10)

        # 遮罩层
        self.overlay = QWidget(self)
        self.overlay.setStyleSheet("""
            QWidget {
                background-color: rgba(0, 0, 0, 150);
            }
            QLabel {
                color: white;
                font-size: 20px;
                background-color: transparent;
            }
        """)
        self.overlay_layout = QVBoxLayout(self.overlay)
        self.overlay_layout.setAlignment(Qt.AlignCenter)
        self.overlay_label = QLabel("正在检测运行环境，请稍候...")
        self.overlay_layout.addWidget(self.overlay_label)
        self.overlay.hide()

        scroll_area.setWidget(main_widget)
        self.layout.addWidget(scroll_area)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.overlay.resize(self.size())

    def run_environment_check(self):
        """执行环境检测"""
        if hasattr(self, 'thread') and self.thread and self.thread.isRunning():
            return

        self.is_checking_env = True
        self.check_env_button.setEnabled(False)
        self.environment_check_started.emit()
        self.overlay.show()

        # 禁用所有功能卡片
        if hasattr(self, 'function_cards'):
            for card in self.function_cards:
                card.setProperty("disabled", True)
                card.setCursor(Qt.ArrowCursor)  # 设置为默认鼠标指针
                card.setStyle(card.style())  # 刷新样式

        # 创建新的线程和工作对象
        self.thread = QThread()
        self.environment_checker = EnvironmentChecker()
        self.environment_checker.moveToThread(self.thread)

        # 连接信号
        self.thread.started.connect(self.environment_checker.run)
        self.environment_checker.output_signal.connect(self.append_output)
        self.environment_checker.structured_result_signal.connect(self.on_structured_results)
        self.environment_checker.finished.connect(self.on_check_finished)
        self.environment_checker.finished.connect(self.cleanup_check)

        self.thread.start()

    def cleanup_check(self):
        """清理检测相关资源"""
<<<<<<< HEAD
        if self.thread and self.thread.isRunning():
            try:
                # 先把对象引用保存下来
                thread = self.thread
                environment_checker = self.environment_checker
                
                # 先将属性设置为None，防止其他地方重复调用
                self.thread = None
                self.environment_checker = None
                
                # 停止线程
                thread.quit()
                # 等待线程结束，如果超时就强制终止
                if not thread.wait(3000):  # 等待最多3秒
                    self.output_text_edit.append("环境检测线程超时，强制终止")
                
                # 断开所有信号连接
                try:
                    thread.started.disconnect()
                    if environment_checker:
                        environment_checker.output_signal.disconnect()
                        environment_checker.structured_result_signal.disconnect()
                        environment_checker.finished.disconnect()
                except (TypeError, RuntimeError):
                    # 忽略已断开连接的异常
                    pass
                
                # 安全删除对象
                if environment_checker:
                    environment_checker.deleteLater()
                thread.deleteLater()
                
                # 清理任何残留的msedgedriver进程 - 使用延迟导入避免循环导入
                try:
                    # 延迟导入kill_msedgedriver函数
                    import importlib
                    main_module = importlib.import_module('main')
                    if main_module.kill_msedgedriver():
                        self.output_text_edit.append("[系统] 已清理残留的msedgedriver进程")
                except Exception as e:
                    self.output_text_edit.append(f"[错误] 清理msedgedriver进程时出错: {str(e)}")
            except Exception as e:
                self.output_text_edit.append(f"[错误] 清理资源时出错: {str(e)}")
=======
        try:
            if self.thread and self.thread.isRunning():
                self.thread.quit()
                self.thread.wait()

                # 断开所有信号连接
                if hasattr(self.environment_checker, 'output_signal'):
                    try:
                        self.environment_checker.output_signal.disconnect()
                    except:
                        pass
                if hasattr(self.environment_checker, 'structured_result_signal'):
                    try:
                        self.environment_checker.structured_result_signal.disconnect()
                    except:
                        pass
                if hasattr(self.environment_checker, 'finished'):
                    try:
                        self.environment_checker.finished.disconnect()
                    except:
                        pass

                # 安全删除对象
                if self.thread:
                    self.thread.deleteLater()
                if self.environment_checker:
                    self.environment_checker.deleteLater()

        except Exception as e:
            print(f"清理资源时出现异常: {str(e)}")
        finally:
            self.thread = None
            self.environment_checker = None
>>>>>>> 79605118dcf5a6a6197f5cf20a3439c6a47db93d

    def append_output(self, message):
        """优化输出信息显示"""
        # 过滤掉不需要显示的结构化结果
        if not any(prefix in message for prefix in ["网络连接检测:", "Edge浏览器检测:", "Edge WebDriver检测:"]):
            # 添加时间戳和美化格式
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {message}"
            self.output_text_edit.append(formatted_message)

    def on_structured_results(self, results):
        """处理结构化结果，不直接输出"""
        # 仅用于内部处理，不输出到界面
        pass

    def on_check_finished(self, has_errors):
        """检测完成的处理"""
        self.is_checking_env = False
        self.check_env_button.setEnabled(True)
        self.overlay.hide()

        # 启用所有功能卡片
        if hasattr(self, 'function_cards'):
            for card in self.function_cards:
                card.setProperty("disabled", False)
                card.setCursor(Qt.PointingHandCursor)  # 恢复指针手型
                card.setStyle(card.style())  # 刷新样式

        # 获取当前日期并记录结果
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.record_env_check_result(current_date, not has_errors)

        # 发送信号
        self.environment_check_finished.emit(has_errors, True)

        # 清理资源
        self.cleanup_check()

    def record_env_check_result(self, date_str, result):
        data = {
            "date": date_str,
            "result": result
        }
        try:
            with open(self.env_result_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.output_text_edit.append(f"记录环境检测结果失败：{str(e)}")

    def on_card_clicked(self, name):
        """处理卡片点击事件"""
        if not self.is_checking_env:  # 只有在非检测状态下才处理点击事件
            self.card_clicked.emit(name)
