from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QTextEdit, QLabel, QFileDialog,
    QProgressBar, QDialog, QPushButton, QMessageBox
)
from PySide6.QtCore import Qt, QThread, Signal, QObject, QTimer
from qfluentwidgets import PrimaryPushButton
from .base_interface import BaseInterface
import sys
import os

# 导入著作权查询功能
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.copyright_query import CopyrightQuery

class LoginConfirmationDialog(QDialog):
    """登录确认对话框"""
    
    def __init__(self, parent=None, is_anti_crawl=False):
        super().__init__(parent)
        self.setWindowTitle("登录确认" if not is_anti_crawl else "反爬机制检测")
        self.resize(400, 150)
        
        layout = QVBoxLayout()
        
        # 提示信息
        if not is_anti_crawl:
            message = "检测到需要登录企查查账号\n请在浏览器中完成登录后，点击'已完成登录'按钮继续"
        else:
            message = "检测到网站反爬机制，需要重新登录或等待一段时间\n请在浏览器中重新登录或解除限制后，点击'已解决'按钮继续"
        
        label = QLabel(message)
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        
        confirm_button = QPushButton("已完成登录" if not is_anti_crawl else "已解决问题")
        confirm_button.clicked.connect(self.accept)
        button_layout.addWidget(confirm_button)
        
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)

class CopyrightQueryWorker(QObject):
    """著作权查询工作线程"""
    
    finished = Signal()
    progress = Signal(str)
    progress_percent = Signal(int)
    login_required = Signal()
    anti_crawl_detected = Signal()  # 新增反爬检测信号
    login_confirmed = Signal(bool)
    error_occurred = Signal(str)
    
    def __init__(self, excel_file):
        super().__init__()
        self.excel_file = excel_file
        self.is_running = True
        self.waiting_for_login = False
        self.copyright_query = None
        self.is_anti_crawl = False  # 标记是否为反爬检测
    
    def run(self):
        """执行著作权查询"""
        try:
            self.progress.emit("初始化著作权查询工具...")
            self.copyright_query = CopyrightQuery()
            
            # 创建进度回调函数
            def progress_callback(message, percent=None):
                if not self.is_running:
                    return False  # 返回False将终止查询
                self.progress.emit(message)
                if percent is not None:
                    self.progress_percent.emit(percent)
                return True  # 返回True继续查询
            
            # 创建用户确认函数
            def login_confirm_callback():
                self.waiting_for_login = True
                
                # 判断是登录需求还是反爬检测
                if "反爬机制" in self.copyright_query.update_progress.__self__.paused_reason:
                    self.is_anti_crawl = True
                    self.anti_crawl_detected.emit()  # 发送反爬检测信号
                else:
                    self.is_anti_crawl = False
                    self.login_required.emit()  # 发送登录需求信号
                
                # 在这里等待用户确认结果
                while self.waiting_for_login and self.is_running:
                    QThread.msleep(100)  # 小暂停，避免高CPU使用
                return self.login_result if hasattr(self, 'login_result') else False
            
            # 设置著作权查询对象的回调函数
            self.copyright_query.set_progress_callback(progress_callback)
            self.copyright_query.set_login_confirm_callback(login_confirm_callback)
            
            # 开始查询
            self.progress.emit("开始处理Excel文件...")
            self.progress_percent.emit(5)
            
            if self.is_running:
                # 如果有保存的索引，从该索引继续处理
                start_index = self.copyright_query.current_game_index
                result_file = self.copyright_query.process_excel(self.excel_file, start_from_index=start_index)
                if result_file:
                    self.progress.emit(f"查询完成！结果已保存到: {result_file}")
                else:
                    self.progress.emit("查询未完成，可能发生了错误")
            else:
                self.progress.emit("查询已被用户取消")
        
        except Exception as e:
            self.error_occurred.emit(f"发生错误: {str(e)}")
            self.progress.emit(f"查询过程中发生错误: {str(e)}")
        
        finally:
            self.progress_percent.emit(100)
            self.finished.emit()
    
    def on_login_confirmation(self, confirmed):
        """处理用户登录确认结果"""
        self.login_result = confirmed
        self.waiting_for_login = False
    
    def stop(self):
        """停止查询"""
        self.is_running = False
        if self.waiting_for_login:
            self.on_login_confirmation(False)
        
        # 如果copyright_query已经初始化，设置其paused为True以终止操作
        if self.copyright_query:
            self.copyright_query.paused = True
            # 确保进度停止更新
            if hasattr(self.copyright_query, 'progress_callback'):
                def terminated_callback(message, percent=None):
                    return False  # 总是返回False以终止查询
                self.copyright_query.progress_callback = terminated_callback

class CopyrightQueryInterface(BaseInterface):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        self.worker = None
        self.thread = None

    def init_ui(self):
        self.layout.setAlignment(Qt.AlignTop)

        header_layout = QHBoxLayout()
        self.upload_button = PrimaryPushButton("选择Excel文件并查询著作权人")
        self.upload_button.clicked.connect(self.handle_upload)
        header_layout.addWidget(self.upload_button)
        
        self.cancel_button = PrimaryPushButton("取消查询")
        self.cancel_button.clicked.connect(self.cancel_query)
        self.cancel_button.setVisible(False)
        header_layout.addWidget(self.cancel_button)
        
        header_layout.addStretch()

        explanation_label = QLabel(
            "说明：请选择包含【游戏名称】和【运营单位】的Excel文件，系统将自动查询并填充著作权人信息。\n"
            "查询结果将保存为新的Excel文件。查询过程中需要登录企查查网站，请按提示操作。"
        )
        explanation_label.setWordWrap(True)

        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("著作权人查询日志将显示在这里...")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)

        self.layout.addLayout(header_layout)
        self.layout.addWidget(explanation_label)
        self.layout.addWidget(self.output_text_edit)
        self.layout.addWidget(self.progress_bar)

    def handle_upload(self):
        dlg = QFileDialog(self)
        dlg.setWindowTitle("选择Excel文件")
        dlg.setNameFilter("Excel Files (*.xlsx *.xls)")
        if dlg.exec():
            files = dlg.selectedFiles()
            if files:
                excel_path = files[0]
                self.output_text_edit.clear()
                self.output_text_edit.append(f"已选择文件: {excel_path}")
                
                # 确认文件存在
                if not os.path.exists(excel_path):
                    QMessageBox.warning(self, "文件错误", f"文件不存在: {excel_path}")
                    return
                
                self.start_query(excel_path)

    def start_query(self, excel_path):
        """开始查询著作权信息"""
        # 创建工作线程
        self.thread = QThread()
        self.worker = CopyrightQueryWorker(excel_path)
        self.worker.moveToThread(self.thread)

        # 连接信号
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.on_query_finished)

        self.worker.progress.connect(self.on_progress)
        self.worker.progress_percent.connect(self.on_percent)
        self.worker.error_occurred.connect(self.on_error)
        self.worker.login_required.connect(self.on_login_required)
        self.worker.anti_crawl_detected.connect(self.on_anti_crawl_detected)

        # 更新UI状态
        self.upload_button.setVisible(False)
        self.cancel_button.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # 启动线程
        self.thread.start()

    def on_login_required(self, is_anti_crawl=False):
        """处理登录需求"""
        dialog = LoginConfirmationDialog(self, is_anti_crawl=is_anti_crawl)
        result = dialog.exec()
        
        # 向工作线程传递用户选择
        if self.worker:
            self.worker.on_login_confirmation(result == QDialog.Accepted)
            
    def on_anti_crawl_detected(self):
        """处理反爬机制检测"""
        self.on_login_required(is_anti_crawl=True)

    def on_progress(self, msg):
        """更新进度消息"""
        self.output_text_edit.append(msg)
        # 滚动到底部
        scrollbar = self.output_text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_percent(self, val):
        """更新进度条"""
        self.progress_bar.setValue(val)
        self.progress_bar.setFormat(f"进度: {val}%")

    def on_error(self, error_msg):
        """处理错误"""
        QMessageBox.critical(self, "错误", error_msg)

    def cancel_query(self):
        """取消查询"""
        if self.worker:
            self.output_text_edit.append("正在取消查询...")
            self.worker.stop()
            # 禁用取消按钮，防止多次点击
            self.cancel_button.setEnabled(False)
            self.cancel_button.setText("正在取消...")
            
            # 确保取消按钮在一定时间后仍然可见，以防进程不能正常结束
            QTimer.singleShot(10000, self.ensure_cancel_complete)
    
    def ensure_cancel_complete(self):
        """确保取消操作完成，防止界面卡住"""
        if self.worker:  # 如果worker仍存在，强制结束
            self.on_query_finished()
            self.output_text_edit.append("查询被强制取消")

    def on_query_finished(self):
        """查询完成后的处理"""
        self.progress_bar.setValue(100)
        self.cancel_button.setVisible(False)
        self.cancel_button.setEnabled(True)
        self.cancel_button.setText("取消查询")
        self.upload_button.setVisible(True)
        
        # 确保线程和工作对象被正确清理
        self.thread = None
        self.worker = None 