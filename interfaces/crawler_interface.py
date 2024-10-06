from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QTextEdit, QLabel, QFrame, QDateEdit
)
from PySide6.QtCore import Qt, QDate, QObject, QThread, Signal
from qfluentwidgets import PrimaryPushButton, ToolButton, FluentIcon as FIF
from .base_interface import BaseInterface
from utils.crawler import crawl_new_games
import traceback

class CrawlWorker(QObject):
    finished = Signal()
    progress = Signal(str)

    def __init__(self, start_date, end_date):
        super().__init__()
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        try:
            def progress_callback(message):
                self.progress.emit(message)
            crawl_new_games(self.start_date, self.end_date, progress_callback)
        except Exception as e:
            # 捕获所有异常并输出
            self.progress.emit(f"发生错误: {str(e)} 疑似网络存在问题，请重试。")
        finally:
            self.finished.emit()

class CrawlerInterface(BaseInterface):
    """新游爬虫界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.layout.setAlignment(Qt.AlignTop)
        header_layout = QHBoxLayout()

        self.start_button = PrimaryPushButton("开始爬取")
        self.start_button.clicked.connect(self.handle_start)

        self.expand_button = ToolButton(FIF.CHEVRON_DOWN_MED)
        self.expand_button.clicked.connect(self.toggle_expand)

        # 添加开始按钮和下拉按钮到header_layout
        header_layout.addWidget(self.start_button)
        header_layout.addWidget(self.expand_button)

        # 创建并添加说明标签
        description_text = (
            "说明：点击下拉可配置爬取时间段，针对taptap不建议爬取时间跨度过长的数据，五天内为佳（默认五天）。"
        )
        self.description_label = QLabel(description_text)
        header_layout.addWidget(self.description_label)

        # 添加伸缩项，使说明标签靠左
        header_layout.addStretch()

        self.config_widget = QFrame()
        self.config_widget.setFrameShape(QFrame.StyledPanel)
        self.config_widget.setVisible(False)
        self.config_layout = QVBoxLayout(self.config_widget)

        # 添加日期选择控件
        start_date_label = QLabel("起始日期（yyyy-MM-dd）：")
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate())

        end_date_label = QLabel("结束日期（yyyy-MM-dd）：")
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate().addDays(4))

        self.config_layout.addWidget(start_date_label)
        self.config_layout.addWidget(self.start_date_edit)
        self.config_layout.addWidget(end_date_label)
        self.config_layout.addWidget(self.end_date_edit)

        # 信息输出区域
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("信息输出区域")

        self.layout.addLayout(header_layout)
        self.layout.addWidget(self.config_widget)
        self.layout.addWidget(self.output_text_edit)

    def toggle_expand(self):
        expanded = self.config_widget.isVisible()
        self.config_widget.setVisible(not expanded)
        if expanded:
            self.expand_button.setIcon(FIF.CHEVRON_DOWN_MED.icon())
        else:
            self.expand_button.setIcon(FIF.CHEVRON_DOWN_MED.icon())

    def handle_start(self):
        start_date = self.start_date_edit.date().toString("yyyy-MM-dd")
        end_date = self.end_date_edit.date().toString("yyyy-MM-dd")

        self.thread = QThread()
        self.worker = CrawlWorker(start_date, end_date)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.report_progress)

        self.thread.start()

        self.start_button.setEnabled(False)
        self.thread.finished.connect(lambda: self.start_button.setEnabled(True))

    def report_progress(self, message):
        self.output_text_edit.append(message)
