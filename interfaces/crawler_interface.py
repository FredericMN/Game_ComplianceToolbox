# interfaces/crawler_interface.py

from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QTextEdit, QLabel, QFrame,
    QDateEdit, QCheckBox, QProgressBar
)
from PySide6.QtCore import Qt, QDate, QObject, QThread, Signal
from qfluentwidgets import PrimaryPushButton, ToolButton, FluentIcon as FIF
from .base_interface import BaseInterface
from utils.crawler import crawl_new_games


class CrawlerWorker(QObject):
    finished = Signal()
    progress = Signal(str)
    progress_percent = Signal(int,int)  # (value, stage)

    def __init__(self, start_date, end_date, enable_version_match=True):
        super().__init__()
        self.start_date = start_date
        self.end_date = end_date
        self.enable_version_match = enable_version_match

    def run(self):
        # 包装回调
        def pcallback(msg):
            self.progress.emit(msg)
        def ppercent(value, stage):
            # stage=0 => 爬虫; stage=1 => 匹配
            self.progress_percent.emit(value, stage)

        try:
            crawl_new_games(
                start_date_str=self.start_date,
                end_date_str=self.end_date,
                progress_callback=pcallback,
                enable_version_match=self.enable_version_match,
                progress_percent_callback=ppercent
            )
        except Exception as e:
            # 其他未捕获异常
            self.progress.emit(f"爬虫出现异常: {e}")
        finally:
            self.finished.emit()

class CrawlerInterface(BaseInterface):
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

        self.match_checkbox = QCheckBox("自动匹配版号")
        self.match_checkbox.setChecked(True)

        header_layout.addWidget(self.start_button)
        header_layout.addWidget(self.expand_button)
        header_layout.addWidget(self.match_checkbox)
        header_layout.addStretch()

        self.desc_label = QLabel(
            "说明：默认爬取从今天到未来4天共5天的数据。\n勾选自动匹配版号则爬取完成后会重置进度条并进行匹配。"
        )
        header_layout.addWidget(self.desc_label)

        # config
        self.config_frame = QFrame()
        self.config_frame.setFrameShape(QFrame.StyledPanel)
        self.config_frame.setVisible(False)
        self.config_layout = QVBoxLayout(self.config_frame)

        sd_label = QLabel("起始日期（yyyy-MM-dd）：")
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate())

        ed_label = QLabel("结束日期（yyyy-MM-dd）：")
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate().addDays(4))

        self.config_layout.addWidget(sd_label)
        self.config_layout.addWidget(self.start_date_edit)
        self.config_layout.addWidget(ed_label)
        self.config_layout.addWidget(self.end_date_edit)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0,100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)

        self.layout.addLayout(header_layout)
        self.layout.addWidget(self.config_frame)
        self.layout.addWidget(self.output_text)
        self.layout.addWidget(self.progress_bar)

    def toggle_expand(self):
        self.config_frame.setVisible(not self.config_frame.isVisible())

    def handle_start(self):
        sdate = self.start_date_edit.date().toString("yyyy-MM-dd")
        edate = self.end_date_edit.date().toString("yyyy-MM-dd")
        enable_match = self.match_checkbox.isChecked()

        self.thread = QThread()
        self.worker = CrawlerWorker(sdate, edate, enable_match)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.worker.progress.connect(self.on_progress)
        self.worker.progress_percent.connect(self.on_percent)

        self.start_button.setEnabled(False)
        self.match_checkbox.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("进度: 0%")

        self.thread.start()
        self.thread.finished.connect(lambda: self.start_button.setEnabled(True))
        self.thread.finished.connect(lambda: self.match_checkbox.setEnabled(True))
        self.thread.finished.connect(self.on_finish)

    def on_finish(self):
        self.output_text.append("\n[提示] 整体任务结束。")

    def on_progress(self, msg):
        self.output_text.append(msg)

    def on_percent(self, value, stage):
        """
        stage=0 => 爬虫阶段
        stage=1 => 版号匹配阶段
        """
        if stage==0:
            # 爬虫进度
            self.progress_bar.setValue(value)
            self.progress_bar.setFormat(f"爬虫进度: {value}%")
        else:
            # 版号匹配阶段
            if value==0:
                # 重置
                self.progress_bar.setValue(0)
                self.progress_bar.setFormat("版号匹配进度: 0%")
                self.output_text.append("\n[提示] 进入版号匹配阶段，已重置进度为 0。\n")
            else:
                self.progress_bar.setValue(value)
                self.progress_bar.setFormat(f"版号匹配进度: {value}%")
