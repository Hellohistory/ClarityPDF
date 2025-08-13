import sys
import os
import re
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
                               QLabel, QLineEdit, QProgressBar, QSpinBox, QMessageBox,
                               QDoubleSpinBox, QGridLayout, QComboBox, QGroupBox, QListWidget)
from PySide6.QtGui import QIcon
from clarity_core import process_pdf


class BatchWorker(QThread):
    """
    处理整个文件批次的Worker。
    这个线程负责协调调用 process_pdf 函数来处理每个文件。
    """
    # 信号定义
    current_file_progress = Signal(int, str)  # 当前文件进度: (百分比, 状态信息)
    overall_progress = Signal(int, int, str)  # 总体进度: (当前文件索引, 文件总数, 当前文件名)
    finished = Signal(bool, str)  # 批处理完成: (是否成功, 结束信息)

    def __init__(self, settings_list, output_folder):
        super().__init__()
        self.settings_list = settings_list
        self.output_folder = output_folder
        self.is_cancelled = False

    def run(self):
        """线程主执行函数"""
        total_files = len(self.settings_list)
        for i, settings in enumerate(self.settings_list):
            if self.is_cancelled:
                break

            input_path = settings["input_pdf_path"]
            filename = os.path.basename(input_path)
            base_name, ext = os.path.splitext(filename)
            output_path = os.path.join(self.output_folder, f"{base_name}_clarity.pdf")
            settings["output_pdf_path"] = output_path

            self.overall_progress.emit(i + 1, total_files, filename)

            try:
                process_pdf(**settings, progress_callback=self.current_file_progress)
            except Exception as e:
                import traceback
                error_details = traceback.format_exc()
                print(error_details)
                self.finished.emit(False, f"处理文件 {filename} 时发生严重错误: {e}")
                return

        if self.is_cancelled:
            self.finished.emit(True, "处理已由用户取消。")
        else:
            self.overall_progress.emit(total_files, total_files, "全部完成！")
            self.finished.emit(True, "所有文件处理完成！")

    def cancel(self):
        """外部调用的取消方法"""
        self.is_cancelled = True


class ClarityPDFApp(QWidget):
    PRESET_MODES = {
        "均衡模式 (Balanced)": {"target_dpi": 200, "window_size": 25, "k": 0.1},
        "快速模式 (Fast)": {"target_dpi": 150, "window_size": 51, "k": 0.2},
        "高清模式 (High Quality)": {"target_dpi": 300, "window_size": 15, "k": 0.05},
        "高级设置 (Advanced)": {}
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ClarityPDF - 批量处理与加速工具")
        self.setGeometry(100, 100, 600, 650)
        if os.path.exists("icon.png"):
            self.setWindowIcon(QIcon("icon.png"))

        self.worker = None
        self.init_ui()
        self.update_advanced_controls_from_preset()

    def init_ui(self):
        """初始化所有UI控件"""
        main_layout = QVBoxLayout(self)

        # 文件列表区域
        list_group = QGroupBox("文件列表")
        list_layout = QVBoxLayout(list_group)
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        list_layout.addWidget(self.file_list_widget)

        list_button_layout = QHBoxLayout()
        btn_add = QPushButton(QIcon.fromTheme("list-add"), "添加文件")
        btn_add.clicked.connect(self.add_files)
        btn_remove = QPushButton(QIcon.fromTheme("list-remove"), "移除选中")
        btn_remove.clicked.connect(self.remove_selected_files)
        btn_clear = QPushButton(QIcon.fromTheme("edit-clear"), "清空列表")
        btn_clear.clicked.connect(self.clear_list)
        list_button_layout.addWidget(btn_add)
        list_button_layout.addWidget(btn_remove)
        list_button_layout.addWidget(btn_clear)
        list_layout.addLayout(list_button_layout)
        main_layout.addWidget(list_group)

        # 输出文件夹选择
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件夹:"))
        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText("请选择一个文件夹用于保存处理后的文件")
        self.output_folder_edit.setReadOnly(True)
        btn_select_folder = QPushButton("浏览...")
        btn_select_folder.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_folder_edit)
        output_layout.addWidget(btn_select_folder)
        main_layout.addLayout(output_layout)

        # 参数设置区域
        settings_group = QGroupBox("处理设置")
        settings_layout = QVBoxLayout(settings_group)

        skip_pages_layout = QHBoxLayout()
        skip_pages_layout.addWidget(QLabel("跳过页面:"))
        self.skip_pages_edit = QLineEdit()
        self.skip_pages_edit.setPlaceholderText("例如: 1, 3-5, 8 (对所有文件生效)")
        skip_pages_layout.addWidget(self.skip_pages_edit)
        settings_layout.addLayout(skip_pages_layout)

        settings_layout.addWidget(QLabel("处理模式:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(self.PRESET_MODES.keys())
        self.mode_combo.currentIndexChanged.connect(self.on_mode_changed)
        settings_layout.addWidget(self.mode_combo)

        self.advanced_group = QGroupBox("高级参数设置")
        advanced_layout = QGridLayout(self.advanced_group)

        # 将缩放输入框改为DPI输入框
        self.dpi_input = QSpinBox()
        self.dpi_input.setRange(100, 600)
        self.dpi_input.setSuffix(" DPI")
        self.dpi_input.setSingleStep(50)
        self.dpi_input.setValue(200)

        self.window_size_input = QSpinBox()
        self.window_size_input.setRange(3, 99)
        self.window_size_input.setSingleStep(2)
        self.window_size_input.setValue(25)

        self.k_input = QDoubleSpinBox()
        self.k_input.setRange(0.01, 0.5)
        self.k_input.setSingleStep(0.01)
        self.k_input.setDecimals(3)
        self.k_input.setValue(0.1)

        advanced_layout.addWidget(QLabel("目标DPI:"), 0, 0)
        advanced_layout.addWidget(self.dpi_input, 0, 1)
        advanced_layout.addWidget(QLabel("窗口:"), 1, 0)
        advanced_layout.addWidget(self.window_size_input, 1, 1)
        advanced_layout.addWidget(QLabel("k值:"), 2, 0)
        advanced_layout.addWidget(self.k_input, 2, 1)

        self.advanced_group.setLayout(advanced_layout)
        settings_layout.addWidget(self.advanced_group)
        main_layout.addWidget(settings_group)

        # 进度与控制区域
        progress_group = QGroupBox("处理进度")
        progress_layout = QVBoxLayout(progress_group)
        self.overall_status_label = QLabel("总体进度: 未开始")
        self.overall_progress_bar = QProgressBar()
        self.current_file_status_label = QLabel("当前文件进度: 未开始")
        self.current_file_progress_bar = QProgressBar()
        progress_layout.addWidget(self.overall_status_label)
        progress_layout.addWidget(self.overall_progress_bar)
        progress_layout.addWidget(self.current_file_status_label)
        progress_layout.addWidget(self.current_file_progress_bar)
        main_layout.addWidget(progress_group)

        self.process_button = QPushButton("开始批量处理")
        self.process_button.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 8px;")
        self.process_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.process_button)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择要处理的PDF文件", "", "PDF 文件 (*.pdf)")
        if files:
            for file in files:
                if not self.file_list_widget.findItems(file, Qt.MatchFlag.MatchExactly):
                    self.file_list_widget.addItem(file)
            if not self.output_folder_edit.text():
                self.output_folder_edit.setText(os.path.dirname(files[0]))

    def remove_selected_files(self):
        for item in self.file_list_widget.selectedItems():
            self.file_list_widget.takeItem(self.file_list_widget.row(item))

    def clear_list(self):
        self.file_list_widget.clear()

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder:
            self.output_folder_edit.setText(folder)

    def start_processing(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            return

        file_count = self.file_list_widget.count()
        output_folder = self.output_folder_edit.text()
        if file_count == 0:
            QMessageBox.warning(self, "列表为空", "请先添加要处理的PDF文件。")
            return
        if not output_folder or not os.path.isdir(output_folder):
            QMessageBox.warning(self, "路径无效", "请选择一个有效的输出文件夹。")
            return

        try:
            skip_pages = self.parse_page_ranges(self.skip_pages_edit.text())
        except ValueError as e:
            QMessageBox.critical(self, "格式错误", str(e))
            return

        base_settings = {"skip_pages": skip_pages}
        mode_text = self.mode_combo.currentText()
        if mode_text == "高级设置 (Advanced)":
            window_size = self.window_size_input.value()
            if window_size % 2 == 0:
                window_size += 1
                self.window_size_input.setValue(window_size)
            base_settings.update({
                "target_dpi": self.dpi_input.value(), "window_size": window_size, "k": self.k_input.value()
            })
        else:
            preset = self.PRESET_MODES[mode_text]
            base_settings.update({
                "target_dpi": preset["target_dpi"], "window_size": preset["window_size"], "k": preset["k"]
            })

        settings_list = []
        for i in range(file_count):
            file_path = self.file_list_widget.item(i).text()
            file_settings = base_settings.copy()
            file_settings["input_pdf_path"] = file_path
            settings_list.append(file_settings)

        self.set_ui_enabled(False)
        self.worker = BatchWorker(settings_list, output_folder)
        self.worker.current_file_progress.connect(self.update_current_file_progress)
        self.worker.overall_progress.connect(self.update_overall_progress)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.start()

    def set_ui_enabled(self, enabled):
        is_running = not enabled
        self.file_list_widget.parent().setEnabled(enabled)
        self.output_folder_edit.parent().setEnabled(enabled)
        self.mode_combo.parent().setEnabled(enabled)

        self.process_button.setText("取消处理" if is_running else "开始批量处理")
        if is_running:
            self.process_button.setStyleSheet(
                "background-color: #dc3545; color: white; font-weight: bold; padding: 8px;")
        else:
            self.process_button.setStyleSheet(
                "background-color: #28a745; color: white; font-weight: bold; padding: 8px;")

    def update_current_file_progress(self, value, message):
        self.current_file_progress_bar.setValue(value)
        self.current_file_status_label.setText(f"当前文件: {message}")

    def update_overall_progress(self, value, total, filename):
        self.overall_progress_bar.setValue(value)
        self.overall_progress_bar.setMaximum(total)
        self.overall_status_label.setText(f"总体进度: {value}/{total} - 正在处理 {filename}")
        self.current_file_progress_bar.setValue(0)
        self.current_file_status_label.setText("当前文件: 未开始")

    def on_processing_finished(self, success, message):
        self.set_ui_enabled(True)
        if success:
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "错误", message)
        self.overall_status_label.setText(f"总体进度: {message}")
        self.current_file_status_label.setText("当前文件: 已结束")

    def parse_page_ranges(self, range_string):
        if not range_string: return set()
        pages = set()
        range_string = re.sub(r'\s+', '', range_string)
        if not range_string: return set()
        parts = range_string.split(',')
        for part in parts:
            if not part: continue
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if start > end:
                        raise ValueError(f"范围起始值不能大于结束值: '{part}'")
                    pages.update(range(start, end + 1))
                except ValueError as e:
                    raise ValueError(f"页面范围格式错误: '{part}' - {e}")
            else:
                try:
                    pages.add(int(part))
                except ValueError:
                    raise ValueError(f"页面号码格式错误: '{part}'")
        return pages

    def update_advanced_controls_from_preset(self):
        """根据当前选择的预设模式更新高级设置中的值"""
        mode_text = self.mode_combo.currentText()
        if mode_text != "高级设置 (Advanced)":
            preset = self.PRESET_MODES[mode_text]
            self.dpi_input.setValue(preset["target_dpi"])
            self.window_size_input.setValue(preset["window_size"])
            self.k_input.setValue(preset["k"])

    def on_mode_changed(self):
        """处理模式下拉框变化时的槽函数"""
        is_advanced = (self.mode_combo.currentText() == "高级设置 (Advanced)")
        self.advanced_group.setVisible(is_advanced)
        self.update_advanced_controls_from_preset()

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ClarityPDFApp()
    window.show()
    sys.exit(app.exec())