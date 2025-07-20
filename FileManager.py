import sys
import os
import shutil
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QComboBox, QSpinBox, QRadioButton, QGroupBox, QVBoxLayout, QHBoxLayout,
    QFormLayout, QMessageBox, QTabWidget, QCheckBox, QDateEdit
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont

# 数字转换函数
def number_to_style(num, style):
    if style == "Arabic":
        return str(num)
    elif style == "Chinese":
        return ''.join("零一二三四五六七八九"[int(d)] for d in str(num))
    elif style == "Chinese_Upper":
        return ''.join("零壹贰叁肆伍陆柒捌玖"[int(d)] for d in str(num))
    elif style == "Circle":
        circle_map = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩',
                      '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱', '⑲', '⑳']
        return circle_map[num - 1] if 1 <= num <= len(circle_map) else str(num)
    elif style == "Roman":
        roman_dict = {
            1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI",
            7: "VII", 8: "VIII", 9: "IX", 10: "X"
        }
        return roman_dict.get(num, str(num))
    return str(num)

class BatchFileGenerator(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量文件生成工具 - by 黄方")
        self.setFont(QFont("微软雅黑", 10))
        self.resize(750, 650)
        self.setStyleSheet(self.get_stylesheet())
        self.init_ui()

    def get_stylesheet(self):
        return """
            QWidget {
                background-color: #f8f9fa;
                font-family: 微软雅黑;
                font-size: 10pt;
            }
            QLabel {
                color: #343a40;
                font-weight: bold;
            }
            QLineEdit, QComboBox, QSpinBox, QDateEdit {
                padding: 5px;
                border: 1px solid #ced4da;
                border-radius: 5px;
                background-color: #ffffff;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 15px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QGroupBox {
                border: 1px solid #ced4da;
                border-radius: 8px;
                margin-top: 10px;
                padding: 10px;
            }
            QGroupBox::title {
                subline-control-position: top center;
                color: #495057;
                font-size: 11pt;
                font-weight: bold;
            }
            QTabBar::tab {
                background-color: #e9ecef;
                padding: 10px 20px;
                border-radius: 5px;
                margin-right: 5px;
            }
            QTabBar::tab:selected {
                background-color: #007bff;
                color: white;
            }
        """

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        # 标题
        title_label = QLabel("批量文件生成工具")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold; color: #343a40; margin-bottom: 10px;")
        main_layout.addWidget(title_label)

        # 标签页
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)

        # 复制文件页面
        copy_tab = QWidget()
        copy_layout = QVBoxLayout(copy_tab)
        self.init_copy_page(copy_layout)
        tab_widget.addTab(copy_tab, "复制文件")

        # 新建文件页面
        create_tab = QWidget()
        create_layout = QVBoxLayout(create_tab)
        self.init_create_page(create_layout)
        tab_widget.addTab(create_tab, "新建文件")

    def init_copy_page(self, layout):
        # 文件路径设置
        path_group = QGroupBox("文件路径设置")
        path_layout = QFormLayout()
        self.source_path = QLineEdit()
        self.source_path.setPlaceholderText("请选择源文件（任意格式）")
        self.btn_select_source = QPushButton("浏览")
        self.btn_select_source.clicked.connect(self.select_source)

        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("请选择输出目录")
        self.btn_select_output = QPushButton("浏览")
        self.btn_select_output.clicked.connect(self.select_output)

        path_layout.addRow("源文件:", self.btn_select_source)
        path_layout.addRow(self.source_path)
        path_layout.addRow("输出目录:", self.btn_select_output)
        path_layout.addRow(self.output_path)
        path_group.setLayout(path_layout)
        layout.addWidget(path_group)

        self.init_common_elements(layout)

    def init_create_page(self, layout):
        # 文件类型设置（仅新建）
        type_group = QGroupBox("文件类型")
        type_layout = QHBoxLayout()
        self.file_type = QComboBox()
        self.file_type.addItems([".docx", ".pptx"])
        type_layout.addWidget(self.file_type)
        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        self.init_common_elements(layout)

    def init_common_elements(self, layout):
        # 文件名规则
        name_group = QGroupBox("文件名规则")
        name_layout = QFormLayout()
        self.filename_template = QLineEdit("文档_{序号}_{数据}")
        self.filename_template.setStyleSheet("padding: 5px;")
        name_layout.addRow("文件名模板:", self.filename_template)
        name_group.setLayout(name_layout)
        layout.addWidget(name_group)

        # 序号设置
        index_group = QGroupBox("序号设置")
        index_layout = QFormLayout()
        self.start_index = QSpinBox()
        self.start_index.setRange(1, 999)
        self.count = QSpinBox()
        self.count.setRange(1, 1000)
        self.number_style = QComboBox()
        self.number_style.addItems(["阿拉伯数字", "汉字小写", "汉字大写", "带圈数字", "罗马数字"])

        self.skip_numbers = QLineEdit()
        self.skip_numbers.setPlaceholderText("1,3,5 或 1-5")
        self.skip_multiples = QSpinBox()
        self.skip_multiples.setRange(2, 100)

        index_layout.addRow("起始序号:", self.start_index)
        index_layout.addRow("生成数量:", self.count)
        index_layout.addRow("序号样式:", self.number_style)
        index_layout.addRow("跳过数字:", self.skip_numbers)
        index_layout.addRow("跳过倍数:", self.skip_multiples)
        index_group.setLayout(index_layout)
        layout.addWidget(index_group)

        # 数据设置
        data_group = QGroupBox("数据设置")
        data_layout = QFormLayout()
        self.data_source_manual = QRadioButton("手动输入数据")
        self.data_source_excel = QRadioButton("从Excel导入")
        self.data_source_manual.setChecked(True)
        self.manual_data = QLineEdit()
        self.manual_data.setPlaceholderText("请输入数据")

        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("请选择Excel文件")
        self.btn_select_excel = QPushButton("浏览")
        self.btn_select_excel.clicked.connect(self.select_excel)

        self.excel_col = QSpinBox()
        self.excel_col.setRange(1, 100)
        self.excel_col.setValue(1)

        data_layout.addRow(self.data_source_manual)
        data_layout.addRow(self.manual_data)
        data_layout.addRow(self.data_source_excel)
        data_layout.addRow("Excel 文件:", self.btn_select_excel)
        data_layout.addRow(self.excel_path)
        data_layout.addRow("列号（从1开始）:", self.excel_col)
        data_group.setLayout(data_layout)
        layout.addWidget(data_group)

        # 日期设置
        date_group = QGroupBox("日期设置")
        date_layout = QFormLayout()
        self.use_date = QCheckBox("在文件名中使用日期")
        self.date_picker = QDateEdit(QDate.currentDate())
        self.date_picker.setCalendarPopup(True)
        self.date_format = QComboBox()
        self.date_format.addItems(["yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyyMMdd"])

        date_layout.addRow(self.use_date)
        date_layout.addRow("选择日期:", self.date_picker)
        date_layout.addRow("日期格式:", self.date_format)
        date_group.setLayout(date_layout)
        layout.addWidget(date_group)

        # 生成按钮
        btn_layout = QHBoxLayout()
        self.btn_generate = QPushButton("生成文件")
        self.btn_generate.clicked.connect(self.generate_files)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_generate)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def select_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择源文件", "", "所有文件 (*.*)")
        if path:
            self.source_path.setText(path)

    def select_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_path.setText(path)

    def select_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel 文件 (*.xlsx *.xls)")
        if path:
            self.excel_path.setText(path)

    def parse_skip_numbers(self, text):
        if not text:
            return []
        numbers = set()
        parts = text.split(",")
        for part in parts:
            if "-" in part:
                start, end = map(int, part.split("-"))
                numbers.update(range(start, end + 1))
            else:
                numbers.add(int(part))
        return list(numbers)

    def generate_files(self):
        # 获取当前标签页
        tab_index = self.parent().parent().currentIndex()
        mode = "copy" if tab_index == 0 else "create"

        source_path = self.source_path.text() if mode == "copy" else ""
        output_path = self.output_path.text()
        file_type = self.file_type.currentText() if mode == "create" else ""

        filename_template = self.filename_template.text()
        start = self.start_index.value()
        count = self.count.value()

        data_source = "excel" if self.data_source_excel.isChecked() else "manual"
        manual_data = self.manual_data.text()
        excel_path = self.excel_path.text()
        excel_col = self.excel_col.value() - 1  # Excel列从0开始

        use_date = self.use_date.isChecked()
        selected_date = self.date_picker.date().toString(self.date_format.currentText())

        skip_numbers = self.parse_skip_numbers(self.skip_numbers.text())
        skip_multiples = self.skip_multiples.value()

        # 验证输入
        if not output_path:
            QMessageBox.critical(self, "错误", "请选择输出目录")
            return

        if mode == "copy" and not source_path:
            QMessageBox.critical(self, "错误", "请选择源文件")
            return

        os.makedirs(output_path, exist_ok=True)

        # 获取数据
        data_list = []
        if data_source == "excel":
            if not excel_path:
                QMessageBox.critical(self, "错误", "请选择Excel文件")
                return
            try:
                df = pd.read_excel(excel_path, header=None)
                data_list = df.iloc[:, excel_col].dropna().tolist()
                if len(data_list) < count:
                    QMessageBox.critical(self, "错误", "Excel 中的数据数量不足")
                    return
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel失败: {e}")
                return
        else:
            if not manual_data:
                QMessageBox.critical(self, "错误", "请输入数据或从Excel导入")
                return
            data_list = [manual_data] * count

        # 映射中文样式为英文标识符
        style_map = {
            "阿拉伯数字": "Arabic",
            "汉字小写": "Chinese",
            "汉字大写": "Chinese_Upper",
            "带圈数字": "Circle",
            "罗马数字": "Roman"
        }
        selected_style = style_map[self.number_style.currentText()]

        # 生成文件
        generated_count = 0
        i = start
        while generated_count < count:
            # 检查是否跳过当前序号
            if i in skip_numbers or (skip_multiples and i % skip_multiples == 0):
                i += 1
                continue

            index_str = number_to_style(i, selected_style)
            current_data = data_list[generated_count]
            filename = filename_template.replace("{序号}", index_str).replace("{数据}", str(current_data))

            if use_date:
                filename = filename.replace("{日期}", selected_date)

            full_path = os.path.join(output_path, filename)

            if mode == "copy":
                if not os.path.exists(source_path):
                    QMessageBox.critical(self, "错误", "源文件不存在")
                    return
                _, ext = os.path.splitext(source_path)
                full_path += ext
                shutil.copy2(source_path, full_path)
            elif mode == "create":
                ext = file_type
                full_path += ext
                if ext == ".docx":
                    doc = Document()
                    doc.add_heading(current_data, 0)
                    doc.save(full_path)
                elif ext == ".pptx":
                    prs = Presentation()
                    slide = prs.slides.add_slide(prs.slide_layouts[0])
                    slide.shapes.title.text = current_data
                    prs.save(full_path)

            generated_count += 1
            i += 1

        QMessageBox.information(self, "完成", f"已生成 {count} 个文件到 {output_path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BatchFileGenerator()
    window.show()
    sys.exit(app.exec_())
