import sys
import os
import shutil
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QComboBox, QSpinBox, QRadioButton, QGroupBox, QVBoxLayout, QHBoxLayout,
    QFormLayout, QMessageBox, QTabWidget, QCheckBox, QDateEdit, QSizePolicy
)
from PyQt5.QtCore import Qt, QDate

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
        self.resize(1200, 700)
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
                padding: 6px 8px;
                border: 1px solid #ced4da;
                border-radius: 5px;
                background-color: #ffffff;
                min-height: 32px;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 15px;
                min-width: 120px;
                min-height: 32px;
                font-size: 10pt;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QGroupBox {
                border: 1px solid #ced4da;
                border-radius: 8px;
                padding: 20px 10px 10px 10px;
                margin-top: 20px;
            }
            QGroupBox::title {
                subline-control-position: top left;
                padding: 0px 10px 0px 10px;
                margin-top: -10px;
                background-color: #f8f9fa;
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
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 标题
        title_label = QLabel("批量文件生成工具")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 20pt; font-weight: bold; color: #343a40;")
        main_layout.addWidget(title_label)

        # 标签页
        self.tab_widget = QTabWidget()
        self.init_tabs(self.tab_widget)
        main_layout.addWidget(self.tab_widget)

        # 主要功能区域
        top_row = QHBoxLayout()
        top_row.setSpacing(20)

        # 文件名规则 + 序号设置
        filename_group = self.create_filename_group()
        index_group = self.create_index_group()
        top_row.addWidget(filename_group, stretch=1)
        top_row.addWidget(index_group, stretch=1)

        # 数据设置 + 日期设置
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(20)

        self.data_group = self.create_data_group()
        self.date_group = self.create_date_group()
        bottom_row.addWidget(self.data_group, stretch=1)
        bottom_row.addWidget(self.date_group, stretch=1)

        # 生成按钮
        btn_layout = QHBoxLayout()
        self.btn_generate = QPushButton("生成文件")
        self.btn_generate.setMinimumHeight(40)
        self.btn_generate.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_generate.clicked.connect(self.generate_files)
        btn_layout.addWidget(self.btn_generate)

        # 主布局
        main_layout.addLayout(top_row)
        main_layout.addLayout(bottom_row)
        main_layout.addLayout(btn_layout)

    def init_tabs(self, tab_widget):
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
        # 文件路径设置（复制文件页显示）
        path_group = QGroupBox("文件路径设置")
        path_layout = QFormLayout()
        path_layout.setSpacing(10)
        path_layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.copy_source_path = QLineEdit()
        self.copy_source_path.setPlaceholderText("请选择源文件（任意格式）")
        self.copy_source_path.setMinimumHeight(32)
        self.copy_source_path.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_copy_select_source = QPushButton("浏览源文件")
        self.btn_copy_select_source.setMinimumWidth(120)
        self.btn_copy_select_source.setMinimumHeight(32)
        self.btn_copy_select_source.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_copy_select_source.clicked.connect(self.select_copy_source)

        self.copy_output_path = QLineEdit()
        self.copy_output_path.setPlaceholderText("请选择输出目录")
        self.copy_output_path.setMinimumHeight(32)
        self.copy_output_path.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_copy_select_output = QPushButton("浏览输出目录")
        self.btn_copy_select_output.setMinimumWidth(120)
        self.btn_copy_select_output.setMinimumHeight(32)
        self.btn_copy_select_output.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_copy_select_output.clicked.connect(self.select_copy_output)

        source_layout = QHBoxLayout()
        source_layout.addWidget(self.btn_copy_select_source)
        source_layout.addWidget(self.copy_source_path, stretch=1)

        output_layout = QHBoxLayout()
        output_layout.addWidget(self.btn_copy_select_output)
        output_layout.addWidget(self.copy_output_path, stretch=1)

        path_layout.addRow("源文件:", source_layout)
        path_layout.addRow("输出目录:", output_layout)
        path_group.setLayout(path_layout)
        layout.addWidget(path_group)

    def init_create_page(self, layout):
        # 文件类型和输出目录设置（新建文件页显示）
        type_output_group = QGroupBox("文件类型和输出目录")
        type_output_layout = QFormLayout()
        type_output_layout.setSpacing(10)
        type_output_layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.create_file_type = QComboBox()
        self.create_file_type.addItems([".docx", ".pptx"])
        self.create_file_type.setMinimumHeight(32)
        self.create_file_type.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.create_output_path = QLineEdit()
        self.create_output_path.setPlaceholderText("请选择输出目录")
        self.create_output_path.setMinimumHeight(32)
        self.create_output_path.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_create_select_output = QPushButton("浏览输出目录")
        self.btn_create_select_output.setMinimumWidth(120)
        self.btn_create_select_output.setMinimumHeight(32)
        self.btn_create_select_output.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_create_select_output.clicked.connect(self.select_create_output)

        output_layout = QHBoxLayout()
        output_layout.addWidget(self.btn_create_select_output)
        output_layout.addWidget(self.create_output_path, stretch=1)

        type_output_layout.addRow("文件类型:", self.create_file_type)
        type_output_layout.addRow("输出目录:", output_layout)
        type_output_group.setLayout(type_output_layout)
        layout.addWidget(type_output_group)

    def create_filename_group(self):
        group = QGroupBox("文件名规则")
        layout = QFormLayout()
        layout.setSpacing(10)
        layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.filename_template = QLineEdit("文档_{序号}_{数据}")
        self.filename_template.setMinimumHeight(32)
        self.filename_template.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        filename_layout = QHBoxLayout()
        filename_layout.addWidget(self.filename_template, stretch=1)

        layout.addRow("文件名模板:", filename_layout)
        group.setLayout(layout)
        return group

    def create_index_group(self):
        group = QGroupBox("序号设置")
        layout = QFormLayout()
        layout.setSpacing(10)
        layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.start_index = QSpinBox()
        self.start_index.setRange(1, 999)
        self.start_index.setMinimumHeight(32)
        self.start_index.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.count = QSpinBox()
        self.count.setRange(1, 1000)
        self.count.setMinimumHeight(32)
        self.count.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.number_style = QComboBox()
        self.number_style.addItems(["阿拉伯数字", "汉字小写", "汉字大写", "带圈数字", "罗马数字"])
        self.number_style.setMinimumHeight(32)
        self.number_style.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.skip_numbers = QLineEdit()
        self.skip_numbers.setPlaceholderText("1,3,5 或 1-5")
        self.skip_numbers.setMinimumHeight(32)
        self.skip_numbers.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.skip_multiples = QSpinBox()
        self.skip_multiples.setRange(2, 100)
        self.skip_multiples.setMinimumHeight(32)
        self.skip_multiples.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        start_index_layout = QHBoxLayout()
        start_index_layout.addWidget(self.start_index, stretch=1)

        count_layout = QHBoxLayout()
        count_layout.addWidget(self.count, stretch=1)

        number_style_layout = QHBoxLayout()
        number_style_layout.addWidget(self.number_style, stretch=1)

        skip_numbers_layout = QHBoxLayout()
        skip_numbers_layout.addWidget(self.skip_numbers, stretch=1)

        skip_multiples_layout = QHBoxLayout()
        skip_multiples_layout.addWidget(self.skip_multiples, stretch=1)

        layout.addRow("起始序号:", start_index_layout)
        layout.addRow("生成数量:", count_layout)
        layout.addRow("序号样式:", number_style_layout)
        layout.addRow("跳过数字:", skip_numbers_layout)
        layout.addRow("跳过倍数:", skip_multiples_layout)
        group.setLayout(layout)
        return group

    def create_data_group(self):
        group = QGroupBox("数据设置")
        layout = QFormLayout()
        layout.setSpacing(10)
        layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.data_source_manual = QRadioButton("手动输入数据")
        self.data_source_excel = QRadioButton("从Excel导入")
        self.data_source_manual.setChecked(True)

        self.manual_data = QLineEdit()
        self.manual_data.setPlaceholderText("请输入数据")
        self.manual_data.setMinimumHeight(32)
        self.manual_data.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("请选择Excel文件")
        self.excel_path.setMinimumHeight(32)
        self.excel_path.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_select_excel = QPushButton("浏览Excel")
        self.btn_select_excel.setMinimumWidth(120)
        self.btn_select_excel.setMinimumHeight(32)
        self.btn_select_excel.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_select_excel.clicked.connect(self.select_excel)

        self.excel_col = QSpinBox()
        self.excel_col.setRange(1, 100)
        self.excel_col.setValue(1)
        self.excel_col.setMinimumHeight(32)
        self.excel_col.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        excel_path_layout = QHBoxLayout()
        excel_path_layout.addWidget(self.btn_select_excel)
        excel_path_layout.addWidget(self.excel_path, stretch=1)

        layout.addRow(self.data_source_manual)
        layout.addRow(self.manual_data)
        layout.addRow(self.data_source_excel)
        layout.addRow("Excel 文件:", excel_path_layout)
        layout.addRow("列号（从1开始）:", self.excel_col)
        group.setLayout(layout)
        return group

    def create_date_group(self):
        group = QGroupBox("日期设置")
        layout = QFormLayout()
        layout.setSpacing(10)
        layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.use_date = QCheckBox("在文件名中使用日期")
        self.use_date.setMinimumHeight(32)

        self.date_picker = QDateEdit(QDate.currentDate())
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setMinimumHeight(32)
        self.date_picker.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.date_format = QComboBox()
        self.date_format.addItems(["yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyyMMdd"])
        self.date_format.setMinimumHeight(32)
        self.date_format.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        date_picker_layout = QHBoxLayout()
        date_picker_layout.addWidget(self.date_picker, stretch=1)

        date_format_layout = QHBoxLayout()
        date_format_layout.addWidget(self.date_format, stretch=1)

        layout.addRow(self.use_date)
        layout.addRow("选择日期:", date_picker_layout)
        layout.addRow("日期格式:", date_format_layout)
        group.setLayout(layout)
        return group

    def select_copy_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择源文件", "", "所有文件 (*.*)")
        if path:
            self.copy_source_path.setText(path)

    def select_copy_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.copy_output_path.setText(path)

    def select_create_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.create_output_path.setText(path)

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
        tab_index = self.tab_widget.currentIndex()
        mode = "copy" if tab_index == 0 else "create"

        # 获取路径和文件类型
        if mode == "copy":
            source_path = self.copy_source_path.text()
            output_path = self.copy_output_path.text()
            file_type = ""
        else:
            source_path = ""
            output_path = self.create_output_path.text()
            file_type = self.create_file_type.currentText()

        filename_template = self.filename_template.text()
        start = self.start_index.value()
        count = self.count.value()

        # 数据源设置
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
                full_path += file_type
                if file_type == ".docx":
                    doc = Document()
                    doc.add_heading(current_data, 0)
                    doc.save(full_path)
                elif file_type == ".pptx":
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
