import os
import sys
import subprocess  # 新增导入subprocess模块
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QProgressBar,
    QPushButton, QAbstractItemView,
    QHeaderView,  # 添加QHeaderView用于列宽设置
    QMenu,  # 新增，用于右键菜单
    QMessageBox  # 新增，用于帮助对话框
)
from PySide6.QtCore import Qt, QThread, Signal
# 导入剪贴板模块
from PySide6.QtGui import QClipboard


class Worker(QThread):
    """工作线程类，处理文件转换"""
    progress_updated = Signal(int)  # 进度更新信号
    conversion_done = Signal(int, str)  # 转换完成信号(行索引, 输出路径)

    def __init__(self, file_paths, parent=None):
        super().__init__(parent)
        self.file_paths = file_paths  # 待转换文件路径列表

    def run(self):
        if not self.file_paths: return

        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        vgmstream_cli = os.path.join(os.getcwd(), "vgmstream", "vgmstream-cli.exe")
        assert os.path.exists(vgmstream_cli), "vgmstream-cli.exe not found"

        for idx, src_path in enumerate(self.file_paths):
            try:
                # ================= 实际转换逻辑 =================
                filename = os.path.basename(src_path)
                output_path = os.path.join(output_dir, filename.replace('.wem', '.wav'))
                cmd = [vgmstream_cli, "-o", output_path, src_path]
                result = subprocess.run(cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode != 0:
                    print(f"Error:{result.stderr}")
                    self.conversion_done.emit(idx, f"Error:{result.stderr}")
                    continue
                # ==============================================

                self.conversion_done.emit(idx, output_path)
            except Exception as e:
                print(f"Error:{str(e)}")
                self.conversion_done.emit(idx, f"Error:{str(e)}")

            # 更新进度 (完成百分比)
            self.progress_updated.emit(int((idx + 1) / len(self.file_paths) * 100))


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WEM转换工具")
        self.setAcceptDrops(True)  # 启用拖拽功能
        self.resize(800, 600)
        # 新增：初始化文件列表
        self.files = []  # 存储拖拽的文件路径列表

        # 主布局
        main_layout = QVBoxLayout()

        # 工具栏
        toolbar = QHBoxLayout()
        self.convert_btn = QPushButton("转换")
        self.play_btn = QPushButton("播放")
        self.copy_btn = QPushButton("复制")
        self.clear_btn = QPushButton("清空")
        # 新增帮助按钮
        self.help_btn = QPushButton("帮助")
        
        toolbar.addWidget(self.convert_btn)
        toolbar.addWidget(self.play_btn)
        toolbar.addWidget(self.copy_btn)
        toolbar.addWidget(self.clear_btn)
        toolbar.addWidget(self.help_btn)  # 添加帮助按钮
        main_layout.addLayout(toolbar)

        # 文件表格
        self.table = QTableWidget(10, 2)  # 10行2列
        self.table.setHorizontalHeaderLabels(["源文件", "输出文件"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)  # 整行选择

        # 设置表格为只读模式
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # 设置列宽：第一列固定宽度，第二列自动填充
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # 第一列固定宽度
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 第二列自动拉伸
        self.table.setColumnWidth(0, 300)  # 设置第一列宽度为300像素

        # 设置上下文菜单策略
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        main_layout.addWidget(self.table)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        main_layout.addWidget(self.progress_bar)

        # 连接信号
        self.convert_btn.clicked.connect(self.start_conversion)
        self.play_btn.clicked.connect(self.play_selected)
        # 连接复制按钮和清空按钮的点击事件
        self.copy_btn.clicked.connect(self.copy_selected_output_paths)
        self.clear_btn.clicked.connect(self.clear_table)
        # 连接帮助按钮的点击事件
        self.help_btn.clicked.connect(self.show_help)
        
        self.setLayout(main_layout)

    def dragEnterEvent(self, event):
        """接受拖拽事件"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if all(url.toLocalFile().endswith('.wem') for url in urls):
                event.acceptProposedAction()

    def dropEvent(self, event):
        """处理文件拖放"""
        # 新增：保存文件路径到self.files，并统一替换路径分隔符为反斜杠
        self.files = sorted(
            [url.toLocalFile().replace('/', '\\') for url in event.mimeData().urls()],
            key=lambda x: os.path.basename(x)
        )

        # 清空表格并设置行数等于文件数量
        self.table.setRowCount(len(self.files))

        # 填充所有文件（只显示文件名，悬停显示完整路径）
        for row in range(len(self.files)):
            # 源文件列
            src_item = QTableWidgetItem(os.path.basename(self.files[row]))
            src_item.setToolTip(self.files[row])  # 设置悬停提示为完整路径
            self.table.setItem(row, 0, src_item)

            # 输出文件列（初始为空）
            out_item = QTableWidgetItem("")
            self.table.setItem(row, 1, out_item)

    def start_conversion(self):
        """开始转换文件"""
        # 修改：直接使用self.files作为转换输入
        if not self.files:
            return

        # 重置进度条
        self.progress_bar.setValue(0)
        self.convert_btn.setEnabled(False)

        # 创建工作线程
        self.worker = Worker(self.files)  # 直接传递原始文件路径列表
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.conversion_done.connect(self.update_table_row)
        self.worker.finished.connect(lambda: self.convert_btn.setEnabled(True))
        self.worker.start()

    def update_table_row(self, row_idx, output_path):
        """更新表格行转换结果"""
        # 只显示文件名，悬停显示完整路径
        item = QTableWidgetItem(os.path.basename(output_path))
        item.setToolTip(output_path)  # 设置悬停提示为完整路径
        self.table.setItem(row_idx, 1, item)

    def play_selected(self):
        """播放选中的输出文件"""
        selected_rows = set(index.row() for index in self.table.selectedIndexes())
        for row in selected_rows:
            output_item = self.table.item(row, 1)
            if output_item and not output_item.text().startswith("Error:"):
                output_path = output_item.toolTip()
                # ================= 用户自定义播放逻辑 =================
                os.startfile(output_path)
                # ====================================================

    def copy_selected_output_paths(self):
        """复制选中行的输出路径到剪贴板"""
        selected_rows = set(index.row() for index in self.table.selectedIndexes())
        output_paths = []
        for row in selected_rows:
            output_item = self.table.item(row, 1)  # 第二列是输出文件
            if output_item and output_item.text():
                output_paths.append(output_item.text())

        if output_paths:
            clipboard = QApplication.clipboard()
            clipboard.setText("\n".join(output_paths))

    def clear_table(self):
        """清空表格"""
        self.table.setRowCount(0)  # 设置行数为0，即清空所有行
        # 新增：同时清空文件列表
        self.files = []  # 清空存储的文件路径

    def show_context_menu(self, pos):
        """显示右键菜单"""
        # 获取点击的行
        row = self.table.indexAt(pos).row()
        if row < 0:  # 点击在表格外
            return

        # 获取输出文件项
        output_item = self.table.item(row, 1)
        if not output_item or not output_item.text():
            return

        # 创建菜单
        menu = QMenu(self)
        copy_action = menu.addAction("复制路径")

        # 显示菜单并等待用户选择
        action = menu.exec(self.table.viewport().mapToGlobal(pos))

        # 处理复制操作
        if action == copy_action:
            self.copy_output_path(row)

    def copy_output_path(self, row):
        """复制指定行的输出路径"""
        output_item = self.table.item(row, 1)
        if output_item and output_item.text():
            clipboard = QApplication.clipboard()
            clipboard.setText(output_item.toolTip())  # 使用toolTip中的完整路径
    
    def show_help(self):
        """显示使用帮助对话框"""
        help_text = """
        <b>WEM转换工具使用说明</b><br><br>
        
        <b>基本功能：</b><br>
        1. 拖放：将.wem文件拖放到窗口区域<br>
        2. 转换：点击"转换"按钮生成.wav文件<br>
        3. 播放：选中文件后点击"播放"按钮测试输出<br>
        4. 复制：复制选中文件的输出路径<br>
        5. 清空：清空当前文件列表<br><br>
        
        <b>高级功能：</b><br>
        - 右键菜单：在文件行上右键点击，可复制完整输出路径<br>
        - 文件信息：鼠标悬停在文件名上可查看完整路径<br><br>
        
        <b>输出位置：</b><br>
        转换后的文件保存在程序目录下的"output"文件夹中<br><br>
        
        <b>注意事项：</b><br>
        - 需要vgmstream-cli.exe在vgmstream目录下<br>
        - 仅支持.wem格式文件输入<br>
        """
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("使用帮助")
        msg_box.setTextFormat(Qt.TextFormat.RichText)
        msg_box.setText(help_text)
        msg_box.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
