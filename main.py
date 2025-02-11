import os
import sys
import json
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QMessageBox,
    QVBoxLayout, QGridLayout, QGroupBox, QWidget, QProgressBar, QStatusBar
)
from PySide6.QtGui import QIcon, QDragEnterEvent, QDropEvent
from PySide6.QtCore import Qt, QSettings
from zipfile import ZipFile, BadZipFile
from os import makedirs, remove, rmdir, walk
from os.path import (join, splitext, exists, basename, dirname,
                     abspath, normpath, commonprefix)


# ======================== 核心功能函数 ========================
def custom_relpath(path, start):
    """
    自定义相对路径计算方法，处理跨平台路径差异
    参数：
        path: 目标路径
        start: 起始路径
    返回：
        标准化的相对路径
    """
    # 规范化路径并转换为绝对路径
    path = abspath(normpath(path))
    start = abspath(normpath(start))
    
    # 如果起始路径不在目标路径中，直接返回原路径
    if start not in path:
        return path
    
    # 计算公共前缀
    common = commonprefix([path, start])
    rel_path = path[len(common):].lstrip(os.sep)
    
    # 处理路径中的特殊符号
    parts = rel_path.split(os.sep)
    rel = []
    for part in parts:
        if part == '..':
            if rel:
                rel.pop()
        elif part != '.':
            rel.append(part)
    return os.sep.join(rel)


def convert_exe_to_sb3(zip_file_name, output_dir, progress_callback=None):
    """
    核心转换函数
    参数：
        zip_file_name: 输入的ZIP文件路径
        output_dir: 输出目录
        progress_callback: 进度回调函数
    返回：
        bool: 转换是否成功
    """
    try:
        # 初始化路径参数
        file_name = basename(zip_file_name)
        target_zip_name = join(output_dir, f"{splitext(file_name)[0]}.sb3")
        temp_dir = join(output_dir, "temp_extract")
        
        # 创建临时目录
        makedirs(temp_dir, exist_ok=True)
        
        # ----------------- 阶段1：解压ZIP文件 -----------------
        _update_progress(progress_callback, 10, "正在解压文件...")
        
        try:
            with ZipFile(zip_file_name, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
        except BadZipFile:
            _update_progress(progress_callback, 0, "错误：无效的ZIP文件格式")
            return False

        # ----------------- 阶段2：验证目录结构 -----------------
        temp_dir = list(temp_dir)
        sum = 0
        for i in temp_dir:
            if i == '/':
                temp_dir[sum] = '\\'
            sum += 1
        temp = temp_dir
        temp_dir = ''
        sum = 0
        for i in temp:
            temp_dir += temp[sum]
            sum += 1
        app_dir = join(temp_dir, "packaged-project", "resources", "app")
        app_dir += '\\'
        if not exists(app_dir):
            _update_progress(progress_callback, 0, f"错误：缺少资源目录 {app_dir}")
            return False

        # ----------------- 阶段3：创建SB3文件 -----------------
        _update_progress(progress_callback, 50, "正在打包SB3文件...")
        
        try:
            with ZipFile(target_zip_name, 'w') as new_zip:
                for root, dirs, files in walk(app_dir):
                    for file in files:
                        src_path = join(root, file)
                        # 计算相对路径并标准化
                        rel_path = join(custom_relpath(root, app_dir), file)
                        new_zip.write(src_path, arcname=rel_path)
        except Exception as e:
            _update_progress(progress_callback, 0, f"写入错误：{str(e)}")
            return False

        # ----------------- 阶段4：清理临时文件 -----------------
        _update_progress(progress_callback, 80, "正在清理临时文件...")
        _cleanup_temp_dir(temp_dir)

        _update_progress(progress_callback, 100, "转换完成！")
        return True

    except Exception as e:
        _update_progress(progress_callback, 0, f"未预期错误：{str(e)}")
        return False


def _update_progress(callback, value, message):
    """安全更新进度"""
    if callback:
        try:
            callback(value, message)
        except RuntimeError:  # 处理界面已关闭的情况
            pass


def _cleanup_temp_dir(temp_dir):
    """安全删除临时目录"""
    try:
        for root, dirs, files in walk(temp_dir, topdown=False):
            for name in files:
                os.remove(join(root, name))
            for name in dirs:
                os.rmdir(join(root, name))
        os.rmdir(temp_dir)
    except Exception as e:
        print(f"清理临时目录失败：{str(e)}")


# ======================== GUI界面类 ========================
class MainWindow(QMainWindow):
    """主窗口类，处理界面交互"""
    
    def __init__(self):
        super().__init__()
        # 窗口基础设置
        self.setWindowTitle("Delicious")
        self.setFixedSize(420, 350)  # 固定窗口尺寸
        self.setWindowIcon(QIcon(self._resource_path("icon.png")))
        
        # 初始化设置存储
        self.settings = QSettings("Bilibili", "EXE2SB3")
        
        # 初始化界面组件
        self._init_ui()
        self._load_history()

    # ----------------- 工具方法 -----------------
    def _resource_path(self, relative_path):
        """解决打包后资源路径问题"""
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(abspath(__file__)))
        return join(base_path, relative_path)

    # ----------------- UI初始化 -----------------
    def _init_ui(self):
        """初始化界面布局"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 启用文件拖放
        self.setAcceptDrops(True)

        # 文件操作区域
        self._setup_file_operations(main_layout)
        
        # 历史记录按钮
        self.history_button = self._create_button(
            "历史记录", "#8e44ad", self._show_history)
        main_layout.addWidget(self.history_button)

        # 进度条
        self._setup_progress_bar(main_layout)

        # 操作按钮组
        self._setup_action_buttons(main_layout)

        # 状态栏
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("color: #7f8c8d;")
        self.setStatusBar(self.status_bar)

    def _setup_file_operations(self, layout):
        """设置文件选择区域"""
        file_group = QGroupBox("文件操作")
        file_group.setStyleSheet("QGroupBox { font-size: 16px; color: #2c3e50; }")
        grid = QGridLayout()

        # 输入文件组件
        self.input_button = self._create_button(
            "选择 ZIP 文件", "#27ae60", self._select_input_file)
        self.input_file_label = self._create_label("未选择文件", "#34495e")
        grid.addWidget(self.input_button, 0, 0)
        grid.addWidget(self.input_file_label, 0, 1)

        # 输出目录组件
        self.output_button = self._create_button(
            "选择输出目录", "#2980b9", self._select_output_dir)
        self.output_dir_label = self._create_label("未选择目录", "#34495e")
        grid.addWidget(self.output_button, 1, 0)
        grid.addWidget(self.output_dir_label, 1, 1)

        file_group.setLayout(grid)
        layout.addWidget(file_group)

    def _setup_progress_bar(self, layout):
        """初始化进度条"""
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                text-align: center;
                background: #ecf0f1;
            }
            QProgressBar::chunk {
                background: #3498db;
                border-radius: 3px;
            }
        """)
        layout.addWidget(self.progress_bar)

    def _setup_action_buttons(self, layout):
        """设置底部操作按钮"""
        button_layout = QGridLayout()
        
        # 转换按钮
        self.convert_button = self._create_button(
            "开始转换", "#e67e22", self._perform_conversion)
        self.convert_button.setFixedHeight(40)
        button_layout.addWidget(self.convert_button, 0, 0)

        # 关于按钮
        self.about_button = self._create_button(
            "关于作者", "#9b59b6", self._show_author_info)
        self.about_button.setFixedHeight(40)
        button_layout.addWidget(self.about_button, 0, 1)

        layout.addLayout(button_layout)

    # ----------------- 组件创建工具 -----------------
    def _create_button(self, text, color, callback):
        """创建风格化按钮"""
        btn = QPushButton(text)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
            }}
            QPushButton:hover {{
                background-color: {color}90;
            }}
        """)
        btn.clicked.connect(callback)
        return btn

    def _create_label(self, text, color):
        """创建风格化标签"""
        label = QLabel(text)
        label.setStyleSheet(f"color: {color}; font-size: 14px;")
        label.setWordWrap(True)
        return label

    # ----------------- 事件处理 -----------------
    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖放进入事件处理"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        """拖放释放事件处理"""
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith(".zip"):
                self.input_file_path = file_path
                self.input_file_label.setText(basename(file_path))
                self._save_history()
                break

    # ----------------- 业务逻辑 -----------------
    def _select_input_file(self):
        """选择输入文件"""
        default_dir = self.settings.value("last_input_dir", "")
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 ZIP 文件", default_dir, "ZIP 文件 (*.zip)")
        if path and self._validate_zip(path):
            self.input_file_path = path
            self.input_file_label.setText(basename(path))
            self.settings.setValue("last_input_dir", dirname(path))
            self._save_history()

    def _select_output_dir(self):
        """选择输出目录"""
        default_dir = self.settings.value("last_output_dir", "")
        path = QFileDialog.getExistingDirectory(
            self, "选择输出目录", default_dir)
        if path:
            self.output_dir_path = path
            self.output_dir_label.setText(path)
            self.settings.setValue("last_output_dir", path)
            self._save_history()

    def _perform_conversion(self):
        """执行转换操作"""
        if not self._validate_inputs():
            return

        self.progress_bar.setValue(0)
        self.status_bar.showMessage("正在转换...")

        def progress_callback(value, message):
            self.progress_bar.setValue(value)
            self.status_bar.showMessage(message)

        success = convert_exe_to_sb3(
            self.input_file_path,
            self.output_dir_path,
            progress_callback
        )

        if success:
            QMessageBox.information(self, "完成", "文件转换成功！")
        else:
            QMessageBox.critical(self, "错误", "转换失败，请检查：\n1. 文件是否为合法Scratch项目\n2. 输出目录写入权限")

    # ----------------- 历史记录处理 -----------------
    def _save_history(self):
        """保存历史记录"""
        history = {
            "input": self.input_file_path,
            "output": self.output_dir_path
        }
        self.settings.setValue("history", json.dumps(history))

    def _load_history(self):
        """加载历史记录"""
        history = json.loads(self.settings.value("history", "{}"))
        if self._validate_path(history.get("input")):
            self.input_file_path = history["input"]
            self.input_file_label.setText(basename(history["input"]))
        if self._validate_path(history.get("output")):
            self.output_dir_path = history["output"]
            self.output_dir_label.setText(history["output"])

    # ----------------- 验证工具 -----------------
    def _validate_inputs(self):
        """验证输入有效性"""
        if not all([self.input_file_path, self.output_dir_path]):
            QMessageBox.warning(self, "警告", "请先选择文件和输出目录！")
            return False
        if not self._validate_zip(self.input_file_path):
            return False
        if not os.access(self.output_dir_path, os.W_OK):
            QMessageBox.critical(self, "错误", "输出目录没有写入权限")
            return False
        return True

    def _validate_zip(self, path):
        """验证ZIP文件有效性"""
        if not path.lower().endswith(".zip"):
            QMessageBox.warning(self, "警告", "请选择ZIP格式文件")
            return False
        if not os.access(path, os.R_OK):
            QMessageBox.critical(self, "错误", "文件不可读")
            return False
        return True

    @staticmethod
    def _validate_path(path):
        """通用路径验证"""
        return path and os.path.exists(path)

    # ----------------- 信息展示 -----------------
    def _show_history(self):
        """显示历史记录"""
        history = json.loads(self.settings.value("history", "{}"))
        msg = (
            "最近使用的文件：\n"
            f"输入文件：{history.get('input', '无')}\n"
            f"输出目录：{history.get('output', '无')}"
        )
        QMessageBox.information(self, "历史记录", msg)

    def _show_author_info(self):
        """显示作者信息"""
        QMessageBox.information(
            self, "关于作者",
            "作者：Easytong\n"
            "B站主页：https://space.bilibili.com/3546576561637431\n\n"
            "✨ 功能特性 ✨\n"
            "- 拖放ZIP文件快速选择\n"
            "- 自动保存历史记录\n"
            "- 实时进度反馈\n"
            "- 智能路径验证\n"
            "基于‘熊孩子728’的开源项目开发"
        )


if __name__ == "__main__":
    # 启动应用程序
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())