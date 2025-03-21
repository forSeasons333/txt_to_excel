import sys
import os
import time
import query_ico_rc
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QPushButton, QLabel, QProgressBar, QFileDialog,
                            QTextEdit, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon

class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, str)  # 进度百分比和当前文件名
    finished = pyqtSignal(pd.DataFrame)      # 完成信号
    error_occurred = pyqtSignal(str)         # 错误信号

    def __init__(self, input_folder):
        super().__init__()
        self.input_folder = input_folder
        self._is_running = True

    def run(self):
        try:
            txt_files = []
            for root, _, files in os.walk(self.input_folder):
                for file in files:
                    if file.lower().endswith('.txt'):
                        txt_files.append(os.path.join(root, file))

            total_files = len(txt_files)
            data = []
            
            for idx, file_path in enumerate(txt_files, 1):
                if not self._is_running:
                    break
                
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read().strip()
                    result = {
                        "文件名": os.path.basename(file_path),
                        "相对路径": os.path.relpath(file_path, self.input_folder),
                        "内容摘要": content[:100] + "..." if len(content)>100 else content,
                        "字符数": len(content)
                    }
                    data.append(result)
                    
                    # 发送进度更新
                    progress = int((idx / total_files) * 100)
                    self.progress_updated.emit(progress, os.path.basename(file_path))
                except Exception as e:
                    self.error_occurred.emit(f"文件 {file_path} 读取失败：{str(e)}")
                
                time.sleep(0.01)  # 模拟处理延时

            if data:
                df = pd.DataFrame(data)
                self.finished.emit(df)
            else:
                self.error_occurred.emit("未找到有效文本文件")

        except Exception as e:
            self.error_occurred.emit(f"处理失败：{str(e)}")

    def stop(self):
        self._is_running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(':/icons/tte.png'))
        self.init_ui()
        self.thread = None

    def init_ui(self):
        self.setWindowTitle("txt_to_excel v1.0.1 🚀")
        self.setGeometry(300, 300, 800, 600)

        # 主控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        # 路径选择
        self.btn_select = QPushButton("选择文件夹")
        self.btn_select.setIcon(QIcon(':/icons/folder_icon.png'))
        self.btn_select.clicked.connect(self.select_folder)
        self.lbl_path = QLabel("未选择文件夹")
        
        # 进度显示
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.lbl_status = QLabel("就绪")

        # 日志显示
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)

        # 按钮组
        self.btn_start = QPushButton("开始处理")
        self.btn_start.setIcon(QIcon(':/icons/start_icon.png'))
        self.btn_start.clicked.connect(self.start_processing)
        self.btn_start.setEnabled(False)

        # 布局
        layout.addWidget(self.btn_select)
        layout.addWidget(self.lbl_path)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.lbl_status)
        layout.addWidget(self.txt_log)
        layout.addWidget(self.btn_start)
        central_widget.setLayout(layout)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.lbl_path.setText(folder)
            self.btn_start.setEnabled(True)

    def start_processing(self):
        if self.thread and self.thread.isRunning():
            QMessageBox.warning(self, "警告", "已有任务正在运行")
            return

        input_folder = self.lbl_path.text()
        self.thread = ProcessingThread(input_folder)
        
        # 信号连接
        self.thread.progress_updated.connect(self.update_progress)
        self.thread.finished.connect(self.on_finished)
        self.thread.error_occurred.connect(self.show_error)
        
        # 重置界面
        self.progress_bar.setValue(0)
        self.txt_log.clear()
        self.txt_log.append(f"{time.strftime('%H:%M:%S')} 开始处理...")
        self.thread.start()

    def update_progress(self, value, filename):
        self.progress_bar.setValue(value)
        self.lbl_status.setText(f"正在处理：{filename}")
        self.txt_log.append(f"{time.strftime('%H:%M:%S')} 已处理：{filename}")

    def on_finished(self, df):
        try:
            output_dir = os.path.join(self.lbl_path.text(), "处理结果")
            os.makedirs(output_dir, exist_ok=True)
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(output_dir, f"整合数据_{timestamp}.xlsx")
            df.to_excel(output_path, index=False)
            
            self.txt_log.append(f"\n✅ 处理完成！文件已保存至：{output_path}")
            self.lbl_status.setText("处理完成")
            
            # 自动打开文件
            if sys.platform == 'win32':
                os.startfile(output_path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', output_path])
            else:
                subprocess.run(['xdg-open', output_path])
                
        except Exception as e:
            self.show_error(f"保存文件失败：{str(e)}")

    def show_error(self, message):
        QMessageBox.critical(self, "错误", message)
        self.txt_log.append(f"\n❌ 错误：{message}")
        self.lbl_status.setText("处理中断")

    def closeEvent(self, event):
        if self.thread and self.thread.isRunning():
            reply = QMessageBox.question(
                self, '确认退出',
                '后台任务正在运行，确定要退出吗？',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.thread.stop()
                self.thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

    