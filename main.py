# 导入python 自带库

# 导入自定义模块

from Ui_fastwork import Ui_mainWindow
from PyQt5 import QtWidgets,QtGui
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QFileDialog
from core import Core

import os


class MergeWorkbooks(QtWidgets.QMainWindow, Ui_mainWindow):
    def __init__(self):
        super(MergeWorkbooks, self).__init__()
        self.setupUi(self)
        self.pushbutton_openfolder.setVisible(False)

        self.setWindowIcon(QIcon('logo.ico'))

        qss = open('./ui.qss').read()
        self.setStyleSheet(qss)

        print()

    def setBrowerPath(self):
        pass

    def init_app(self):
        self.excel_dirpath = self.lineEdit_filepath.text()
        self.is_check = self.checkBox.isChecked()
        self.textEdit.clear()

    @pyqtSlot()
    def on_open_filepath_clicked(self):
        filename = self.open_file_dialog()
        self.lineEdit_filepath.setText(filename)

    def open_file_dialog(self):
        dir_path = QFileDialog.getExistingDirectory(self,
                                                    "请选择需要合并的excel文件所在的目录",
                                                    r"%s" % os.getcwd())

        # if not os.path.exists(dir_path):
        #     return

        dir_path = dir_path.replace('/', '\\')  # windows下需要进行文件分隔符转换
        return dir_path

    @pyqtSlot()
    def on_pushbutton_openfolder_clicked(self):
        self.open_folder()

    @pyqtSlot()
    def on_pushbutton_start_clicked(self):
        self.init_app()
        if os.path.exists(self.excel_dirpath):
            self.to_merge()
        else:
            self.textEdit.setText("当前输入的文件路径有误！请检查后重新输入！")

    def to_merge(self):
        try:
            self.textEdit.clear()
            if self.is_check:
                keep_sheetname_lst = None
            else:
                keep_sheetname_lst = 0

            self.merge = Core(excel_dirpath=self.excel_dirpath, keep_sheetname_lst=keep_sheetname_lst)

            self.merge.progress_signal.connect(self.progress_signal_display)
            self.merge.start()
            self.pushbutton_openfolder.setVisible(True)
        except Exception as e:
            print(e)

    def progress_signal_display(self, log_info):
        self.textEdit.append(log_info)

    def error_message(self, error_info):
        self.label_progress.setText(error_info)

    def open_folder(self):
        abs_filepath = os.path.abspath(self.excel_dirpath)
        folder_path = os.path.dirname(abs_filepath)
        print(folder_path)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        os.system('explorer.exe /n, %s' % folder_path)


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    merge = MergeWorkbooks()
    merge.show()
    sys.exit(app.exec_())
