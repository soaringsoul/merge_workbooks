from fastwork.merge import MergeExcel
from PyQt5 import QtCore
import pandas as pd
import os


class Core(QtCore.QThread):
    progress_signal = QtCore.pyqtSignal(str)

    def __init__(self, excel_dirpath, keep_sheetname_lst=None):
        super(Core, self).__init__()
        self.excel_dirpath = excel_dirpath
        self.keep_sheetname_lst = keep_sheetname_lst

    def run(self):
        self.progress_signal.emit("开始合并!")
        merge = MergeExcel(folder_path=self.excel_dirpath, sheetname_lst=self.keep_sheetname_lst)
        try:
            df_all = merge.merge_workbooks()
        except Exception as e:
            self.progress_signal.emit("合并出错了！请检查你输入的文件路径是否有误，"
                                      "或者微信搜索【人文互联网】公众号，复制以下错误详情发送给作者！")
            self.progress_signal.emit("错误详情：%s" %e)

        if df_all is not None:
            merge.to_excel(df_all)
            new_filename = "%s_合并完成.xlsx" % os.path.basename(self.excel_dirpath)
            abs_filepath = os.path.abspath(self.excel_dirpath)
            new_filepath = os.path.join(os.path.dirname(abs_filepath), new_filename)
            self.progress_signal.emit('*' * 30)
            self.progress_signal.emit('*' * 30)
            self.progress_signal.emit("【合并完成】，合并后的工作表共计%s行" % df_all.shape[0])
            self.progress_signal.emit("请到以下目录获取合并后的excel文件：\n【%s】" % new_filepath)

        # self.progress_signal.emit("合并完成")
