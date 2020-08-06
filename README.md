---
typora-root-url: images
---

## 一键合并excel工作簿

功能：一键将相同目录下的所有excel文件合并至同一个工作表



![主界面](/images/main.png)

### 普通使用

如果你只是想使用这个小工具，请直接下载windows可执行文件压缩包。

链接：https://pan.baidu.com/s/1Bd0QB1SE4hUU6_sPog8Q1w 
提取码：6688
选择 一键合并excel工作簿.7z

解压后找到exe可执行文件，直接打开即可使用。这也是写这个小工具的契机。

### 开发者使用

如果你熟悉python,以下是启动流程

1 使用以下指令clone到本地

`git clone https://github.com/xugongli/merge_workbooks.git `

```python
│  core.py               # 这是合并文件的核心程序
│  imgcode_rc.py   # ui界面上的微信公众号素材图片
│  LICENSE
│  logo.ico            # 打包exe的logo图标
│  main.py         #  主程序，用于启动pyqt5界面
│  requirements.txt             # 依赖库
│  ui.qss                             # pyqt5界面美化
│  Ui_fastwork.py         # pyqt5窗口程序
│  __init__.py      
├─images               图片
│      imgcode.qrc
│      imgcode_rc.py
│      qrcode.jpg
│      wechat_search.png
│      
├─ui             # eric6 中使用qt-designer创建的ui工作文件，可以直接根据你的需要修改
│      fastwork.ui
│      fastwork_ui.e4p
│      Mainwindow.qss
│      Ui_fastwork.py
│      __init__.py
```



2 安装依赖

` pip install -r requirements.txt`

3 启动

`python main.py`


## 详细说明

[知乎-一键合并N个excel文件至同一工作表](https://zhuanlan.zhihu.com/p/168793811)

1 运行`python main.py` 打开一键合并excel工作簿小工具，点击 选择文件夹 按钮，选择excel文件主目录，如图中的 小文收藏的书单

![](/images/open_folder.gif)

2 选择是否合并单个文件中的工作表
默认不合并excel文件中的工作表，对于每个excel文件，只合并第一个工作表。
如果需要合并单个excel文件中的所有工作表，勾选合并工作表即可。
举例来说，例如某个excel文件中包含了科普、经典、生活等多个工作表，如果你不勾选合并工作表 ，只会将这个excel文件中的第一个工作表，即名为科普的这个工作表进行合并，其他的工作表就不会再读取了。

3 点击 开始合并 按钮，执行合并
合并完成后按照提示打开结果文件夹即可。

![](/images/start_merge.gif)









