from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout, QFileDialog, QLineEdit, QPushButton\
    , QMessageBox
from PyQt5.QtGui import QIcon
import sys
import pyperclip
import mammoth
from os.path import exists
from base64 import b64decode
from tools_png import img


class Windows(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()  # 垂直布局
        self.h_box = QHBoxLayout()
        self.label = QLabel()   # 标签1
        self.line = QLineEdit()  # 输入框
        self.button1 = QPushButton()  # 按钮1:选择文件
        self.button2 = QPushButton()  # 按钮2：开始转换
        self.doc_path = None
        self.file_dialog = None
        self.init_ui()

    def init_ui(self):
        self.resize(300, 100)
        self.setWindowTitle('word一键转html')
        self.setWindowIcon(QIcon('tools.png'))
        # --  设置布局 -- #
        # 设置标签1
        self.label.setText('请选择word文件')
        self.layout.addWidget(self.label)
        # 设置输入框
        self.line = QLineEdit()
        self.layout.addWidget(self.line)
        # 设置按钮1
        self.button1.setText('选择文件')
        self.button1.clicked.connect(self.select_file)
        self.h_box.addWidget(self.button1)
        # 设置按钮2
        self.button2.setText('开始转换')
        self.button2.clicked.connect(self.convert_file)
        self.h_box.addWidget(self.button2)
        self.layout.addLayout(self.h_box)
        # 设置总布局
        self.setLayout(self.layout)
        self.show()

    def select_file(self):
        self.file_dialog = QFileDialog()
        self.file_dialog.resize(300, 300)
        self.file_dialog.setWindowIcon(QIcon('tools.png'))
        file = self.file_dialog.getOpenFileName(self, '选择文件', '', 'Word文件(*.docx , *.doc)')
        if bool(file[0]):
            self.line.setText(file[0])
            self.doc_path = file[0]

    def convert_file(self):
        # print(self.doc_path)
        if self.doc_path is not None:
            with open(self.doc_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html = result.value  # The generated HTML
                print(html)
                text = '亲，文件已经给你转换好啦！<br>点击“Yes”即可将内容复制到剪辑版，点击“No”则导出到html文件'
                replay = QMessageBox.warning(self, '提示', text, QMessageBox.No, QMessageBox.Yes)
                print(replay)
            if replay == 65536:
                print(self.doc_path.replace('.', '_') + '.html')
                f1 = open(self.doc_path.replace('.', '_') + '.html', 'wt', encoding='utf-8')
                f1.write(html)
                f1.close()
            else:
                pyperclip.copy(html)  # 复制到剪辑板
        else:
            text = '亲，还没有选择word文件'
            QMessageBox.warning(self, '错误', text)


if __name__ == '__main__':
    # 构建图标图片
    if not exists('tools.png'):
        img_data = b64decode(img)
        f = open('tools.png', 'wb')
        f.write(img_data)
        f.close()
    app = QApplication(sys.argv)
    win = Windows()
    sys.exit(app.exec_())