import requests
from http.cookiejar import CookieJar
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import re
import datetime
import os
from lxml import html
from weasyprint import HTML

host = 'http://c.gb688.cn'
dom = '''
<html>
<head>
<meta name="viewport" content="width=1200, initial-scale=1, maximum-scale=5">
<style>
@page {
    size: 1191px 1684px;
    margin: 0px;
}
html, body {
    margin: 0;
    padding: 0;
}
.pdfViewer {
    transform-origin: center top;
    background-color: #ffffff;
}
.page {
    direction: ltr;
    margin: 0 auto;
    background-color: #ffffff;
    position: relative;
    overflow: visible;
    width: 100%;
    box-sizing: content-box;
}
[class^="pdfImg"] {
    width: 10%;
    height: 10%;
    display: block;
    position: absolute;
}
[class^="pdfImg-1-"] {
    left: 10%;
}
[class^="pdfImg-2-"] {
    left: 20%;
}
[class^="pdfImg-3-"] {
    left: 30%;
}
[class^="pdfImg-4-"] {
    left: 40%;
}
[class^="pdfImg-5-"] {
    left: 50%;
}
[class^="pdfImg-6-"] {
    left: 60%;
}
[class^="pdfImg-7-"] {
    left: 70%;
}
[class^="pdfImg-8-"] {
    left: 80%;
}
[class^="pdfImg-9-"] {
    left: 90%;
}
[class^="pdfImg-"][class$="-1"] {
    top: 10%;
}
[class^="pdfImg-"][class$="-2"] {
    top: 20%;
}
[class^="pdfImg-"][class$="-3"] {
    top: 30%;
}
[class^="pdfImg-"][class$="-4"] {
    top: 40%;
}
[class^="pdfImg-"][class$="-5"] {
    top: 50%;
}
[class^="pdfImg-"][class$="-6"] {
    top: 60%;
}
[class^="pdfImg-"][class$="-7"] {
    top: 70%;
}
[class^="pdfImg-"][class$="-8"] {
    top: 80%;
}
[class^="pdfImg-"][class$="-9"] {
    top: 90%;
}
</style>
<head>
<body>
</body>
</html>
'''


class App(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

        self.cookie_jar = CookieJar()
        self.session = requests.Session()
        self.session.cookies = self.cookie_jar

    def initUI(self):
        self.setWindowTitle('国家标准文档PDF导出工具')

        grid = QtWidgets.QGridLayout()

        self.label = QtWidgets.QLabel('1、请输入文档地址')
        self.label.setStyleSheet("color: red;")
        grid.addWidget(self.label, 0, 0, 1, 2, QtCore.Qt.AlignCenter)

        self.link_a_entry = QtWidgets.QLineEdit(self)
        self.link_a_entry.setStyleSheet("height: 32px;border: 1px solid red;")
        grid.addWidget(self.link_a_entry, 1, 0, 1, 2)

        self.button = QtWidgets.QPushButton('2、获取验证码', self)
        self.button.setStyleSheet("height: 32px;")
        self.button.clicked.connect(self.get_captcha)
        grid.addWidget(self.button, 2, 0, 1, 2, QtCore.Qt.AlignCenter)

        self.captcha_pad = QtWidgets.QLabel('3、输入验证码')
        self.captcha_pad.setStyleSheet("color: red;")
        grid.addWidget(self.captcha_pad, 3, 0, 4, 2, QtCore.Qt.AlignCenter)

        self.captcha_label = QtWidgets.QLabel(self)
        self.captcha_label.setAlignment(QtCore.Qt.AlignCenter)
        grid.addWidget(self.captcha_label, 7, 0, 1, 2)

        self.result_label = QtWidgets.QLabel(self)
        self.result_label.setStyleSheet("color: yellow;")
        self.result_label.setAlignment(QtCore.Qt.AlignCenter)
        self.result_label.setWordWrap(True)
        self.result_label.setMaximumWidth(380)
        grid.addWidget(self.result_label, 8, 0, 1, 2)

        self.download_button = QtWidgets.QPushButton('4、下载文档到本地', self)
        self.download_button.setStyleSheet("height: 32px;")
        self.download_button.clicked.connect(self.download_document)
        self.download_button.setEnabled(False)
        grid.addWidget(self.download_button, 9, 0, 1, 2, QtCore.Qt.AlignCenter)

        self.setLayout(grid)

        # 设置窗口宽高
        self.setFixedSize(400, 440)
        self.show()

    def get_link_a(self):
        link_a = self.link_a_entry.text()
        pattern = r'^http://c\.gb688\.cn/bzgk/gb/showGb\?type=online&hcno=[A-F0-9]+'
        match = re.match(pattern, link_a)
        if match:
            return link_a
        else:
            self.result_label.setText("输入的链接有误")
            return False

    def get_captcha(self):
        link_a = self.get_link_a()
        if not link_a:
            return
        link_b = host + "/bzgk/gb/gc"

        # 记录cookie
        self.session.get(link_a)

        # 获取验证码图片
        response = self.session.get(link_b)
        img_data = response.content

        image = QtGui.QImage.fromData(img_data)
        pixmap = QtGui.QPixmap.fromImage(image)

        if not pixmap.isNull():
            self.captcha_label.setPixmap(pixmap)
            self.result_label.setText("获取验证码成功，请在弹窗中输入验证码")
            self.input_captcha()
        else:
            self.result_label.setText("验证码图片加载失败")

    def input_captcha(self):
        captcha, ok = QtWidgets.QInputDialog.getText(self, '输入验证码', '3、请输入图片中的验证码：')
        if ok:
            self.verify_captcha(captcha)
        else:
            self.result_label.clear()
            self.captcha_label.clear()

    def verify_captcha(self, captcha):
        link_c = host + "/bzgk/gb/verifyCode"
        payload = {"verifyCode": captcha}

        # 调用接口C进行验证
        resp = self.session.post(link_c, data=payload)
        if resp.text == 'error':
            self.result_label.setText("输入的验证码错误")
            self.captcha_label.clear()
            return

        link_a = self.get_link_a()
        if not link_a:
            return

        # 重新请求A，保存返回值到本地文件html中
        response = self.session.get(link_a)
        with open("response.html", "w", encoding="utf-8") as file:
            file.write(response.text)

        self.result_label.setText("验证成功")
        self.download_button.setEnabled(True)
        self.download_button.setText('4、下载文档到本地')

    def download_document(self):
        # 经过系列处理把response.html转换为pdf 并保存到指定位置
        # 读取 response.html 文件
        html_file = "response.html"

        # 检查文件是否存在
        if not os.path.exists(html_file):
            self.result_label.setText("response.html 文件未找到")
            return

        # 弹出保存文件对话框
        pdf_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "保存 PDF 文件",
            datetime.datetime.now().strftime('%Y%m%d_%H%M%S'),
            "PDF 文件 (*.pdf);;所有文件 (*)"
        )

        if not pdf_path:
            self.result_label.setText("未选择文件保存路径")
            return

        # 转换 HTML 到 PDF
        self.result_label.setText("文档正在转换中，请耐心等待，具体时长取决于文档大小。")
        self.download_button.setText('文档转换中...')
        self.download_button.setEnabled(False)
        self.setEnabled(False)

        # 启动后台线程执行 PDF 生成任务
        self.thread = WorkerThread(html_file, pdf_path)
        self.thread.finished.connect(self.on_finished)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_finished(self, message):
        self.result_label.setText(message)
        self.download_button.setText('文档保存成功')
        self.setEnabled(True)

    def on_error(self, message):
        self.result_label.setText(message)
        self.download_button.setText('转换失败')
        self.setEnabled(True)


class WorkerThread(QtCore.QThread):
    # 定义信号用于通知任务完成或出错
    finished = QtCore.pyqtSignal(str)
    error = QtCore.pyqtSignal(str)

    def __init__(self, html_file, pdf_path):
        super().__init__()
        self.html_file = html_file
        self.pdf_path = pdf_path

    def run(self):
        try:
            HTML(string=self.handle_html(self.html_file)).write_pdf(self.pdf_path)
            os.remove(self.html_file)
            self.finished.emit("文档已保存至: {}".format(self.pdf_path))
        except Exception as e:
            self.error.emit("转换为PDF失败: {}".format(str(e)))

    def handle_html(self, html_file):
        # 读取 HTML 内容
        with open(html_file, "r", encoding="utf-8") as file:
            html_content = file.read()

        tree = html.fromstring(html_content)
        viewer_div = tree.xpath('//div[@id="viewer" and @class="pdfViewer"]')[0]
        # 替换 span 背景
        for page in viewer_div.xpath('.//div[@class="page"]'):
            bg_image = host + '/bzgk/gb/' + page.get('bg')
            spans = page.xpath('.//span')
            for span in spans:
                # 替换 span 的背景
                span_style = 'background-image: url(' + bg_image + ');' + span.get('style', '')
                span.set('style', span_style)
        new_page = html.fromstring(dom)
        body = new_page.find('.//body')
        body.append(viewer_div)
        new_html = html.tostring(new_page, pretty_print=True, encoding='utf-8').decode('utf-8')

        # with open("demo.html", "w", encoding="utf-8") as file:
        #     file.write(new_html)

        return new_html


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
