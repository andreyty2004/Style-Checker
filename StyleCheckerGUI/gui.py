import os
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog, QApplication, QMainWindow, QMessageBox, QTextEdit
from PyQt5.QtGui import QIcon, QFont
import sys
from docx import Document
sys.path.append('./')
from page_numbering import page_numbering
os.chdir('StyleCheckerGUI')

class MyWindow(QMainWindow):
    fname = ""
    fpath = ""
    check_text = ""

    def __init__(self, width = 500, height = 500):
        super(MyWindow, self).__init__()
        self.resize(width, height)
        self.setFixedSize(width, height)
        self.center()
        self.setWindowTitle("Style Checker")
        self.setWindowIcon(QIcon('./assets/logo2.png'))
        # background = QPixmap("./assets/logo2.png").scaled(width, height)
        # pal = self.palette()
        # pal.setBrush(QPalette.Background, QBrush(background))
        # self.setPalette(pal)
        self.setStyleSheet("background-color: #C29A60;")
        self.initUI()

    def initUI(self):
        self.title = QtWidgets.QLabel(self)
        self.title.setText("Style Checker")
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.title.setFont(QFont('Helvetica Bold', 30))
        self.title.adjustSize()
        self.title.move((self.size().width() - self.title.size().width()) // 2, 50)
        
        self.upload = QtWidgets.QPushButton(self)
        self.upload.setText("ВЫБРАТЬ\nФАЙЛ")
        self.upload.setFont(QFont('Helvetica Bold', 12))
        self.upload.setToolTip("Выберете файл")
        self.upload.setGeometry(self.title.pos().x(), self.title.pos().y() + 100, 120, 120)
        self.upload.setStyleSheet("QPushButton::hover" "{""background-color : lightgreen;"";}"
                                  "QPushButton" "{"
                                                 "background : #8C8C8C;"
                                                 "border-radius: 15px;" 
                                                 "border-style: solid;" 
                                                 "border-color: black;"
                                                 "border-width: 5px;"
                                                 "}")
  
        self.upload.clicked.connect(self.browse_files)

        self.download = QtWidgets.QPushButton(self)
        self.download.setText("СКАЧАТЬ\nФАЙЛ")
        self.download.setFont(QFont('Helvetica Bold', 12))
        self.download.setToolTip("Скачать результаты проверки в формате .txt")
        self.download.setGeometry(self.title.pos().x() + self.title.size().width() - 
                                  120, self.title.pos().y() + 100, 120, 120)
        self.download.setStyleSheet("QPushButton::hover" "{""background-color : lightgreen;"";}"
                                    "QPushButton" "{"
                                                   "background : #8C8C8C;"
                                                   "border-radius: 15px;" 
                                                   "border-style: solid;" 
                                                   "border-color: black;"
                                                   "border-width: 5px;"
                                                   "}")
        # self.download.setGraphicsEffect(QGraphicsOpacityEffect().setOpacity(0.3))
        self.download.setEnabled(False)
        self.download.clicked.connect(self.download_clicked)

        self.view = QtWidgets.QPushButton(self)
        self.view.setIcon(QtWidgets.QApplication.style().standardIcon(QtWidgets.QStyle.SP_FileDialogDetailedView))
        self.view.setIconSize(QtCore.QSize(50, 50))
        self.view.setToolTip("Посмотреть результаты проверки")
        self.view.setGeometry((self.size().width() // 2 - self.view.iconSize().width() // 2),
                                    self.upload.pos().y() + (self.upload.height() - self.view.iconSize().height()) // 2, 
                                    self.view.iconSize().width(), self.view.iconSize().height())
        self.view.setStyleSheet("QPushButton" "{""border-style: solid""}")
        self.view.setEnabled(False)
        self.view.clicked.connect(self.view_clicked)

        self.label = QtWidgets.QLabel(self)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setFont(QFont('Helvetica Bold', 10))
        self.label.move(self.upload.pos().x(), 2 * self.upload.pos().y() + 10)
        self.label.setGeometry(self.upload.pos().x(), 2 * self.upload.pos().y() + 10, 
                               self.size().width() - self.upload.pos().x() * 2, 30)

    def view_clicked(self):
        self.view_window = TextWindow(800, 800, window_title = self.fname, check_text = self.check_text)
        self.view_window.show()

    def browse_files(self):
        path_file = open("PATHS.txt", 'r+', encoding = 'utf-8')
        browse_path = path_file.readline()
        self.fpath = QFileDialog.getOpenFileName(self, 'Выберете файл', browse_path, 'DOCX files (*.docx)')
        self.replace_line("PATHS.txt", 0, self.fpath[0])
        browse_path = self.fpath[0]
        
        # test part
        if(browse_path != ""):
            try:
                doc = Document(browse_path)
                self.download.setEnabled(True)
                self.view.setEnabled(True)
                self.fname = os.path.normpath(self.fpath[0]).split(os.path.sep)[-1]
                self.check_text = page_numbering(self.fpath[0])
                self.label.setText(self.fname)
            except:
                msg = QMessageBox()
                msg.setWindowTitle("Error")
                msg.setText("Произошла ошибка")
                msg.setIcon(QMessageBox.Critical)
                msg.setStandardButtons(QMessageBox.Retry)
                msg.setDefaultButton(QMessageBox.Retry)
                msg.exec_()

    def download_clicked(self):
        path_file = open("PATHS.txt", 'r', encoding = 'utf-8')
        download_path = self.read_line("PATHS.txt", 1)
        self.dpath = QFileDialog.getExistingDirectory(self, 'Выберете папку', download_path)
        self.replace_line("PATHS.txt", 1, self.dpath)
        download_path = self.dpath
        if(download_path != ""):
            download_file = open(download_path + '/' + self.fname[0:len(self.fname) - 5] + '_SC.txt', "w+", encoding = 'utf-8')
            download_file.write(self.check_text)

    def center(self):
        frameGm = self.frameGeometry()
        centerPoint = QtWidgets.QDesktopWidget().availableGeometry().center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def read_line(self, filename, line_number):
        file = open(filename, 'r+', encoding = 'utf-8')
        lines = file.readlines()
        total_lines = len(lines)
        if(line_number > total_lines - 1):
            return ""
        else:
            line = lines[line_number].rstrip('\n')
            return line

    def replace_line(self, filename, line_number, text):
        with open(filename, encoding = 'utf-8') as file:
            lines = file.readlines()
        if (line_number <= len(lines) - 1):
            lines[line_number] = text + "\n"
        else:
            lines.append(text + '\n')
        if (len(lines) == 0):
            lines[0] = text + "\n"
        with open(filename, "w", encoding = 'utf-8') as file:
            for line in lines:
                file.write(line)

class TextWindow(QMainWindow):
    def __init__(self, width = 700, height = 700, window_title = "", check_text = ""):
        super(TextWindow, self).__init__()
        self.resize(width, height)
        self.setWindowTitle(window_title)
        self.setWindowIcon(QIcon('./assets/logo3.png'))
        self.initUI(check_text)

    def initUI(self, check_text = ""):
        self.text = QTextEdit(self)
        self.text.insertHtml(open("HTML.html", "r", encoding = "utf-8").read())
        self.text.insertPlainText(check_text)
        self.text.setReadOnly(True)
        self.text.resize(self.width(), self.height())

def window():
    app = QApplication(sys.argv)
    win = MyWindow(500, 500)
    win.show()
    sys.exit(app.exec_())

window()
# добавить русификацию
# QLabel - название открытого файла
