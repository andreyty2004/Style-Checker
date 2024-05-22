from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QDialog, QFileDialog, QApplication, QMainWindow, QMessageBox, QTextEdit, QGraphicsOpacityEffect
from PyQt5.QtGui import QPainter, QColor, QIcon, QFont, QPixmap, QPalette, QBrush
import sys
from docx import Document

class MyWindow(QMainWindow):
    def __init__(self, width = 500, height = 500):
        super(MyWindow, self).__init__()
        self.resize(width, height)
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
        self.upload.setText("UP\nLOAD")
        self.upload.setFont(QFont('Helvetica Bold', 15))
        self.upload.setToolTip("Select file to process")
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
        self.download.setText("DOWN\nLOAD")
        self.download.setFont(QFont('Helvetica Bold', 15))
        self.download.setToolTip("Download Result")
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
        self.download.clicked.connect(self.browse_files)

        self.view = QtWidgets.QPushButton(self)
        self.view.setIcon(QtWidgets.QApplication.style().standardIcon(QtWidgets.QStyle.SP_FileDialogDetailedView))
        self.view.setIconSize(QtCore.QSize(50, 50))
        self.view.setToolTip("View Result")
        self.view.setGeometry((self.size().width() // 2 - self.view.iconSize().width() // 2),
                                    self.upload.pos().y() + (self.upload.height() - self.view.iconSize().height()) // 2, 
                                    self.view.iconSize().width(), self.view.iconSize().height())
        self.view.setStyleSheet("QPushButton" "{""border-style: solid""}")
        self.view.setEnabled(False)
        self.view.clicked.connect(self.view_clicked)
    
    def view_clicked(self):
        self.view_window = TextWindow()
        self.view_window.show()

    def browse_files(self):
        path_file = open("PATHS.txt", "w+", encoding = "utf-8")
        browse_path = path_file.readline()
        fname = QFileDialog.getOpenFileName(self, 'Select file', browse_path, 'DOCX files (*.docx)')
        path_file.writelines(fname[0])
        browse_path = fname[0]

        # test part
        if(browse_path != ""):
            try:
                doc = Document(browse_path)
                self.download.setEnabled(True)
                self.view.setEnabled(True)
            except:
                msg = QMessageBox()
                msg.setWindowTitle("Error")
                msg.setText("Package Not Found")
                msg.setIcon(QMessageBox.Critical)
                msg.setStandardButtons(QMessageBox.Retry)
                msg.setDefaultButton(QMessageBox.Retry)
                msg.exec_()
 
    # def downloaf_file(self):
    #     self.label.adjustSize()   

    def center(self):
        frameGm = self.frameGeometry()
        centerPoint = QtWidgets.QDesktopWidget().availableGeometry().center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

class TextWindow(QMainWindow):
    def __init__(self, width = 700, height = 700):
        super(TextWindow, self).__init__()
        self.resize(width, height)
        self.setWindowTitle("Checking logs")
        self.setWindowIcon(QIcon('./assets/logo3.png'))
        self.initUI()

    def initUI(self):
        self.text = QTextEdit(self)
        self.text.insertHtml(open("HTML.html", "r", encoding = "utf-8").read())

        # self.text.setReadOnly(True)
        self.text.resize(self.width(), self.height())

         


def window():
    app = QApplication(sys.argv)
    win = MyWindow(500, 500)
    win.show()
    sys.exit(app.exec_())

window()

# add ways to remember upload and download paths
# add lock to download button if no upload file chosen
# add tooltips at the bottom 
# add view button inbetween to vire output without downloading txt file
# add logo checking logs and then text 


# добавить русификацию
# status bar
# QLabel - название открытого файла
