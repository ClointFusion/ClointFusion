from PyQt5.QtWidgets import QApplication, QDesktopWidget
from PyQt5 import QtWidgets, QtCore, QtGui, QtWidgets
import os, sys

path = os.path.dirname(os.path.abspath(__file__))

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(300, 200)
        MainWindow.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        MainWindow.setAcceptDrops(True)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap((os.path.join(path, 'icon.png'))), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStatusTip("")
        MainWindow.setAutoFillBackground(True)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setMinimumSize(QtCore.QSize(300, 150))
        self.centralwidget.setStyleSheet("background-color: rgb(0, 0, 36);")
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setSpacing(0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.run_tab = QtWidgets.QTabWidget(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Roboto Black")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.run_tab.setFont(font)
        self.run_tab.setAcceptDrops(True)
        self.run_tab.setObjectName("run_tab")
        self.run = QtWidgets.QWidget()
        self.run.setAcceptDrops(True)
        self.run.setStyleSheet("background-color: rgb(0, 0, 36);\n"
"border-color: rgb(0, 0, 36);\n"
"selection-color: rgb(0, 0, 36);")
        self.run.setObjectName("run")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.run)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        spacerItem = QtWidgets.QSpacerItem(506, 30, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout.addItem(spacerItem)
        self.label = QtWidgets.QLabel(self.run)
        self.label.setMinimumSize(QtCore.QSize(300, 80))
        self.label.setMaximumSize(QtCore.QSize(10000, 10000))
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap((os.path.join(path, "CLOINTFUSION.svg"))))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        spacerItem1 = QtWidgets.QSpacerItem(40, 30, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout.addItem(spacerItem1)
        self.run_tab.addTab(self.run, "")
        self.delete = QtWidgets.QWidget()
        self.delete.setObjectName("delete")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.delete)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_3 = QtWidgets.QLabel(self.delete)
        self.label_3.setMinimumSize(QtCore.QSize(100, 100))
        self.label_3.setMaximumSize(QtCore.QSize(500, 500))
        self.label_3.setText("")
        self.label_3.setPixmap(QtGui.QPixmap((os.path.join(path, "DELETE.svg"))))
        self.label_3.setScaledContents(True)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_3.addWidget(self.label_3)
        self.run_tab.addTab(self.delete, "")
        self.verticalLayout_2.addWidget(self.run_tab)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.run_tab.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "DOST Client"))
        self.run_tab.setTabText(self.run_tab.indexOf(self.run), _translate("MainWindow", "Run"))
        self.run_tab.setTabText(self.run_tab.indexOf(self.delete), _translate("MainWindow", "Delete"))


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)

        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        tab = self.ui.run_tab.currentIndex()
        if tab == 0:
          for file in files:
            import threading
            name = f"thread_{files.index(file)}"
            name = threading.Thread(target=self.run_script, args=(file,))
            name.start()
            name.join()
        if tab == 1:
          for file in files:
            if os.path.isfile(file):
              os.remove(file)
            else:
              print("File not found")

    def run_script(self, script_path):
      import subprocess
      subprocess.run(["python", script_path])
      
      
    def location_on_the_screen(self):
        ag = QDesktopWidget().availableGeometry()
        sg = QDesktopWidget().screenGeometry()

        widget = self.geometry()
        x = ag.width() - widget.width()
        y = 2 * ag.height() - sg.height() - widget.height()
        self.move(x, y)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = Ui()
    ui.location_on_the_screen()
    ui.show()
    sys.exit(app.exec_())