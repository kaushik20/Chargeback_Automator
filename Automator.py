import os
import sys
import glob
import traceback
from datetime import datetime
from PyQt6 import QtCore, QtGui
from PyQt6.QtGui import QMovie, QIcon
from ITSM import WorkOrderReportProcessor
from PyQt6.QtWidgets import QMessageBox, QApplication, QMainWindow, QPushButton, QLineEdit, QLabel, QWidget

base_dir = os.path.dirname(__file__)
icon_path = os.path.join(base_dir, "Circular_Chargeback_Automator.png")
gif_path = os.path.join(base_dir, "ae7cd05d9438e3a42f955718affa1c9b.gif")

class Ui_MainWindow(object):
    def __init__(self):
        pass
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(363, 171)

        try:
            icon = QIcon(icon_path)
            if icon.isNull():
                raise ValueError("Icon file could not be loaded. Check the path and file.")
            MainWindow.setWindowIcon(icon)
        except Exception as e:
            print(f"Failed to set icon: {str(e)}")
        
        self.centralwidget = QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.backgroundLabel = QLabel(parent=self.centralwidget)
        self.backgroundLabel.setGeometry(QtCore.QRect(0, 0, MainWindow.width(), MainWindow.height()))
        self.backgroundLabel.setScaledContents(True)
        self.backgroundLabel.setObjectName("backgroundLabel")

        self.movie = QMovie(gif_path)
        self.backgroundLabel.setMovie(self.movie)
        self.movie.start()

        self.overlay = QLabel(self.centralwidget)
        self.overlay.setGeometry(QtCore.QRect(0, 0, MainWindow.width(), MainWindow.height()))
        self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 128);")

        self.lineEdit = QLineEdit(self.overlay)
        self.lineEdit.setGeometry(QtCore.QRect(140, 60, 141, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setStyleSheet("color: white; background-color: rgba(255, 255, 255, 50);")
        self.lineEdit.setReadOnly(False)
        self.lineEdit.setAcceptDrops(True)

        self.lineEdit.dragEnterEvent = self.dragEnterEvent
        self.lineEdit.dropEvent = self.dropEvent
        
        self.label_2 = QLabel(self.overlay)
        self.label_2.setGeometry(QtCore.QRect(10, 60, 120, 16))
        self.label_2.setStyleSheet("color: white;")
        self.label_2.setText("Select Document:")
        
        self.Automate = QPushButton("Automate", self.overlay)
        self.Automate.setGeometry(QtCore.QRect(145, 139, 71, 21))
        self.Automate.setStyleSheet("""QPushButton {background-color: #e74c3c; color: white; border-radius: 10px; border: 1px solid #c0392b; padding: 5px 10px;}QPushButton:hover {background-color: #c0392b;}QPushButton:pressed {background-color: #e67e22;}""")
        self.Automate.setObjectName("Automate")
        self.Automate.clicked.connect(self.run_automation)
        
        MainWindow.setCentralWidget(self.centralwidget)
        self.automate_file_selection()

        MainWindow.setWindowTitle("Chargeback Automator")
        self.lineEdit.setReadOnly(False)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        MainWindow.resizeEvent = self.resize_event
        self.resize_event(QtGui.QResizeEvent(MainWindow.size(), MainWindow.size()))

    def resize_event(self, event):
        self.backgroundLabel.setGeometry(0, 0, event.size().width(), event.size().height())
        self.overlay.setGeometry(0, 0, event.size().width(), event.size().height())

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        self.label_2.setText(_translate("MainWindow", "Select the Automation Document"))
        self.Automate.setText(_translate("MainWindow", "Automate"))

    def automate_file_selection(self):
        folder_path = os.path.join(os.path.expanduser("~"), "Downloads")
        list_of_files = glob.glob(os.path.join(folder_path, 'WO Report.xlsx*'))
        if list_of_files:
            latest_file = max(list_of_files, key=os.path.getctime)
            self.lineEdit.setText(latest_file)
            self.lineEdit.setReadOnly(True)
            self.lineEdit.setAcceptDrops(False)
            self.lineEdit.setStyleSheet("color: white; background-color: rgba(255, 255, 255, 50); font-style: normal;")
        else:
            self.lineEdit.setText("Drag and drop an Excel sheet here")
            self.lineEdit.setReadOnly(False)
            self.lineEdit.setAcceptDrops(True)
            self.lineEdit.setStyleSheet("color: gray; background-color: rgba(255, 255, 255, 50); font-style: italic;")
            self.lineEdit.dragEnterEvent = self.dragEnterEvent
            self.lineEdit.dropEvent = self.dropEvent
    
    def run_automation(self):
        self.input_file_path = self.lineEdit.text()
        if self.input_file_path and os.path.exists(self.lineEdit.text()):
            try:
                processor = WorkOrderReportProcessor(input_file_path=self.input_file_path)
                processor.automate_ITSM()
                QMessageBox.information(None, "Success", "Automation completed successfully.")
            except Exception as e:
                self.log_error(str(e))
                QMessageBox.critical(None, "Automation Error", f"An error occurred during automation: {str(e)}")
        else:
            QMessageBox.warning(None, "Invalid File", "Please select a valid file to automate.")

    def log_error(self, error_message):
        log_folder = os.path.join(os.path.expanduser("~"), "Desktop", "Logs")
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
        log_file_path = os.path.join(log_folder, "error_log.txt")
        with open(log_file_path, "a") as log_file:
            log_file.write(f"{datetime.now()}:{error_message}\n")
            log_file.write(f"Traceback: {traceback.format_exc()}\n")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            self.lineEdit.setText(file_path)
            self.lineEdit.setStyleSheet("color: white; background-color: rgba(255, 255, 255, 50); font-style: normal;")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())