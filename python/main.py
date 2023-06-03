import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from main_window import Ui_ConverterWindow
from handler import Handler


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.handler_ = Handler()

        # Set up the user interface from Designer.
        self.ui = Ui_ConverterWindow()
        self.ui.setupUi(self)
        self.__setConnections()

    def __setConnections(self):
        self.ui.pbLoadOldWord.clicked.connect(self.__onLoadOldWord)
        self.ui.pbLoadNewWord.clicked.connect(self.__onLoadNewWord)
        self.ui.pbCompareAndExport.clicked.connect(self.__onCompareAndExport)
        self.ui.pbClean.clicked.connect(self.__onClean)

    def __onLoadOldWord(self):
        pathToOldWord = self.__getOpenFileName("docx")
        if pathToOldWord == "":
            return

        self.handler_.setPathToOldWord(pathToOldWord)
        self.handler_.parseOldWord()
        if self.ui.chbxExportOld2Excel.isChecked():
            self.handler_.exportOldWordData()
        self.ui.lblOldWord.setText(" : " + pathToOldWord)

    def __onLoadNewWord(self):
        pathToNewWord = self.__getOpenFileName("docx")
        if pathToNewWord == "":
            return

        self.handler_.setPathToNewWord(pathToNewWord)
        self.handler_.parseNewWord()
        if self.ui.chbxExportNew2Excel.isChecked():
            self.handler_.exportNewWordData()
        self.ui.lblNewWord.setText(" : " + pathToNewWord)

    def __onCompareAndExport(self):
        self.handler_.compareHashTables()
        self.__exportComparedData()

    def __onClean(self):
        self.handler_.clearData()
        self.ui.lblOldWord.setText(" : ")
        self.ui.lblNewWord.setText(" : ")

    def __getOpenFileName(self, extension: str):
        options = QFileDialog.Options()
        file, _ = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "",
                                              f'All Files (*);;{extension} Files (*.{extension})', options=options)
        return file

    def __exportComparedData(self) -> None:
        pathToExcel = self.__getSaveFileName("xlsx")
        if "" == pathToExcel:
            return

        self.handler_.exportDataToExcel(pathToExcel)

    def __getSaveFileName(self, extension: str) -> None:
        options = QFileDialog.Options()
        file, _ = QFileDialog.getSaveFileName(None, "QFileDialog.getSaveFileName()", "",
                                              f'All Files (*);;{extension} Files (*.{extension})', options=options)
        return file


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
