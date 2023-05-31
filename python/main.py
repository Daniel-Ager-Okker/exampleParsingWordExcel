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
        self.ui.pbLoadPDBExcel.clicked.connect(self.__onLoadPDBExcel)
        self.ui.pbCompareWithPDBAndExport.clicked.connect(self.__onCompareWithPDBAndExport)

    def __onLoadOldWord(self):
        pathToOldWord = self.__getOpenFileName("docx")
        if pathToOldWord == "":
            return

        self.handler_.setPathToOldWord(pathToOldWord)
        # try:
        self.handler_.parseOldWord()
        if self.ui.chbxExportOld2Excel.isChecked():
            self.handler_.exportOldWordData()
        self.ui.lblOldWord.Text = " : " + pathToOldWord
        '''except Exception as exc:
            msgBox = QMessageBox(exc.Message)
            msgBox.exec_()'''

    def __onLoadNewWord(self):
        # implementation for loading new word file
        pass

    def __onCompareAndExport(self):
        # implementation for comparing and exporting files
        pass

    def __onClean(self):
        # implementation for cleaning files
        pass

    def __onLoadPDBExcel(self):
        # implementation for loading PDB Excel file
        pass

    def __onCompareWithPDBAndExport(self):
        # implementation for comparing with PDB Excel and exporting file
        pass

    def __getOpenFileName(self, extension: str):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file, _ = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "",
                                              f'All Files (*);;{extension} Files (*.{extension})', options=options)
        return  file


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
