from PyQt6 import QtWidgets, QtCore, QtGui
from PyQt6.QtWidgets import QWidget, QFileDialog
import UploadCerts, os, shutil

class UploadCertsController(UploadCerts.Ui_UploadCerts, QWidget):
    def __init__(self, window):
        super().__init__()
        self.setupUi(window)
        window.setWindowIcon(QtGui.QIcon("MTUCLogo.png"))
        self.Paths = {}
        self.list = [
            {"label": self.Brackets, "button": self.BracketsUploadBtn, "pathLabel": self.path_1},
            {"label": self.NoBrackets, "button": self.NoBracketsUploadBtn, "pathLabel": self.path_2},
            {"label": self.BST, "button": self.BSTUploadBtn, "pathLabel": self.path_3},
            {"label": self.GMDSS, "button": self.GMDSSUploadBtn, "pathLabel": self.path_4},
            {"label": self.ECDISman, "button": self.ECDISmanUploadBtn, "pathLabel": self.path_5},
            {"label": self.ECDISop, "button": self.ECDISopUploadBtn, "pathLabel": self.path_6},
            {"label": self.CSO, "button": self.CSOUploadBtn, "pathLabel": self.path_7},
            {"label": self.PFSO, "button": self.PFSOUploadBtn, "pathLabel": self.path_8},
        ]
        for i in self.list:
            self.checkTemplates(i["label"], i["button"])
        
        self.BracketsUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[0]))
        self.NoBracketsUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[1]))
        self.BSTUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[2]))
        self.GMDSSUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[3]))
        self.ECDISmanUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[4]))
        self.ECDISopUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[5]))
        self.CSOUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[6]))
        self.PFSOUploadBtn.clicked.connect(lambda: self.uploadCert(self.list[7]))

        self.SaveBtn.clicked.connect(lambda: self.save(window))
        self.CancelBtn.clicked.connect(window.close)

    def checkTemplates(self, label, button):
        file = os.path.exists("Certificates Templates/" + label.text() + ".docx")
        if file:
            self.Paths[label.text()] = "Certificates Templates/" + label.text() + ".docx"
            label.setStyleSheet("color: green")
            button.setStyleSheet("color: green")
            button.setText("Update")
        else:
            label.setStyleSheet("color: red")
            button.setStyleSheet("color: red")

    def uploadCert(self, obj):
        label, button, pathLabel = obj["label"], obj["button"], obj["pathLabel"]
        path = QFileDialog.getOpenFileName(self, "Open File", ".", "Word (*.docx *.doc)")
        if not path[0]:
            return
        self.Paths[label.text()] = path[0]
        pathLabel.setText(path[0])
        label.setStyleSheet("color: blue")
        button.setStyleSheet("color: blue")


    def save(self, window):
        if not os.path.exists("Certificates Templates"):
            os.mkdir("Certificates Templates")

        for i in self.Paths:
            if self.Paths[i] == None or self.Paths[i] == "" or self.Paths[i] == "Certificates Templates/" + i + ".docx":
                continue
            try:
                shutil.copy2(self.Paths[i], "Certificates Templates/" + i + ".docx")
            except:
                continue
        window.close()

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = QtWidgets.QWidget()
    controller = UploadCertsController(window)
    window.show()
    app.exec()
