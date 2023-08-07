from PyQt6 import QtWidgets, QtCore, QtGui
from PyQt6.QtWidgets import  QWidget, QMessageBox, QFileDialog
import MainWindow, WebController, CertificateController, os, win32api
from datetime import datetime
import LoadData as DataLoader

class MainWindowController(MainWindow.Ui_MainWindow, QWidget):
    def __init__(self):
        super().__init__()
        self.setupUi(MainApp)
        MainApp.setWindowIcon(QtGui.QIcon("MTUCLogo.png"))
        self.WorkingDirectoryPath = ""
        self.CourseHasExpireDate = True
        self.pickDirectory.clicked.connect(self.SelectDirectory)
        self.defaultLang = win32api.GetKeyboardLayout(0)
        # Date Initialization
        self.To.setDate(QtCore.QDate.currentDate())
        self.From.setDate(QtCore.QDate.currentDate())
        self.Issue_Date.setDate(datetime.now())
        self.warning = "⚠Warning: Some Templates are missing."
        if os.path.exists("Certificates Templates/") and len(os.listdir("Certificates Templates/")) == 8:
            self.warning = ""
        self.CertError.setText(self.warning)
        self.Expire_Date.setDate(
            QtCore.QDate(
                datetime.strptime(self.To.text(), "%d/%m/%Y").year + 5,
                datetime.strptime(self.To.text(), "%d/%m/%Y").month,
                datetime.strptime(self.To.text(), "%d/%m/%Y").day - 1,
            )
        )
        self.autoCaps.clicked.connect(self.autoFormat)
        self.Name_En.textChanged.connect(self.autoFormat)
        self.Place_of_Birth_En.textChanged.connect(self.autoFormat)

        self.loader = DataLoader.LoadData()
        self.course.addItems(self.loader.getCoursesList())

        app.focusChanged.connect(self.on_focusChanged)

        self.course.currentIndexChanged.connect(self.setCourseCode)

        self.To.dateChanged.connect(self.autoSetExpireDate)

        self.CreateButton.clicked.connect(self.Validate)

        self.isPersonal.clicked.connect(self.styleFromWhere)

        self.uploadCerts.clicked.connect(self.openUploadWindow)

    def on_focusChanged(self):
        widget = app.focusWidget()
        self.focusedWidget = widget
        if widget == self.Name_En:
            self.changLanguage("en")
        elif widget == self.Name_Ar:
            self.changLanguage("ar")
        elif widget == self.Place_of_Birth_En:
            self.changLanguage("en")
        elif widget == self.Place_of_Birth_Ar:
            self.changLanguage("ar")
        elif widget == self.fromWhere_en:
            self.changLanguage("en")
        elif widget == self.fromWhere_ar:
            self.changLanguage("ar")
        else:
            self.changLanguage("default")

    def changLanguage(self, lang):
        if lang == "en":
            # load English US keyboard
            win32api.LoadKeyboardLayout("00000409", 1)
        elif lang == "ar":
            # load Arabic Egypt keyboard
            win32api.LoadKeyboardLayout("00000401", 1)

        elif lang == "default":
            # load default keyboard
            win32api.LoadKeyboardLayout(str(self.defaultLang), 1)

    def styleFromWhere(self):
        if self.isPersonal.isChecked():
            self.fromWhere_en.setEnabled(False)
            self.fromWhere_ar.setEnabled(False)
        else:
            self.fromWhere_en.setEnabled(True)
            self.fromWhere_ar.setEnabled(True)
        self.fromWhere_en.setStyleSheet("")
        self.fromWhere_ar.setStyleSheet("")

    def autoFormat(self):
        if self.autoCaps.isChecked():
            self.Name_En.setText(self.Name_En.text().upper())
            self.Place_of_Birth_En.setText(self.Place_of_Birth_En.text().upper())

    def SelectDirectory(self):
        path = QFileDialog.getExistingDirectory(
            self, "Select Directory", self.WorkingDirectoryPath
        )
        if not path == "":
            self.WorkingDirectoryPath = path
        self.workingPath.setText(self.WorkingDirectoryPath)

    def setCourseCode(self):
        code = self.courseNametoCode()
        self.courseCode.setText(code)
        noneExpireableCourses = self.loader.coursesWithNoExpireDate()
        self.rulesCode.setText(self.loader.getRulesCode(self.Course_Name_En.text()))
        if code in noneExpireableCourses:
            self.CourseHasExpireDate = False
            self.Expire_Date.setDate(QtCore.QDate(1900, 1, 1))
        else:
            self.CourseHasExpireDate = True
            self.autoSetExpireDate()

    def checkTemplateExists(self, msg, courseName=""):
        if os.path.exists("Certificates Templates/") and len(os.listdir("Certificates Templates/")) == 8:
            self.warning = ""
        if os.path.exists("Certificates Templates/"+self.loader.getCertificateType(courseName)+".docx"):
            self.CertError.setText(self.warning + " ✔️ Template Exists")
            self.CertError.setStyleSheet("")
            return True
        else:
            self.CertError.setText(msg)
            self.CertError.setStyleSheet("color: red;")
            return False

    def courseNametoCode(self):
        courseName_en = self.course.currentText().split(" - ")[0]
        self.Course_Name_En.setText(courseName_en)
        self.checkTemplateExists("⚠Certificate Template Not Found", courseName_en)
        arabicCourseName = self.course.currentText().split(" - ")[1].split(" ")
        code = arabicCourseName[-1].split("(")[1].split(")")[0]
        del arabicCourseName[-1]
        self.Course_Name_Ar.setText(" ".join(arabicCourseName))
        return code

    def autoSetExpireDate(self):
        if self.CourseHasExpireDate:
            toDate = datetime.strptime(self.To.text(), "%d/%m/%Y")
            self.Expire_Date.setDate(
                QtCore.QDate(toDate.year + 5, toDate.month, toDate.day - 1)
            )
        else:
            self.Expire_Date.setDate(QtCore.QDate(1900, 1, 1))

    def Validate(self):
        error = False
        Certificate_Data = {}
        if self.workingPath.text() == "":
            self.pickDirectory.setStyleSheet("border: 1px solid red;")
            error = True
        else:
            self.pickDirectory.setStyleSheet("")
        print("trace error 1:",error)
        if self.course.currentText() == "":
            self.course.setStyleSheet("border: 1px solid red;")
            Certificate_Data["CertNo"] = ""
            error = True
        else:
            self.course.setStyleSheet("")
            Certificate_Data["courseCode"] = self.courseNametoCode()
            Certificate_Data["CertNo"] = self.CertNo.text()
            Certificate_Data["Course_En"] = self.course.currentText().split(" - ")[0]
            arabicCourseName = self.course.currentText().split(" - ")[1].split(" ")
            code = arabicCourseName[-1].split("(")[1].split(")")[0]
            del arabicCourseName[-1]
            Certificate_Data["Course_Ar"] = " ".join(arabicCourseName)
            error = not self.checkTemplateExists("⚠Certificate Template Not Found", Certificate_Data["Course_En"])

        print("trace error 2:",error)
        if self.isPersonal.isChecked():
            Certificate_Data["FromWhere_En"] = "PERSONAL"
            Certificate_Data["FromWhere_Ar"] = "على نفقته الشخصية"
        else:
            if self.fromWhere_en.text() == "":
                self.fromWhere_en.setStyleSheet("border: 1px solid red;")
                error = True
            else:
                self.fromWhere_en.setStyleSheet("")
                Certificate_Data["FromWhere_En"] = self.fromWhere_en.text()
            if self.fromWhere_ar.text() == "":
                self.fromWhere_ar.setStyleSheet("border: 1px solid red;")
                error = True
            else:
                self.fromWhere_ar.setStyleSheet("")
                Certificate_Data["FromWhere_Ar"] = self.fromWhere_ar.text()
        #TODO: check if in the database (Future Work)
        # if self.CertNo.text() == "":
        #     self.CertNo.setStyleSheet("border: 1px solid red;")
        #     error = True
        # else:
        #     self.CertNo.setStyleSheet("")
        #     Certificate_Data["CertNo"] = Certificate_Data["CertNo"] + self.CertNo.text()
        print("trace error 3:",error)
        if self.Name_En.text() == "":
            self.Name_En.setStyleSheet("border: 1px solid red;")
            error = True
        else:
            self.Name_En.setStyleSheet("")
            Certificate_Data["Name_En"] = self.Name_En.text()
        print("trace error 4:",error)
        if self.Name_Ar.text() == "":
            self.Name_Ar.setStyleSheet("border: 1px solid red;")
            error = True
        else:
            self.Name_Ar.setStyleSheet("")
            Certificate_Data["Name_Ar"] = self.Name_Ar.text()
        print("trace error 5:",error)
        if self.Place_of_Birth_En.text() == "":
            self.Place_of_Birth_En.setStyleSheet("border: 1px solid red;")
            error = True
        else:
            self.Place_of_Birth_En.setStyleSheet("")
            Certificate_Data["Place_of_Birth_En"] = self.Place_of_Birth_En.text()
        print("trace error 6:",error)
        if self.Place_of_Birth_Ar.text() == "":
            self.Place_of_Birth_Ar.setStyleSheet("border: 1px solid red;")
            error = True
        else:
            self.Place_of_Birth_Ar.setStyleSheet("")
            Certificate_Data["Place_of_Birth_Ar"] = self.Place_of_Birth_Ar.text()
        print("trace error 7:",error)
        BirthDate = datetime(
            int(self.Date_of_Birth.text().split("/")[2]),
            int(self.Date_of_Birth.text().split("/")[1]),
            int(self.Date_of_Birth.text().split("/")[0]),
        )
        if BirthDate >= datetime.now() or BirthDate < datetime(1900, 1, 1):
            self.Date_of_Birth.setStyleSheet("border: 1px solid red;")
            error = True
        else:
            self.Date_of_Birth.setStyleSheet("")
            Certificate_Data["Date_of_Birth"] = self.Date_of_Birth.text()
        print("trace error 8:",error)
        #TODO: check if in the database (Future Work)
        # if self.regNo.text() == "":
        #     self.regNo.setStyleSheet("border: 1px solid red;")
        #     error = True
        # else:
        #     self.regNo.setStyleSheet("")
        #     Certificate_Data["RegNo"] = self.regNo.text()

        FromDate = datetime(
            int(self.From.text().split("/")[2]),
            int(self.From.text().split("/")[1]),
            int(self.From.text().split("/")[0]),
        )
        ToDate = datetime(
            int(self.To.text().split("/")[2]),
            int(self.To.text().split("/")[1]),
            int(self.To.text().split("/")[0]),
        )

        dateError1 = False
        dateError2 = False

        if (
            FromDate >= datetime.now()
            or FromDate < datetime(2020, 1, 25)
            or FromDate > ToDate
            or ToDate >= datetime.now()
            or ToDate < datetime(2020, 1, 25)
        ):
            self.From.setStyleSheet("border: 1px solid red;")
            self.To.setStyleSheet("border: 1px solid red;")
            dateError1 = True
        else:
            self.From.setStyleSheet("")
            self.To.setStyleSheet("")
            Certificate_Data["From"] = self.From.text()
            Certificate_Data["To"] = self.To.text()
            dateError1 = False

        IssueDate = datetime(
            int(self.Issue_Date.text().split("/")[2]),
            int(self.Issue_Date.text().split("/")[1]),
            int(self.Issue_Date.text().split("/")[0]),
        )
        if IssueDate > datetime.now() or IssueDate < datetime(2020, 1, 25):
            self.Issue_Date.setStyleSheet("border: 1px solid red;")
            dateError2 = True
        else:
            self.Issue_Date.setStyleSheet("")
            Certificate_Data["Issue_date"] = self.Issue_Date.text()
            dateError2 = False

        if error == False and dateError1 == False and dateError2 == False:
            try:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Question)
                msg.setText("Do you want to open the certificate in MS Word?")
                msg.setWindowIcon(QtGui.QIcon("MTUCLogo.png"))
                msg.setWindowTitle("Open Certificate")
                msg.setStandardButtons(
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                Certificate_Data["Expire_date"] = self.Expire_Date.text()
                Certificate_Data["RegNo"] = "H" + f"{int(self.regNo.text()):07d}"
                Certificate_Data["Rules"] = self.rulesCode.text()
                self.resultLabel.setText("Created Successfully")
                self.resultLabel.setStyleSheet("color: green;")
                certificateGenerator = CertificateController.CertificateController(
                    Certificate_Data, self.WorkingDirectoryPath
                )
                path = certificateGenerator.generateCertificate()
                try:
                    print(self.WorkingDirectoryPath)
                    recordAdder = WebController.WebController(self.WorkingDirectoryPath)
                    recordAdder.addDateToExcel(Certificate_Data)

                    x = msg.exec()
                    if x == QMessageBox.StandardButton.Yes:
                        os.startfile(path)
                    else:
                        pass
                except Exception as e:
                    print(e)
                    os.remove(path)
            except Exception as e:
                print(e)
                self.resultLabel.setText("⚠Error Creating Certificate")
                self.resultLabel.setStyleSheet("color: red;")
        else:
            try:
                self.resultLabel.setText("⚠Check your Entries")
                self.resultLabel.setStyleSheet("color: red;")
            except Exception as e:
                print(e)
    def openUploadWindow(self):
        from UploadCerts import Ui_UploadCerts
        import UploadCertsController
        self.uploadWindow = QtWidgets.QWidget()
        self.ui = UploadCertsController.UploadCertsController(self.uploadWindow, self.CertError)
        self.uploadWindow.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.uploadWindow.show()

def render():
    import sys
    global app
    app = QtWidgets.QApplication(sys.argv)
    global MainApp
    MainApp = QtWidgets.QWidget()
    ui = MainWindowController()
    MainApp.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    render()