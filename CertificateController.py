from docxtpl import DocxTemplate, RichText
from pathlib import Path
import datetime, os
import LoadData as ld


class CertificateController:
    def __init__(self, data, workingDir):
        self.data = data
        self.workingDir = workingDir
        self.template = DocxTemplate("Certificate Template.docx")
        self.Loader = ld.LoadData()

    def generateCertificate(self):
        # get my current path
        TemplatePath = (
            str(Path(__file__).parent.absolute())
            + "/Certificates Templates/"
            + self.selectTemplate(self.data["Course_En"])
        )
        print("Templates Path", TemplatePath)
        doc = DocxTemplate(TemplatePath)

        def e2p(s):
            return (
                s.replace("0", "٠")
                .replace("1", "١")
                .replace("2", "٢")
                .replace("3", "٣")
                .replace("4", "٤")
                .replace("5", "٥")
                .replace("6", "٦")
                .replace("7", "٧")
                .replace("8", "٨")
                .replace("9", "٩")
            )

        def rev(s):
            s = s.split("/")
            s = s[::-1]
            s = "/".join(s)
            return s

        nonExpireCourses = self.Loader.coursesWithNoExpireDate()
        if self.data["courseCode"] in nonExpireCourses:
            self.data["Expire_date"] = "--/--/----"
            
        def formatName(str, length):
            if len(str) > length:
                size = 10*2
            else:
                size = 14*2
            return RichText(str, size=size)
        
        #TODO: uncomment the following lines when adding {{r <var> }} in the template
        # ask for the max length of the name in English and Arabic
        context = {
            "CertNo": self.data["courseCode"] + str(self.data["CertNo"]),
            # "Name_En": formatName(self.data["Name_En"], 30),
            # "Name_Ar": formatName(self.data["Name_Ar"], 35),
            "Name_En": self.data["Name_En"],
            "Name_Ar": self.data["Name_Ar"],
            "Place_of_Birth_En": self.data["Place_of_Birth_En"],
            "Place_of_Birth_Ar": self.data["Place_of_Birth_Ar"],
            "Date_of_Birth_En": self.data["Date_of_Birth"],
            "Date_of_Birth_Ar": e2p(rev(self.data["Date_of_Birth"])),
            "FromWhere_En": self.data["FromWhere_En"],
            "FromWhere_Ar": self.data["FromWhere_Ar"],
            "Course_En": self.data["Course_En"],
            "Course_Ar": self.data["Course_Ar"],
            "From_En": self.data["From"],
            "To_En": self.data["To"],
            "From_Ar": e2p(rev(self.data["From"])),
            "To_Ar": e2p(rev(self.data["To"])),
            "Rules_En": self.data["Rules"],
            "Rules_Ar": self.data["Rules"],
            "Expire_date": self.data["Expire_date"],
            "Issue_date": self.data["Issue_date"],
            "RegNo": self.data["RegNo"],
        }
        doc.render(context)
        savingDir = self.workingDir
        print(savingDir)
        path = (
            savingDir + "/" + self.data["courseCode"] + self.data["Name_Ar"] + ".docx"
        )
        doc.save(path)
        return path

    def selectTemplate(self, courseName):
        certTypes = self.Loader.getCertificateType(courseName)

        return certTypes + ".docx"

    def check_certs_folder(self):
        if not os.path.exists(os.getcwd() + "/Certificates"):
            os.mkdir(os.getcwd() + "/Certificates")
        
    # def selectTemplate(self, code):
    #     if code in self.Loader.getCertificateType()["Brackets"]:
    #         return "Brackets.docx"
    #     elif code in self.Loader.getCertificateType()["No Brackets"]:
    #         return "No Brackets.docx"
    #     elif code in self.Loader.getCertificateType()["BST"]:
    #         return "BST.docx"
    #     elif code in self.Loader.getCertificateType()["ECDIS OP"]:
    #         return "its ECDIS OP"
    #     elif code in self.Loader.getCertificateType()["CSO"]:
    #         return "its CSO"
    #     elif code in self.Loader.getCertificateType()["PFSC"]:
    #         return "its PFSO"
    #     elif code in self.Loader.getCertificateType()["GMDSS"]:
    #         return "its GMDSS"
    #     elif code in self.Loader.getCertificateType()["ECDIS OP"]:
    #         return "its ECDIS Operation"
    #     elif code in self.Loader.getCertificateType()["ECDIS MANA"]:
    #         return "its ECDIS Management"


# data = {
#     "courseCode": "XNFF",
#     "CertNo": 123,
#     "Name_En": "Omar Ahmed Mahmoud Ahmed Mahmoud Shalaby".upper(),
#     "Name_Ar": "عمر أحمد محمود أحمد محمود شلبي",
#     "Place_of_Birth_En": "Alex".upper(),
#     "Place_of_Birth_Ar": "الإسكندرية",
#     "Date_of_Birth": "01/01/1999",
#     "FromWhere_En": "Personal",
#     "FromWhere_Ar": "شخصي",
#     "Course_En": "Fire Fighting",
#     "Course_Ar": "الإطفاء",
#     "From": "01/01/2020",
#     "To": "02/01/2020",
#     "Rules": "A - VII / 1-1",
#     "Expire_date": "01/01/2025",
#     "Issue_date": "04/01/2020",
#     "RegNo": "H00123456",
# }


# loader = ld.LoadData()
# courseslist = loader.getCoursesList()
# # print(courseslist)
# Courses_En = [x.split(" - ")[0] for x in courseslist]
# Courses_Ar = [x.split(" - ")[1].split(" (")[0] for x in courseslist]
# Courses_Code = [x.split(" - ")[1].split(" (")[-1].split(")")[0] for x in courseslist]
# # print(Courses_Code)

# for i in range(len(Courses_En)):
#     data["courseCode"] = Courses_Code[i]
#     data["Course_En"] = Courses_En[i]
#     data["Course_Ar"] = Courses_Ar[i]
#     data["Rules"] = loader.getRulesCode(Courses_En[i])
#     generator = CertificateController(data, "D:\Projects\Python\MTUC Project\Results")
#     generator.generateCertificate()
#     print(f"{i} Done")
# generator = CertificateController(
#     data, "D:\Projects\Python\MTUC Project\\fire fighting"
# ).generateCertificate()
