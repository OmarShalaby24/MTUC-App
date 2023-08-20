from docxtpl import DocxTemplate


class LetterController:
    def __init__(self):
        self.templatePath = "Letter Template.docx"

    def generate_letter(self, data, working_dir="."):
        doc = DocxTemplate(self.templatePath)
        doc.render(data)
        doc.save(working_dir + "/الخطاب.docx")


if __name__ == "__main__":
    letter = LetterController()

    def rev(s):
        s = s.split("/")
        s = s[::-1]
        s = "/".join(s)
        return s

    data = {
        "CourseName_ar": "منع الحرائق ومكافحتها",
        "Expire": rev("12/8/2028"),
        "From": rev("11/8/2023"),
        "To": rev("13/8/2023"),
        "students": [
            {
                "status": "مدني",
                "passport": "S000123",
                "nationality": "مصري",
                "serial": "XNFF1",
                "regno": "H0025361",
                "name": "كريم أحمد محمود أحمد محمود",
                "i": "1",
            },
            {
                "status": "عسكري",
                "passport": "S000223",
                "nationality": "مصري",
                "serial": "XNFF2",
                "regno": "H0025362",
                "name": "عمر أحمد محمود أحمد محمود شلبي",
                "i": "2",
            },
            {
                "status": "مدني",
                "passport": "S000113",
                "nationality": "مصري",
                "serial": "XNFF3",
                "regno": "H0025363",
                "name": "أحمد محمود أحمد محمود",
                "i": "3",
            },
        ],
        "Issue_date": rev("17/8/2023"),
        "Issuer": "رشا صبرى",
    }
    data["Count"] = len(data["students"])
    data["Civil_count"] = 2
    data["Military_count"] = 1

    letter.generate_letter(data)
