from ExcelDao import ExcelDao
from ExcelWrited import ExcelWrited

if __name__ == "__main__":
    dao = ExcelDao("chuqin.xlsx")
    result = dao.readExcel()
    write = ExcelWrited("chuqin.xlsx", result)
    write.toWriteData()
    print(write.CList)
    print(write.ZList)
    print(write.JList)
    write.toWriteExcel()
