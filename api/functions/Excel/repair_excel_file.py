
def repair_excel_file(filename):
    import  jpype     
    import  asposecells     
    jpype.startJVM() 
    from asposecells.api import Workbook
    
    workbook = Workbook("newfile.xlsx")
    workbook.Save("newfile.xlsx")
    jpype.shutdownJVM()
repair_excel_file('a')