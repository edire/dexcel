
import os
from win32com.client import Dispatch


directory = os.path.dirname(os.path.realpath(__file__))


class Excel():
    
    macro_path = os.path.join(directory, 'MiscFiles', 'SharedVBA.bas')
    xl = Dispatch("Excel.Application")
    
    
    def __init__(self, file_path, import_vba=False, visible=False):
        self.directory = os.path.dirname(file_path)
        self.workbook = os.path.basename(file_path)
        
        self.wb = Excel.xl.Workbooks.Open(file_path)
        Excel.xl.Visible = visible
        
        if import_vba == True:
            xlm = self.wb.VBProject.VBComponents.Import(Excel.macro_path)
            xlm.Name = "SharedVBA"
        
        
    def __enter__(self):
        return self
    
    
    def __cleanup(self):
        num_instances = Excel.xl.Application.Workbooks.Count
        try:
            self.close()
            num_instances -= 1
        except:
            pass
        if num_instances <= 0:
            Excel.xl.Application.Quit()
            
    
    def __exit__(self, *args):
        self.__cleanup()
        
        
    def __del__(self):
        self.__cleanup()
        
        
    def save(self, new_file_path=None):
        Excel.xl.Application.DisplayAlerts = False
        Excel.xl.Application.Calculate()
        if new_file_path != None:
            self.wb.SaveAs(new_file_path)
        else:
            self.wb.Save()
        Excel.xl.Application.DisplayAlerts = True
            
        
    def close(self, save=False):
        if save == True:
            self.save()
        self.wb.Close(SaveChanges=0)
        

    def add_macro(self, macro, module_name=None):
        xlm = self.wb.VBProject.VBComponents.Add(1)
        if module_name != None:
            xlm.Name = module_name
        xlm.CodeModule.AddFromString(macro)
            
            
    def run(self, macro, module='Module1'):
        Excel.xl.Application.Run(f"{self.workbook}!{module}.{macro}")


    def refresh_all(self):
        self.run("RefreshWorkbook", 'SharedVBA')
        
        
    def save_pdf(self, pdf_sheet, pdf_range, pdf_path):
        Excel.xl.Application.Run(f"{self.workbook}!SharedVBA.SavePDF", pdf_sheet, pdf_range, pdf_path)
        
        
    def range_to_image(self, pic_sheet, pic_range, pic_path):
        Excel.xl.Application.Run(f"{self.workbook}!SharedVBA.RangetoImage", pic_sheet, pic_range, pic_path)