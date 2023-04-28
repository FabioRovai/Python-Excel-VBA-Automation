import win32com.client

class ModifyExcelWorkbook:
    def __init__(self, file_path):
        self.file_path = file_path
        self.xl = win32com.client.Dispatch("Excel.Application")
        self.workbook = None
        self.xl.DisplayAlerts = False
        self.xl.Application.EnableEvents = False

    def __enter__(self):
        self.workbook = self.xl.Workbooks.Open(self.file_path)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.save_and_close()

    def modify_macro_code(self):
        try:
            macros = self.workbook.VBProject.VBComponents

            # Print the names of the macros
            for macro in macros:
                if macro.Name == 'ThisWorkbook':
                    code_module = macro.CodeModule
                    code_module.DeleteLines(1, code_module.CountOfLines)

                    if "dst" in self.file_path:
                        new_code = "Private Sub Workbook_Open()" + '\n' + "Call run" + '\n' + "End Sub"
                    else:
                        new_code = "Private Sub Workbook_Open()" + '\n' + "Call run2" + '\n' + "End Sub"
                    code_module.AddFromString(new_code)
                    print(code_module)

                    break
        except Exception as e:
            print(f"An error occurred while modifying the macro code: {str(e)}")
            pass

    def save_and_close(self):
        try:
            self.workbook.Save()
        except Exception as e:
            print(f"An error occurred while saving the workbook: {str(e)}")
            pass
        finally:
            try:
                self.workbook.Close(False)
            except Exception as e:
                print(f"An error occurred while closing the workbook: {str(e)}")
                pass
            finally:
                try:
                    self.xl.Quit()
                except Exception as e:
                    print(f"An error occurred while quitting Excel: {str(e)}")
                    pass
