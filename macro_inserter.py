import xlwings as xw

def insert_macro_to_excel(file_path):
    if not file_path.endswith(".xlsm"):
        print("⚠️ Macro only works in .xlsm files. Please save your file as .xlsm")
        return

    app = xw.App(visible=False)
    wb = app.books.open(file_path)

    vba_code = '''
Sub AutoFilterColumnA()
    Columns("A:A").AutoFilter
End Sub
'''

    module = wb.api.VBProject.VBComponents.Add(1)  # 1 = Module
    module.CodeModule.AddFromString(vba_code)

    wb.save()
    wb.close()
    app.quit()
    print("✅ Macro added successfully.")

# ✅ Example usage
# insert_macro_to_excel("C:/Users/admin/Desktop/sample.xlsm")

def insert_macro_to_excel(file_path):
    if not file_path.endswith(".xlsm"):
        print("⚠️ Please save your file as .xlsm for macro support.")
        return
