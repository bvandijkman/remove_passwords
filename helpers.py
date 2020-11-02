import win32com.client
from pathlib import Path
import subprocess

def remove_password_xlsx(filename, pw_str):
        """ Removes password from excel file

        Args:
            filename (str): Full path to the excel file
            pw_str (str): The password to be removed
        """        
        xcl = win32com.client.Dispatch("Excel.Application")
        wb = xcl.Workbooks.Open(filename, False, False, None, pw_str)
        wb.Unprotect(pw_str)
        wb.UnprotectSharing(pw_str)
        xcl.DisplayAlerts = False
        wb.SaveAs(filename, None, '', '')
        xcl.Quit()


def set_password(excel_file_path, pw):

    excel_file_path = Path(excel_file_path)

    vbs_script = \
    f"""' Save with password required upon opening

    Set excel_object = CreateObject("Excel.Application")
    Set workbook = excel_object.Workbooks.Open("{excel_file_path}")

    excel_object.DisplayAlerts = False
    excel_object.Visible = False

    workbook.SaveAs "{excel_file_path}",, "{pw}"

    excel_object.Application.Quit
    """

    # write
    vbs_script_path = excel_file_path.parent.joinpath("set_pw.vbs")
    with open(vbs_script_path, "w") as file:
        file.write(vbs_script)

    #execute
    subprocess.call(['cscript.exe', str(vbs_script_path)])

    # remove
    vbs_script_path.unlink()

    return None