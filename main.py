import win32com
import streamlit as st
import os
import helpers

def main():

    st.title(":closed_lock_with_key: Password Manager")
    path_to_excel_files = st.text_input("Insert the path of the folder in which your excel files reside")
    choice_action = st.sidebar.radio("What do you want do do?", ['Set Password', 'Remove Workbook Protection'])
    password_string = st.text_input("Type the password here")

    if path_to_excel_files and password_string and st.checkbox("Run the program"):

        # Get list of files in folder
        files = os.listdir(path_to_excel_files)            
        
        # Look for excel files 
        files_xls = [f for f in files if f[-4:] == 'xlsx' or f[-3:] == 'xls' or f[-4:] == 'xlsm']
        
        for file in files_xls:
            # Append filename to directory path
            file = path_to_excel_files + "\\" + file
            
            if choice_action == 'Remove Workbook Protection':
                try: 
                    # Remove password using function from helpers script
                    helpers.remove_password_xlsx(file, password_string)
                    st.success(f"The password has been successfully removed from {file}")
                except:
                    st.error("This is not the right password")
                    st.stop()

            elif choice_action == 'Set Password':
                helpers.set_password(file, password_string)
                st.success("You have successfully set the password")

if __name__ == "__main__":
    main()