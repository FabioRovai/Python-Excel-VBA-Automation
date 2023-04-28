import csv
import json
import os
import time
import win32api
import win32con
# do not cancel import below
import comtypes.client
from comtypes.client import CreateObject
from datetime import datetime, timedelta
from typing import List, Optional
from workbook_overwrite import ModifyExcelWorkbook
import pandas as pd
import os
import win32api
import win32gui


def enable_vba_access() -> None:
    registry_key = win32api.RegOpenKeyEx(
        win32con.HKEY_CURRENT_USER,
        "Software\\Microsoft\\Office\\16.0\\Excel\\Security",
        0, win32con.KEY_ALL_ACCESS
    )
    win32api.RegSetValueEx(
        registry_key, "AccessVBOM", 0, win32con.REG_DWORD, 1
    )
    print('access obtained')


def open_excel_workbook(path_to_wb: str) -> tuple:
    with ModifyExcelWorkbook(path_to_wb) as excel:
        excel.modify_macro_code()
    print('object created')
    excel = CreateObject("Excel.Application")

    try:
        excel.EnableEvents = False
        print('event not enabled')
    except:
        excel.EnableEvents = True
        print('event enabled')
    print('workbook opened')
    time.sleep(20)
    wb = excel.Workbooks.Open(path_to_wb)
    print('excel visible')
    excel.Visible = True
    return excel, wb


def remove_vba_macros(wb) -> None:
    print('remove old macros')
    for vb_component in wb.VBProject.VBComponents:
        if vb_component.Type in [1, 2, 3]:
            try:
                wb.VBProject.VBComponents.Remove(vb_component)
            except:
                None


def add_vba_macros(excel, wb, needed_modules: List[str]) -> None:
    print('declaring modules')
    modules = needed_modules
    print('declaring path')
    base_path = r"N:\path"
    for module in modules:
        print('adding macros')
        print('module')
        vba_module = wb.VBProject.VBComponents.Import(base_path + "\\" + module)

    print('done')

    # Turn off Excel alerts and display alerts
    excel.DisplayAlerts = False
    excel.AlertBeforeOverwriting = False

    try:
        excel.DisplayAlerts = False
        excel.AlertBeforeOverwriting = False
        wb.SaveAs(wb.FullName, FileFormat=wb.FileFormat)
        print('saved')
    except:
        None
    wb.Close()
    print('closed')
    excel.Quit()
    print('Quit')


def execute_vba_macros(path_to_wb: str, needed_modules: List[str], should_execute: Optional[str] = 'y') -> None:
    if should_execute == 'y':
        enable_vba_access()
        excel, wb = open_excel_workbook(path_to_wb)
        remove_vba_macros(wb)
        add_vba_macros(excel, wb, needed_modules)
    else:
        return None


def run_file_links(max_exe, opening_buffer, inter_op_buffer, json_file_path: str, should_execute_macro: bool = True) -> None:
    needed_modules = ["UpdateAllStaticDataAndRefreshBBGWorkbook.bas", "AutoSaveAndClose.bas", "AutoRefreshSheet.bas",
                      "CloseAllWorkbooks.bas", 'StartProcessingRealTime.bas', 'fill_freq_list_combobox.bas',
                      'RTExcelAddInEvents.bas', 'RefreshDataAndPerformNADS.bas', 'todaydate.bas', 'removedata.bas',
                      'RefreshDataAndPerformNADST.bas']

    timeout = timedelta(minutes=int(max_exe))
    with open(json_file_path, "r") as file:
        file_paths = json.load(file)
        print(json_file_path)

        for file_name, file_path in file_paths.items():
            print(file_name, file_path)
            if file_name.endswith('.xlsm'):
                last_folder=file_path.split('\\')[0].split('/')[-1]
                print(last_folder)
                status_path = fr"N:path\run_status_{last_folder}.csv"
                print()



    # Create the directory if it doesn't already exist
    directory = os.path.dirname(status_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Create the file if it doesn't already exist
    if not os.path.isfile(status_path):
        with open(status_path, 'w+') as file:
            file.write('')  # write an empty string to create the file



    with open(json_file_path, "r") as file, open(status_path, "a", newline="") as csvfile:
        writer = csv.writer(csvfile)
        file_paths = json.load(file)

        for file_name, file_path in file_paths.items():
            if file_name.endswith('.xlsm'):
                print(f"Working with file: {file_path}")
                if should_execute_macro:
                    try:
                        execute_vba_macros(file_path, needed_modules, should_execute="y")
                        '''writer.writerow([
                            file_name,
                            "Macro Updated",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ])'''
                    except Exception as e:
                        execute_vba_macros(file_path, needed_modules, should_execute="n")
                        print(f"Macro Failed: {e}")
                        '''writer.writerow([
                            file_name,
                            "Macro Not Updated",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ])'''
                else:
                    cooldown_seconds = int(opening_buffer) if "dst" in file_path else int(opening_buffer)
                    start_time = datetime.now()

                    while (datetime.now() - start_time) < timeout:
                        try:

                            os.startfile(file_path)
                            time.sleep(cooldown_seconds)
                            writer.writerow([file_name,
                                "Scraping Successful",
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            ])
                            break
                        except Exception as e:
                            print(f"Failed to open file: {e}")
                            writer.writerow([file_name,
                                "Scraping Failure",
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            ])
                            break
                    else:
                        writer.writerow([
                            file_name,
                            "Timed out",
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ])



                time.sleep(int(inter_op_buffer))
    import pandas as pd

    # Read the file into a pandas dataframe
    df = pd.read_csv(fr"N:path\run_status_{last_folder}.csv",header=None,
                     names=['file_name','API_status','API_date'])
    df = df.drop_duplicates(subset=['file_name'], keep='last')
    df_filtered=df
    df_filtered.reset_index()

    df_filtered.to_csv(fr"N:path\run_status_{last_folder}.csv", header=False, index=False)



if __name__ == "__main__":

    print('pass')

