import os
import subprocess
from datetime import date, datetime, timezone, timedelta
import holidays
import argparse
import time
import subprocess
import psutil
from run_automated_process import run_file_links

import subprocess
import os
import time
from typing import List, Any

from typing import List


def levenshtein_distance(s1: str, s2: str) -> int:
    # Swap the strings to ensure that s1 is always longer than s2
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)

    # If s2 is empty, then the Levenshtein distance is the length of s1
    if len(s2) == 0:
        return len(s1)

    # Initialize the previous row to [0, 1, 2, ..., len(s2)]
    previous_row: List[int] = list(range(len(s2) + 1))

    # Iterate over each character in s1
    for i, c1 in enumerate(s1):
        # Initialize the current row to [i+1]
        current_row: List[int] = [i + 1]

        # Iterate over each character in s2
        for j, c2 in enumerate(s2):
            # Calculate the cost of inserting c2 into s1
            insertions: int = previous_row[j + 1] + 1

            # Calculate the cost of deleting c1 from s1
            deletions: int = current_row[j] + 1

            # Calculate the cost of substituting c1 with c2
            substitutions: int = previous_row[j] + (c1 != c2)

            # Append the minimum cost to the current row
            current_row.append(min(insertions, deletions, substitutions))

        # Set the previous row to the current row
        previous_row = current_row

    # Return the last element in the previous row
    return previous_row[-1]



def switch_macros(SWITCH: str) -> None:
    if SWITCH == 'OFF':
        refresh_data_and_perform_nads_file = r'N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\macros\RefreshDataAndPerformNADS.bas'
        refresh_data_and_perform_nadst_file = r'N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\macros\RefreshDataAndPerformNADST.bas'
        _replace_string_in_file(refresh_data_and_perform_nads_file, 16, "    Call BR3AK")
        _replace_string_in_file(refresh_data_and_perform_nadst_file, 16, "    Call BR3AK")
    elif SWITCH == 'ON':
        refresh_data_and_perform_nads_file = r'N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\macros\RefreshDataAndPerformNADS.bas'
        refresh_data_and_perform_nadst_file = r'N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\macros\RefreshDataAndPerformNADST.bas'
        _replace_string_in_file(refresh_data_and_perform_nads_file, 16, "    'Call BR3AK")
        _replace_string_in_file(refresh_data_and_perform_nadst_file, 16, "    'Call BR3AK")


def _replace_string_in_file(file_path: str, dig:int, new_string: str) -> None:
    with open(file_path, 'r') as file:
        lines = file.readlines()
    with open(file_path, 'w') as file:
        for i, line in enumerate(lines):
            if i != dig-1:
                file.write(line)
            else:
                file.write(new_string + '\n')


def run_automation(json_file_path: str, arguments: any) -> None:
    """
    Opens an Excel file using `os.startfile`, waits for a cooldown period,
    and executes a macro if `execute_macro` or `main_automation` is provided.
    If the `arguments` parameter includes the `terminate` argument (bool), it terminates the Excel process.

    Parameters:
        json_file_path (str): The path to the JSON file.
        arguments (argparse.Namespace): The arguments for the function.

    Returns:
        None.

    Raises:
        None.
    """
    # If the 'terminate' argument is provided, kill the Excel process and print a message.
    if arguments.terminate:
        subprocess.call("taskkill /f /im excel.exe", shell=True)
        print('Excel process terminated')
        return

#



    # Use os.startfile to open the Excel file.
    os.startfile(r"N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\init.xlsm")
    print('Excel process launched')

    # Print a message that the cooldown period has started and sleep for 20 seconds.
    print('Cool down period started')
    time.sleep(20)
    print('Cool down period finished')
    os.startfile(r"N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\_On-Start\On_Start_AUTO_ds.xlsm")
    time.sleep(10)

    # If the 'execute_macro' argument is provided, execute the macro with 'should_execute_macro=True'.
    if arguments.execute_macro:
        switch_macros(SWITCH=str(arguments.macro_switch))
        print(arguments.macro_switch)
        run_file_links(arguments.max_exe, arguments.opening_buffer, arguments.inter_op_buffer,
                       json_file_path=json_file_path, should_execute_macro=True)

    # If the 'main_automation' argument is provided, execute the macro with 'should_execute_macro=False'.
    elif arguments.main_automation:
        run_file_links(arguments.max_exe, arguments.opening_buffer, arguments.inter_op_buffer,
                       json_file_path=json_file_path, should_execute_macro=False)

    # Otherwise, print a message that no argument was provided and exit the script.
    else:
        print('No argument provided, exiting script...')






def run_script(uk_holidays: holidays.HolidayBase,
               json_file_path: str,
               script_type: str,
               args: dict) -> None:
    """
    Runs a script based on the `script_type` parameter, taking into account the `trigger_day` parameter for monthly scripts,
    and checking against UK holidays.

    Parameters:
        uk_holidays (holidays.HolidayBase): A holiday calendar for the UK.
        json_file_path (str): The path to the JSON file.
        script_type (str): The type of script to run (e.g. "monthly", "daily", "weekly").
        args (dict): A dictionary of arguments for the script.

    Returns:
        None.

    Raises:
        None.
    """
    now = datetime.now(timezone.utc)
    today = now.date()
    weekday = now.weekday()
    trigger_date = args.trigger_date

    json_file_path=args.json_file_path
    trigger_day = str(args.trigger_day).lower()

    weekdays = {
        'monday': 0,
        'tuesday': 1,
        'wednesday': 2,
        'thursday': 3,
        'friday': 4,
        'saturday': 5,
        'sunday': 6
    }



    closest_match = min(weekdays.keys(), key=lambda x: levenshtein_distance(x, trigger_day))

    trigger_day = weekdays[closest_match]

    if script_type == "monthly":
        if args.skip_wk_hol==False:
            if int(now.day) == int(trigger_date) and today not in uk_holidays and weekday < 5:
                run_automation(json_file_path, args)
        elif int(now.day) == int(trigger_date):
            run_automation(json_file_path, args)


    elif script_type == "daily":
        if args.skip_wk_hol == False:
            if today not in uk_holidays and weekday < 5:
                run_automation(json_file_path, args)
        else: run_automation(json_file_path, args)

    elif script_type == "weekly":
        if args.skip_wk_hol == False:
            if today not in uk_holidays and weekday == int(trigger_day):
                run_automation(json_file_path, args)
        elif weekday == int(trigger_day):
            run_automation(json_file_path, args)

    if args.kill_process:
        print(f'schedulerdebugger{args.kill_process }')
        subprocess.call("taskkill /f /im excel.exe", shell=True)



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Automatisation Script')
    parser.add_argument('--terminate', action='store_true', help='Terminate Excel process')
    parser.add_argument('--main-automation', action='store_true', help='Run main automation')
    parser.add_argument('--execute-macro', action='store_true', help='Execute macro')
    parser.add_argument('--macro_switch', choices=['ON', 'OFF'], help='Specify whether to turn macros "on" or "off".')
    parser.add_argument('--json_file_path', required=False, help='Path to JSON file')
    parser.add_argument('--script_type', required=True, choices=['monthly', 'weekly', 'daily'],
                        help='monthly, weekly or daily')
    parser.add_argument('--skip_wk_hol',help='run on holiday and weekends')
    parser.add_argument('--trigger_date', help='Trigger date for monthly script (month day)')
    parser.add_argument('--trigger_day', help='Trigger day for weekly script weekday')
    parser.add_argument('--kill_process', help='kill process')

    parser.add_argument('--max_exe', help='max exe time')
    parser.add_argument('--opening_buffer', help='opening buffer')
    parser.add_argument('--inter_op_buffer', help='operations buffer')


    args = parser.parse_args()
    print(args)
    uk_holidays: holidays.HolidayBase = holidays.UnitedKingdom()
    run_script(uk_holidays, args.json_file_path, args.script_type, args)
