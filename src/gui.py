# Local Imports
import os
import sys
import random
import subprocess
# External imports
from datetime import datetime as dt

from PyQt5.QtCore import (QThread,
                          pyqtSignal)
from PyQt5.QtGui import (QIcon,
                         QPixmap)
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
)
# Local Imports
from synchro_folder import create_file_dict


class ScriptRunner(QThread):
    finished: pyqtSignal = pyqtSignal()

    def __init__(self, command: str) -> None:
        super().__init__()

        self.command = command

    def run(self) -> None:
        process = subprocess.run(
            self.command, shell=True, capture_output=True, text=True
        )
        print(process.stdout)
        if process.stderr:
            print(process.stderr)
        self.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VBA CARD Scheduler")
        self.setGeometry(100, 100, 470, 330)

        # Create a pixmap from an image file
        pixmap = QPixmap("N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\src\logo.jpg")
        # Create an icon from the pixmap
        icon = QIcon(pixmap)

        # Set the icon as the window icon
        self.setWindowIcon(icon)

        self.setStyleSheet("""QFrame { border-radius: 10px;}""")

        # Set UI elements
        self.setStyleSheet("QMainWindow {background-color: #287444;}")

        # Labels
        self.select_json_label = QLabel("If first time updating:", self)
        self.select_json_label.move(170, 30)
        self.select_json_label.setStyleSheet("background-color: #287444;color: white;")

        self.select_json_label = QLabel("If json already exist:", self)
        self.select_json_label.move(170, 90)
        self.select_json_label.setStyleSheet("background-color: #287444;color: white;")

        self.macro_status_label = QLabel("Enable Macro:", self)
        self.macro_status_label.move(20, 30)
        self.macro_status_label.setStyleSheet("background-color: #287444;color: white;")

        self.action_type_label = QLabel("Action Type", self)
        self.action_type_label.move(20, 90)
        self.action_type_label.setStyleSheet("background-color: #287444;color: white;")

        self.action_type_label = QLabel("Execution Type", self)
        self.action_type_label.move(20, 150)
        self.action_type_label.setStyleSheet("background-color: #287444;color: white;")

        self.select_json_label = QLabel("Frequency:", self)
        self.select_json_label.move(170, 150)
        self.select_json_label.setStyleSheet("background-color: #287444;color: white;")

        # Buttons
        self.browse_json_button = QPushButton("Browse", self)
        self.browse_json_button.move(170, 120)
        self.browse_json_button.setStyleSheet("background-color: white ;color: black;")
        self.browse_json_button.clicked.connect(self.choose_json)

        self.exx = QPushButton("Execute", self)
        self.exx.move(320, 260)
        self.saving_option = QComboBox(self)
        self.saving_option.addItems(["Run", "Schedule"])
        self.saving_option.move(20, 180)
        self.saving_option.setStyleSheet("background-color: white ;color: black;")

        self.sync_folder_button = QPushButton("Synchronise folder", self)
        self.sync_folder_button.move(170, 60)
        self.sync_folder_button.setStyleSheet("background-color: white ;color: black;")
        self.sync_folder_button.clicked.connect(self.choose_xlsm)

        # Combo Boxes
        self.switch_macro_filter = QComboBox(self)
        self.switch_macro_filter.addItems(["ON", "OFF"])
        self.switch_macro_filter.move(20, 60)
        self.switch_macro_filter.setStyleSheet("background-color: white ;color: black;")

        self.timeframe = QComboBox(self)
        self.timeframe.addItems(["daily", "monthly", "weekly"])
        self.timeframe.move(170, 180)
        self.timeframe.setStyleSheet("background-color: white ;color: black;")

        self.action_type_combo_box = QComboBox(self)
        self.action_type_combo_box.addItems(["Inject Macro", "Terminate Excel", "API Refresh"])
        self.action_type_combo_box.move(20, 120)
        self.action_type_combo_box.setStyleSheet("background-color: white ;color: black;")

        self.timeframe.currentTextChanged.connect(self.on_date_type_changed)

        self.exx.setStyleSheet("background-color: white ;color: black;")
        self.exx.clicked.connect(self.run_script)

        # tick box
        self.holiday = QCheckBox('Run on Holiday', self)
        self.holiday.move(20, 220)
        self.holiday.setStyleSheet("background-color: #287444;color: white;")

        self.kill = QCheckBox('Close Excel after Run', self)
        self.kill.move(20, 240)
        self.kill.setStyleSheet("background-color: #287444; color: white;")
        self.kill.setFixedSize(130, 30)





        self.max_label = QLabel('Max execution per file (m):', self)
        self.max_label.move(320, 90)
        self.max_label.setStyleSheet("color: white;")
        self.max_label.setFixedSize(130, 30)

        self.max_dropdown = QComboBox(self)
        self.max_dropdown.move(320, 120)
        self.max_dropdown.addItems([str(i) for i in range(2, 31)])
        self.max_dropdown.setCurrentText('15')
        self.max_dropdown.setFixedSize(100, 30)
        self.max_dropdown.setStyleSheet("background-color: white ;color: black;")

        self.op_label = QLabel('Opening Buffer (S):', self)
        self.op_label.move(320, 150)
        self.op_label.setStyleSheet("color: white;")
        self.op_label.setFixedSize(130, 30)

        self.op_dropdown = QComboBox(self)
        self.op_dropdown.move(320, 180)
        self.op_dropdown.addItems(['10','60','120', '180', '240','300','600','900'])
        self.op_dropdown.setCurrentText('120')
        self.op_dropdown.setFixedSize(100, 30)
        self.op_dropdown.setStyleSheet("background-color: white ;color: black;")

        self.inter_label = QLabel('Inter-operation buffer (S):', self)
        self.inter_label.move(320, 30)
        self.inter_label.setStyleSheet("color: white;")
        self.inter_label.setFixedSize(130, 30)


        self.inter_dropdown = QComboBox(self)
        self.inter_dropdown.move(320, 60)
        self.inter_dropdown.addItems(['5', '10', '20', '30', '40', '50', '60'])
        self.inter_dropdown.setCurrentText('20')
        self.inter_dropdown.setFixedSize(100, 30)
        self.inter_dropdown.setStyleSheet("background-color: white ;color: black;")



        # Variables

        self.script_type = ""
        self.json_file_path = ""
        self.xlsm_file_path = ""

    def on_date_type_changed(self, text):
        if text == "weekly":
            # Display a message box asking for the user's input
            day_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            input_text, ok_pressed = QInputDialog.getItem(self, "Select a day", "Day of the week:", day_list, 0, False)

            # Check if the user pressed OK
            if ok_pressed:
                # Store the selected day in a variable
                selected_day = str(input_text)
                msg_box = QMessageBox()
                msg_box.setText(f"Selected day: {selected_day}")
                msg_box.exec_()
            else:
                selected_day = None

        elif text == "monthly":
            # Get the current year and month
            current_year = dt.now().year
            current_month = dt.now().month

            # Get the last day of the current month
            if current_month == 12:
                last_day = 31
            else:
                last_day = (dt(current_year, current_month + 1, 1) - dt(current_year, current_month, 1)).days

            # Display a message box asking for the user's input
            input_text, ok_pressed = QInputDialog.getText(self, "Enter a number", f"Number 1 - {last_day}:")

            # Check if the user pressed OK and if the input is a number between 1 and the last day of the month
            if ok_pressed and input_text.isdigit() and 1 <= int(input_text) <= last_day:
                # Store the selected day in a variableF
                selected_day = int(input_text)
                msg_box = QMessageBox()
                msg_box.setText(f"Selected day: {selected_day}")
                msg_box.exec_()
            else:
                # Display an error message if the input is not valid
                msg_box = QMessageBox()
                msg_box.setText(f"Invalid input. Please enter a number between 1 and {last_day}.")
                msg_box.exec_()
                selected_day = None
        elif text == "daily":
            msg_box = QMessageBox()
            msg_box.setText(f"You are running daily tasks")
            msg_box.exec_()
            selected_day = None
        if selected_day is not None:
            self.selected_day = selected_day
        else:
            self.selected_day = ''

    def choose_json(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        initial_dir = "N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main/data/jsons"
        file_dialogq = QFileDialog()
        file_dialogq.setOptions(options)
        file_dialogq.setDirectory(initial_dir)
        file_dialogq.setNameFilter("JSON Files (*.json)")
        if file_dialogq.exec_():
            self.json_file_path = file_dialogq.selectedFiles()[0]

    def choose_xlsm(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        initial_dir = "N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main/data"
        folder_dialog = QFileDialog()
        folder_dialog.setOptions(options)
        folder_dialog.setDirectory(initial_dir)
        folder_path = folder_dialog.getExistingDirectory()
        if folder_path:
            create_file_dict(folder_path, f"N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\jsons\{folder_path.split('/')[-1]}.json")
            QMessageBox.information(self, "Done", "Script execution is completed.")
            print(os.path.join(folder_path, 'test.json'))
            self.json_file_path = f"N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\jsons\{folder_path.split('/')[-1]}.json"

    def run_script(self):

        ex = self.action_type_combo_box.currentText()
        time_frame = self.timeframe.currentText()
        save = self.saving_option.currentText()



        if ex == "Terminate Excel":

            self.json_file_path = r'N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\data\do_not_delete.JSON'

        if not self.json_file_path:
            print("Invalid JSON file path")
            QMessageBox.information(self, "Close", "Invalid JSON file path")
            return

        if time_frame == '':
            print("Script type is required")
            return

        script_type = time_frame
        print('test')


        arguments = ["python", "N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\src\scheduler.py",
                     "--json_file_path", self.json_file_path, "--script_type", script_type]

        try:
            if isinstance(self.selected_day, int) and '--trigger_day' not in arguments:
                arguments.append(f"--trigger_date {self.selected_day}")
            elif isinstance(self.selected_day, str) and '--trigger_day' not in arguments:
                arguments.append(f"--trigger_day {self.selected_day}")
        except:
            print('daily')
        if self.kill.isChecked():
            arguments.append(f"--kill_process {self.kill.isChecked()}")
        if ex == "Inject Macro":
            arguments.append("--execute-macro")

        macro_switch = self.switch_macro_filter.currentText()
        if macro_switch:
            arguments.extend(["--macro_switch", macro_switch])

        if ex == "API Refresh":
            arguments.append("--main-automation")

        if ex == "Terminate Excel":
            arguments.append("--terminate")

        if self.holiday.isChecked():
            arguments.append(f"--skip_wk_hol {self.holiday.isChecked()}")

        arguments.append(f"--max_exe {str(self.max_dropdown.currentText())}")


        arguments.append(f"--opening_buffer {str(self.op_dropdown.currentText())}")

        arguments.append(f"--inter_op_buffer {str(self.inter_dropdown.currentText())}")


        print(arguments)




        if save == 'Schedule':

            src = self.json_file_path.split("/")[-1].replace('.json','')
            print(src)

            file_path = f'N:\Research\AUTO\VBA_CARD-main\AUTO-RUG\AUTO-main\src\{src}.bat'
            with open(file_path, 'w') as file:
                file.write(' '.join(arguments))
            QMessageBox.information(self, "Done", "Script execution is running.")
            QMessageBox.information(self, "Done", "Script execution is completed.")
        elif save == 'Run':

            command = ' '.join(arguments)

            self.runner = ScriptRunner(command)

            self.runner.started.connect(lambda: print("Running script..."))
            QMessageBox.information(self, "Done", "Script execution is running.")
            self.runner.finished.connect(self.on_finished)
            self.runner.start()






    def on_finished(self):
        print("Done!")
        self.runner.finished.connect(self.on_finished)
        QMessageBox.information(self, "Close", "Script execution is completed.")
        self.json_file_path=''
        '''if self.kill.isChecked() == True:
            subprocess.call("taskkill /f /im excel.exe", shell=True)'''


        #QApplication.quit()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())