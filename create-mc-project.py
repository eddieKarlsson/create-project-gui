import os
import shutil
import json
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk
from editpyxl import Workbook
import re


class CreateProjGui:
    def __init__(self, root):

      # Define Excel file paths relative to the new folder
        self.excel_files_to_edit = {
            "Automation/Test document/TD YYMMDD.xlsx": {
                "project_number": "F1",
                "project_name": "F2",
                "sheet": "Overview"
            },
            "Automation/Manuals/AD YYMMDD.xlsx": {
                "project_number": "W1",
                "project_name": "W2",
                "sheet": "AD_Machine 1"

            },
            "Automation/YYxxxx WB.xlsx": {
                "project_number": "C1",
                "project_name": "C2",
                "sheet": "Overview"

            },
            "Project management/23xxxx PM V1.1.xlsx": {
                "project_number": "C1",
                "project_name": "C2",
                "sheet": "Overview"

            }

        }

        self.root = root
        self.root.title("Create project")
        self.user_info = self.load_user_info()

        self.setup_ui()

    def load_user_info(self):
        if os.path.exists("user_info.json"):
            with open("user_info.json", "r") as file:
                return json.load(file)
        return {}

    def save_user_info(self):
        with open("user_info.json", "w") as file:
            json.dump(self.user_info, file, indent=4)

    def browse_folder(self, entry):
        selected_folder = filedialog.askdirectory(title="Select a folder")
        if selected_folder:
            entry.delete(0, tk.END)
            entry.insert(0, selected_folder)

    def update_available_folders(self):
        project_template_path = self.project_template_path_entry.get()
        if project_template_path and os.path.exists(project_template_path):
            available_folders = [folder for folder in os.listdir(
                project_template_path) if os.path.isdir(os.path.join(project_template_path, folder))]
            self.selected_folder_var.set("")  # Clear the selected folder
            self.selected_folder_dropdown['values'] = available_folders
            self.user_info["project_template_path"] = project_template_path
            self.user_info["available_folders"] = available_folders
            self.save_user_info()
        else:
            print("Invalid project template path")

    def browse_and_update_template_path(self):
        self.browse_folder(self.project_template_path_entry)
        self.update_available_folders()

    def browse_and_update_path(self):
        self.browse_folder(self.destination_folder_entry)
        self.user_info["destination_folder"] = self.destination_folder_entry.get()
        self.save_user_info()

    def setup_ui(self):
        self.project_template_path_frame = tk.Frame(self.root)
        self.project_template_path_frame.pack(fill=tk.X, padx=15, pady=15)

        self.project_template_path_label = tk.Label(
            self.project_template_path_frame, text="Project Template Path:")
        self.project_template_path_label.pack(side=tk.LEFT)

        self.project_template_path_entry = tk.Entry(
            self.project_template_path_frame)
        self.project_template_path_entry.pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        self.project_template_path_entry.insert(
            0, self.user_info.get("project_template_path", ""))

        self.project_template_browse_button = tk.Button(
            self.project_template_path_frame, text="Browse", command=self.browse_and_update_template_path)
        self.project_template_browse_button.pack(side=tk.LEFT)

        self.update_template_button = tk.Button(
            self.root, text="Update Available Folders", command=self.update_available_folders)
        self.update_template_button.pack(pady=5)

        self.project_name_label = tk.Label(self.root, text="Project Name:")
        self.project_name_label.pack(anchor="w")

        self.project_name_entry = tk.Entry(self.root)
        self.project_name_entry.pack(fill=tk.X, padx=5, pady=2)
        self.project_name_entry.insert(
            0, self.user_info.get("project_name", ""))

        self.project_number_label = tk.Label(self.root, text="Project Number:")
        self.project_number_label.pack(anchor="w")

        self.project_number_entry = tk.Entry(self.root)
        self.project_number_entry.pack(fill=tk.X, padx=5, pady=2)
        self.project_number_entry.insert(
            0, self.user_info.get("project_number", ""))

        self.selected_folder_label = tk.Label(
            self.root, text="Select Project Folder Structure:")
        self.selected_folder_label.pack(anchor="w")

        available_folders = self.user_info.get("available_folders", [])
        self.selected_folder_var = tk.StringVar(
            value=self.user_info.get("selected_folder", ""))
        self.selected_folder_dropdown = ttk.Combobox(
            self.root, textvariable=self.selected_folder_var, values=available_folders)
        self.selected_folder_dropdown.pack(fill=tk.X, padx=5, pady=2)

        self.destination_folder_label = tk.Label(
            self.root, text="Destination Folder:")
        self.destination_folder_label.pack(anchor="w")

        self.destination_frame = tk.Frame(self.root)
        self.destination_frame.pack(fill=tk.X, padx=10)

        self.destination_folder_entry = tk.Entry(self.destination_frame)
        self.destination_folder_entry.pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        self.destination_folder_entry.insert(
            0, self.user_info.get("destination_folder", ""))

        self.browse_button = tk.Button(
            self.destination_frame, text="Browse", command=self.browse_and_update_path)
        self.browse_button.pack(side=tk.LEFT)

        self.run_button = tk.Button(self.root, text="Run", command=self.run)
        self.run_button.pack(pady=10)

    def edit_excel_files(self, new_folder_path):
        for rel_path, excel_data in self.excel_files_to_edit.items():
            # Get the path of the Excel file in the new directory
            excel_file_path = os.path.join(new_folder_path, rel_path)

            # Check if the Excel file exists
            if os.path.exists(excel_file_path):
                # Extract information from excel_data
                sheet_name = excel_data.get("sheet")

                # Get project number and project name from self.user_info
                project_number = self.user_info["project_number"]
                project_number_cell = excel_data.get("project_number")
                project_name = self.user_info["project_name"]
                project_name_cell = excel_data.get("project_name")

                try:
                    # Edit the Excel file using editpyxl
                    workbook = Workbook()
                    workbook.open(excel_file_path)
                    worksheet = workbook[sheet_name]
                    worksheet[project_number_cell] = project_number
                    worksheet[project_name_cell] = project_name
                    workbook.save(excel_file_path)
                    workbook.close()

                    print(f"Excel file updated: {excel_file_path}")
                except Exception as e:
                    print(
                        f"Error editing Excel file {excel_file_path}: {e}")
            else:
                print(f"Excel file not found: {excel_file_path}")

    def replace_projnr_in_filenames(self, new_folder_path):
        project_number = self.user_info["project_number"]

        # Define the patterns to search for in filenames
        patterns = ['YYxxxx', 'xxxxxx', '250xxx',
                    '240xxx', '230xxx', '220xxx', '25xxxx', '24xxxx', '23xxxx', '22xxxx']

        for root, dirs, files in os.walk(new_folder_path):
            for filename in files:
                for pattern in patterns:
                    if re.search(pattern, filename):
                        new_filename = filename.replace(
                            pattern, project_number)
                        file_path = os.path.join(root, filename)
                        new_file_path = os.path.join(root, new_filename)

                        try:
                            os.rename(file_path, new_file_path)
                            # print(f"File renamed: {
                            #      file_path} -> {new_file_path}")
                        except Exception as e:
                            print(f"Error renaming file {file_path}: {e}")

        print("**File renaming completed**")

    def run(self):
        self.user_info["project_name"] = self.project_name_entry.get()
        self.user_info["project_number"] = self.project_number_entry.get()
        self.user_info["selected_folder"] = self.selected_folder_var.get()
        self.user_info["destination_folder"] = self.destination_folder_entry.get()
        self.save_user_info()
        # self.root.destroy()

        selected_folder = self.user_info["selected_folder"]
        destination_folder = self.user_info["destination_folder"]

        if selected_folder and destination_folder:
            source_folder = os.path.join(
                self.user_info["project_template_path"], selected_folder)
            new_folder_path = os.path.join(
                destination_folder, self.user_info["project_name"])
            shutil.copytree(source_folder, new_folder_path)
            print("Content copied to:", new_folder_path)

            self.edit_excel_files(new_folder_path)
            self.replace_projnr_in_filenames(new_folder_path)
        print("**Completed**")


if __name__ == "__main__":
    root = tk.Tk()
    app = CreateProjGui(root)
    root.mainloop()
