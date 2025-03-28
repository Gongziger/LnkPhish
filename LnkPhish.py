import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client

class ShortcutGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shortcut Generator Tool")
        self.root.geometry("600x350")
        self.root.resizable(True, True)
        self.create_widgets()

    def create_widgets(self):
        self.shortcut_name_label = tk.Label(self.root, text="Shortcut Name")
        self.shortcut_name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.shortcut_name_entry = tk.Entry(self.root, width=40)
        self.shortcut_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.icon_choice_var = tk.IntVar()
        self.icon_choice_var.set(1)

        self.default_icon_radio = tk.Radiobutton(self.root, text="Use Default Icon", variable=self.icon_choice_var, value=1, command=self.update_icon_input)
        self.default_icon_radio.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.custom_icon_radio = tk.Radiobutton(self.root, text="Choose Custom Icon", variable=self.icon_choice_var, value=2, command=self.update_icon_input)
        self.custom_icon_radio.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        self.icon_combobox = tk.StringVar(self.root)
        self.icon_combobox.set("PDF")
        self.icon_dropdown = tk.OptionMenu(self.root, self.icon_combobox, "PDF", "Excel", "Word", "Text", "PowerPoint", "HTML", "Image", "Video", "Java", "Archive")
        self.icon_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.custom_icon_button = tk.Button(self.root, text="Choose Custom Icon", command=self.select_icon)
        self.custom_icon_button.grid(row=2, column=1, padx=5, pady=5)
        self.custom_icon_button.config(state=tk.DISABLED)

        self.custom_icon_entry = tk.Entry(self.root, width=40)
        self.custom_icon_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.custom_icon_entry.config(state=tk.DISABLED)

        self.separator_line = tk.Frame(self.root, height=2, bd=1, relief="solid")
        self.separator_line.grid(row=4, column=0, columnspan=3, pady=10, padx=5, sticky="ew")

        self.file_choice_var = tk.IntVar()
        self.file_choice_var.set(1)

        self.local_file_radio = tk.Radiobutton(self.root, text="Choose Local File", variable=self.file_choice_var, value=1, command=self.update_input_fields)
        self.local_file_radio.grid(row=5, column=0, padx=5, pady=5, sticky="w")

        self.file_entry = tk.Entry(self.root, width=40)
        self.file_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

        self.browse_button = tk.Button(self.root, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=5, column=2, padx=5, pady=5)

        self.remote_link_radio = tk.Radiobutton(self.root, text="Enter Remote Link", variable=self.file_choice_var, value=2, command=self.update_input_fields)
        self.remote_link_radio.grid(row=6, column=0, padx=5, pady=5, sticky="w")

        self.link_entry = tk.Entry(self.root, width=40)
        self.link_entry.grid(row=6, column=1, padx=5, pady=5, sticky="ew")
        self.link_entry.config(state=tk.DISABLED)

        self.parameters_label = tk.Label(self.root, text="Execution Parameters")
        self.parameters_label.grid(row=7, column=0, padx=5, pady=5, sticky="w")

        self.parameters_entry = tk.Entry(self.root, width=40)
        self.parameters_entry.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

        self.save_button = tk.Button(self.root, text="Save Shortcut", command=self.save_shortcut)
        self.save_button.grid(row=8, column=0, columnspan=3, padx=5, pady=10, sticky="ew")
        self.save_button.config(state=tk.DISABLED)

        self.selected_icon = None
        self.selected_file = None
        self.selected_link = None

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(8, weight=1)

    def update_input_fields(self):
        if self.file_choice_var.get() == 1:
            self.file_entry.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.NORMAL)
            self.link_entry.config(state=tk.DISABLED)
        else:
            self.file_entry.config(state=tk.DISABLED)
            self.browse_button.config(state=tk.DISABLED)
            self.link_entry.config(state=tk.NORMAL)

        self.check_save_button_state()

    def update_icon_input(self):
        if self.icon_choice_var.get() == 1:
            self.icon_dropdown.config(state=tk.NORMAL)
            self.custom_icon_button.config(state=tk.DISABLED)
            self.custom_icon_entry.config(state=tk.DISABLED)
        else:
            self.icon_dropdown.config(state=tk.DISABLED)
            self.custom_icon_button.config(state=tk.NORMAL)
            self.custom_icon_entry.config(state=tk.NORMAL)

        self.check_save_button_state()

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.selected_file = file_path

        self.check_save_button_state()

    def select_icon(self):
        icon_path = filedialog.askopenfilename(filetypes=[("Icon Files", "*.ico"), ("All Files", "*.*")])
        if icon_path:
            self.custom_icon_entry.delete(0, tk.END)
            self.custom_icon_entry.insert(0, icon_path)
            self.selected_icon = icon_path

        self.check_save_button_state()

    def check_save_button_state(self):
        shortcut_name = self.shortcut_name_entry.get()
        if self.file_choice_var.get() == 1:
            self.selected_file = self.file_entry.get()
            self.selected_link = None
        else:
            self.selected_link = self.link_entry.get()
            self.selected_file = None

        if shortcut_name and (self.selected_file or self.selected_link):
            self.save_button.config(state=tk.NORMAL)
        else:
            self.save_button.config(state=tk.DISABLED)

    def save_shortcut(self):
        shortcut_name = self.shortcut_name_entry.get()
        parameters = self.parameters_entry.get()
        icon_selection = self.icon_combobox.get()

        if not shortcut_name:
            messagebox.showerror("Error", "Please enter shortcut name")
            return

        if not self.selected_file and not self.selected_link:
            messagebox.showerror("Error", "Please select a file or enter a remote link")
            return

        if self.selected_icon:
            icon_name = os.path.basename(self.selected_icon)
            exe_folder = os.path.join(os.getcwd(), f"{shortcut_name}_Folder", "$Recycle.Bin")
            os.makedirs(exe_folder, exist_ok=True)
            icon_copy_path = os.path.join(exe_folder, icon_name)

            try:
                shutil.copy(self.selected_icon, icon_copy_path)
            except Exception as e:
                print(f"Icon copy failed: {e}")

            icon_dest_path = os.path.join(exe_folder, icon_name)

        else:
            icon_dest_path = self.get_default_icon_path()

        target_folder = os.path.join(os.getcwd(), f"{shortcut_name}_Folder")
        os.makedirs(target_folder, exist_ok=True)

        exe_folder = os.path.join(target_folder, "$Recycle.Bin")
        os.makedirs(exe_folder, exist_ok=True)

        exe_dest_path = os.path.join(exe_folder, f"{shortcut_name}.exe")
        shutil.copy(self.selected_file, exe_dest_path)

        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(os.path.join(target_folder, f"{shortcut_name}.lnk"))

            shortcut.TargetPath = exe_dest_path
            shortcut.IconLocation = icon_dest_path

            if parameters:
                shortcut.Arguments = parameters

            shortcut.Save()

            self.set_folder_hidden(exe_folder)

            messagebox.showinfo("Success", f"Shortcut saved to: {target_folder}\\{shortcut_name}.lnk")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create shortcut: {str(e)}")

    def set_folder_hidden(self, folder_path):
        try:
            os.system(f'attrib +h "{folder_path}"')
        except Exception as e:
            print(f"Failed to hide folder {folder_path}: {str(e)}")

    def get_default_icon_path(self):
        icons = {
            "PDF": "icons/pdf.ico",
            "Excel": "icons/excel.ico",
            "Word": "icons/word.ico",
            "Text": "icons/text.ico",
            "PowerPoint": "icons/powerpoint.ico",
            "HTML": "icons/html.ico",
            "Image": "icons/image.ico",
            "Video": "icons/video.ico",
            "Java": "icons/java.ico",
            "Archive": "icons/archive.ico"
        }

        selected_icon = self.icon_combobox.get()

        current_directory = os.getcwd()
        icon_path = os.path.join(current_directory, icons.get(selected_icon, "icons/pdf.ico"))

        return icon_path


if __name__ == "__main__":
    root = tk.Tk()
    app = ShortcutGeneratorApp(root)
    root.mainloop()
