from tkinter import simpledialog, messagebox, Tk, Entry, Label, filedialog
import os

class CustomDialog(simpledialog.Dialog):
    def __init__(self, parent, title, prompt, width):
        self.prompt = prompt
        self.width = width
        simpledialog.Dialog.__init__(self, parent, title)

    def body(self, master):
        Label(master, text=self.prompt).grid(row=0, column=0, columnspan=2)
        self.entry = Entry(master, width=self.width)
        self.entry.grid(row=1, column=0, columnspan=2)

        return self.entry  # return the entry widget

    def apply(self):
        self.result = self.entry.get()

def get_project_name(parent=None):
    while True:
        # Ask the user to select the directory for saving the file
        directory = filedialog.askdirectory(title="Select Save Directory")

        if not directory:
            return None, None  # User canceled

        # Ask the user for the file name separately
        dialog = CustomDialog(parent, "File Name", "Enter Excel Name:", width=40)
        project_name = dialog.result

        if not project_name:
            return None, None  # User canceled or provided an empty project name

        # Construct the full file path
        file_path = os.path.join(directory, f"{project_name}.xlsx")

        # Check if the Excel file already exists
        if os.path.isfile(file_path):
            messagebox.showerror("Error", "File with the same name already exists. Please choose a different name.")
        else:
            return project_name, file_path