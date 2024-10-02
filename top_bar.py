import tkinter as tk
import tkinter.messagebox as messagebox
import sys
import os
from admin import AdminPopup

import tkinter as tk
from admin import AdminPopup, PasswordDialog

class TopBar(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.create_widgets()

    def new_file(self):
        self.restart_application()

    def show_options(self):
        # Open the password entry dialog first
        password_dialog = PasswordDialog(self.parent, callback=self.open_admin_window)
        password_dialog.transient(self.parent)
        password_dialog.grab_set()
        self.parent.wait_window(password_dialog)

    def open_admin_window(self, entered_password):
        # Check the entered password and open admin window if it's correct
        if entered_password == "mario1710":  # Replace with your actual password check
            admin_popup = AdminPopup(self.parent)
            admin_popup.transient(self.parent)
            admin_popup.grab_set()
            self.parent.wait_window(admin_popup)
        else:
            tk.messagebox.showwarning("Incorrect Password", "Invalid password. Access denied.")

    def show_help(self):
        tk.messagebox.showinfo("Help", "This Application was created for W&M by Ali El Fekki.")

    def restart_application(self):
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def create_widgets(self):
        menu_bar = tk.Menu(self.parent)
        self.parent.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.parent.destroy)

        options_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Options", menu=options_menu)
        options_menu.add_command(label="Admin", command=self.show_options)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_help)

if __name__ == "__main__":
    root = tk.Tk()
    top_bar = TopBar(root)
    top_bar.pack(side=tk.TOP, fill=tk.X)
    root.mainloop()
