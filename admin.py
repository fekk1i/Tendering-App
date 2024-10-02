import json
import os
import customtkinter as ctk
import tkinter as tk
import datetime
import tkinter.messagebox as messagebox
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend

class PasswordDialog(tk.Toplevel):
    def __init__(self, master=None, callback=None):
        super().__init__(master)
        self.title("Password Entry")
        self.password_var = tk.StringVar()
        self.create_widgets(callback)

        self.configure(bg='#242424')
        self.geometry('300x200')

    def create_widgets(self, callback):
        password_label = ctk.CTkLabel(self, text="Enter Password:")
        password_entry = ctk.CTkEntry(self, show="*", textvariable=self.password_var)
        password_label.pack(pady=10)
        password_entry.pack(pady=10)

        submit_button = ctk.CTkButton(self, text="Submit", command=lambda: self.on_submit(callback))
        submit_button.pack(pady=10)

    def on_submit(self, callback):
        entered_password = self.password_var.get()
        self.destroy()  # Close the password entry dialog
        callback(entered_password)

class AdminPopup(tk.Toplevel):
    CONFIG_FILE = "config_encrypted.json"
    STATIC_KEY = b'StaticKey16Bytes'  # Replace with your own static key (16 bytes)

    def __init__(self, master=None):
        super().__init__(master)
        self.title("Admin Window")
        self.usd_price_var = tk.DoubleVar()
        self.steel_price_var = tk.DoubleVar()
        self.updated_pergola = tk.DoubleVar()
        self.updated_decking = tk.DoubleVar()
        self.updated_cladding = tk.DoubleVar()
        self.shipping = tk.DoubleVar()
        self.current_config = {}  # Store the current configuration
        self.load_prices()  # Load prices from the configuration file
        self.create_widgets()

        self.configure(bg='#242424')
        self.geometry('300x500')

    def create_widgets(self):
        usd_label = ctk.CTkLabel(self, text="USD Price:")
        usd_entry = ctk.CTkEntry(self, textvariable=self.usd_price_var)
        steel_label = ctk.CTkLabel(self, text="Steel Price:")
        steel_entry = ctk.CTkEntry(self, textvariable=self.steel_price_var)
        updated_pergola_label = ctk.CTkLabel(self, text="Pergola Install:")
        updated_entry_pergola = ctk.CTkEntry(self, textvariable=self.updated_pergola)
        updated_decking_label = ctk.CTkLabel(self, text="Decking Install:")
        updated_entry_decking = ctk.CTkEntry(self, textvariable=self.updated_decking)
        updated_cladding_label = ctk.CTkLabel(self, text="Cladding Install:")
        updated_entry_cladding = ctk.CTkEntry(self, textvariable=self.updated_cladding)
        shipping_label = ctk.CTkLabel(self, text="Shipping Fees:")
        entry_shipping = ctk.CTkEntry(self, textvariable=self.shipping)

        usd_label.pack(pady=5)
        usd_entry.pack(pady=5)
        steel_label.pack(pady=5)
        steel_entry.pack(pady=5)
        updated_pergola_label.pack(pady=5)
        updated_entry_pergola.pack(pady=5)
        updated_decking_label.pack(pady=5)
        updated_entry_decking.pack(pady=5)
        updated_cladding_label.pack(pady=5)
        updated_entry_cladding.pack(pady=5)
        shipping_label.pack(pady=5)
        entry_shipping.pack(pady=5)

        apply_button = ctk.CTkButton(self, text="Apply Changes", command=self.apply_changes)
        apply_button.pack(pady=10)

    def apply_changes(self):
        current_datetime = datetime.datetime.now().strftime("%d-%m-%Y")

        price_vars = {
            "usd_price": self.usd_price_var,
            "steel_price": self.steel_price_var,
            "updated_pergola": self.updated_pergola,
            "updated_decking": self.updated_decking,
            "updated_cladding": self.updated_cladding,
            "shipping": self.shipping,
        }

        for key, var in price_vars.items():
            new_value = var.get()
            if key not in self.current_config or self.current_config[key]['value'] != new_value:
                self.current_config[key] = {"value": new_value, "last_updated": current_datetime}

        self.encrypt_config_file()

        messagebox.showinfo("Changes Applied", "Price changes have been applied successfully!")

    def load_prices(self):
        try:
            with open(self.CONFIG_FILE, "rb") as file:
                ciphertext = file.read()

            key = self.STATIC_KEY
            cipher = Cipher(algorithms.AES(key), modes.ECB(), backend=default_backend())
            decryptor = cipher.decryptor()
            decrypted_data = decryptor.update(ciphertext) + decryptor.finalize()

            # Remove PKCS7 padding
            padding_length = decrypted_data[-1]
            decrypted_data = decrypted_data[:-padding_length]

            self.current_config = json.loads(decrypted_data.decode('utf-8'))

            # Set the Tkinter variable values
            self.usd_price_var.set(self.current_config.get("usd_price", {}).get("value", 100.0))
            self.steel_price_var.set(self.current_config.get("steel_price", {}).get("value", 50.0))
            self.updated_pergola.set(self.current_config.get("updated_pergola", {}).get("value", 0.0))
            self.updated_decking.set(self.current_config.get("updated_decking", {}).get("value", 0.0))
            self.updated_cladding.set(self.current_config.get("updated_cladding", {}).get("value", 0.0))
            self.shipping.set(self.current_config.get("shipping", {}).get("value", 0.0))
        except FileNotFoundError:
            pass  # Ignore if the file is not found
        except Exception as e:
            print(f"Error loading prices: {e}")

    def encrypt_config_file(self):
        # Convert the current_config to JSON string
        json_data = json.dumps(self.current_config).encode('utf-8')

        # Add PKCS7 padding
        block_size = algorithms.AES.block_size // 8
        padding_length = block_size - (len(json_data) % block_size)
        padded_json_data = json_data + bytes([padding_length] * padding_length)

        key = self.STATIC_KEY
        cipher = Cipher(algorithms.AES(key), modes.ECB(), backend=default_backend())
        encryptor = cipher.encryptor()
        ciphertext = encryptor.update(padded_json_data) + encryptor.finalize()

        # Save the encrypted data to the file
        with open(self.CONFIG_FILE, 'wb') as file:
            file.write(ciphertext)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Optionally hide the main window
    app = AdminPopup(master=root)
    app.mainloop()
