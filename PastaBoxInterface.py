import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pandas as pd
import serial
from serial.tools.list_ports import comports
from threading import Thread
from time import sleep, strftime, localtime

class SerialController:
    def __init__(self, port, console, app):
        self.port = port
        self.console = console
        self.app = app
        self.ser = serial.Serial(baudrate=9600, timeout=1)
        self.running = False
        self.paused = False
        self.current_index = 0

    def start(self, filename):
        self.ser.port = self.port
        self.ser.open()
        self.running = True
        self.paused = False
        self.filename = filename
        self.app.update_console("Script started.")
        self.thread = Thread(target=self.send_data)
        self.thread.start()

    def stop(self):
        self.running = False
        if self.ser.is_open:
            self.ser.write(b'D')
            self.ser.close()
        self.app.update_console("Script stopped.")
        self.app.reset_buttons()

    def pause(self):
        self.paused = True
        self.app.update_console("Script paused.")

    def resume(self):
        self.paused = False
        self.app.update_console("Script resumed.")

    def send_data(self):
        data = self.load_csv(self.filename)
        if data is None:
            self.stop()
            return

        for index, row in data.iterrows():
            self.current_index = index
            if not self.running:
                return

            if self.paused:
                while self.paused:
                    sleep(0.1)

            time_value = row['Time (min)']
            output_value = row['Output']
            current_time = strftime("%H:%M:%S", localtime())
            msg = f"{output_value} | Csv time: {time_value} | Current time: {current_time}"
            print(msg)
            self.console.insert(tk.END, msg + '\n')
            self.console.see(tk.END)

            if self.ser.is_open:
                self.ser.write(output_value.encode())

            if index < len(data) - 1:
                next_time_value = data.at[index + 1, 'Time (min)']
                time_difference = (next_time_value - time_value) * 60
                sleep(time_difference)
        
        # Si la lecture du fichier CSV est terminée, arrêter le script et réinitialiser les boutons
        self.stop()

    def load_csv(self, filename):
        if filename:
            try:
                return pd.read_csv(filename)
            except Exception as e:
                print("Error loading CSV file:", e)
        return None

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Pasta Box")

        # Définir la couleur de fond
        self.bg_color = "#f0f0f0"  # Couleur de fond
        self.button_bg_color = "#FFFFFF"  # Couleur de fond des boutons
        self.button_disabled_color = "#D3D3D3"  # Couleur de fond désactivée des boutons
        self.button_border_color = "#808080"  # Couleur du contour des boutons

        self.root.configure(bg=self.bg_color)  # Appliquer la couleur de fond à la fenêtre

        # Définir l'icône
        self.root.iconbitmap('favicon.ico')

        # Centrer la fenêtre au lancement
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x_coordinate = (screen_width - 600) // 2
        y_coordinate = (screen_height - 400) // 2
        self.root.geometry(f"600x400+{x_coordinate}+{y_coordinate}")

        # Définir une taille minimale
        self.root.minsize(600, 400)

        self.port_frame = tk.Frame(root, bg=self.bg_color)
        self.port_frame.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.port_label = tk.Label(self.port_frame, text="Select COM port:", bg=self.bg_color)
        self.port_label.grid(row=0, column=0, padx=(0, 5))

        self.ports = self.get_serial_ports()
        self.port_var = tk.StringVar()
        self.port_dropdown = tk.OptionMenu(self.port_frame, self.port_var, *self.ports)
        self.port_dropdown.config(bg=self.button_bg_color, bd=1, relief=tk.FLAT)

        # Appliquer la couleur de fond blanc aux propositions du menu déroulant
        for option in self.port_dropdown.children.values():
            option.config(bg="#FFFFFF", relief=tk.FLAT)

        self.port_dropdown.grid(row=0, column=1)

        self.load_button = tk.Button(root, text="Load CSV", command=self.load_csv, bg=self.button_bg_color, bd=1, relief=tk.FLAT)
        self.load_button.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.start_stop_button = tk.Button(root, text="Start", command=self.start_stop, width=10, highlightbackground="green", bg=self.button_bg_color, activebackground=self.button_bg_color, relief=tk.FLAT, borderwidth=1, state=tk.DISABLED)
        self.start_stop_button.grid(row=0, column=1, padx=10, pady=5, sticky="e")

        self.pause_button = tk.Button(root, text="Pause", command=self.pause_resume, width=10, highlightbackground="red", bg=self.button_bg_color, activebackground=self.button_bg_color, relief=tk.FLAT, borderwidth=1, state=tk.DISABLED)
        self.pause_button.grid(row=1, column=1, padx=10, pady=5, sticky="e")

        self.console_label = tk.Label(root, text="Console:")
        self.console_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.console = scrolledtext.ScrolledText(root, wrap=tk.WORD)
        self.console.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        self.clear_button = tk.Button(root, text="Clear Console", command=self.clear_console, bg=self.button_bg_color, bd=1, relief=tk.FLAT)
        self.clear_button.grid(row=4, column=0, padx=10, pady=5, sticky="w")

        self.serial_controller = None
        self.paused = False

        root.grid_rowconfigure(3, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # Gestion de la fermeture de la fenêtre
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def get_serial_ports(self):
        ports = []
        for port_info in comports():
            ports.append(f"{port_info.device} - {port_info.description}")
        return ports

    def load_csv(self):
        filename = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV files", "*.csv")])
        if filename:
            print("Selected CSV file:", filename)
            self.filename = filename
            self.start_stop_button.config(state=tk.NORMAL, text="Start", fg="white", bg="green")
            self.update_console("CSV file loaded successfully: " + filename)

    def start_stop(self):
        if not self.port_var.get():
            self.update_console("Please select a COM port.")
            return

        if self.serial_controller is None:
            self.serial_controller = SerialController(self.port_var.get().split(" - ")[0], self.console, self)

        if self.start_stop_button.cget("text") == "Start":
            self.serial_controller.start(self.filename)
            self.start_stop_button.config(text="Stop", bg="red")
            self.pause_button.config(state=tk.NORMAL)
        else:
            self.serial_controller.stop()
            self.start_stop_button.config(text="Start", bg="green")
            self.pause_button.config(state=tk.DISABLED)

    def pause_resume(self):
        if self.serial_controller is not None:
            if self.paused:
                self.paused = False
                self.pause_button.config(text="Pause")
                self.serial_controller.resume()
            else:
                self.paused = True
                self.pause_button.config(text="Resume")
                self.serial_controller.pause()

    def clear_console(self):
        self.console.delete(1.0, tk.END)

    def update_console(self, message):
        if "| Csv time:" not in message and "| Current time:" not in message:
            num_dashes = 69  # Nombre de tirets souhaité
            self.console.insert(tk.END, '-' * num_dashes + '\n')  # Ajout de la ligne séparatrice
        self.console.insert(tk.END, message + '\n')
        self.console.see(tk.END)

    def reset_buttons(self):
        self.start_stop_button.config(text="Start", bg="green")
        self.pause_button.config(state=tk.DISABLED)

    def on_close(self):
        if self.serial_controller is not None:
            if messagebox.askokcancel("Quit", "Do you want to quit?"):
                self.serial_controller.stop()
                self.root.destroy()
        else:
            if messagebox.askokcancel("Quit", "Do you want to quit?"):
                self.root.destroy()

root = tk.Tk()
app = App(root)
root.mainloop()
