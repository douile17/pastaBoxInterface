import pandas as pd
from tkinter import ttk
import serial
import serial.tools.list_ports
from tkinter import *
from tkinter import filedialog, messagebox
from threading import Thread, Event
from time import sleep
from datetime import datetime

# Variables globales pour suivre l'état de la connexion au port COM, du chargement du fichier CSV et de l'envoi de données
connected = False
csv_loaded = False
sending = False
send_thread = None

# Fonction pour rechercher les ports COM disponibles
def find_serial_ports():
    ports = serial.tools.list_ports.comports()
    available_ports = []
    for port in ports:
        available_ports.append((port.device, port.description))
    return available_ports

# Fonction pour se connecter au port série
def connect_serial():
    global ser, connected, send_button
    try:
        serial_port = port_combo.get().split()[0]  # Récupérer le port sélectionné depuis la ComboBox
        ser = serial.Serial(serial_port, baudrate=9600, timeout=1)
        connected = True
        connect_button.config(text="Disconnect", command=disconnect_serial)
        add_to_console(f"Connected to {serial_port}")
        if csv_loaded:  # Si le fichier CSV est chargé, activer le bouton "Start"
            root.after(1000, lambda: send_button.config(state="normal", bg="green", fg="white", text="Start Sequence"))
    except serial.SerialException as e:
        messagebox.showerror("Error", str(e))

# Fonction pour se déconnecter du port série
def disconnect_serial():
    global ser, connected, sending, send_button
    if connected:
        try:
            ser.close()
            connected = False
            connect_button.config(text="Connect", command=connect_serial)
            send_button.config(state="disabled", bg="gray", fg="black", text="Start Sequence")
            add_to_console("Disconnected from serial port")
            if sending:  # Si une séquence est en cours, arrêter l'envoi
                stop_data()
        except serial.SerialException as e:
            messagebox.showerror("Error", str(e))
    else:
        print("Already disconnected")

# Fonction pour charger un fichier CSV
def load_csv():
    global data, csv_loaded
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        try:
            data = pd.read_csv(file_path)
            csv_loaded = True
            add_to_console(f"CSV file loaded: {file_path}")
            print(f"CSV file loaded: {file_path}")
            if connected:  # Si le port série est connecté, activer le bouton "Start"
                root.after(1000, lambda: send_button.config(state="normal", bg="green", fg="white", text="Start Sequence"))
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Fonction principale pour démarrer ou arrêter l'envoi de données au port série
def toggle_send_data():
    global sending, send_thread
    if sending:
        stop_data()
    else:
        if connected and csv_loaded:
            sending = True
            send_thread = Thread(target=send_data_thread)
            send_thread.start()
            send_button.config(text="Stop Sequence", bg="red", fg="white")
        else:
            messagebox.showwarning("Warning", "Please connect to a serial port and load a CSV file first.")

def stop_data():
    global sending
    sending = False
    if connected:
        ser.write(b"D")  # Envoie la lettre "D" lorsque le bouton "Stop" est pressé
    send_button.config(text="Start Sequence", bg="green" if connected else "gray", fg="white")

# Fonction principale pour envoyer les données au port série
def send_data_thread():
    global stop_event, send_button, sending
    stop_event = Event()  # Crée un nouvel événement pour chaque cycle d'envoi de données
    sending = True  # Définir l'état d'envoi comme True au début de l'envoi
    for index, row in data.iterrows():
        if not sending:  # Vérifie si l'envoi est toujours en cours
            break  # Sort de la boucle si l'envoi est arrêté

        current_time = datetime.now().strftime("%H:%M:%S")
        time_value = row['Time (min)']
        output_value = row['Output']
        add_to_console(f"Current time: {current_time} | CSV time: {time_value} | Output: {output_value}", separate=False)

        ser.write(output_value.encode())

        next_time_value = data.at[index + 1, 'Time (min)'] if index < len(data) - 1 else None
        if next_time_value is not None:
            time_difference = next_time_value - time_value
            if time_difference > 0:
                sleep(time_difference * 60)

        if not connected:  # Si la connexion au port COM est perdue, arrêter l'envoi
            stop_data()
            break

    # Vérifier si l'envoi est terminé et mettre à jour le bouton en conséquence
    if sending:
        send_button.config(text="Start Sequence", bg="green" if connected else "gray", fg="white")
        sending = False  # Mettre à jour l'état d'envoi comme False à la fin de l'envoi

    add_to_console("Sequence finished")  # Ajouter un message à la console une fois la séquence terminée

# Fonction pour rafraîchir les ports COM dans la ComboBox
def refresh_ports():
    ports = find_serial_ports()
    port_combo['values'] = [f"{port} - {description}" for port, description in ports]
    if ports:
        port_combo.current(0)
    else:
        port_combo.set("No COM Ports Available")

# Fonction pour ajouter un message à la console
def add_to_console(message, separate=True):
    console.config(state="normal")
    console.insert(END, message + "\n")
    if separate:
        console.insert(END, "-" * 72 + "\n")  # Ajouter des tirets pour séparer les messages
    console.config(state="disabled")
    console.see(END)  # Fait défiler automatiquement vers le bas pour afficher le message le plus récent

# Fonction pour quitter proprement
def on_closing():
    global sending
    if connected:
        if messagebox.askokcancel("Quitter", "Voulez-vous vraiment quitter ?"):
            if sending:
                stop_data()  # Arrêter l'envoi de données s'il est en cours
            disconnect_serial()  # Déconnecter proprement le port série
            root.destroy()
    else:
        if messagebox.askokcancel("Quitter", "Voulez-vous vraiment quitter ?"):
            root.destroy()

# Charger les données du fichier CSV par défaut
data = pd.DataFrame()

# Créer une fenêtre
root = Tk()
root.title("Pasta Box")
root.geometry("600x400")
root.iconbitmap('mipi.ico')
# Combobox pour sélectionner le port COM
port_combo = ttk.Combobox(root, width=30)
port_combo.place(x=10,y=10)

# Bouton pour refresh les ports COM disponibles
refresh_button = Button(root, text="Refresh Ports",relief='flat',bg="white", command=refresh_ports)
refresh_button.place(x=310,y=8)
refresh_ports()  # Rafraîchir la liste des ports dès le démarrage

# Bouton pour charger un fichier CSV
load_csv_button = Button(root, text="Load CSV",relief='flat',bg="white", command=load_csv)
load_csv_button.place(x=10,y=50)

# Bouton pour se connecter ou se déconnecter
connect_button = Button(root, text="Connect",relief='flat',bg="white", command=connect_serial)
connect_button.place(x=230,y=8)

# Bouton pour démarrer ou arrêter l'envoi de données
send_button = Button(root, text="Start Sequence", command=toggle_send_data, state="disabled", bg="gray", fg="white",relief='flat')
send_button.place(y=10, x=500)

# Console pour afficher les messages
console = Text(root, height=17,width=72, state="disabled", wrap="word")
console.place(x=10 , y=100)

# Fonction pour quitter proprement
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
