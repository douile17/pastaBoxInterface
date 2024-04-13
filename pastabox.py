import pandas as pd
import serial
from time import sleep
from datetime import datetime

# Charger les données du fichier CSV
data = pd.read_csv('data.csv')

# Spécifier le port série où est connecté l'Arduino
serial_port = "COM9"
ser = serial.Serial(serial_port, baudrate=9600, timeout=1)

# Attendre que l'Arduino soit prêt
sleep(2)

# Boucle sur les lignes du DataFrame
for index, row in data.iterrows():
    # Récupérer le temps de l'horloge Windows en minutes
    current_time = datetime.now().strftime("%H:%M:%S")  # Format sans millisecondes

    # Récupérer les données du CSV
    time_value = row['Time (min)']
    output_value = row['Output']
    print(f"Current time: {current_time} | CSV time: {time_value} | Output: {output_value}")

    # Envoyer la sortie correspondante à l'Arduino
    ser.write(output_value.encode())

    # Attendre jusqu'au prochain moment spécifié dans le CSV
    next_time_value = data.at[index + 1, 'Time (min)'] if index < len(data) - 1 else None
    if next_time_value is not None:
        time_difference = next_time_value - time_value
        if time_difference > 0:
            sleep(time_difference * 60)

# Fermer le port série
ser.close()

