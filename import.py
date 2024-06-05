import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk
import xml.etree.ElementTree as ET
import pandas as pd
import os
import time


SERIAL_PORT = None
BAUD_RATE = 9600


XML_FILENAME = None


def list_serial_ports():
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]


def select_serial_port():
    port_window = tk.Toplevel(root)
    port_window.title("Select Serial Port")
    
    port_listbox = tk.Listbox(port_window)
    port_listbox.pack(fill=tk.BOTH, expand=True)
    
    available_ports = list_serial_ports()
    for port in available_ports:
        port_listbox.insert(tk.END, port)
    
    def on_select(event):
        global SERIAL_PORT
        selected_index = port_listbox.curselection()
        if selected_index:
            SERIAL_PORT = available_ports[selected_index[0]]
            port_window.destroy()  
            initialize_serial()  
    
    port_listbox.bind('<<ListboxSelect>>', on_select)
    
    port_window.mainloop()

def initialize_serial():
    global ser, XML_FILENAME
    if SERIAL_PORT:
        ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=1)
        
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        XML_FILENAME = f"ph_values_{timestamp}.xml"
        
        read_data()  
last_value = None

def read_data():
    global last_value
    if 'ser' in globals() and ser.in_waiting > 0:
        line = ser.readline().decode('utf-8').rstrip()
        try:
            last_value = float(line)
            update_gui(last_value)
        except ValueError:
            print("Received non-numeric data:", line)
    root.after(100, read_data)

def update_gui(value):
    label_var.set(f"Ph Değeri: {value:.2f}")
    pointer_position = (value / 14.00) * canvas_width
    canvas.coords(pointer, pointer_position, 0, pointer_position, canvas_height)
    category_label_var.set(categorize_ph(value))

def save_to_xml():
    global XML_FILENAME
    if last_value is not None:
        ph_value_element = ET.Element("ph_value")
        ph_value_element.text = str(last_value)
        
        if os.path.exists(XML_FILENAME):
            tree = ET.parse(XML_FILENAME)
            root = tree.getroot()
        else:
            root = ET.Element("root")
        
        root.append(ph_value_element)
        
        tree = ET.ElementTree(root)
        tree.write(XML_FILENAME, xml_declaration=True)

def output_to_excel():
    global XML_FILENAME
    if os.path.exists(XML_FILENAME):
        tree = ET.parse(XML_FILENAME)
        root = tree.getroot()
        data = []
        for ph_value in root.findall('ph_value'):
            data.append(float(ph_value.text))
        df = pd.DataFrame(data, columns=['Ph Values'])
        df.to_excel('ph_values.xlsx', index=False)

def interpolate_color(start_color, end_color, factor):
    r1, g1, b1 = start_color
    r2, g2, b2 = end_color
    r = int(r1 + (r2 - r1) * factor)
    g = int(g1 + (g2 - g1) * factor)
    b = int(b1 + (b2 - b1) * factor)
    return r, g, b

def rgb_to_hex(rgb):
    return '#%02x%02x%02x' % rgb

def categorize_ph(ph):
    if ph < 4:
        return "Very Acidic"
    elif 4 <= ph < 5:
        return "Acidic"
    elif 5 <= ph < 7:
        return "Acidic-ish"
    elif 7 <= ph < 8:
        return "Neutral"
    elif 8 <= ph < 10:
        return "Alkaline-ish"
    elif 10 <= ph < 11:
        return "Alkaline"
    else:
        return "Very Alkaline"

root = tk.Tk()
root.title("Ph Ölçer")

label_var = tk.StringVar()
label = ttk.Label(root, textvariable=label_var)
label.pack(pady=10)

category_label_var = tk.StringVar()
category_label = ttk.Label(root, textvariable=category_label_var)
category_label.pack(pady=5)

canvas_width = 400
canvas_height = 20
canvas = tk.Canvas(root, width=canvas_width, height=canvas_height + 50)  
canvas.pack()

rainbow_colors = [
    (255, 0, 0),  # Red
    (255, 127, 0),  # Orange
    (255, 255, 0),  # Yellow
    (0, 255, 0),  # Green
    (0, 0, 255),  # Blue
    (75, 0, 130),  # Indigo
    (148, 0, 211)  # Violet
]

segment_width = canvas_width / (len(rainbow_colors) - 1)  
for i in range(len(rainbow_colors) - 1):
    start_color = rainbow_colors[i]
    end_color = rainbow_colors[i + 1]
    for j in range(int(segment_width)):
        factor = j / segment_width
        color = interpolate_color(start_color, end_color, factor)
        color_hex = rgb_to_hex(color)
        canvas.create_rectangle(i * segment_width + j, 0, (i + 1) * segment_width, canvas_height, fill=color_hex, outline=color_hex)

pointer = canvas.create_line(0, 0, 0, canvas_height, fill="black", width=10)

acidity_labels = ["Acidic", "Neutral", "Alkaline"]
for i, acidity in enumerate(acidity_labels):
    label = ttk.Label(canvas, text=acidity)
    if i == 0:  
        x_position = 0
    elif i == len(acidity_labels) - 1:  
        x_position = canvas_width - segment_width / 2
    else:  
        x_position = canvas_width / 2 - segment_width / 2
    label.place(x=x_position, y=canvas_height + 20)  

save_button = ttk.Button(root, text="Save", command=save_to_xml)
save_button.pack(side=tk.LEFT, padx=(10, 5))

output_button = ttk.Button(root, text="Output", command=output_to_excel)
output_button.pack(side=tk.RIGHT, padx=(5, 10))

read_data()

root.after(100, initialize_serial)
root.after(200, read_data)

select_serial_port()

root.mainloop()

if 'ser' in locals() or 'ser' in globals():
    ser.close()