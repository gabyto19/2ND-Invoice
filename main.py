import os
import sys
import fitz  # PyMuPDF
import datetime
import tkinter as tk
from tkinter import messagebox

# Get the path of the existing PDF and font in the bundled executable
def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Function to modify the existing PDF and add input data
def create_pdf(name, deposit, vehicle, vin, vin_last4):
    try:
        # Load the existing PDF from a fixed path
        input_pdf = resource_path("invoice.pdf")  # The fixed local path to the PDF file
        doc = fitz.open(input_pdf)
        
        # Choose the first page to insert text
        page = doc[0]
        
        # Define the position where you want to insert text
        name_position = (185, 298)
        vehicle_position = (235, 324)
        date_position = (420, 98)
        deposit_position = (500, 403)
        depositPos = (500, 430)
        vin_position = (160, 351)
        vin_last4_position = (503, 80)  # Position for the last 4 characters of VIN
        validation_date_position = (420, 116)  # Define the position for validation date
        
        # Insert system date
        current_date = datetime.datetime.now().strftime("%d/%m/%Y")
        validation_date = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%d/%m/%Y")
        
        # Load the ArialMT font
        font_path = resource_path("ArialMT.ttf")
        page.insert_font(fontname="ArialMT", fontfile=font_path, encoding="unicode")

        # Set the text color to gray (RGB: 85, 85, 85)
        gray_color = (85 / 255, 85 / 255, 85 / 255)  # RGB values for #555555

        # Simulate subtle bold by drawing the text a couple of times with a slight offset
        def insert_bold_text(position, text):
            for x_offset, y_offset in [(0, 0), (0.2, 0.2)]:
                x, y = position
                page.insert_text((x + x_offset, y + y_offset), text, fontsize=12, fontname="ArialMT", color=gray_color)

        # Insert the text with gray color and subtle bold effect
        insert_bold_text(name_position, f"{name}")
        insert_bold_text(vehicle_position, f"{vehicle}")
        insert_bold_text(date_position, f"{current_date}")
        insert_bold_text(deposit_position, f"{deposit}")
        insert_bold_text(depositPos, f"{deposit}")
        insert_bold_text(vin_position, f"{vin}")
        insert_bold_text(vin_last4_position, f"{vin_last4}")  # Insert last 4 characters of VIN
        insert_bold_text(validation_date_position, f"{validation_date}")
        # Save the modified PDF to a new file
        output_pdf = "output_modified.pdf"
        doc.save(output_pdf)
        doc.close()

        messagebox.showinfo("Success", f"PDF modified successfully and saved as '{output_pdf}'!")
        
        # Close the app after PDF generation
        root.quit()  # Closes the Tkinter window
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to handle the "Generate PDF" button click
def generate_pdf():
    name = entry_name.get()
    vehicle = entry_vehicle.get() + "$"
    deposit = entry_deposit.get()
    vin = entry_vin.get()
    
    if not name or not vehicle or not deposit or not vin:
        messagebox.showwarning("Input Error", "All fields are required!")
        return
    
    # Extract the last 4 characters from the VIN
    vin_last4 = vin[-4:] if len(vin) >= 4 else vin  # Ensure VIN has at least 4 characters
    
    # Call the function to create the PDF
    create_pdf(name, vehicle, deposit, vin, vin_last4)

# Create the tkinter GUI
root = tk.Tk()
root.title("PDF Modifier")

# Create and place the input fields and labels
label_name = tk.Label(root, text="Enter your Name and ID:")
label_name.grid(row=0, column=0, padx=10, pady=5, sticky="w")

entry_name = tk.Entry(root, width=30)
entry_name.grid(row=0, column=1, padx=10, pady=5)

label_vehicle = tk.Label(root, text="Enter Money Amount:")
label_vehicle.grid(row=3, column=0, padx=10, pady=5, sticky="w")

entry_vehicle = tk.Entry(root, width=30)
entry_vehicle.grid(row=3, column=1, padx=10, pady=5)

label_vin = tk.Label(root, text="Enter VIN Code:")
label_vin.grid(row=2, column=0, padx=10, pady=5, sticky="w")

entry_vin = tk.Entry(root, width=30)
entry_vin.grid(row=2, column=1, padx=10, pady=5)

label_deposit = tk.Label(root, text="Enter your Vehicle Model:")
label_deposit.grid(row=1, column=0, padx=10, pady=5, sticky="w")

entry_deposit = tk.Entry(root, width=30)
entry_deposit.grid(row=1, column=1, padx=10, pady=5)

# Create and place the Generate PDF button
generate_button = tk.Button(root, text="Generate PDF", command=generate_pdf)
generate_button.grid(row=4, column=0, columnspan=2, pady=10)

# Start the GUI event loop
root.mainloop()


#pyinstaller --onefile --noconsole --add-data "invoice.pdf;." --add-data "arialmt.ttf;." main.py
