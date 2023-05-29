from tkinter import *
import tkinter.font as font
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
import matplotlib.pyplot as plt
from geopy.geocoders import Nominatim
from PIL import Image, ImageTk

# Declare global variables for entry fields in the survey window
site_name_entry = None
city_entry = None
length_entry = None
width_entry = None
roof_type_entry = None

def validateLogin():
    myFont = font.Font(family='Helvetica', size=12)
    if usernameEntry.get() == "shivasagar" and passwordEntry.get() == "solar":
        resultLabel.config(fg="green", font=myFont)
        resultLabel["text"] = "Valid User"
        open_survey_window()
    else:
        resultLabel.config(fg="red", font=myFont)
        resultLabel["text"] = "Invalid User"


def open_survey_window():
    # Close the login window
    login_window.destroy()

    # Create the survey window
    survey_window = tk.Tk()
    survey_window.title("Solar Site Survey")

    # Create labels and entry fields for site information
    site_name_label = tk.Label(survey_window, text="SITE NAME:", font=("Arial", 14), bg="yellow", padx=3, pady=3)
    site_name_label.pack()
    global site_name_entry
    site_name_entry = tk.Entry(survey_window)
    site_name_entry.pack()

    city_label = tk.Label(survey_window, text="CITY:",font=("Arial",14),bg="Green",padx=3,pady=3)
    city_label.pack()
    global city_entry
    city_entry = tk.Entry(survey_window)
    city_entry.pack()

    length_label = tk.Label(survey_window, text="ROOF LENGTH (m):",font=("Arial",14),bg="pink",padx=3,pady=3)
    length_label.pack()

    global length_entry
    length_entry = tk.Entry(survey_window)
    length_entry.pack()

    width_label = tk.Label(survey_window, text="ROOf WIDTH (m):",font=("Arial",14),bg="red",padx=3,pady=3)
    width_label.pack()
    global width_entry
    width_entry = tk.Entry(survey_window,font=("Arial",13))
    width_entry.pack()

    roof_type_label = tk.Label(survey_window, text="ROOF TYPE:",font=("Arial",14),padx=3,pady=3)
    roof_type_label.pack()
    global roof_type_entry
    roof_type_entry = tk.Entry(survey_window)
    roof_type_entry.pack()   

    # Create a button to generate the report
    generate_button = tk.Button(survey_window, text="Generate Report", command=generate_report)
    generate_button.pack()

def generate_report():
    site_name = site_name_entry.get()
    city = city_entry.get()
    length = float(length_entry.get())
    width = float(width_entry.get())
    roof_type = roof_type_entry.get()

    # Get coordinates of the city
    latitude, longitude = get_coordinates(city)

    if latitude is None or longitude is None:
        messagebox.showerror("Error", "Invalid city name. Please enter a valid city name.")
        return

    # Calculate max capacity and solar potential
    max_capacity = calculate_max_capacity(length, width)
    roof_area = length * width
    solar_potential = calculate_solar_potential(roof_area, max_capacity)

    # Generate monthly generation yields
    months = ["January", "February", "March", "April", "May", "June", "July",
              "August", "September", "October", "November", "December"]
    generation_yields = [roof_area * max_capacity * 4, roof_area * max_capacity * 4.5,
                         roof_area * max_capacity * 5, roof_area * max_capacity * 5.5,
                         roof_area * max_capacity * 6, roof_area * max_capacity * 6.5,
                         roof_area * max_capacity * 7, roof_area * max_capacity * 6,
                         roof_area * max_capacity * 6, roof_area * max_capacity * 5,
                         roof_area * max_capacity * 4.5, roof_area * max_capacity * 4]

    # Create an Excel workbook and add data
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Survey Report"
    sheet['A1'] = "Site Name"
    sheet['B1'] = "City"
    sheet['C1'] = "Latitude"
    sheet['D1'] = "Longitude"
    sheet['E1'] = "Roof Length (m)"
    sheet['F1'] = "Roof Width (m)"
    sheet['G1'] = "Roof Type"
    sheet['H1'] = "Max Capacity (kW)"
    sheet['I1'] = "Solar Potential (kWh)"
    sheet['J1'] = "Month"
    sheet['K1'] = "Generation Yield (kWh)"

    sheet.append([site_name, city, latitude, longitude, length, width, roof_type,
                  max_capacity, solar_potential])

    for month, yield_value in zip(months, generation_yields):
        sheet.append(["", "", "", "", "", "", "", "", "", month, yield_value])

    # Save the workbook
    workbook.save("survey_report.xlsx")

     # Create and save charts
    plt.plot(months, generation_yields, marker='o', linestyle='-')
    plt.xlabel("Month")
    plt.ylabel("Generation Yield (kWh)")
    plt.title("Monthly Generation Yields")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("generation_yields.png")

    # Show report status
    messagebox.showinfo("Report Generated", "Survey report generated successfully.")

    # Show report status
    messagebox.showinfo("Report Generated", "Survey report generated successfully.")

def get_coordinates(city):
    geolocator = Nominatim(user_agent="solar_app")
    location = geolocator.geocode(city)

    if location is None:
        return None, None

    return location.latitude, location.longitude

def calculate_max_capacity(length, width):
    roof_area = length * width
    max_capacity = roof_area / 10  # Assumption: 1 kW capacity per 10 square meters of roof area
    return max_capacity

def calculate_solar_potential(roof_area, max_capacity):
    solar_potential = roof_area * max_capacity
    return solar_potential

# Create the login window
login_window = tk.Tk()
login_window.geometry('400x250')
login_window.title('Solar Site Survey')

# Create a colorful background
backgroundCanvas = Canvas(login_window, width=400, height=250, bg="lightblue")
backgroundCanvas.pack(fill="both", expand=True)

# Load and resize the photo
photo = Image.open("1.png")
resized_photo = photo.resize((150, 150), Image.ANTIALIAS)
image = ImageTk.PhotoImage(resized_photo)

# Create a label for the photo
photoLabel = Label(login_window, image=image, bg="lightblue")
photoLabel.place(x=1300, y=50)

# Calculate the center coordinates
center_x = backgroundCanvas.winfo_width() // 2
center_y = backgroundCanvas.winfo_height() // 2

# Create a frame for the login form
loginFrame = Frame(backgroundCanvas, bg="white",borderwidth=9)
loginFrame.place(relx=0.5, rely=0.4, anchor=CENTER)

# Label above login frame
titleLabel = Label(backgroundCanvas, text="FUSION SOLAR", font=("Arial", 25, "bold"), bg="Pink", borderwidth=4,
                  relief=SUNKEN)
titleLabel.place(relx=0.5, rely=0.2, anchor=CENTER)

# Username label and text entry box
usernameLabel = Label(loginFrame, text="Username", font=("Arial", 12), bg="yellow", borderwidth=7, relief=RIDGE)
usernameLabel.grid(row=0, column=0, padx=20, pady=15)

usernameEntry = Entry(loginFrame, font=("Arial", 12))
usernameEntry.grid(row=0, column=1, padx=10, pady=5)

# Password label and password entry box
passwordLabel = Label(loginFrame, text="Password", font=("Arial", 12), bg="red", borderwidth=7, relief=RIDGE)
passwordLabel.grid(row=1, column=0, padx=10, pady=5)

passwordEntry = Entry(loginFrame, show='*', font=("Arial", 12))
passwordEntry.grid(row=1, column=1, padx=10, pady=5)

# Result label
resultLabel = Label(loginFrame, font=("Arial", 14))
resultLabel.grid(row=2, columnspan=2, padx=10, pady=20)

# Login button
loginButton = Button(loginFrame, text="Login", font=("Arial", 12), command=validateLogin)
loginButton.grid(row=3, columnspan=2, padx=5, pady=5)

login_window.mainloop()