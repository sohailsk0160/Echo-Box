import tkinter as tk
from tkinter import Label, Button
from PIL import Image, ImageTk


# Function to fade in text
def fade_in(label, text, index=0):
    if index < len(text):
        label.config(text=text[:index + 1])  # Display characters gradually
        root.after(100, fade_in, label, text, index + 1)


# Function to append names one by one
def append_name(label, name_list, index=0):
    if index < len(name_list):
        current_text = label.cget("text")  # Get current text
        label.config(text=current_text + "\n" + name_list[index])  # Append new name
        root.after(500, append_name, label, name_list, index + 1)  # Add next name


# Create Splash Screen Window
root = tk.Tk()
root.title("Splash Screen")
root.geometry("800x500")
root.overrideredirect(False)  # Allows user to move/close window

# Load Background Image (background2.jpeg)
try:
    bg_main = Image.open("background2.jpg")
    bg_main = bg_main.resize((800, 500))  # Resize to fit window
    bg_main_image = ImageTk.PhotoImage(bg_main)

    bg_main_label = Label(root, image=bg_main_image)
    bg_main_label.place(relwidth=1, relheight=1)  # Cover full window
except FileNotFoundError:
    print("Main background image not found. Using white background.")
    root.config(bg="white")

# Load College Name Image (background.png) at the top
try:
    bg_top = Image.open("background.png")
    bg_top = bg_top.resize((800, 100))  # Resize to fit top section
    bg_top_image = ImageTk.PhotoImage(bg_top)

    bg_top_label = Label(root, image=bg_top_image)
    bg_top_label.place(relx=0.5, rely=0.0, anchor="n")  # Set image at the top
except FileNotFoundError:
    print("College name image not found.")

# Display Project Name
project_label = Label(root, text="", font=("Arial", 30, "bold"), bg="white", fg="black")
project_label.place(relx=0.5, rely=0.85, anchor="center")

# Animate text appearance
root.after(300, fade_in, project_label, "Project Name: Echo Box !!!")

# Run the Splash Screen
root.mainloop()
