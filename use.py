import os
import cv2
import pandas as pd
import numpy as np
from tkinter import Tk, filedialog, Label, Entry, Button, StringVar
from tkinter import ttk  
import shutil


# Specify the path to the Excel file
excel_file_path = "C:\\Users\\kokne\\OneDrive\\Pictures\\combined_extraction.xlsx"

def open_image(image_path):
    os.startfile(image_path)

def search_images(event=None):  # Modified to accept an event parameter
    search_text = entry_search.get()

    # Filter rows containing the search text
    filtered_df = df[df['Extracted Text (Method 1)'].str.contains(search_text, case=False) |
                     df['Extracted Text (Method 2)'].str.contains(search_text, case=False)]

    # Display matching images with click functionality
    if not filtered_df.empty:
        folder_path = os.path.dirname(excel_file_path)
        result_folder_path = os.path.join(folder_path, "search_results")
        
        # Create a subfolder for result images if it doesn't exist
        if not os.path.exists(result_folder_path):
            os.makedirs(result_folder_path)

        images = []
        image_paths = []

        for index, row in filtered_df.iterrows():
            image_name = row['Image Name']
            image_path = os.path.join(folder_path, image_name)

            # Copy the original image to the result subfolder
            result_image_path = os.path.join(result_folder_path, image_name)
            shutil.copy(image_path, result_image_path)

            images.append(image_path)
            image_paths.append(result_image_path)

        def on_click(event, x, y, flags, param):
            # Check if the left mouse button is clicked
            if event == cv2.EVENT_LBUTTONDOWN:
                # Determine the clicked image
                col_idx = x // (400 + margin)
                row_idx = y // (400 + margin)
                image_index = row_idx * num_cols + col_idx

                # Open the clicked image using os.startfile for handling special characters
                clicked_image_path = image_paths[image_index]
                open_image(clicked_image_path)

        # Create a larger grid of images with a border, padding, and margin
        num_cols = 2  # Number of columns in the grid
        padding = 20  # Padding around the entire grid
        margin = 10  # Margin between images

        num_rows = -(-len(images) // num_cols)  # Calculate the number of rows

        # Initialize the larger grid canvas with a border, padding, and margin
        grid_canvas = 255 * np.ones(
            ((num_rows * (400 + margin) + padding), (num_cols * (400 + margin) + padding), 3),
            dtype=np.uint8
        )

        # Populate the grid with resized images and border, padding, and margin
        for i, img_path in enumerate(images):
            img = cv2.imread(img_path)
            resized_image = cv2.resize(img, (400, 400))  # Adjust the size as needed

            row_idx = i // num_cols
            col_idx = i % num_cols
            grid_canvas[
                (row_idx * (400 + margin) + padding):((row_idx + 1) * (400 + margin) + padding) - margin,
                (col_idx * (400 + margin) + padding):((col_idx + 1) * (400 + margin) + padding) - margin,
                :
            ] = resized_image

        # Display the larger grid
        cv2.imshow("Search Results", grid_canvas)
        cv2.setMouseCallback("Search Results", on_click)  # Set the mouse click callback
        cv2.waitKey(0)
        cv2.destroyAllWindows()
    else:
        # Display message in the GUI when no images are found
        no_images_label = Label(root, text="No images found containing the specified text.", font=("Helvetica", 10), pady=10)
        no_images_label.pack()

# Rest of the code remains the same
# Check if the Excel file exists
if os.path.exists(excel_file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file_path)

    # Create Tkinter GUI for text search
    root = Tk()
    root.title("Visual Eureka")

    label_welcome = Label(root, text="WELCOME TO VISUAL EUREKA ",font=("Helvetica", 14, "bold"), pady=10)
    label_welcome.pack()
    
    label_search = Label(root, text="Enter text to search:", font=("Helvetica",10),pady=10)
    label_search.pack()

    entry_search = Entry(root)
    entry_search.pack(pady=10)

    button_search = Button(root, text="Search", command=search_images)
    button_search.pack()

    # Set the window dimensions
    window_width = 400
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2

    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Added binding to handle Enter key press
    root.bind('<Return>', search_images)

    root.mainloop()
else:
    print("Excel file not found at the specified path.")