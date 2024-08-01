import os
import logging
from transformers import TrOCRProcessor, VisionEncoderDecoderModel
from PIL import Image
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm  # Import the tqdm library for progress bar

# Set the logging level to ERROR for the transformers library
logging.getLogger("transformers").setLevel(logging.ERROR)

# Create a Tkinter window and hide it
root = tk.Tk()
root.withdraw()

# Ask the user for the directory containing images
directory_path = filedialog.askdirectory(title="Select a Directory")

if not directory_path:
    print("No directory selected. Exiting.")
else:
    # Initialize lists to store image names and extracted text
    image_names = []
    extracted_texts = []

    # Load processor and model
    processor = TrOCRProcessor.from_pretrained('microsoft/trocr-base-handwritten')
    model = VisionEncoderDecoderModel.from_pretrained('microsoft/trocr-base-handwritten')

    # Get a list of image files in the selected directory
    image_files = [filename for filename in os.listdir(directory_path) if filename.endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))]

    # Create a tqdm progress bar
    progress_bar = tqdm(image_files, unit="image")

    # Process each image in the selected directory
    for filename in progress_bar:
        image_path = os.path.join(directory_path, filename)
        image = Image.open(image_path).convert("RGB")

        # Generate text with a specific max_length (e.g., 100)
        pixel_values = processor(images=image, return_tensors="pt").pixel_values
        generated_ids = model.generate(pixel_values, max_length=200)
        generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

        # Store the image name and extracted text
        image_names.append(filename)
        extracted_texts.append(generated_text)

        # Update the progress bar
        # progress_bar.set_description(f"Processing {filename}")

    # Close the tqdm progress bar
    progress_bar.close()

    # Create a DataFrame with image names and extracted texts
    data = {'Image Name': image_names, 'Extracted Text': extracted_texts}
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    excel_file_path = os.path.join(directory_path, 'extracted_text1.xlsx')
    df.to_excel(excel_file_path, index=False)

    print("Text extraction complete. Results saved to", excel_file_path)