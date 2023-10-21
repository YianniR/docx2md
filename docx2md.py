import os
import argparse
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from PIL import Image, ImageTk
import pypandoc

def count_files_and_folders(folder_path, recursive):
    if recursive:
        num_folders = sum([len(dirs) for _, dirs, _ in os.walk(folder_path)])
        num_files = sum([len(files) for _, _, files in os.walk(folder_path)])
    else:
        num_folders = len([f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))] )
        num_files = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])

    return num_folders, num_files

def convert_docx_to_md(folder_path, recursive, media_extraction):
    num_folders, num_files = count_files_and_folders(folder_path, recursive)

    if recursive:
        docx_files = []
        for root, dirs, files in os.walk(folder_path):
            docx_files += [os.path.join(root, f) for f in files if f.endswith('.docx')]
    else:
        docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.docx')]

    num_docx_files = len(docx_files)

    print(f'Total folders found: {num_folders}')
    print(f'Total files found: {num_files}')
    print(f'Total DOCX files found: {num_docx_files}')

    # Specify the media extraction folder
    media_folder = os.path.join(folder_path, "media")
    os.makedirs(media_folder, exist_ok=True)

    # with tqdm(total=num_docx_files, desc='Converting DOCX to MD', unit='file') as pbar:
    for i, filename in enumerate(docx_files):
        for docx_file in docx_files:
            if os.path.isdir(docx_file):
                continue

            md_path = os.path.splitext(docx_file)[0] + '.md'
            if media_extraction:
                extra_args=['--wrap=none', f'--extract-media={media_folder}']
            else:
                extra_args=['--wrap=none']
            pypandoc.convert_file(docx_file, 'md', format='docx', outputfile=md_path, extra_args=extra_args )
            yield i + 1, len(docx_files)
            # pbar.update(1)

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, "end")
    folder_entry.insert(0, folder_path)

def update_progress_bar(value):
    progress_bar['value'] = value

def convert():
    folder_path = folder_entry.get()
    recursive = recursive_var.get()
    media_extraction = media_extraction_var.get()

    for i, total in convert_docx_to_md(folder_path, recursive, media_extraction):
        progress_num=int(100 * i / total)
        percentage_label.config(text=f"{progress_num}%")
        update_progress_bar(progress_num)
        root.update_idletasks()

    percentage_label.config(text=f"Done!")

# Create the main GUI window
root = tk.Tk()
root.title("DOCX to Markdown Converter")

# Set a custom icon (replace 'icon.png' with your image file)
icon = Image.open("logo.png")
root.iconphoto(True, ImageTk.PhotoImage(icon))

# Create and configure the frame
frame = ttk.Frame(root, padding=10)
frame.grid(column=0, row=0, padx=20, pady=20)

# Open the logo image
logo_image = Image.open("logo.png")

# Define the new size for your logo
new_width = 100  # Replace with the desired width
new_height = 100  # Replace with the desired height

# Resize the logo image
logo_image = logo_image.resize((new_width, new_height))

# Convert the resized image to PhotoImage
logo_photo = ImageTk.PhotoImage(logo_image)

# Create a label with the resized logo image
logo_label = ttk.Label(frame, image=logo_photo)
logo_label.grid(column=0, row=0, columnspan=3, pady=(0, 20))

# Create a section for folder selection
folder_label = ttk.Label(frame, text="Select Folder:")
folder_label.grid(column=0, row=1, columnspan=3, sticky=tk.W)

folder_frame = ttk.Frame(frame)
folder_frame.grid(column=0, row=2, columnspan=3, sticky=(tk.W, tk.E))

folder_entry = ttk.Entry(folder_frame)
folder_entry.grid(column=0, row=0, padx=5, pady=5, sticky=(tk.W, tk.E))
folder_entry.insert(0, os.getcwd())

browse_button = ttk.Button(folder_frame, text="Browse", command=browse_folder)
browse_button.grid(column=1, row=0, padx=5)

# Create a section for options
options_label = ttk.Label(frame, text="Options:")
options_label.grid(column=0, row=3, columnspan=3, sticky=tk.W)

recursive_var = tk.IntVar()
recursive_check = ttk.Checkbutton(frame, text="Search subfolders", variable=recursive_var)
recursive_check.grid(column=0, row=4, columnspan=3, sticky=tk.W)

media_extraction_var = tk.IntVar()
media_extraction_check = ttk.Checkbutton(frame, text="Extract media", variable=media_extraction_var)
media_extraction_check.grid(column=0, row=5, columnspan=3, sticky=tk.W)

# Create a section for the conversion button
convert_button = ttk.Button(frame, text="Convert", command=convert)
convert_button.grid(column=0, row=6, columnspan=3, pady=(20, 0))

# Create a section for result and progress
result_label = ttk.Label(frame, text="")
result_label.grid(column=0, row=7, columnspan=3, pady=10)

# Create a frame to hold the progress bar and percentage label
progress_frame = ttk.Frame(root)
progress_frame.grid(column=0, row=8, columnspan=3, pady=(10, 20), padx=20, sticky=(tk.W, tk.E))

progress_bar = ttk.Progressbar(progress_frame, length=200, mode='determinate')
progress_bar.grid(column=0, row=0, padx=5, pady=5, columnspan=3)

percentage_label = ttk.Label(progress_frame, text="")
percentage_label.grid(column=0, row=1, columnspan=3, pady=5)


root.mainloop()