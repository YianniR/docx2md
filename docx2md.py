import os
import argparse
import PySimpleGUI as sg
from PIL import Image, ImageTk
import pypandoc

def count_files_and_folders(folder_path, recursive):
    if recursive:
        num_folders = sum([len(dirs) for _, dirs, _ in os.walk(folder_path)])
        num_files = sum([len(files) for _, _, files in os.walk(folder_path)])
    else:
        num_folders = len([f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))])
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
    for i, filename in enumerate(docx_files):
        for docx_file in docx_files:
            if os.path.isdir(docx_file):
                continue

            md_path = os.path.splitext(docx_file)[0] + '.md'
            if media_extraction:
                extra_args=['--wrap=none', f'--extract-media={media_folder}']
            else:
                extra_args=['--wrap=none']
            pypandoc.convert_file(docx_file, 'md', format='docx', outputfile=md_path, extra_args=extra_args)
            yield i + 1, len(docx_files)

def browse_folder():
    folder_path = sg.popup_get_folder("Select a folder")
    folder_entry.update(value=folder_path)

def update_progress_bar(value):
    progress_bar.update(value)

def convert():
    folder_path = folder_entry.get()
    recursive = recursive_var.get()
    media_extraction = media_extraction_var.get()
    for i, total in convert_docx_to_md(folder_path, recursive, media_extraction):
        progress_num = int(100 * i / total)
        percentage_label.update(value=progress_num)
        update_progress_bar(progress_num)
    percentage_label.update(value="Done!")

# Create the main GUI window
layout = [
    [sg.Image('logo.png')],
    [sg.Text("Select Folder:")],
    [sg.Input(), sg.FolderBrowse(), sg.Button("Convert")],
    [sg.Checkbox("Search subfolders"), sg.Checkbox("Extract media")],
    [sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')],
    [sg.Text("0%", key='percentage_label')]
]

window = sg.Window("DOCX to Markdown Converter", layout)
window.set_icon("logo.png")

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    if event == "Convert":
        folder_path = values[1]
        recursive = values[2]
        media_extraction = values[3]
        for i, total in convert_docx_to_md(folder_path, recursive, media_extraction):
            progress_num = int(100 * i / total)
            window['progressbar'].update(progress_num)
            window['percentage_label'].update(f"{progress_num}%")
        window['percentage_label'].update("Done!")

window.close()
