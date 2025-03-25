import tkinter as tk
from tkinter import filedialog, messagebox
import os
from docx import Document
from itertools import groupby

#App Folder Selection Component
root = tk.Tk()
root.title("Transcription Highlight Widget")
root.geometry("400x400")

labela = tk.Label(root, text="Data Folder", font=("Arial", 12))
labela.pack(pady=20)
label1 = tk.Label(root, text="NONE SELECTED", font=("Arial", 12))
label1.pack(pady=20)
labelb = tk.Label(root, text="Output Folder", font=("Arial", 12))
labelb.pack(pady=20)
label2 = tk.Label(root, text="NONE SELECTED", font=("Arial", 12))
label2.pack(pady=10)

out_num = 0
folder_paths = {'folder1': None, 'folder2': None}
with open("file_num.txt", "r") as f:
    out_num = int(f.read())

def open_folder_picker(folder_num, label):
    try:
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path:
            label.config(text=f"Folder: {folder_path}")
            folder_paths[folder_num] = folder_path
        else:
            label.config(text="NONE SELECTED")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while selecting the folder: {e}")
        label.config(text="NONE SELECTED")
    print(folder_paths)
    output_folder_path = tk.StringVar()

def compare_time(out_num):
    #Grab All Files
    if folder_paths['folder1'] is None or folder_paths['folder2'] is None:
        messagebox.showerror("Error", "Please select both folders before comparing.")
        return
    file_reads = {}
    input = folder_paths['folder1']
    for file in os.listdir(input):
        if file.lower().endswith(('.doc', '.docx')):
            details = file.split("-")
            key = details[0] + details[1]
            if key not in file_reads.keys():
                file_reads[key] = []
            file_reads[key].append(file)
            print(file_reads)
            print(details)
    print(input, file_reads)

    for key, values in file_reads.items():
        output_path = folder_paths['folder2'] + "/COMBINED_" + key + ".docx"
        print(type(values), values)
        if len(values) == 2:
            compare_2(values, output_path)
        elif len(values) == 3:
            print(3)
        else:
            print(key, "needs to have 2 or 3 values")

def compare_2(values, output_path):
    file1 = Document(values[0])
    file2 = Document(values[1])
    outfile = Document()
    outfile._body.clear_content()

    #Loop through the paragraphs and compare them
    for para1, para2 in zip(file1.paragraphs, file2.paragraphs):
        runs1 = [run.text for run in para1.runs]
        runs2 = [run.text for run in para2.runs]
        final_text = ''.join(runs1)
        final_map1 = ""
        final_map2 = ""
        
        for run in para1.runs:
            if run.font.highlight_color == None:
                final_map1 += "N" * len(run.text)
            else:
                final_map1 += "Y" * len(run.text)

        for run in para2.runs:
            if run.font.highlight_color == None:
                final_map2 += "N" * len(run.text)
            else:
                final_map2 += "Y" * len(run.text)
        map_FINAL = ""

        #Calculate overlaps
        for char1, char2 in zip(final_map1, final_map2):
            if char1 == char2:
                if char1=="Y":
                    map_FINAL += "G"
                elif char1=="N":
                    map_FINAL += "N"
            else:
                map_FINAL += "Y"
        
        final_group = [(char, len(list(group))) for char, group in groupby(map_FINAL)]

        #Now that we know the highlighting, create the output paragraph

        p = outfile.add_paragraph('')

        for group in final_group:
            r = p.add_run(final_text[:group[1]])
            final_text = final_text[group[1]:]
            if group[0] == "N":
                r.font.highlight_color = None
            elif group[0] == "Y":
                r.font.highlight_color = 7  # Yellow
            elif group[0] == "G":
                r.font.highlight_color = 4  # Green

    #Save the output file
    outfile.save(output_path)
        
    messagebox.showinfo("Success", f"Comparison saved to {output_file_path}")
    out_num += 1
    with open("file_num.txt", "w") as f:
        f.write(str(out_num))
    
#Decision Buttons
button1 = tk.Button(root, text="Select File 1", command=lambda: open_folder_picker('folder1', label1))
button1.pack(pady=10)
button2 = tk.Button(root, text="Select File 2", command=lambda: open_folder_picker('folder2', label2))
button2.pack(pady=10)
button3 = tk.Button(root, text="Create Comparison", command=lambda: compare_time(out_num))
button3.pack(pady=10)

#Run
root.mainloop()