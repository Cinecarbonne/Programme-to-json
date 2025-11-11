import tkinter
from tkinter import ttk
from tkinter  import filedialog
import os
import shutil

#Local Import
import normalize
import enrich
import excel_to_json

TMDB_API_KEY=os.getenv ("TMDB_API_KEY","4b400d47b0a36eed006040846feebaf5")
selected_program_file=""
INPUT_PATH  = "input/source.xlsx"


if not os.path.isdir("input") :
    os.mkdir("input")

if not os.path.isdir("work") :
    os.mkdir("work")


# Function for opening the
# file explorer window
def browseFiles():
    global selected_program_file
    selected_program_file = filedialog.askopenfilename(initialdir="",
                                          title="Select a File",
                                          filetypes=(("Excel2007 files",
                                                      "*.xlsx"),
                                                     ("all files",
                                                      "*.*")))

    # Change label contents
    label_file_explorer.configure(text="Programme: " + selected_program_file)

def convert():
    if selected_program_file != "":
        print ("Copy selected file to input/source.xlsx")
        shutil.copy(selected_program_file,INPUT_PATH)
    else :
        print("ATTENTION : pas de fichier d'entrée selectionné le fichier source.xlsx courant sera utilisé si existant!! ")
    print ("Normalize")
    normalize.main()
    print ("ajout TMDB data witk key %s __" %input_TMDB.get('1.0',"end"))
    os.environ["TMDB_API_KEY"]=TMDB_API_KEY
    enrich.main(True)
    print ("convert to json")
    excel_to_json.main()

    print (" TBD .. move to site GitHub (and Commit ? )")


# Create the root window
window = tkinter.Tk()

#define Styles
ttk.Style().configure("TButton", font=("helvetica", 14), padding=6, relief="raised", lighcolor="blue")
ttk.Style().configure("TLabel", font=("helvetica", 14), padding=6)

# Set window title
window.title('CinéCarbonne - Site - Programme')

# Set window size
window.geometry("800x200")

# Set window background color
#window.config(background="lightgrey")

# Create a File Explorer label
label_file_explorer = ttk.Label(window,
                            text="Programme (.xlsx): ?",
                            style="BW.TLabel")
label_TMDB = ttk.Label(window,
                            text="Clée TMDB: ",
                            style="BW.TLabel")
input_TMDB = tkinter.Text(window, width=40, height=1)
input_TMDB.insert(1.0,TMDB_API_KEY)

button_explore = ttk.Button(window,
                        text="Selection du fichier Programme source",
                        command=browseFiles)

button_convert = ttk.Button(window,
                     text="Conversion Json",
                     command=convert)

# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns

button_explore.grid(column=1, row=1, padx=5, pady=10)
label_file_explorer.grid(column=2, row=1, padx=5, pady=10)

label_TMDB.grid(column=1,row=3, padx=5, pady=10)
input_TMDB.grid(column=2,row=3, padx=5, pady=10)
button_convert.grid(column=1, row=6, padx=5, pady=10)

window.mainloop()

