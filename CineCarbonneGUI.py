import tkinter
from tkinter import ttk,scrolledtext,filedialog
import os,sys
import shutil
from tkinter.constants import HORIZONTAL

#Local Import
import normalize
import enrich
import excel_to_json

TMDB_API_KEY=os.getenv ("TMDB_API_KEY","4b400d47b0a36eed006040846feebaf5")
selected_program_file=""

mode_standalone=os.path.isfile("./CineCarbonneGUI.exe")

if mode_standalone :
    os.makedirs("CineCarbonne", exist_ok=True)
    os.chdir("CineCarbonne")

INPUT_PATH  = "input/source.xlsx"


if not os.path.isdir("input") :
    os.mkdir("input")

if not os.path.isdir("work") :
    os.mkdir("work")

if not os.path.isdir("public") :
    os.mkdir("public")


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
    label_file_explorer.configure(text=selected_program_file)

def convert():
    print(f"Répertoire courant début convert: {os.getcwd()}")
    if selected_program_file != "":
        print ("Copy selected file to input/source.xlsx")
        shutil.copy(selected_program_file,INPUT_PATH)
    else :
        print("ATTENTION : pas de fichier d'entrée selectionné le fichier source.xlsx courant sera utilisé si existant!! ")

    print(f"Répertoire courant début normalize: {os.getcwd()}")
    if options["normalize"].get():
     print ("Normalize")
     normalize.main()

     if options["enrich"].get():
        print(f"Répertoire courant début enrich: {os.getcwd()}")
        print ("ajout TMDB data witk key %s __" %input_TMDB.get('1.0',"end"))
        os.environ["TMDB_API_KEY"]=TMDB_API_KEY
        enrich.main(window)

    if options["export"].get():
        print ("convert to json")
        excel_to_json.main()

    print (" TBD .. move to site GitHub (and Commit ? )")


# Create the root window
window = tkinter.Tk()

#define Styles
ttk.Style().configure("TButton", font=("helvetica", 12), padding=6, relief="raised", lighcolor="blue")
ttk.Style().configure("TLabel", font=("helvetica", 12), padding=6)

# Set window title
window.title('CinéCarbonne - Site - Programme')

# Set window size
window.geometry("800x250")

# Set window background color
#window.config(background="lightgrey")

# Create a File Explorer label

button_explore = ttk.Button(window,
                        text="Selection du fichier Programme source",
                        command=browseFiles)

label_file_explorer = ttk.Label(window,
                            text="          Programme (.xlsx)            ",
                            style="BW.TLabel",
                            width=50)

# champ texte pour Modifier la clée TMDB par defaut
label_TMDB = ttk.Label(window,
                            text="Clée TMDB: ",
                            style="BW.TLabel")
input_TMDB = tkinter.Text(window, width=40, height=1)
input_TMDB.insert(1.0,TMDB_API_KEY)

#checkBoxe pour le  choix des etapes de conversion
options = {"normalize": tkinter.BooleanVar(),
           "enrich": tkinter.BooleanVar(),
           "export": tkinter.BooleanVar()}

options["normalize"].set(True)
options["enrich"].set(os.path.isfile('work/normalized.xlsx'))
options["export"].set(False)

normalizeRB = ttk.Checkbutton(window, text="normalisation du fichier Excell brut ", variable=options["normalize"])
enrichRB = ttk.Checkbutton(window, text="enrichissement auto (synopsis, Lien allociné,..) ", variable=options["enrich"])
exportRB = ttk.Checkbutton(window, text="export Site CineCarbonne", variable=options["export"])

#Bouuton pour lancer la conversion du fichier d'entrée
button_convert = ttk.Button(window,
                     text="convert",
                     command=convert)

button_quit = ttk.Button(window,
                     text="Quit",
                     command=window.destroy)

# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns

button_explore.grid(column=1, row=1, padx=15, pady=10, columnspan=2, sticky="e")
label_file_explorer.grid(column=3, row=1, pady=10,columnspan=2, sticky="ew")

label_TMDB.grid(column=1,row=3, padx=5, pady=10, columnspan=2)
input_TMDB.grid(column=3,row=3, pady=10,columnspan=2)

normalizeRB.grid(column=2,row=4,sticky="w", pady=5)
ttk.Separator(window, orient=HORIZONTAL).grid(column=1, row=5, columnspan=5, pady=5, sticky="we"  )
enrichRB.grid(column=2,row=6,sticky="w")
exportRB.grid(column=2,row=7,sticky="w")
button_convert.grid(column=3, row=6, padx=5, pady=10)
button_quit.grid(column=4, row=6, padx=5, pady=10)

window.attributes("-topmost", True)
window.mainloop()

