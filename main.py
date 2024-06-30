from tkinter import *
from tkinter import filedialog, messagebox
from docx2pdf import convert


word = ""
save = ""
wordCheck = False
saveCheck = False
#functions
def enableConvert():
    if wordCheck and saveCheck:
        redButton.config(state=ACTIVE)

def get_word_file():
    global word, wordCheck
    word = filedialog.askopenfilename(filetypes=[("Word files", "*.doc;*.docx")])
    if word:
        wordCheck = True
    enableConvert()

def get_save_path():
    global save, saveCheck
    save = filedialog.askdirectory()
    if save:
        saveCheck = True
    enableConvert()

def convert_to_pdf():
    convert(word, save)
    messagebox.showinfo(title="", message="Successfully converted!")

#Colors
BACKGROUND_COLOR = "#333333"
RED_BUTTON_COLOR = "red"
FOREGROUND = "white"
BLUE_BUTTON_COLOR = "#5f9cd9"
GREEN_BUTTON_COLOR = "#07b56a"
BUTTON_WIDTH = 15

#widget creation
window = Tk()
icon = PhotoImage(file="Icon.png")
window.geometry("800x450")
window.resizable(False, False)
window.title("W2PDF")
window.iconphoto(True, icon)
window.config(background=BACKGROUND_COLOR)

frame = Frame(pady=20,
                background=BACKGROUND_COLOR)
header = Label(window, 
                text="Word 2 PDF", 
                background=BACKGROUND_COLOR, 
                foreground=FOREGROUND, 
                font=("sans-serif", 40, "bold"))
blueButton = Button(frame,
                    activebackground=BLUE_BUTTON_COLOR,
                    background=BLUE_BUTTON_COLOR,
                    border=0,
                    text="Select Word file",
                    foreground=FOREGROUND,
                    font=("sans-serif", 20, "bold"),
                    width=BUTTON_WIDTH,
                    cursor="hand2",
                    command=get_word_file)
greenButton = Button(frame,
                    activebackground=GREEN_BUTTON_COLOR,
                    background=GREEN_BUTTON_COLOR,
                    border=0,
                    text="Select save path",
                    foreground=FOREGROUND,
                    font=("sans-serif", 20, "bold"),
                    width=BUTTON_WIDTH,
                    cursor="hand2",
                    command=get_save_path)
redButton= Button(frame,
                    activebackground=RED_BUTTON_COLOR, 
                    background=RED_BUTTON_COLOR, border=0, 
                    text="Convert 2 PDF", 
                    foreground=FOREGROUND, 
                    font=("sans-serif", 20, "bold"),
                    width=BUTTON_WIDTH,
                    cursor="hand2",
                    state=DISABLED,
                    command=convert_to_pdf)
        
#widget placing
header.pack(pady=10)
frame.pack()
blueButton.pack(pady=20)
greenButton.pack(pady=20)
redButton.pack(pady=20)

window.mainloop()