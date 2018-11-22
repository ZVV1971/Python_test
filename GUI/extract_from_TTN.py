from tkinter import Tk, Label, Button, Menu, Menubutton
import tkinter.filedialog as Dialog

root = Tk()

class MyFirstGUI:
        
    def __init__(self, master):
        self.master = master
        master.title("Сертификаты из ТТН")
        master.geometry("300x100")
        menu = Menu(master, tearoff=0)

        menubar=Menu(menu, tearoff=0)
        menu.add_command(label="Open", command=self.open)
        menu.add_separator()
        menu.add_command(label="Exit", command=master.destroy)

        menubar.add_cascade(label="Test", menu=menu)
        
        master.config(menu=menubar)

    def open(self):
        print("Opening files")
        filename = Dialog.askopenfilename(initialdir = "/",
                                          title="Файлы с ТТН",
                                          #neede to leave comma to build 1-x tuple
                                          filetypes = (("Excel files","*.xls"),),
                                          multiple=True
                                          )
        print(filename)

my_gui = MyFirstGUI(root)
root.mainloop()

