from tkinter import Tk, Label, Button, Menu, Menubutton
import tkinter.filedialog as Dialog

root = Tk()

class MyFirstGUI:
        
    def __init__(self, master):
        self.master = master
        master.title("Сертификаты из ТТН")
        master.geometry("300x100")
        menu = Menu(master, tearoff=0)
        menu2 = Menu(master, tearoff=0)

        menubar=Menu(menu, tearoff=0)
        menu.add_command(label="Open", command=self.open)
        menu.add_separator()
        menu.add_command(label="Exit", command=master.destroy)

        menubar.add_cascade(label="File", menu=menu)

        menu2.add_command(label="Copies")
        
        self.m = menubar.add_cascade(label="Operations", menu=menu2)

        master.config(menu=menubar)

    def open(self):
        print("Opening files")

        filename = Dialog.askopenfilename(initialdir = "/",
                                          title="Файлы с ТТН",
                                          #need to leave comma to build 1-x tuple
                                          filetypes = (("Excel files","*.xls"),),
                                          multiple=True
                                          )
        for nm in filename:
            print(nm)


my_gui = MyFirstGUI(root)
root.mainloop()

