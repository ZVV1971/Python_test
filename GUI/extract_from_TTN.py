from tkinter import Tk, Label, Button, Menu, Menubutton

root = Tk()

class MyFirstGUI:
        
    def __init__(self, master):
        self.master = master
        master.title("Сертификаты из ТТН")
        master.geometry("300x100")
        main_menu = Menu(master)

    def hello(self):
        pass

my_gui = MyFirstGUI(root)
root.mainloop()

