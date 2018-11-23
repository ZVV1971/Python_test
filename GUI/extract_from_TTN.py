from tkinter import Tk, Label, Button, Menu, Menubutton
import tkinter.filedialog as Dialog
import xlrd
import re

root = Tk()

class CreateToolTip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)
    def enter(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                       background='yellow', relief='solid', borderwidth=1,
                       font=("times", "8", "normal"))
        label.pack(ipadx=1)
    def close(self, event=None):
        if self.tw:
            self.tw.destroy()


class MyFirstGUI:
        
    def __init__(self, master):
        self.master = master
        master.title("Сертификаты из ТТН")
        master.geometry("300x100")
        menu = Menu(master, tearoff=0)
        menu2 = Menu(master, tearoff=0)

        self.menubar=Menu(menu, tearoff=0)
        menu.add_command(label="Open", command=self.open)
        menu.add_separator()
        mm=menu.add_command(label="Exit", command=master.destroy)
        
        self.menubar.add_cascade(label="File", menu=menu, state="normal")

        menu2.add_command(label="Copies")
        
        self.menubar.add_cascade(label="Operations", menu=menu2,
                                     state="disabled")

        master.config(menu=self.menubar)

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
        self.menubar.entryconfig("Operations", state="normal")

my_gui = MyFirstGUI(root)
root.mainloop()

