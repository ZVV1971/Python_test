from tkinter import Tk, Label, Button, Menu, Menubutton, messagebox, Frame, Scrollbar
from tkinter import Listbox as Listbox
import tkinter.filedialog as Dialog
import xlrd
import pprint
import re
import os
import dateparser

#Vlues to test if XLS is really TTN
TEST_ROW = 22
TEST_COLUM = 6
TEST_VALUE = 'I. ТОВАРНЫЙ РАЗДЕЛ'
#Begin of cells with additional info
INFO_ROW = 28          #Row where positions begin to appear
INFO_COLUMN = 13       #Column where additional info about positions is stored
CONSIGNEE_ROW= 16      #Row where consignee is stored
CONSIGNEE_COLUMN = 3   #Column where consignee is stored
AGREEMENT_ROW = 17     #Row where agreement inscription is stored
AGREEMENT_COLUMN = 3   #Column where agreement inscription is stored
TTN_NUMBER_ROW = 3     #Row where TTN number is stored
TTN_NUMBER_COLUMN = 13 #Column where TTN number is stored
DATE_ROW = 10          #Row with TTN Date
DATE_COLUMN = 1        #Column with TTN date
MATCH_PATTERN = r'(\d{8})' #QC match pattern
LAST_ROW_IDENTIFIER = 'С товаром переданы документы:'
COPY_TEST_COLUMN = 1 #colum where inscription of
                     #complacency certificates copy is stored
COPY_TEST_START_ROW = 29 #Begin of the possible last row
COPY_TEST_END_ROW = 1000 #Supposed end of the TTN
ADDRESS_BEGINNING_PATTERN = 'г\.|Минская\w|Витебская\w|Могилевская\w|Гомельская\w|Гродненская\w|Брестская\w|Республика\w'
COMPLACENCY_CERT_PATTERN = 'BY\/\d{3}\s(?:\d{2}\.){2}\d{3}\s(\d{5})'

#Need to define dictionary throwing an exception
#when duplicate key is added
#https://stackoverflow.com/questions/4999233/how-to-raise-error-if-duplicates-keys-in-dictionary
class RejectingDict(dict):
    def __setitem__(self, k, v):
        if k in self.keys():
            raise ValueError("Key is already present")
        else:
            return super(RejectingDict, self).__setitem__(k, v)

class MyFirstGUI:

    fileData= RejectingDict()
    last_statusbar_value=''
        
    def __init__(self, master):
        self.master = master
        master.title("Сертификаты из ТТН")
        master.geometry("300x100")
        master.grid_columnconfigure(index = 0, minsize = 350, weight = 5)
        #master.grid_columnconfigure(index = 1, weight = 0)
        menu = Menu(master, tearoff=0)
        #need to store submenu to be able to address it from others subs
        self.submenu2 = Menu(master, tearoff=0)
        self.submenu3 = Menu(master, tearoff=0)
        self.submenu4 = Menu(master, tearoff=0)

        self.menubar=Menu(menu, tearoff=0)
        
        menu.add_command(label="Open", command=self.open)
        menu.add_separator()
        menu.add_command(label="Exit", command=master.destroy)
        #cascade index = 0
        self.menubar.add_cascade(label="File", menu=menu, state="normal")

        self.submenu2.add_command(label="Copies", command=self.copies, state="disabled")

        #cascade index = 1
        self.menubar.add_cascade(label="Operations", menu=self.submenu2)

        self.submenu4.add_command(label="Options", command=self.options)
        #cascade index = 2
        self.menubar.add_cascade(label="Settings", menu=self.submenu4)

        self.commandAbout = self.submenu3.add_command(label="About", command=self.about)
        #cascade index = 3
        self.menubar.add_cascade(label="Help", menu=self.submenu3)

        self.lbMain = Listbox(master, selectmode='extended',
                              state= "disabled", width =50, #width in characters
                              height = 5 #number of lines
                              )
        self.scrollbar = Scrollbar(master, orient = 'vertical')
        self.lbMain.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command= self.lbMain.yview)
        self.scrollbar.grid(row = 0, column =0, ipady = 0,
                            sticky ='e' # just to the right side East
                            )
        
        self.lbMain.grid(row = 0, column = 0,
                         sticky = 'n, s, e, w' #to all sides
                         )

        # status bar
        status_frame = Frame(master, height = 10)
        self.status = Label(status_frame, text="this is the status bar")
        self.status.pack(fill="both", expand = True)
        status_frame.grid(row = 1, column = 0)
        self.lbMain.bind('<<ListboxSelect>>', self.on_lbSelect)
        self.submenu3.bind('<<MenuSelect>>', self.about_status)
        self.submenu3.bind('<Enter>', self.about_status)
        self.submenu4.bind('<<MenuSelect>>', self.options_status)
        self.submenu4.bind('<Enter>', self.options_status)
        self.submenu3.bind('<Leave>', self.status_leave)
        self.submenu4.bind('<Leave>', self.status_leave)

        master.config(menu=self.menubar)

    def about_status(self, event):
        self.last_statusbar_value = self.status['text']
        self.status['text'] = 'About'
        print(event.widget)

    def options_status(self, event):
        self.last_statusbar_value = self.status['text']
        self.status['text'] = 'Options'
        print(event.widget)

    def status_leave(self,event):
        self.status['text'] = self.last_statusbar_value
        
    def about(self):
        messagebox.showinfo(title="About", message="Program reads TTNs from Excel files " +
                            "and prints either quality certificates Consignee-wise or prepares" +
                            " PDFs with complacency certificates on the same basis")
    def options(self):
        pass
    
    def open(self):
        print("Opening files")
 
        filename = Dialog.askopenfilename(initialdir = os.path.dirname(__file__),
                                          title="Файлы с ТТН",
                                          #need to leave comma to build 1-x tuple
                                          filetypes = (("Excel files","*.xls"),),
                                          multiple=True
                                          )
        if len(filename) == 0: return
        
        for nm in filename:
            try:
                ob = TTNReader(nm)
                self.fileData[ob.TTN_data["TTN_Number"]] = ob
                #filling in the listbox
                if self.lbMain.size() == 0:
                    self.lbMain.config(state='normal')
                self.lbMain.insert('end', str(ob.TTN_data["TTN_Number"]))
            except:
                continue

    def on_lbSelect(self, evt):
        if len(self.lbMain.curselection()) !=0:
            self.submenu2.entryconfig('Copies', state="normal")
        else:
            self.submenu2.entryconfig('Copies', state="disabled")

    def copies(self):
        """
        Prepares PDFs with inscriptions
        """
        for ttn in self.lbMain.curselection():
            print(self.fileData[int(self.lbMain.get(ttn))].TTN_data["path"])
        
class TTNReader(object):
    
    def __init__(self, path):
        self.TTN_data = {}
        if self.CheckCorrectExcelFile(path):
            self.TTN_data["path"]=path
            self.ReadDataFromTTN(path)
        else:
            raise ValueError()

    def __repr__(self):
        return pprint.pformat(self.TTN_data)
    
    def CheckCorrectExcelFile(self, path_to_file):
        """
        Checks if the opened Excel file is a file with TTN

        returns True if correct
        otherwise returns False
        """

        res =False

        try:
            with xlrd.open_workbook(path_to_file,
                                    encoding_override='cp1251') as book:
                sheet = book.sheets()[0]
                #print(sheet.cell(TEST_ROW, TEST_COLUM).value)
                if sheet.cell(TEST_ROW, TEST_COLUM).value == TEST_VALUE:
                    res = True
                else:
                    res = False
        except:
            pass
        finally:
            return res
    def ReadDataFromTTN(self, path_to_file):
        """
        Reads information from TTN
        and stores it in the dictionary
        """

        cert_array={}
        comp_cert_array={}

        try:
            with xlrd.open_workbook(path_to_file,
                                    encoding_override='cp1251') as book:
                sheet = book.sheets()[0]
                #reading other data from the file
                self.TTN_data["Agreement"] = sheet.cell(AGREEMENT_ROW, AGREEMENT_COLUMN).value
                self.TTN_data["TTN_Number"] = int(sheet.cell(TTN_NUMBER_ROW, TTN_NUMBER_COLUMN).value)
                Cons = sheet.cell(CONSIGNEE_ROW, CONSIGNEE_COLUMN).value
                matches = re.finditer(ADDRESS_BEGINNING_PATTERN, Cons, re.M)
                for match in matches:
                     self.TTN_data["Consignee"] = Cons[:match.start() - 1]
                     #need only the first
                     break
                self.TTN_data["TTN_Date"] = dateparser.parse(sheet.cell(DATE_ROW, DATE_COLUMN).value)
                self.TTN_data["Certificates"] = {}
                self.TTN_data["CompCertificates"] = {}

                #reading addtional info containg information about certificates
                delta = 0
                while True:
                    cell_value = sheet.cell(INFO_ROW + delta, INFO_COLUMN).value
                    if cell_value != '' and cell_value[0] in 'Сс':
                        cert_search_res = re.findall(MATCH_PATTERN, cell_value, re.M)
                        try:
                            #only unique cert numbers needed for the file
                            for cer in cert_search_res:
                                cert_array[cer] = ''
                        except:
                            pass
                        delta += 1
                    else:
                        break
                    
                #finding cell with inscription of complacency certificates copy
                l = len(LAST_ROW_IDENTIFIER)
                for i in range(COPY_TEST_START_ROW , COPY_TEST_END_ROW):
                    if str(sheet.cell(i, COPY_TEST_COLUMN).value)[:l] == LAST_ROW_IDENTIFIER:
                        inscr = sheet.cell(i, COPY_TEST_COLUMN).value
                        break
                matches2 = re.findall(COMPLACENCY_CERT_PATTERN, inscr, re.M)
                for match2 in matches2:
                    try:
                        comp_cert_array[match2]=''
                    except:
                        pass
                                
        except:
            pass
        finally:
            self.TTN_data["Certificates"] = cert_array
            self.TTN_data["CompCertificates"] = comp_cert_array

###############################
# Main Part
###############################
root = Tk()
my_gui = MyFirstGUI(root)
root.mainloop()
