from tkinter import Tk, Label, Button, Menu, Menubutton
import tkinter.filedialog as Dialog
import xlrd
import pprint
import re

root = Tk()

TEST_ROW = 22
TEST_COLUM = 6
TEST_VALUE = 'I. ТОВАРНЫЙ РАЗДЕЛ'
#Begin of cells with additional info
INFO_ROW = 28
INFO_COLUMN = 13
CONSIGNEE_ROW= 16
CONSIGNEE_COLUMN = 3
AGREEMENT_ROW = 17
AGREEMENT_COLUMN = 3
TTN_NUMBER_ROW = 3
TTN_NUMBER_COLUMN = 13
MATCH_PATTERN = r'(\d{8})'


class MyFirstGUI:

    fileData= []
        
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

        menu2.add_command(label="Copies", command=self.copies)
        
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
        if len(filename) == 0: return
        
        for nm in filename:
            try:
                ob = TTNReader(nm)
                self.fileData.append(ob)
            except:
                continue
        if len(self.fileData)==0: return    
        self.menubar.entryconfig("Operations", state="normal")

    def copies(self):
        """
        Prepares PDFs with inscriptions
        """
        for ttn in self.fileData:
            print(ttn)
        
class TTNReader(object):

    
    
    def __init__(self, path):
        TTN_data = {}
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
        """

        cert_array={}

        try:
            with xlrd.open_workbook(path_to_file,
                                    encoding_override='cp1251') as book:
                sheet = book.sheets()[0]
                #reading other data from the file
                self.TTN_data["Agreement"] = sheet.cell(AGREEMENT_ROW, AGREEMENT_COLUMN)
                self.TTN_data["TTN_Number"] = sheet.cell(TTN_NUMBER_ROW, TTN_NUMBER_COLUMN)
                self.TTN_data["Consignee"] = sheet.cell(CONSIGNEE_ROW, CONSIGNEE_COLUMN)
                self.TTN_data["Certificates"] = {}
                #reading addtional info containg informatino about certificates
                delta = 0
                while True:
                    cell_value = sheet.cell(INFO_ROW + delta, INFO_COLUMN).value
                    if cell_value[0] in 'Сс':
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
        except:
            pass
        finally:
            self.TTN_data["Certificates"] = cert_array

my_gui = MyFirstGUI(root)
root.mainloop()

