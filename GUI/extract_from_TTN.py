from tkinter import Tk, Label, Button, Menu, Menubutton, messagebox
import tkinter.filedialog as Dialog
import xlrd
import pprint
import re
import os
import dateparser

root = Tk()

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

class MyFirstGUI:

    fileData= []
        
    def __init__(self, master):
        self.master = master
        master.title("Сертификаты из ТТН")
        master.geometry("300x100")
        menu = Menu(master, tearoff=0)
        menu2 = Menu(master, tearoff=0)
        menu3 = Menu(master, tearoff=0)

        self.menubar=Menu(menu, tearoff=0)
        menu.add_command(label="Open", command=self.open)
        menu.add_separator()
        mm=menu.add_command(label="Exit", command=master.destroy)
        self.menubar.add_cascade(label="File", menu=menu, state="normal")

        menu2.add_command(label="Copies", command=self.copies, state="disabled")
        self.menubar.add_cascade(label="Operations", menu=menu2)

        menu3.add_command(label="About", command=self.about)
        self.menubar.add_cascade(label="Help", menu=menu3)

        master.config(menu=self.menubar)

    def about(self):
        messagebox.showinfo(title="About", message="Program reads TTNs from Excel files " +
                            "and prints either quality certificates Consignee-wise or prepares" +
                            " PDFs with complacency certificates on the same basis")
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
                self.fileData.append(ob)
            except:
                continue
        if len(self.fileData)==0: return    
        self.menubar.entryconfig("Copies", state="normal")

    def copies(self):
        """
        Prepares PDFs with inscriptions
        """
        for ttn in self.fileData:
            print(ttn)
        
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

my_gui = MyFirstGUI(root)
root.mainloop()
