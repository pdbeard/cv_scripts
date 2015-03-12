import os
import sys
import mmap
import imp
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import xlrd
from xml.sax.saxutils import quoteattr as xml_quoteattr

def stage():
    #Initilize window
    gui = Tk()
    gui.geometry('350x200+500+300')
    gui.title('Import Library')

    def load_dir():  #Directory Selection Location
        global sel_dir
        sel_dir = filedialog.askdirectory()
        dir_title = Label(gui,text = "Selected Directory : " ).grid(row = 1, column = 0, sticky = W)
        chosen_dir = Label(gui,text = sel_dir ).grid(row = 1, column = 1, columnspan = 3, sticky = W)
        return

    def load_excel():  #Excel Selection Location
        global sel_excel
        sel_excel = filedialog.askopenfile()
        excel_title = Label(gui,text = "Selected Excel File : " ).grid(row = 3, column = 0, sticky = W)
        chosen_excel = Label(gui,text = sel_excel.name ).grid(row = 3, column = 1, columnspan = 3, sticky = W)
        return

    def opt_set():
        g_img        = c_img.get().lower()
        g_title      = c_title.get().lower()
        g_desc       = c_desc.get().lower()
        g_collection = c_collection.get().lower()
        g_subcollection = c_subcollection.get().lower()

        global ex_opt
        ex_opt = [g_img, g_title, g_desc, g_collection, g_subcollection]

        if ex_opt[0]=="":
            messagebox.showwarning(title = "Error", message = "You must select a column for at least the file names")

        try:
            sel_dir
            sel_excel

        except NameError:
            messagebox.showwarning(title = "Error", message = "You must choose a directory , excel sheet, and at least the File Names!")
        else:
            for i in range (0,5):
                if ex_opt[i]=="":
                    ex_opt[i] = -1
                else:
                    ex_opt[i] = ord(ex_opt[i])-97
            parse()
            gui.quit()
        return

    #Choose image library directory
    labelDir  = Label(gui,text ="Find Image Directory").grid(row = 0, column = 0, sticky = W)
    buttonDir = Button(gui, text ="Browse", command = load_dir).grid(row = 0, column = 1, sticky = W)

    #Choose Excel sheet
    labelExcel  = Label(gui,text ="Find Excel File").grid(row = 2, column = 0, sticky = W)
    buttonExcel = Button(gui, text ="Browse", command = load_excel).grid(row = 2, column = 1, sticky = W)

    #Excel Column Options
    labelMatch = Label(gui,text = "Match excel column letters to desired image attributes.").grid(row = 4, column = 0, columnspan = 4, sticky = W)

    d_img        = Label(gui,text = "File names").grid(row = 5, column = 0, sticky = W)
    d_title      = Label(gui,text = "Titles").grid(row = 5, column = 1, sticky = W)
    d_desc       = Label(gui,text = "Descriptions").grid(row = 7, column = 0, sticky = W)
    d_collection = Label(gui,text = "Collections").grid(row = 7, column = 1, sticky = W)
    d_subcollection = Label(gui,text = "Subcollection").grid(row = 9, column = 0, sticky = W)

    c_img = StringVar()
    c_title = StringVar()
    c_desc = StringVar()
    c_collection = StringVar()
    c_subcollection = StringVar()

    e_img        = Entry(gui, textvariable = c_img).grid(row = 6, column = 0, sticky = W)
    e_title      = Entry(gui, textvariable = c_title).grid(row = 6, column = 1, sticky = W)
    e_desc       = Entry(gui, textvariable = c_desc).grid(row = 8, column = 0, sticky = W)
    e_collection = Entry(gui, textvariable = c_collection).grid(row = 8, column = 1, sticky = W)
    e_subcollection = Entry(gui, textvariable = c_subcollection).grid(row = 10, column = 0, sticky = W)

    buttonDone = Button(gui, text = "Generate CML File", command = opt_set).grid(row =14, column = 0)
    mainloop()
    return 1

def parse():
    #Open excel file and sets variable sh to the first worksheet
    wb=xlrd.open_workbook(sel_excel.name)
    sh = wb.sheet_by_index(0)


    #Stores the data from colums of the selected row
    def find_info(row):
        img=[]

        for i in range (1,5):
            if ex_opt[i] < 0:
                img.append("")
            else:
                try:
                    img.append(sh.cell(rowx=row, colx=ex_opt[i]).value)
                except UnicodeEncodeError:
                    img_app = img.append(sh.cell(rowx=row, colx=ex_opt[i]).value)
                    img_app.encode('ascii','xmlcharrefreplace')

        info_list = [img[0], img[1], img[2], img[3]]
        return info_list

    #Loops through excel sheet rows and generates data content for each row.
    def gencon(path):
		#Determines how many rows there are
        column = len(sh.col_values(ex_opt[0]))
        sourceId = 0
        result=""

        for row in range(column):
            #Increments SourceID
            sourceId = sourceId + 1
            result += '<img>\n<id>%d</id>\n' % sourceId

            #Appends Picture name to Directory Path
            cell_value = sh.cell(rowx=row,colx=ex_opt[0]).value
            localPath = path+'/'+cell_value

            #Insurance Parantheses are correct
            parenSwitch = localPath.replace('\\','/')
            info_list = find_info(row)

            result += '	<src>%s</src>\n' % (parenSwitch[cleanPath:])
            result += '''\
		<title>%s</title><!-- Sets the image title -->
		<description>%s</description><!-- Sets the image description-->
		<collection>%s</collection><!-- must match collection buttons. to include multiple collections, separate values with commas -->
		<subcollection>%s</subcollection><!-- Sets the image author data -->
	</img>\n''' % (info_list[0],info_list[1],info_list[2],info_list[3])
            result += '\n'
        return result


    #Number of characters to delete for path to start at local dir (location of the script).
    def refinePath(path):
        fullPath = path
        startPath = os.path.basename(path)
        changeNum = len(fullPath) - len(startPath)
        #localPath = fullPath[changeNum:]
        return changeNum


    #Default settings for the collection Viewer
    end = '''\

</!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!>

    '''
    outfile = open('../ImageViewer_Config_template.cml','w')
    global cleanPath
    cleanPath = refinePath(os.getcwd())
    print ('Creating XML Template...')
    print ('<?xml version="1.0" encoding="UTF-8"?>\n<cml>\n <!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!>\n' + gencon(sel_dir) + end, file = outfile)
    print ('\nDone!')
    return

if __name__ == '__main__':
    stage()



