import os
import sys
import mmap
import imp 
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

#newImp = os.listdir(os.getcwd())
#for dir in newImp:
#    if dir == "xlrd-0.6.1":
#        itempath = os.path.join(os.getcwd(), dir)
#        sys.path.append(itempath)
#        break

import xlrd
from xml.sax.saxutils import quoteattr as xml_quoteattr

sourceId = 0 
   
def stage():
    #Initilize window
    gui = Tk()
    gui.geometry('450x350+500+300')
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
        g_img        = int(str(c_img.get())) - 1 
        g_title      = int(str(c_title.get())) - 1
        g_author     = int(str(c_author.get())) - 1
        g_desc       = int(str(c_desc.get())) -1 
        g_year       = int(str(c_year.get())) -1 
        g_location   = int(str(c_location.get())) -1 
        g_publish    = int(str(c_publish.get())) - 1
        g_collection = int(str(c_collection.get())) - 1
    
        global ex_opt
        ex_opt = g_img, g_title, g_author, g_desc, g_year, g_location, g_publish, g_collection
        
        try: 
            sel_dir
            sel_excel
            ex_opt[0] 
            ex_opt[1]
            ex_opt[2] 
            ex_opt[3] 
            ex_opt[4]
            ex_opt[5]
            ex_opt[6] 
            ex_opt[7] 
        except NameError:
            messagebox.showwarning(title = "Error", message = "You must choose a directory , excel sheet, and choose your column numbers")  
        else: 
            parse()
            gui.quit()
        return
    
 
    #Choose image library directory
    LabelDir  = Label(gui,text = "Find Image Directory").grid(row = 0, column = 0, sticky = W)
    buttonDir = Button(gui, text ="Browse", command = load_dir).grid(row = 0, column = 1, sticky = W)
   
    #Choose Excel sheet
    LabelExcel  = Label(gui,text = "Find Excel File").grid(row = 2, column = 0, sticky = W)
    buttonExcel = Button(gui, text ="Browse", command = load_excel).grid(row = 2, column = 1, sticky = W)
    
    #Excel Column Options
    labelMatch = Label(gui,text = "Match Excel Column Numbers to Desired Image attributes.").grid(row = 4, column = 0, columnspan = 4, sticky = W) 
    
    d_img        = Label(gui,text = "File names").grid(row = 5, column = 0, sticky = W)
    d_title      = Label(gui,text = "Titles").grid(row = 5, column = 1, sticky = W)
    d_author     = Label(gui,text = "Authors").grid(row = 7, column = 0, sticky = W)
    d_desc       = Label(gui,text = "Descriptions").grid(row = 7, column = 1, sticky = W)
    d_year       = Label(gui,text = "Years").grid(row = 9, column = 0, sticky = W)
    d_location   = Label(gui,text = "Locations").grid(row = 9, column = 1, sticky = W)
    d_publish    = Label(gui,text = "Publishers").grid(row = 11, column = 0, sticky = W)
    d_collection = Label(gui,text = "Collections").grid(row = 11, column = 1, sticky = W)
    
    c_img = StringVar()
    c_title = StringVar()
    c_author = StringVar()
    c_desc = StringVar()
    c_year = StringVar()
    c_location = StringVar()
    c_publish = StringVar()
    c_collection = StringVar()
    
    e_img        = Entry(gui, textvariable = c_img).grid(row = 6, column = 0, sticky = W)
    e_title      = Entry(gui, textvariable = c_title).grid(row = 6, column = 1, sticky = W)
    e_author     = Entry(gui, textvariable = c_author).grid(row = 8, column = 0, sticky = W)
    e_desc       = Entry(gui, textvariable = c_desc).grid(row = 8, column = 1, sticky = W)
    e_year       = Entry(gui, textvariable = c_year).grid(row = 10, column = 0, sticky = W)
    e_location   = Entry(gui, textvariable = c_location).grid(row = 10, column = 1, sticky = W)
    e_publish    = Entry(gui, textvariable = c_publish).grid(row = 12, column = 0, sticky = W)
    e_collection = Entry(gui, textvariable = c_collection).grid(row = 12, column = 1, sticky = W)
    
    buttonDone = Button(gui, text = "Generate XML File", command = opt_set).grid(row =14, column = 0)
    mainloop()
    return 1

def parse(): 
    #Open excel file and sets variable sh to the first worksheet
    wb=xlrd.open_workbook(sel_excel.name)
    sh = wb.sheet_by_index(0)

    #Counts how many source IDs are located in the former XML file     
    def findSource(item):
        infile = open('%s' % item, 'r')
        for line in infile:
            if '<Source id =' in infile :
                global sourceId 
                sourceId = sourceId + 1
        return 1
        
    #Matches directory filename to the filename used in excel sheet and returns the row   
    def match_file(item): 
        retrow = 0
        file_name = os.path.splitext(item)
        column = len(sh.col_values(0))
        for row in range(column):
            #print(row)
            cell_value = sh.cell(rowx=row,colx=ex_opt[0]).value
            #print (cell_value+' AND '+ file_name[0] )
            if cell_value == file_name[0]:
                print ('File Matched:')
                print (row)
                retrow = row
        return retrow
 
    #Stores the data from colums of the selected row      
    def find_info(row):
        img_title      = sh.cell(rowx=row, colx=ex_opt[1]).value
        img_author     = sh.cell(rowx=row, colx=ex_opt[2]).value
        img_desc       = sh.cell(rowx=row, colx=ex_opt[3]).value
        img_year       = sh.cell(rowx=row, colx=ex_opt[4]).value
        img_location   = sh.cell(rowx=row, colx=ex_opt[5]).value    
        img_publish    = sh.cell(rowx=row, colx=ex_opt[6]).value
        img_collection = sh.cell(rowx=row, colx=ex_opt[7]).value
    
        info_list = [img_title, img_desc, img_author, img_publish, img_collection, img_year, img_location]
        return info_list
    
    #Prints Directory names as "collections"  
    def fetchDir(path):
        print('.')
        result = '<!--Searching Directory:  %s -->\n' % xml_quoteattr(os.path.basename(path))
        for item in os.listdir(path):
            itempath = os.path.join(path, item)
            if os.path.isdir(itempath):
                result += '\n'.join('  ' + line for line in 
                    fetchDir(os.path.join(path, item)).split('\n'))
            #Determines if there is a file present, and if it is a valid picture format
            elif os.path.isfile(itempath):
                if os.path.splitext(itempath)[1].lower() in ('.jpg', '.jpeg', '.png','.gif','.bmp'): #Add any valid fromat to this list
                    global sourceId
                    sourceId = sourceId + 1
                    result += '<Source id = \'%d\'>\n' % sourceId
                    localPath = path+'/'+item
                    parenSwitch = localPath.replace('\\','/')
                    info_list = find_info(match_file(item))
                    result += '	<url>%s</url>\n' % (parenSwitch[cleanPath:])
                    result += '''\
		<qrCodeTag></qrCodeTag><!-- Sets and creates the QR Code tag.  Can be text or url.  If left empty, QR Code Tag is not created. -->
		<title>%s</title><!-- Sets the image title -->
		<description>%s</description><!-- Sets the image description-->
		<author>%s</author><!-- Sets the image author data -->
		<publish>%s</publish><!-- Sets the copyright information. Note: The code for the copyright symbol is " &#169; " so you will want to leave that in the tag if you wish to use the copyright symbol-->
		<collection>%s</collection><!-- must match collection buttons. to include multiple collections, separate values with commas -->
		<year>%s</year> 
		<location>%s</location> 
		<type>Image</type>
	</Source>\n''' % (info_list[0],info_list[1],info_list[2],info_list[3],info_list[4],info_list[5],info_list[6])
        result += '\n'
        return result
        
    #Number of characters to delete for path to start at local dir (location of the script).
    def refinePath(path):  
        fullPath = path
        startPath = os.path.basename(path)
        changeNum = len(fullPath) - len(startPath)
        #localPath = fullPath[changeNum:]
        return changeNum
    end = '''\
    		
	<GlobalSettings><!-- Sets the initial, min and max scale values for the video objects. !!Important: Either <scale> or <imagesNormalize> should be used. Not both !! --
		<!-- <scale></scale> --><!-- Sets the images's full size (the size at which the control buttons appear) as relative to the images's actual size. So, for instance, if an image is 500px x 400px, setting <scale> to 1 would mean its full size equals its actual size. Setting scale to 2 would make it twice as large. -->
		
		<amountToShow>0</amountToShow>
		
		<globalScale>0.25</globalScale><!-- The size images start at before they are double-tapped or zoomed to their full size, which is the size at which the control buttons appear. This variable modulates either <scale> or <imageNormalize>, depending on whether the user wants to set the full size relative to the actual size or as an absolute value-->
		<imagesNormalize></imagesNormalize><!-- Sets the full size as an absolute pixel width; the height of the images are adjusted accordingly. If left blank, will be calculated automatically. -->
		<maxScale>2.5</maxScale><!-- The maximum size the images can zoom to, set relative to their full size within the application (set by <scale> or <imagesNormalize>). For instance, setting maxScale to 2.5 will allow the user to zoom the videos to 2.5 times their full size. -->
		<minScale>.1</minScale><!-- The minimum size the images can zoom to, set relative to their full size within the application (set by <scale> or <imagesNormalize>). For instance, setting minScale to .5 will allow the user to shrink the images to half their full size. -->
		<infoPadding>18</infoPadding><!-- Sets the padding around the description information when the info button is selected, and the spacing between the thumbnail, the main description box, and author info. -->
	</GlobalSettings>
	
	<Gestures><!-- Sets whether or not the named gesture can be used in the application. Value must be either true or false. -->
		<rotate>true</rotate>
		<scale>true</scale>
		<drag>true</drag>
		<doubleTap>true</doubleTap>
		<flick>true</flick>
	</Gestures>

	<ControlBtns>
		<fillColor1>0x333333</fillColor1><!-- Sets the color of the bar behind the control buttons. The preceding Ox ensures that the program changes hexadecimal colors to integers -->
		<outlineColor>0xFFFFFF</outlineColor><!-- Sets the color of the border that surrounds the button control bar. -->
		<cornerRadius>18</cornerRadius>
		<outlineStroke>1</outlineStroke><!-- Sets the width of the border that surrounds the button control bar.. -->

		<buttonRadius>21</buttonRadius><!-- Sets the button radius in pixels. NB: Radius is one-half of total button size (diameter) -->
		<buttonPadding>5</buttonPadding><!-- Sets the padding in pixels on all four sides of each button -->
		<buttonOutlineStroke>1</buttonOutlineStroke><!-- Sets the width of the button outline in pixels -->
		<buttonOutlineColor>0x009966</buttonOutlineColor><!-- Sets the color of the button outline -->
		<buttonColorPassive>0x999999</buttonColorPassive><!-- Sets the background color of the buttons when they aren't active (haven't been selected). -->
		<buttonColorActive>0x555555</buttonColorActive><!-- Sets the background color of the buttons when they are active (have been selected). -->
		<buttonSymbolColor>0x000000</buttonSymbolColor><!-- Sets the color of the symbol or text on each button -->
	</ControlBtns>
	
	<BackgroundGraphic><!-- Sets the application background to be either a solid color (if only <fillColor1> is given a value) or a two-color horizontal gradient (if both color values are filled in).   -->
		<fillColor1>0xFFFFFF</fillColor1>
		<fillColor2>0xCCCCCC</fillColor2>
	</BackgroundGraphic>
	
	<BackgroundOutline><!-- Sets the style of the outline that surrounds each image object -->
		<outlineColor>0xFFFFFF</outlineColor><!-- Sets the color of the outline surrounding each image object. -->
		<outlineStroke>1</outlineStroke><!-- Sets the width of the border around each image object in pixels -->
	</BackgroundOutline>
		
	<InfoText><!-- Sets the font size and color for each metadata field. -->
		<TitleText>
			<fontColor>0x000000</fontColor>
		</TitleText>
		<DescriptionText>
			<fontColor>0x000000</fontColor>
		</DescriptionText>
		<AuthorText>
			<fontColor>0x222222</fontColor>
		</AuthorText>
		<PublishText>
			<fontColor>0x333333</fontColor>
		</PublishText>
	</InfoText>
		
</ImageViewer>
    
    '''
    outfile = open('../ImageViewer_Config_template.xml','w')
    global cleanPath  
    cleanPath = refinePath(os.getcwd())
    print ('Creating XML Template...')
    print ('<?xml version="1.0" encoding="UTF-8"?>\n<ImageViewer>\n <Content>\n' + fetchDir(sel_dir) + '\n</Content> \n' + end, file = outfile)
    print ('\nDone!')    
    return 
    
if __name__ == '__main__':
    stage()

    
        
