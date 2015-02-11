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
        g_img        = c_img.get().lower()
        g_title      = c_title.get().lower()
        g_author     = c_author.get().lower() 
        g_desc       = c_desc.get().lower()
        g_year       = c_year.get().lower()
        g_location   = c_location.get().lower()
        g_publish    = c_publish.get().lower()
        g_collection = c_collection.get().lower()
        
        global ex_opt
        ex_opt = [g_img, g_title, g_desc, g_author, g_publish, g_collection, g_year, g_location]
        
        if ex_opt[0]=="":
            messagebox.showwarning(title = "Error", message = "You must select a column for at least the file names")  
        
        try: 
            sel_dir
            sel_excel

        except NameError:
            messagebox.showwarning(title = "Error", message = "You must choose a directory , excel sheet, and at least the File Names!")  
        else: 
            for i in range (0,8):
                if ex_opt[i]=="":
                    ex_opt[i] = -1
                else: 
                    ex_opt[i] = ord(ex_opt[i])-97 
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


    #Stores the data from colums of the selected row      
    def find_info(row):
        img=[]
        
        for i in range (1,8): 
            if ex_opt[i] < 0:
                img.append("") 
            else:
                try: 
                    img.append(sh.cell(rowx=row, colx=ex_opt[i]).value)
                except UnicodeEncodeError:
                    img_app = img.append(sh.cell(rowx=row, colx=ex_opt[i]).value)
                    img_app.encode('ascii','xmlcharrefreplace')
        
        info_list = [img[0], img[1], img[2], img[3], img[4], img[5], img[6]]
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
            result += '<Source id = \'%d\'>\n' % sourceId
            
            #Appends Picture name to Directory Path
            cell_value = sh.cell(rowx=row,colx=ex_opt[0]).value
            localPath = path+'/'+cell_value
            
            #Insurance Parantheses are correct
            parenSwitch = localPath.replace('\\','/')
            info_list = find_info(row)
            
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
		<type>Video</type>
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
        
        
    #Default settings for the collection Viewer     
    end = '''\
    
    <!--
		
		Note: You can override the global InfoText style for a source by inserting local style tags as shown below.
		
		<InfoText>
			<TitleText>
				<fontColor>0xFF0000</fontColor>
				<fontSize>50</fontSize>
			</TitleText>
			<DescriptionText>
				<fontColor>0xFF0000</fontColor>
				<fontSize>50</fontSize>
			</DescriptionText>
			<AuthorText>
				<fontColor>0xFF0000</fontColor>
				<fontSize>50</fontSize>
			</AuthorText>
			<PublishText>
				<fontColor>0xFF0000</fontColor>
				<fontSize>50</fontSize>
			</PublishText>
		</InfoText>
		
	-->

    
    
    </Content>		
		
	<GlobalSettings><!-- Sets the initial, min and max scale values for the video objects. !!Important: Either <scale> or <imagesNormalize> should be used. Not both !! --
		<!-- <scale></scale> --><!-- Sets the video's full size (the size at which the control buttons appear) as relative to the video's actual size. So, for instance, if an image is 500px x 400px, setting <scale> to 1 would mean its full size equals its actual size. Setting scale to 2 would make it twice as large. -->
		
		<amountToShow>0</amountToShow>
		
		<globalScale>.25</globalScale><!-- The size videos start at before they are double-tapped or zoomed to their full size, which is the size at which the control buttons appear. This variable modulates either <scale> or <imageNormalize>, depending on whether the user wants to set the full size relative to the actual size or as an absolute value-->
		<imagesNormalize></imagesNormalize><!-- Sets the full size as an absolute pixel width; the height of the videos are adjusted accordingly.  If left blank, will be calculated automatically. -->
		<maxScale>2.5</maxScale><!-- The maximum size the videos can zoom to, set relative to their full size within the application (set by <scale> or <imagesNormalize>). For instance, setting maxScale to 2.5 will allow the user to zoom the videos to 2.5 times their full size. -->
		<minScale>.1</minScale><!-- The minimum size the videos can zoom to, set relative to their full size within the application (set by <scale> or <imagesNormalize>). For instance, setting minScale to .5 will allow the user to shrink the videos to half their full size. -->
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
		<buttonPadding>6</buttonPadding><!-- Sets the padding in pixels on all four sides of each button -->
		<buttonOutlineStroke>2</buttonOutlineStroke><!-- Sets the width of the button outline in pixels -->
		<buttonOutlineColor>0x000000</buttonOutlineColor><!-- Sets the color of the button outline -->
		<buttonColorPassive>0x333333</buttonColorPassive><!-- Sets the background color of the buttons when they aren't active (haven't been selected). -->
		<buttonColorActive>0x000000</buttonColorActive><!-- Sets the background color of the buttons when they are active (have been selected). -->
		<buttonSymbolColor>0xFFFFFF</buttonSymbolColor><!-- Sets the color of the symbol or text on each button -->
		<timeFontSize>15</timeFontSize><!--Sets the text size for the timecode that displays on video objects. -->
	</ControlBtns>
	
	<BackgroundGraphic><!-- Sets the application background to be either a solid color (if only <fillColor1> is given a value) or a two-color horizontal gradient (if both color values are filled in).   -->
		<fillColor1>0xCCCCCC</fillColor1>
		<fillColor2>0xCCCCCC</fillColor2>
	</BackgroundGraphic>
	
	<BackgroundOutline><!-- Sets the style of the outline that surrounds each video object -->
		<outlineColor>0xFFFFFF</outlineColor><!-- Sets the color of the outline surrounding each media object. -->
		<outlineStroke>1</outlineStroke><!-- Sets the width of the border around each media object in pixels -->
	</BackgroundOutline>
	
	<InfoText><!-- Sets the font size and color for each metadata field. -->
		<TitleText>
			<fontColor>0x000000</fontColor>
			<fontSize></fontSize>
		</TitleText>
		<DescriptionText>
			<fontColor>0x333333</fontColor>
			<fontSize></fontSize>
		</DescriptionText>
		<AuthorText>
			<fontColor>0x666666</fontColor>
			<fontSize></fontSize>
		</AuthorText>
		<PublishText>
			<fontColor>0x009966</fontColor>
			<fontSize></fontSize>
		</PublishText>
	</InfoText>
	
</VideoViewer>
    
    '''
    outfile = open('../VideoViewer_Config_template.xml','w')
    global cleanPath  
    cleanPath = refinePath(os.getcwd())
    print ('Creating XML Template...')
    print ('<?xml version="1.0" encoding="UTF-8"?>\n<ImageViewer>\n <Content>\n' + gencon(sel_dir) + end, file = outfile)
    print ('\nDone!')    
    return 
    
if __name__ == '__main__':
    stage()

    
        
