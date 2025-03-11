import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import img2pdf as i2f
import comtypes.client 
from tkinter import messagebox as mb
import PyPDF2
import os
import sys



class PDFMergerApp(ctk.CTk):
    paths = []
    paths_check = []
    frame = []
    folder =""
    custom_font =("Arial",15,'bold')
    
    def __init__(self):
        icon_path = self.resource_path("PDFmc.ico")
        super().__init__()
        
        # get screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # calculate position x and y coordinates
        x = (screen_width/2) - (800/2)
        y = (screen_height/2) - (600/2)
        self.geometry('%dx%d+%d+%d' % (846, 400, x, y))
        self.resizable(False,False)

        self.wm_iconbitmap(icon_path)
        self.title("PDFmc")

        # self.resizable(False, False)
        ctk.set_appearance_mode("dark")

        # Layout Frames
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.rowconfigure(0,weight=1,uniform='a') 
        self.file_frame.pack(side="left", fill="y", padx=5, pady=10)
        
        # File List
        text_var = tk.StringVar(value="SELECTED:")
        self.label = ctk.CTkLabel(master=self,
                                    textvariable=text_var,
                                    width=100,
                                    height=17,
                                    fg_color="silver",
                                    text_color="#000000",
                                    corner_radius=4)
        self.label.place(relx=0.05, rely=0.03)  
        
        self.selected_path = ctk.CTkScrollableFrame(self.file_frame, width=600)
        self.selected_path.pack(fill="y", expand=True, padx=10, pady=27)

        self.merge_button = ctk.CTkButton(self, text="DESELECT ALL", command=self.deseelect_all, border_color="red",border_width=1,text_color="black",fg_color="white",hover_color="#bebebe")
        self.merge_button.pack(fill="x", ipady=0,pady=10)
 
        # Buttons
        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(fill="both",pady=10)

        self.merge_button = ctk.CTkButton(self.button_frame, text="Browse files", command=self.browse_files, fg_color="#5c83b4",hover_color="#504f4f")
        self.merge_button.pack(padx=20, pady=5, fill="x")  

        self.merge_button = ctk.CTkButton(self.button_frame, text="MERGE!", command=self.merge_pdfs, fg_color="#8e8e8e", width=150,height=100,hover_color="#504f4f")
        self.merge_button.configure(font= self.custom_font)
        self.merge_button.pack(padx=20, pady=5, fill="x")

        self.merge_button = ctk.CTkButton(self.button_frame, text="CONVERT!", command=self.convert_files, fg_color="#8e8e8e", width=150,height=100,hover_color="#504f4f")
        self.merge_button.configure(font= self.custom_font)
        self.merge_button.pack(padx=20, pady=5, fill="x")

        self.app_name = ctk.CTkLabel(self,text="PDFmc")
        self.app_name.configure(font = ("VT323",45,'bold'))
        self.app_name.pack()

        #deselect all
    def deseelect_all(self):
        for i in (self.paths_check):
            if i.get()=="on":
                i.deselect()

        # File Browse                                             
    def browse_files(self):                
        file = filedialog.askopenfilenames(title="Choose files") 
        self.add_files_to_list(file)

        
    def add_files_to_list(self,file):  
        for i in file:
            if i:
                #frame for list
                file_frame = ctk.CTkFrame(self.selected_path)
                self.frame.append(file_frame)
                
                #checkbox for list
                checkbox = ctk.CTkCheckBox(file_frame, onvalue="on", offvalue="off",text="")   
                checkbox.select() 
                self.paths_check.append(checkbox)

                #label to show path in list
                file_label = ctk.CTkLabel(file_frame, text=i, anchor="w")
                self.paths.append(file_label)

        for i in range(len(self.paths)):

            # Create a frame inside the scrollable frame for each entry       
            self.frame[i].pack(fill="x", padx=5, pady=2)

            # Checkbox to select/deselect the file
            self.paths_check[i].pack(side="left", padx = (5,2))

            # Label to display the file path
            self.paths[i].pack(side="left", padx=2, fill="x", expand=True)

    def browse_folder(self):
        self.folder = filedialog.askdirectory() 
        self.dir_view.insert("1.0", self.folder) 
        self.dir_view.configure(state="disabled")

    def merge_pdfs(self):
        err = 0

        #This part Check if there's file to execute
        count=0
        for i in self.paths_check:
            if i.get()=="on":
                count+=1
        if (len(self.paths)<2) or (count==0):
            mb.showwarning(title="ERROR", message="Please Select File!")
            return
        
        merge = PyPDF2.PdfMerger()
        for i in range (len(self.paths)):
            if (self.paths_check[i].get() == "on") and (self.paths[i].cget("text")[-4:]==".pdf"):
                merge.append(self.paths[i].cget("text"))   
            elif(self.paths[i].cget("text")[-4:]!=".pdf"):
                err+=1
        if err>0:mb.showwarning(title="WARNING!", message=(err,"Of your file can't be merged"))
        savepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=(("pdf file", "*.pdf"),("All Files", "*.*") ))
        if not savepath:return
        merge.write(savepath)
        merge.close()
            
    def convert_files(self):
        converting=[]
        err = 0
        #This part Check if there's file to execute
        count=0
        for i in self.paths_check:
            if i.get()=="on":
                count+=1
        if (self.paths==[]) or (count==0):
            mb.showwarning(title="ERROR", message="Please Select File!")
            return
        
        word = comtypes.client.CreateObject("word.application")
        for i in range (len(self.paths)):
            if (self.paths_check[i].get() == "on"):
                if self.paths[i].cget("text").rsplit(".",1)[-1].strip() in ("jpg","jpeg","png","tiff","tif","bmp","ppm","pgm","doc","docx","rtf","odt","txt","xls","xlsx","xlsm","csv","ppt","pptx","pps","ppsx","vsd","vsdx"):
                    converting.append(self.paths[i].cget("text"))
                else:
                    err+=1

        savepath = filedialog.askdirectory()

        for i in converting:
            filename = i.rsplit("/", 1)[-1].rsplit(".", 1)[0]
            if i.rsplit(".",1)[-1].strip() in ("jpg","jpeg","png","tiff","tif","bmp","ppm","pgm"):
                with open(savepath+"/"+filename+".pdf","wb") as converted:
                    converted.write(i2f.convert(i))
            elif i.rsplit(".",1)[-1].strip() in ("doc","docx","rtf","odt","txt","xls","xlsx","xlsm","csv","ppt","pptx","pps","ppsx","vsd","vsdx"):
                path = i.replace("/","\\")
                print(path)
                doc = word.Documents.Open(path)
                # filename need fix every space replaced by %20
                savefilepathname = os.path.normpath(savepath+"/"+filename+".pdf")
                doc.SaveAs(savefilepathname,FileFormat=17)
            print("CONVERT!")
        doc.Close()
        word.quit()
        if err>0:mb.showwarning(title="WARNING!", message=(err,"Of your file can't be merged"))
    
    @staticmethod  
    def resource_path(relative_path):
        # """ Get the absolute path to a resource, works for dev and PyInstaller """
        if hasattr(sys, '_MEIPASS'):  # If running as an EXE
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)  # Running as script
    
if __name__ == "__main__":
    app = PDFMergerApp()
    app.mainloop()