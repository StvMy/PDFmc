import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox as mb
import pypdf
import os


class PDFMergerApp(ctk.CTk):
    paths = []
    paths_check = []
    folder =""
    custom_font =("Arial",15,'bold')
    def __init__(self):
        super().__init__()
        
        # get screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # calculate position x and y coordinates
        x = (screen_width/2) - (800/2)
        y = (screen_height/2) - (600/2)
        self.title("PDFM")
        self.geometry('%dx%d+%d+%d' % (846, 446, x, y))
        # self.resizable(False, False)
        ctk.set_appearance_mode("dark")

        # Layout Frames
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(side="left", fill="y", padx=10, pady=10)
        
        # File List
        text_var = tk.StringVar(value="SELECTED:")
        self.label = ctk.CTkLabel(master=self,
                                    textvariable=text_var,
                                    width=200,
                                    height=25,
                                    fg_color="silver",
                                    text_color="#000000",
                                    corner_radius=4)
        self.label.place(relx=0.05, rely=0.03)  
        
        self.selected_path = ctk.CTkScrollableFrame(self.file_frame, width=600)
        self.selected_path.pack(fill="y", expand=True, padx=10, pady=36)

        # Buttons
        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(fill="x",side="bottom", pady=10)

        self.merge_button = ctk.CTkButton(self.button_frame, text="Browse files", command=self.browse_files)
        self.merge_button.pack(padx=20, pady=5, fill="x")  

        self.merge_button = ctk.CTkButton(self.button_frame, text="MERGE!", command=self.merge_pdfs, fg_color="#008000", width=150,height=100)
        self.merge_button.configure(font= self.custom_font)
        self.merge_button.pack(padx=20, pady=5, fill="x")

        self.merge_button = ctk.CTkButton(self.button_frame, text="CONVERT!", command=self.convert_files, fg_color="#FF0000", width=150,height=100)
        self.merge_button.configure(font= self.custom_font)
        self.merge_button.pack(padx=20, pady=5, fill="x")

        self.merge_button = ctk.CTkButton(self.button_frame, text="test", command=self.test)
        self.merge_button.pack(padx=20, pady=5, fill="x")  

        # Select Dir Frames
        self.dir_frame = ctk.CTkFrame(self)
        self.dir_frame.pack(side = "bottom",ipady=50,fill="x")
            
        # "Save file"
        textsave = tk.StringVar(value="Save Path:")
        self.save = ctk.CTkLabel(self.dir_frame,textvariable=textsave)
        self.save.place(relx=0.05,rely=0.2)

        # Dir view
        self.dir_view = ctk.CTkTextbox(self.dir_frame,text_color="White",height=20,width=143)
        self.dir_view.place(relx=0.11,rely=0.48)

        # Browse dir button
        self.dir_browse = ctk.CTkButton(self.dir_frame, text="Browse", command=self.browse_folder, fg_color="#FFFFFF",text_color="black")
        self.dir_browse.pack(side="bottom",pady=15)  

        #TEST
    def test(self):
            
        print(self.paths_check)

        # File Browse                                             _________________________ NEED FIX HERE!___________________________
    def browse_files(self):                                          
        self.clearFrame()
        file = filedialog.askopenfilename() 
        self.add_files_to_list(file)

    def add_files_to_list(self,file):  
        if file.endswith(".pdf") or file.endswith(".docx"):
            self.paths.append(file)
        
        elif (file!=""):
            mb.showwarning(title="ERROR", message="Can't use this file")
        # self.selected_path.configure(state="normal")
        # self.selected_path.delete(0.0, 'end')                
        for i in range(len(self.paths)):

            # Create a frame inside the scrollable frame for each entry
            file_frame = ctk.CTkFrame(self.selected_path)
            file_frame.pack(fill="x", padx=5, pady=2)

            # Checkbox to select/deselect the file
            checkbox = ctk.CTkCheckBox(file_frame, onvalue="on", offvalue="off",text="")
            checkbox.pack(side="left", padx = (5,2))
            # self.paths_check.append(checkbox.get())

            # Label to display the file path
            file_label = ctk.CTkLabel(file_frame, text=self.paths[i], anchor="w")
            file_label.pack(side="left", padx=2, fill="x", expand=True)

        if (len(self.paths_check)!=len(self.paths)):
            delta = len(self.paths_check)-len(self.paths)
            
                



    def browse_folder(self):
        self.folder = filedialog.askdirectory()
        self.dir_view.insert("1.0", self.folder) 
        self.dir_view.configure(state="disabled")
        print(self.folder)

    def merge_pdfs(self):
        if self.paths==[]:
            mb.showwarning(title="ERROR", message="Please Select File!")
        elif self.folder=="":
            mb.showwarning(title="ERROR", message="Please Select Save Path!")

    def convert_files(self):
        if self.paths==[]:
            mb.showwarning(title="ERROR", message="Please Select File!")
        elif self.folder=="":
            mb.showwarning(title="ERROR", message="Please Select Save Path!")
  
    def clearFrame(self):
        # destroy all widgets from frame
        for widget in self.selected_path.winfo_children():
            widget.destroy()
        # this will clear frame and frame will be empty

if __name__ == "__main__":
    app = PDFMergerApp()
    app.mainloop()