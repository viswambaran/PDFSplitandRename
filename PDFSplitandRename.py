#!/usr/bin/env python
# coding: utf-8

# In[1]:


## Import packages 

from PyPDF2 import PdfReader, PdfFileWriter
import os 
import re 
import shutil
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo


# In[60]:


### Functions 
def RunApp():

   
   
   root = tk.Tk()
   root.title('PDF Split and Rename')
   root.resizable(False, False)
   root.geometry('400x200')
   
   def _close_app():
       root.destroy()
   
   # open button
   # command section uses lambda to call the function needed and then to close the 
   # window once the task has completed
   open_button = ttk.Button(
       root,
       text = 'Open Files',
       command = lambda:[RenamePDF(), root.destroy()])

   # close button
   close_button = ttk.Button(
       root, 
       text = 'Close', 
       command=root.destroy)

   open_button.pack(expand = False)
   close_button.pack(expand = False)

   # Bring window to front of users screen
   root.lift()
   root.attributes('-topmost', True)
   root.after_idle(root.attributes, '-topmost', False)

   root.mainloop()

def select_files():
   
   file_chosen = tk.messagebox.askokcancel(title = "Choose files to rename", message = "Please select the files you wish you rename!")
   
 
   
   filenames = fd.askopenfilenames(
       title='Open files',
       initialdir='/',
       filetypes=(("PDF Files", "*.pdf"),))
   
   if not filenames: 
       tk.messagebox.showwarning(title = "No files were selected!", message = "No files were selected. Aborting application!")
   
   return(filenames)

def select_folder():
   """
   Function used for user to select the final folder to store renamed PDF files.
   """ 
   tk.messagebox.askokcancel(title = "Choose folder to save renamed files", message = "Please select the folder"                                                                                         " you wish to save the renamed files to")
   select_folder = fd.askdirectory()
   return(select_folder)

def RenamePDF():
   
   # User to select files to rename and folder to output the renamed files
   filenames = select_files()
   New_Folder = select_folder()
   
   for file in filenames:
       #reader = PdfReader("data/Brighton.pdf")
       reader = PdfReader(file)
       
       for i in range(0,reader.getNumPages()):
           page = reader.getPage(i)
           text = page.extract_text() + "\n"
           
           ## Getting Invoice Number
           InvNo = re.search('(?<=Invoice\sNo\s:\s)(\d+)', text).group(0).strip()
           new_file = New_Folder + "/" + InvNo + ".pdf"
           output = PdfFileWriter()
           
           with open(new_file, "wb") as output_file:
               output.addPage(reader.getPage(i))
               output.write(output_file)
   
   tk.messagebox.showinfo(title = "PDF split and rename complete!", message = str(reader.getNumPages()) + " files were created to " + New_Folder)


# In[61]:


RunApp()

