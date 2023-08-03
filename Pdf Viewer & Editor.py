import os
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkPDFViewer import tkPDFViewer as pdf
import aspose.words as aw
import tkinter.messagebox as messagebox 
import PyPDF2
class PDFReader(tk.Frame):
    # Configuring GUI Frame
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry('650x750+400+10')
        self.master.title('PDF Reader & Editor')
        self.master.resizable(False, False)
        self.fullscreen = False
        self.recent_files = []
        self.pack()
        self.second_pdf_file = None 
        self.create_widgets()

    def create_widgets(self):
        self.select_button = tk.Button(self, text="Select PDF", width=500, font='NewTimesRoman 16',
                                       command=self.open_pdf)
        self.select_button.pack()
        self.recent_button = tk.Button(self,text="Recently opened",width=500,font='NewTimesRoman 16' ,
                                       command=self.show_recent_files)
        self.recent_button.pack()
        self.read_button = tk.Button(self, text="Read PDF", width=500, font='NewTimesRoman 16',
                                      command=self.show_pdf, state=DISABLED)
        self.read_button.pack()
        self.edit_button = tk.Button(self, text="Edit PDF", width=500, font='NewTimesRoman 16',
                                      command=self.edit_pdf, state=DISABLED)
        self.edit_button.pack()    
       
    def shortcuts(self):
        self.master.bind("<F11>", self.toggle_fullscreen)
        self.master.bind("<Control-s>", self.split_pdf)
        self.master.bind("<Control-x>", self.extract_one_page)
        self.master.bind("<Control-c>", self.pdf_to_word)
        self.master.bind("<Control-o>", self.open_pdf)
        self.master.bind("<Control-q>", self.quit)
        self.master.bind("<Control-r>", self.show_pdf)
        self.master.bind("<Control-e>", self.extract_Range)
        self.master.bind("<Control-a>", self.extract_all)
        self.master.bind("<Control-t>", self.rotate_pdf)
    # Adding a Function to select and open PDF files

    def open_pdf(self):
        file_path = filedialog.askopenfilename(title="Select file", filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
        if file_path:
            self.pdf_file = open(file_path, 'rb')
            self.file_path = file_path
            self.recent_files.append(file_path) 
            self.read_button.config(state=NORMAL)
            self.edit_button.config(state=NORMAL)
        else:
            pass
    
    # Show a list of recent opened files

    def show_recent_files(self):
        recent_files_window = tk.Toplevel(self.master)
        recent_files_window.title("Recent Files")
        recent_files_window.geometry("300x300")

        if len(self.recent_files) == 0:
            no_recent_files_label = tk.Label(recent_files_window, text="No recent files")
            no_recent_files_label.pack()
        else:
            for file_path in self.recent_files:
                button = tk.Button(recent_files_window, text=file_path, command=lambda path=file_path: self.open_recent_file(path))
                button.pack()

    # Select and open on of recent files

    def open_recent_file(self, file_path):
        self.pdf_file = open(file_path, 'rb')
        self.file_path = file_path
        self.read_button.config(state=NORMAL)
        self.edit_button.config(state=NORMAL)
        self.show_pdf(file_path)
    
    # Display PDF file in GUI Frame

    def show_pdf(self):
        v1 = pdf.ShowPdf()
        v1.img_object_li.clear()
        v2 = v1.pdf_view(self, pdf_location=self.file_path)
        v2.config(background='white')
        v2.config(width=500,height=750)
        v2.pack()

    # Show a list of options to edit selected PDF file

    def edit_pdf(self):
        edit_window = Toplevel(self.master)
        edit_window.geometry('300x300')
        edit_window.title('PDF Editor')
        edit_window.resizable(False, False)
        options_label = Label(edit_window,text="Select an editing option:")
        options_label.pack()
        option_listbox = Listbox(edit_window,height=5)
        option_listbox.insert(1, "Convert to")
        option_listbox.insert(2, "Split PDF")
        option_listbox.insert(3, "Merge PDF")
        option_listbox.insert(4, "Extract one page")
        option_listbox.insert(5, "Extract all pages")
        option_listbox.pack()
        def select_option():
            option = option_listbox.curselection()
            if option == (0,):
                self.Conventor()
            elif option == (1,):
                self.split_pdf()
            elif option == (2,):
                self.merge()
            elif option == (3,):
                self.extract_one_page()
            elif option == (4,):
                self.extract_all()
            else:
                pass
        select_button = Button(edit_window, text="Select", command=select_option)
        select_button.pack()
    def Conventor(self):
        Convertor_window = Toplevel(self.master)
        Convertor_window.geometry('300x300')
        Convertor_window.title('Conventor To')
        Convertor_window.resizable(False, False)
        options_label = Label(Convertor_window,text="Select an Converting option:")
        options_label.pack()
        option_listbox = Listbox( Convertor_window,height=5)
        option_listbox.insert(1, " Word")
        option_listbox.insert(2, "Png")
        option_listbox.insert(3, "Jpeg")
        option_listbox.insert(4, "ppt")
        option_listbox.pack()
        def select_option():
            option = option_listbox.curselection()
            if option ==(0,):
                self.pdf_to_word()
            elif option == (1,):
                self.pdf_to_png()
            elif option == (2,):
                self.pdf_to_jpeg()
            elif option == (3,):
                self.pdf_to_ppt()
            else:
                pass
        select_button = Button(Convertor_window, text="Select", command=select_option)
        select_button.pack()
    def pdf_to_word(self):
        doc = aw.Document(self.file_path)
        docx_path = os.path.join(os.path.dirname(self.file_path), f"Word version of {os.path.basename(self.file_path)}.docx")
        doc.save(docx_path, aw.SaveFormat.DOCX)
        os.startfile(docx_path)
    def pdf_to_png(self):
        doc = aw.Document(self.file_path)
        docx_path = os.path.join(os.path.dirname(self.file_path), f"Png version of {os.path.basename(self.file_path)}.Png")
        doc.save(docx_path, aw.SaveFormat.PNG)
        os.startfile(docx_path)
    def pdf_to_jpeg(self):
        doc = aw.Document(self.file_path)
        docx_path = os.path.join(os.path.dirname(self.file_path), f"jpeg version of {os.path.basename(self.file_path)}.Jpeg")
        doc.save(docx_path, aw.SaveFormat.JPEG)
    def pdf_to_ppt(self):
        presentation = os.path.join(os.path.dirname(self.file_path), f"presentation version of {os.path.basename(self.file_path)}.pptx")
        pdf2pptx.convert_pdf2pptx(self.file_path, presentation , 300, 0, None, False)
        os.startfile(presentation)
    def split_pdf(self):
        split_window = Toplevel(self.master)
        split_window.geometry('300x300')
        split_window.title('Split')
        split_window.resizable(False, False)
        start_label = Label(split_window, text="Start page")
        start_label.pack()
        start_entry = Entry(split_window)
        start_entry.pack()
        end_label = Label(split_window, text="End page")
        end_label.pack()
        end_entry = Entry(split_window)
        end_entry.pack()
        split_button = Button(split_window, text="Split", command=lambda: self.split(start_entry.get(), end_entry.get()))
        split_button.pack()
    def split(self,start_entry,end_entry):
        try:
            start = int(start_entry)
            end = int(end_entry)
            if start > end:
                messagebox.showerror("Error", "Start page must be less than end page")
            else:
                doc = aw.Document(self.file_path)
                for i in range(0,doc.page_count):
                    extracted = doc.extract_pages(start,end)
                    extracted.save("PDF"+f"[{start}-{end}]", aw.SaveFormat.PDF)
                    os.startfile("PDF"+f"[{start}-{end}]")
        except ValueError:
            messagebox.showerror("Error", "Please enter a number")
            pass
    def extract_one_page(self):
        extract_window = Toplevel(self.master)
        extract_window.geometry('300x300')
        extract_window.title('Extract')
        extract_window.resizable(False, False)
        page_label = Label(extract_window, text="Page")
        page_label.pack()
        page_entry = Entry(extract_window)
        page_entry.pack()
        extract_button = Button(extract_window, text="Extract", command=lambda: self.extract(page_entry.get()))
        extract_button.pack()
    def extract(self,page_entry):
        page = int(page_entry)
        doc = aw.Document(self.file_path)
        extracted = doc.extract_pages(page,page)
        extracted.save("Page"+f"[{page}].pdf", aw.SaveFormat.PDF)
        os.startfile("Page"+f"[{page}].pdf")
    def extract_all(self):
        try:
            doc = aw.Document(self.file_path)
            for page in range(1, doc.page_count + 1):
                extracted = doc.extract_pages(page, page + 1)
                extracted.save(f"Page[{page}].pdf", aw.SaveFormat.PDF)
            messagebox.showinfo("Extraction Complete", "All pages have been extracted successfully.")
        except Exception as e:
           messagebox.showerror("Extraction Error", f"An error occurred during extraction: {str(e)}")
    def merge_pdf(self):
        merge_window = Toplevel(self.master)
        merge_window.geometry('300x300')
        merge_window.title('Merge')
        merge_window.resizable(False, False)
        file_label = Label(merge_window, text="File")
        file_label.pack()
        file_button = Button(merge_window, text='Select', command=self.select_and_merge)
        file_button.pack()
        merge_button = Button(merge_window, text="Merge", command=self.merge)
        merge_button.pack()
    def merge(self):
        file_path2 = filedialog.askopenfilename(title="Select file", filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
        files_list = [self.file_path]
        files_list.append(file_path2)
        output = aw.Document()
        output.remove_all_children()
        for fileName in files_list:
            input = aw.Document(fileName)
            output.append_document(input, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        output.save("merged.pdf")
        os.startfile("merged.pdf")
root = tk.Tk()
app = PDFReader(master=root)
app.mainloop()
