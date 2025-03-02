import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from ttkthemes import ThemedStyle
from PIL import Image
from PyPDF2 import PdfMerger

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image & PDF Merger")
        self.root.geometry("500x450")

        #add favicon
        self.root.iconbitmap(os.path.join(os.path.dirname(__file__), 'favicon.ico'))

        #set dark theme
        self.style = ThemedStyle(root)
        self.style.set_theme("equilux")
        self.root.configure(bg="#2b2b2b")
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.drop_files)

        #file list
        self.file_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=60, height=15, bg="#3b3b3b", fg="white",
                                       selectbackground="#555", highlightbackground="#444", relief="flat")
        self.file_listbox.pack(pady=10)

        #buttons frame
        btn_frame = tk.Frame(root, bg="#2b2b2b")
        btn_frame.pack()

        #buttons
        self.add_btn = tk.Button(btn_frame, text="Add Files", command=self.add_files, bg="#444", fg="white", relief="flat")
        self.add_btn.pack(side=tk.LEFT, padx=5)

        self.remove_btn = tk.Button(btn_frame, text="Remove Selected", command=self.remove_selected, bg="#444", fg="white", relief="flat")
        self.remove_btn.pack(side=tk.LEFT, padx=5)

        self.move_up_btn = tk.Button(btn_frame, text="Move Up", command=lambda: self.move_item(-1), bg="#444", fg="white", relief="flat")
        self.move_up_btn.pack(side=tk.LEFT, padx=5)

        self.move_down_btn = tk.Button(btn_frame, text="Move Down", command=lambda: self.move_item(1), bg="#444", fg="white", relief="flat")
        self.move_down_btn.pack(side=tk.LEFT, padx=5)

        self.merge_btn = tk.Button(root, text="Merge to PDF", command=self.merge_files, fg="white", bg="green", relief="flat")
        self.merge_btn.pack(pady=10)

        #progress bar
        self.progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=10)

    def add_files(self):
        file_paths = filedialog.askopenfilenames(title="Select images and/or PDFs",
                                                 filetypes=[("Image & PDF files", "*.jpg;*.jpeg;*.png;*.pdf")])
        for file in file_paths:
            self.file_listbox.insert(tk.END, file)

    #file dropper - allows only selected file types
    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if file.lower().endswith((".jpg", ".jpeg", ".png", ".pdf")):
                self.file_listbox.insert(tk.END, file)

    #file remover
    def remove_selected(self):
        selected = self.file_listbox.curselection()
        for index in reversed(selected):
            self.file_listbox.delete(index)

    #file mover - todo: not fully functional
    def move_item(self, direction):
        selected = self.file_listbox.curselection()
        if not selected:
            return

        for index in selected:
            new_index = index + direction
            if 0 <= new_index < self.file_listbox.size():
                file = self.file_listbox.get(index)
                self.file_listbox.delete(index)
                self.file_listbox.insert(new_index, file)
                self.file_listbox.selection_set(new_index)

    #converter function
    def convert_images_to_pdfs(self, image_paths):
        """Convert images to temporary PDFs with a progress bar."""
        pdf_paths = []
        total_files = len(image_paths)
        self.progress["value"] = 0
        self.progress["maximum"] = total_files

        #loop images, open each one and make into temp pdf
        for index, image_path in enumerate(image_paths, start=1):
            img = Image.open(image_path)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")

            #save as temp pdf - todo: make sure access is there for the user to save the temp files
            pdf_path = image_path.rsplit(".", 1)[0] + "_temp.pdf"
            img.save(pdf_path, "PDF", resolution=100.0)
            pdf_paths.append(pdf_path)

            #progress bar increment
            self.progress["value"] = index
            self.root.update_idletasks()

        return pdf_paths

    #merger function
    def merge_files(self):
        file_paths = self.file_listbox.get(0, tk.END)

        if not file_paths:
            messagebox.showerror("Error", "No files selected!")
            return

        #separate images and pdf in order to convert images into pdfs and then append them to the pdf_paths array
        #todo: support word files
        image_paths = [f for f in file_paths if f.lower().endswith((".jpg", ".jpeg", ".png"))]
        pdf_paths = [f for f in file_paths if f.lower().endswith(".pdf")]

        #convert images to PDFs
        pdf_paths.extend(self.convert_images_to_pdfs(image_paths))

        #nothing was converted
        if not pdf_paths:
            messagebox.showerror("Error", "No valid files to merge.")
            return

        #pop up window for filename
        output_filename = simpledialog.askstring("Save PDF", "Enter the output PDF filename:", initialvalue="") #initialvalue is placeholder
        if not output_filename:
            return #user cancels

        #add .pdf extension if user did not include it
        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"

        output_pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                       title="Save Merged PDF", initialfile=output_filename)
        if not output_pdf_path:
            return #user cancels

        #initialize merger progress bar
        self.progress["value"] = 0
        self.progress["maximum"] = len(pdf_paths)

        #merge all temp pdf files through generic library
        merger = PdfMerger()
        for index, pdf in enumerate(pdf_paths, start=1):
            merger.append(pdf)

            #update progress bar
            self.progress["value"] = index
            self.root.update_idletasks()

        #save final file
        merger.write(output_pdf_path)
        merger.close()

        #cleanup temp pdf files - todo: make sure user has appropriate rights
        for temp_pdf in pdf_paths:
            if temp_pdf.endswith("_temp.pdf"):
                os.remove(temp_pdf)

        messagebox.showinfo("Success", f"Merged PDF saved: {output_pdf_path}")


if __name__ == "__main__":
    root = TkinterDnD.Tk() #this will allow drag and drop into UI
    app = PDFMergerApp(root)
    root.mainloop()
