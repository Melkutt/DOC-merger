import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import win32com.client
import os
import time
#from  pathlib import Path

class DocMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOC-filsammanfogare")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        self.setup_ui()

    def setup_ui (self):
        
        # Main frame
        main_frame  = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        
        # Config grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        
        # Input folder selection
        ttk.Label(main_frame, text="Mapp med DOC-filer:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.folder_var = tk.StringVar()
        folder_entry = ttk.Entry(main_frame, textvariable=self.folder_var, width=50)
        folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        browse_btn = ttk.Button(main_frame, text="Bläddra...", command=self.browse_folder)
        browse_btn.grid(row=0, column=2, pady=5)
        
        
        # Output file selection
        ttk.Label(main_frame, text="Resultatfil:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar()
        output_entry = ttk.Entry(main_frame, textvariable=self.output_var, width=50)
        output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        save_btn = ttk.Button(main_frame, text="Spara som...", command=self.save_as)
        save_btn.grid(row=1, column=2, pady=5)
        
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        
        # Status label
        self.status_var = tk.StringVar(value="Redo att starta...")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        
        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Logg", padding="5")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W,  tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        self.merge_btn  = ttk.Button(button_frame, text="Starta sammanfogning", command=self.start_merge)
        self.merge_btn.grid(row=0, column=0, padx=5)
        
        ttk.Button(button_frame, text="Rensa logg", command=self.clear_log).grid(row=0, column=1, padx=5)
        
        
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Välj mapp med DOC-filer")
        if folder:
            self.folder_var.set(folder)
           
            
    def save_as(self):
        filename =  filedialog.asksaveasfilename(title="Spara sammanfogad fil", defultextension="*.doc", filetypes=[("Word Documents", "*.doc")])
        if filename:
            self.output_var.set(filename)
         
            
    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        
        
    def update_status(self, message, progress=None):
        self.status_var.set(message)
        if progress is not None:
            self.progress['value'] = progress
        self.root.update_idletasks()
        
        
    def start_merge(self):
        if not self.folder_var.get():
            messagebox.showerror("Fel", "Välj en mapp med DOC-filer!")
            return
        
        if not self.output_var.get():
            messagebox.showerror("Fel", "Välj var resultatfilen skall sparas!")
            
            return
        
        
        #start merging in separate thread to keep UI responsive
        threading.Thread(target=self.merge_documents, daemon=True).start()
        
    def merge_documents(self):
        try:
            input_folder = self.folder_var.get()
            output_file = self.output_var.get()
            
            
        # Get all .doc files
            files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) 
                 if f.lower().endswith(".doc") and not f.startswith("~$")]
        
            if not files:
                self.log_message("Inga  DOC-filer hittades i vad mapp!")
                self.update_status("Inga filer hittades!")
                return
        
            files.sort(key=lambda f: os.path.basename(f)) # Sort numeric
            total_files = len(files)
        
            self.log_message(f"Hittade {total_files} DOC-filer att sammanfoga")
            self.update_status(f"Hittade {total_files} filer...", 0)
        
        
        # Initialize Word application
            self.log_message("Startar MS Words")
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
        
        
        # Open first document
            self.log_message(f"Öppnar första dokumentet: {os.path.basename(files[0])}")
            first_file = os.path.abspath(files[0])
            doc = word.Documents.Open(first_file)
            doc.Activate()
        
        
        # Merge remaining documents
            for i, file in enumerate(files[1:], 1):
                filename = os.path.basename(file)
                self.log_message(f"Lägger till ({i}/{total_files-1}): {filename}")
                self.update_status(f"Bearbetas fil {i+1} av {total_files}", (i / total_files) * 100)

                try:
                    # Inser section break and then the next file
                    range_end = doc.Content
                    range_end.Collapse(0)
                    
                    range_end.InsertBreak(7)
                    range_end.InsertFile(os.path.abspath(file))
                    
                    time.sleep(0.05)  # <-- PAUS HÄR (50 ms)
                    
                except Exception as e:
                    self.log_message(f"Fel vid infogning av {filename}:{str(e)}")
                    continue
            
            
            # Save final document
            self.update_status("Sparar resultatfil...", 95)
            self.log_message("Sparar sammanfogen fil...")
            
            output_path = os.path.abspath(output_file)
            doc.SaveAs(output_path)
            doc.Close()
            word.Quit()
            
            
            self.update_status("Färdig", 100)
            self.log_message(f"Sammanfogningen klar! Filen sparades som: {output_file}")
            
            messagebox.showinfo("Färdig", f"Sammafogningen är klar!\nFile sparad som:\n{output_file}")
            
        except Exception as e:
            error_msg = f"Fel vid sammanfogning: {str(e)}"
            self.log_message(error_msg)
            
            self.update_status("Fel inträffade")
            messagebox.showerror("Fel", error_msg)
            
            
            # Try to clean up Word if it's still running
            try:
                word = win32com.client.GetActiveObject("Word.Application")
                word.Quit()
            except:
                pass
            
def main():
    root = tk.Tk()
    _ = DocMergerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()