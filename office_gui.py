import tkinter as tk
from tkinter import ttk, messagebox
from office_core import create_script_document

class OfficeScriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("The Office Script Generator")
        self.root.geometry("400x300")
        
        # Create main frame with padding
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Season selection
        ttk.Label(main_frame, text="Season:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.season_var = tk.StringVar()
        season_combo = ttk.Combobox(main_frame, textvariable=self.season_var, width=5)
        season_combo['values'] = tuple(range(1, 10))  # Seasons 1-9
        season_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Episodes input
        ttk.Label(main_frame, text="Episodes (comma-separated):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.episodes_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.episodes_var, width=15).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Generate button
        ttk.Button(main_frame, text="Generate Scripts", command=self.generate_scripts).grid(row=2, column=0, columnspan=2, pady=20)
        
        # Status label
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var, wraplength=350).grid(row=3, column=0, columnspan=2, pady=5)

    def generate_scripts(self):
        season = self.season_var.get()
        episodes = self.episodes_var.get().strip()
        
        if not season or not episodes:
            messagebox.showerror("Error", "Please enter both season and episode numbers")
            return
            
        try:
            episodes_list = episodes.split(',')
            # Validate input
            season_num = int(season)
            [int(ep.strip()) for ep in episodes_list]  # Validate all episodes are numbers
            
            self.status_var.set("Generating scripts...")
            self.root.update()
            
            success, message, _ = create_script_document(str(season_num), episodes_list)
            self.status_var.set(message)
            if not success:
                messagebox.showerror("Error", message)
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers")

def main():
    root = tk.Tk()
    app = OfficeScriptGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()