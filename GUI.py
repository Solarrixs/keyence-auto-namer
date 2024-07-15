import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import main  # Import your main script

class KeyenceAutoNamerGUI:
    def __init__(self, master):
        self.master = master
        master.title("Keyence Auto Namer")
        master.geometry("600x800")

        self.run_configs = []
        self.channel_orders_list = []
        self.create_widgets()

    def create_widgets(self):
        # Create a main frame
        main_frame = ttk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=1)

        # Create a canvas
        self.canvas = tk.Canvas(main_frame)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        # Add a scrollbar to the canvas
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure the canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Create another frame inside the canvas
        self.second_frame = ttk.Frame(self.canvas)

        # Add that new frame to a window in the canvas
        self.canvas.create_window((0, 0), window=self.second_frame, anchor="nw")

        # Run Configuration Frame
        run_frame = ttk.LabelFrame(self.second_frame, text="Run Configuration")
        run_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(run_frame, text="Run Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.run_name = ttk.Entry(run_frame)
        self.run_name.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(run_frame, text="Stitch Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.stitch_type = ttk.Combobox(run_frame, values=["Full", "Load"])
        self.stitch_type.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(run_frame, text="Overlay Image:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.overlay = ttk.Combobox(run_frame, values=["Yes", "No"])
        self.overlay.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(run_frame, text="Naming Template:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.naming_template = ttk.Entry(run_frame)
        self.naming_template.grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(run_frame, text="Save Filepath:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.filepath = ttk.Entry(run_frame)
        self.filepath.grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(run_frame, text="Browse", command=self.browse_filepath).grid(row=4, column=2, padx=5, pady=5)

        ttk.Label(run_frame, text="Start XY:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.start_xy = ttk.Entry(run_frame)
        self.start_xy.grid(row=5, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(run_frame, text="End XY:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        self.end_xy = ttk.Entry(run_frame)
        self.end_xy.grid(row=6, column=1, sticky="ew", padx=5, pady=5)

        ttk.Button(run_frame, text="Add Run", command=self.add_run).grid(row=7, column=0, columnspan=2, pady=10)

        # Run Configurations List
        self.run_listbox = tk.Listbox(self.second_frame, width=70, height=10)
        self.run_listbox.pack(padx=10, pady=10)

        # Channel Configuration Frame
        channel_frame = ttk.LabelFrame(self.second_frame, text="Channel Configuration")
        channel_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(channel_frame, text="Number of Channels:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.channel_count = ttk.Entry(channel_frame)
        self.channel_count.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        ttk.Button(channel_frame, text="Set Channels", command=self.set_channels).grid(row=1, column=0, columnspan=2, pady=10)

        self.channel_frame = ttk.Frame(self.second_frame)
        self.channel_frame.pack(padx=10, pady=10, fill="x")

        # Run Button
        ttk.Button(self.second_frame, text="Run Program", command=self.run_program).pack(pady=20)

    def browse_filepath(self):
        filepath = filedialog.askdirectory()
        self.filepath.delete(0, tk.END)
        self.filepath.insert(0, filepath)

    def add_run(self):
        run_config = (
            self.run_name.get(),
            self.stitch_type.get()[0].upper(),
            self.overlay.get()[0].upper(),
            self.naming_template.get(),
            self.filepath.get(),
            int(self.start_xy.get()),
            int(self.end_xy.get())
        )
        self.run_configs.append(run_config)
        self.run_listbox.insert(tk.END, f"Run: {run_config[0]}, XY: {run_config[5]}-{run_config[6]}")

        # Clear all fields
        self.run_name.delete(0, tk.END)
        self.stitch_type.set('')
        self.overlay.set('')
        self.naming_template.delete(0, tk.END)
        self.filepath.delete(0, tk.END)
        self.start_xy.delete(0, tk.END)
        self.end_xy.delete(0, tk.END)

    def set_channels(self):
        count = int(self.channel_count.get())
        
        # Clear existing channel entries
        for widget in self.channel_frame.winfo_children():
            widget.destroy()

        self.channel_entries = []
        for i in range(count):
            ttk.Label(self.channel_frame, text=f"Channel {i+1}:").grid(row=i, column=0, sticky="w", padx=5, pady=5)
            entry = ttk.Entry(self.channel_frame)
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
            self.channel_entries.append(entry)

    def get_placeholder_values(self, naming_template, start_child, end_child):
        placeholder_values = {}
        xy_names = [f"XY{i+1:02}" for i in range(start_child - 1, end_child)]

        placeholder_count = sum(1 for char in naming_template if char == '{') - naming_template.count('{C}')

        for xy_name in xy_names:
            placeholder_values[xy_name] = {}
            for placeholder in range(1, placeholder_count + 1):
                value = simpledialog.askstring("Input", f"Enter placeholder {{{placeholder}}} for {xy_name}")
                placeholder_values[xy_name][str(placeholder)] = value

        return placeholder_values

    def run_program(self):
        self.channel_orders_list = [entry.get() for entry in self.channel_entries]
        self.channel_orders_list.reverse()
        main.channel_orders_list = self.channel_orders_list

        try:
            main.app = main.pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer")
            main.main_window = main.app.window(title="BZ-X800 Analyzer")
            main.stitch_button = main.app.window(title="BZ-X800 Analyzer").child_window(title="Stitch")
        except Exception:
            messagebox.showerror("Error", "BZ-X800 Analyzer not found. Please open the application and try again.")
            return

        failed = []
        for run_config in self.run_configs:
            run_name, stitchtype, overlay, naming_template, filepath, start_child, end_child = run_config
            placeholder_values = self.get_placeholder_values(naming_template, start_child, end_child)

            try:
                main.process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, start_child, end_child, placeholder_values, filepath)
            except Exception as e:
                messagebox.showerror("Error", f"Failed on running {run_name}. Error: {e}")

        if failed:
            messagebox.showinfo("Process Complete", f"All XY sequences have been processed except for: {failed}")
        else:
            messagebox.showinfo("Process Complete", "All XY sequences have been processed successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = KeyenceAutoNamerGUI(root)
    root.mainloop()