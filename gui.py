import tkinter as tk
from tkinter import ttk
import main

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.run_name_label = tk.Label(self, text="Run Name:")
        self.run_name_label.pack()
        self.run_name_entry = tk.Entry(self)
        self.run_name_entry.pack()

        self.xy_count_label = tk.Label(self, text="Number of XY Sequences:")
        self.xy_count_label.pack()
        self.xy_count_entry = tk.Entry(self)
        self.xy_count_entry.pack()

        self.naming_template_label = tk.Label(self, text="Naming Template:")
        self.naming_template_label.pack()
        self.naming_template_entry = tk.Entry(self)
        self.naming_template_entry.pack()

        self.channel_count_label = tk.Label(self, text="Number of Channels:")
        self.channel_count_label.pack()
        self.channel_count_entry = tk.Entry(self)
        self.channel_count_entry.pack()

        self.run_button = tk.Button(self, text="Run", command=self.run_script)
        self.run_button.pack()

        self.output_text = tk.Text(self, height=10, width=50)
        self.output_text.pack()

    def run_script(self):
        run_name = self.run_name_entry.get()
        xy_count = int(self.xy_count_entry.get())
        naming_template = self.naming_template_entry.get()
        channel_count = int(self.channel_count_entry.get())

        placeholder_values = main.get_placeholder_values(xy_count, naming_template)
        channel_orders_list = main.define_channel(channel_count)

        self.output_text.delete('1.0', tk.END)
        self.output_text.insert(tk.END, f"Run Name: {run_name}\n")
        self.output_text.insert(tk.END, f"Number of XY Sequences: {xy_count}\n")
        self.output_text.insert(tk.END, f"Naming Template: {naming_template}\n")
        self.output_text.insert(tk.END, f"Channel Orders: {channel_orders_list}\n")
        self.output_text.insert(tk.END, "Running script...\n")

        main.main(run_name, xy_count, naming_template, placeholder_values, channel_orders_list, self.output_text)

root = tk.Tk()
app = Application(master=root)
app.mainloop()