import tkinter as tk
from tkinter import filedialog
import os

class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

        self.current_master_file = ''
        self.current_chosen_dir = ''

    def open_directory(self):
        dir = filedialog.askdirectory()
        print(dir)
        self.current_chosen_dir = str(dir)
        tk.Label(self, text = 'The package selected is located at: ' + self.current_chosen_dir).pack(padx = 10, pady = 5)

    def open_file(self):
        file = filedialog.askopenfile(mode = 'r')

        if file:
            filepath = os.path.abspath(file.name)
            self.current_master_file = str(filepath)
            print(self.current_master_file)
            tk.Label(self, text = 'The file selected is located at: ' + str(file.name)).pack(padx = 10, pady = 5)
            

    def create_widgets(self):

        # self.geometry('1000x500')
        # self.title("Scorecard Compiler")
        # self.minsize(500, 400)

        # Select operating system radiobutton
        self.os_label = tk.Label(self, text = 'Select your operating system:').pack(side = 'top', padx = 10, pady = (20, 8))
        self.os_var = tk.StringVar(value = 'Windows')
        self.os_button1 = tk.Radiobutton(self, text = 'Windows', variable = self.os_var, value = 'Windows').pack(side = 'top', padx = 10 , pady = (0, 4))
        self.os_button2 = tk.Radiobutton(self, text = 'Mac', variable = self.os_var, value = 'Mac').pack(side = 'top', padx = 10 , pady = (0, 12))

        # Master file input
        self.master_file_label = tk.Label(self, text = 'Which file would you like to add a package of scorecard sheets to?').pack(padx = 10, pady = (10, 4))
        tk.Button(self, text = "Select Master File", command = self.open_file).pack(padx = 10, pady = (0, 10))

        self.select_directory_label = tk.Label(self, text = 'What folder are your scorecards located in?').pack( padx = 10, pady = (10, 4))
        tk.Button(self, text = "Select Package", command = self.open_directory).pack(padx = 10, pady = (0, 20))
        



        
if __name__ == '__main__':
    root = tk.Tk()
    app = App(master = root)
    app.mainloop()