import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
from MapController import MapController


class MainUI:

    def __init__(self):
        # setup main win
        self.win = tk.Tk()
        self.win.title("Keystone Product Code Mapping")
        self.win.geometry("800x800")
        self.win.resizable(0, 0)
        defaultFontObj = font.nametofont("TkDefaultFont")
        defaultFontObj.config(size=14)
        # setup main card
        self.maincard = MainCard(self.win)
        self.maincard.main.pack(fill="both")
        # setup register card
        self.registercard = RegisterCard(self.win)
        # setup controller
        self.controller = MapController()
        self.controller.open_map_json()
        # setup btn commands
        self.build_btn_commands()

    def build_btn_commands(self):
        self.maincard.btn_query.config(command=self.query)
        self.maincard.btn_register.config(command=self.go_to_register)
        self.maincard.btn_export_excel.config(command=self.export_excel)
        self.maincard.btn_import_excel.config(command=self.import_excel)
        self.maincard.btn_batch.config(command=self.batch_mapping)
        self.registercard.btn_back.config(command=self.back_to_home)
        self.registercard.btn_register.config(command=self.register)

    def back_to_home(self):
        self.registercard.main.pack_forget()
        self.maincard.main.pack(fill="x")

    def go_to_register(self):
        code = self.maincard.entry.get()
        self.maincard.main.pack_forget()
        self.registercard.regcode.delete(0, tk.END)
        self.registercard.regcode.insert(0, code)
        self.registercard.main.pack(fill="x")

    def query(self):
        code = self.maincard.entry.get()
        result = self.controller.query(code)
        if result == False:
            self.maincard.result_entry.delete(0, tk.END)
            self.maincard.result_entry.insert(0, "No records")
        elif result == True:
            self.maincard.result_entry.delete(0, tk.END)
            self.maincard.result_entry.insert(0, "This is a standard code")
        else:
            self.maincard.result_entry.delete(0, tk.END)
            self.maincard.result_entry.insert(0, result)

    def register(self):
        code = self.registercard.regcode.get()
        choose = self.registercard.var.get()
        if choose == 0:
            check = self.controller.add_standard(code)
            if check == False:
                messagebox.showerror(
                    title="Error", message="Error on adding standard code.\nYou may register a non-standard code to standard code.")
            else:
                messagebox.showinfo(
                    title="Add Standard Code",
                    message="Add Success!"
                )
        elif choose == 1:
            mapcode = self.registercard.entry.get()
            check = self.controller.add_non_standard(code, mapcode)
            if check == False:
                messagebox.showerror(
                    title="Error", message="Error on adding standard code.\nYou may mapping a standard code before register it.")
            else:
                messagebox.showinfo(
                    title="Add Non Standard Code",
                    message="Add Success!"
                )
        self.back_to_home()

    def export_excel(self):
        path = filedialog.asksaveasfilename(filetypes=(
            ("Excel Files", "*.xlsx"),), defaultextension=".xlsx")
        if not path == None and not path == "":
            check = self.controller.export_excel(path)
            if check == True:
                messagebox.showinfo(title="Export to Excel",
                                    message="Export Success!")
            elif check == False:
                messagebox.showerror(
                    title="Export to Excel", message="Failed to export.")

    def import_excel(self):
        path = filedialog.askopenfilename(
            filetypes=(("Excel Files", "*.xlsx"),))
        if not path == None and not path == "":
            check = self.controller.import_from_excel(path)
            if check == True:
                messagebox.showinfo(title="Import from Excel",
                                    message="Import Success!")
            elif check == False:
                messagebox.showerror(
                    title="Import from Excel", message="Failed to import.\nIt may because you are using unrecogized format of excel.\nPlease use the format same as exporting excel file.")
            else:
                messagebox.showerror(
                    title="Import from Excel", message="Error:\n{}".format(check)
                )

    def batch_mapping(self):
        path = filedialog.askopenfilename(
            filetypes=(("Excel Files", "*.xlsx"),))
        if not path == None and not path == "":
            check = self.controller.batch_mapping(path)
            if check == True:
                messagebox.showinfo(title="Batch Mapping",
                                    message="Batch Mapping Done.")
            else:
                messagebox.showerror(
                    title="Batch Mapping", message="Error on batch mapping.\tThis error may cause by user opening the target excel file.")

    def mainloop(self):
        self.win.mainloop()


class MainCard:

    def __init__(self, master):

        self.main = ttk.Frame(master)
        self.sub = ttk.Frame(self.main)
        self.sub.pack(fill="x", pady=100)

        lbl = ttk.Label(self.sub, text="Keystone Product Code Mapping")
        lbl.pack()

        lbl = ttk.Label(self.sub, text="Please enter a code:")
        lbl.pack(fill="x")

        self.entry = ttk.Entry(self.sub, width=50, font=(18))
        self.entry.pack()

        f = ttk.Frame(self.sub)
        f.pack()
        self.btn_query = ttk.Button(f, text="Query")
        self.btn_query.grid(row=0, column=0)
        self.btn_register = ttk.Button(f, text="Register")
        self.btn_register.grid(row=0, column=1)

        lbl = ttk.Label(self.sub, text="Query Result:")
        lbl.pack(fill="x")

        self.result_entry = ttk.Entry(self.sub, width=50, font=(18))
        self.result_entry.pack()

        f = ttk.Frame(self.sub)
        f.pack(fill="x", pady=30)
        self.btn_export_excel = ttk.Button(f, text="Export to excel")
        self.btn_export_excel.grid(row=0, column=0)
        self.btn_import_excel = ttk.Button(f, text="Import from excel")
        self.btn_import_excel.grid(row=0, column=1)

        f = ttk.Frame(self.sub)
        f.pack(fill="x", pady=30)
        lbl = ttk.Label(f, text="Batch Codes Mapping")
        lbl.pack(pady=5)
        lbl = ttk.Label(f, text="How to Use:\n\t1. open a blank excel book.\n\t2. place all codes you want to map in column A. (no header is needed)\n\t3. press button below to open the excel book.\n\t4. the results will be written in column B.")
        lbl .pack(pady=5)
        self.btn_batch = ttk.Button(f, text="Batch Mapping")
        self.btn_batch.pack()


class RegisterCard:

    def __init__(self, master):

        self.main = ttk.Frame(master)
        self.sub = ttk.Frame(self.main)
        self.sub.pack(fill="x", pady=100)

        f = ttk.Frame(self.sub)
        f.pack(fill="x", pady=15)
        self.btn_back = ttk.Button(f, text="Back")
        self.btn_back.grid(row=0, column=0)

        lbl = ttk.Label(self.sub, text="The code to be register:")
        lbl.pack(fill="x", pady=15)
        self.regcode = ttk.Entry(
            self.sub, width=50, font=(18))
        self.regcode.pack()

        self.var = tk.IntVar()
        self.var.set(0)
        radio = ttk.Radiobutton(
            self.sub, variable=self.var, value=0, text="Register as Standard Code", command=self.disable_entry)
        radio.pack()
        radio = ttk.Radiobutton(
            self.sub, variable=self.var, value=1, text="Register as Non-Standard Code", command=self.enable_entry)
        radio.pack()

        self.lbl = ttk.Label(
            self.sub, text="The code to be map: (MUST be a registered standard code)", state="disable")
        self.lbl.pack(fill="x", pady=15)
        self.entry = ttk.Entry(self.sub, width=50, font=(18), state="disable")
        self.entry.pack()

        self.btn_register = ttk.Button(self.sub, text="Register")
        self.btn_register.pack()

        f = ttk.Frame(self.sub)
        f.pack(fill="x", pady=15)
        lbl = ttk.Label(f, text="Guidance")
        lbl.pack()
        lbl = ttk.Label(f, text="\t1. 'Standard' Code is the code in M18.\n\t2. 'Non-standard' Code is the code you want to map with M18 Codes.\n\t3. The code can only be standard or non-standard, but cannot be both.")
        lbl.pack()

    def disable_entry(self):
        self.entry.config(state="disable")
        self.lbl.config(state="disable")

    def enable_entry(self):
        self.entry.config(state="normal")
        self.lbl.config(state="normal")
