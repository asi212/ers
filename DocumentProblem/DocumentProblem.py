import tkinter as tk                # python 3
from tkinter import font  as tkfont # python 3
import pandas
import numpy
import pdfrw
from pandas import ExcelWriter
import openpyxl
from tkinter import messagebox
import math
import sys
import os
import getpass
from datetime import datetime

global dontUseActiveList
dontUseActiveList = True # make this true if you want people to be able to enter a complaint on any LF, not just on LFs
                          # that have already been created in the EndtestData program

try:
    root = tk.Tk()
    root.geometry("800x750")
    label = tk.Label(text="Instructions: Select or Enter LF nummer. Press Next. Fill out all fields and then save.")
    label.grid(row=1, column=1, columnspan=7, padx=5, pady=10)

    if dontUseActiveList == False:
        # This program looks for active LF's in the same spreadsheet as used by the data entry app
        master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1', dtype = str)
        inactive = master_list[master_list.iloc[:, 19] == '0']
        inactive = inactive.iloc[:,0]
        inactive_list = inactive.values.tolist()
        active = master_list[master_list.iloc[:, 19] == '1']
        active = active.iloc[:,0]
        active_list = ["Select LF"] + active.values.tolist()

        active_label = tk.Label(text="Active LFs")
        active_label.grid(row=2, column=2)
        active_var = tk.StringVar()
        active_var.set("Select LF")
        active_optmenu = tk.OptionMenu(root, active_var, *active_list)
        active_optmenu.grid(row=3, column=2, padx=5, pady=10, sticky="E")

    enter_other_label = tk.Label(text="Other LF:")
    enter_other_label.grid(row=4, column=1, pady=5, padx=5)
    enter_other_entry = tk.Entry(width=10)
    enter_other_entry.grid(row=4, column=2, pady=5, padx=5)

    root.component_label = tk.Label( text="Select Component")
    root.component_label.grid(row=2, column=3)
    root.component_var = tk.StringVar()
    root.component_var.set("         ")
    root.component_optmenu = tk.OptionMenu(root, root.component_var, [])
    root.component_optmenu.grid(row=3, column=3, padx=5, pady=10, sticky="E")
    root.component_label.grid_remove()
    root.component_optmenu.grid_remove()

    error_list = ['BOMB','Drawing', 'Software', 'Bad Material', 'Missing Material', 'Other']
    root.error_label = tk.Label( text="Select Error Type")
    root.error_label.grid(row=2, column=4)
    root.error_var = tk.StringVar()
    root.error_var.set("         ")
    root.error_optmenu = tk.OptionMenu(root, root.error_var, *error_list)
    root.error_optmenu.grid(row=3, column=4, padx=5, pady=10, sticky="E")
    root.error_label.grid_remove()
    root.error_optmenu.grid_remove()

    root.description_label = tk.Label( text="Error Description")
    root.description_label.grid(row=2, column=5, padx=5, pady=10)
    root.description_box = tk.Text()
    root.description_box.configure(width = 20, height = 5)
    root.description_box.grid(row=3, column=5, padx=5, pady=10)
    root.description_label.grid_remove()
    root.description_box.grid_remove()

    root.next_button = tk.Button(text="Next", command=lambda: next())
    root.next_button.grid(row=10, column=1, padx=5, pady=10)
    root.next_button.configure(width=10)

    root.save_button = tk.Button(text="Save", command=lambda: save(lauf_num))
    root.save_button.grid(row=10, column=3, padx=5, pady=10)
    root.save_button.configure(width=10)
    root.save_button.grid_remove()

    root.reset_button = tk.Button(text="Reset", command=lambda: reset())
    root.reset_button.grid(row=10, column=4, padx=5, pady=10)
    root.reset_button.configure(width=10)
    root.reset_button.grid_remove()

    def next():
        global lauf_num
        if dontUseActiveList == False:
            if active_var.get() != "Select LF" and enter_other_entry.get().strip() == "":
                lauf_num = active_var.get()
            elif len(enter_other_entry.get().strip()) > 3:
                lauf_num = enter_other_entry.get().strip()
            elif active_var.get() == "Select LF" and len(enter_other_entry.get().strip()) < 4:
                messagebox.showerror("Error", "Please select or enter a valid LF")
        else:
            if len(enter_other_entry.get().strip()) > 3:
                lauf_num = enter_other_entry.get().strip()
            else:
                messagebox.showerror("Error", "Please enter a valid LF")

        if dontUseActiveList == False:
            # Define paths #
            snxls_path = '//fileserver/produktion/Endtest/30_Seriennummern/Seriennummern.xlsm'  ## path of seriennummern spreadsheet

            # Import XLS #
            D = pandas.read_excel(snxls_path, sheet_name='Serien_Nummern aufsteigend',
                                  index_col=0)  # imports Seriennummern spreadsheet

            alpha = False
            i = 0
            while i < len(lauf_num):
                if lauf_num[i].isalpha() == True:
                    alpha = True
                i = i + 1

            if alpha == True:
                if lauf_num not in D.index.values:
                    messagebox.showerror('Error' , 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                    x = 1 / 0
                else:
                    D = D.loc[str(lauf_num)]
            if alpha == False:
                if len(lauf_num) == 6:
                    if str(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        messagebox.showerror('Error' , 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        x = 1 / 0
                    else:
                        D = D.loc[str(lauf_num)]
                elif len(lauf_num) != 6:
                    if int(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        messagebox.showerror('Error',  'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        x = 1 / 0
                    else:
                        D = D.loc[int(lauf_num)]

            # cancel script if lauf number is listed more than once in the excel file
            if len(
                    D) < 28:  # we say less than 28, because if only 1 exits then it has length 28, but if 2 exist then length 2
                messagebox.showerror("Error", "Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt")
                x = 1 / 0

            #       Index of D:  'Serien Nr.', 'PO-Nr.' 'SN', 'SN Kompl.', 'Quartal', 'Jahr',
            #       'best. Liefertermin', 'Auslieferung', 'Kunde', 'Lieferort',
            #       'Serie gesamt', 'Steuergerät', 'Serie', 'Option I', 'Serie.1', 'Chuck',
            #      'Serie.2', 'Chiller', 'Serie.3', 'Option II', 'Serie.4',
            #       'Softwareversion', 'Bemerkungen', 'Zubehör Option', 'Serie.5',
            #      'Temperatur Bereich', 'Unnamed: 26', 'Unnamed: 27']


            # create list of components
            if D['Steuergerät'] == D['Steuergerät']:
                component_list = [D['Steuergerät']]
            if D['Option I'] == D['Option I']:
                component_list = component_list + [D['Option I']]
            if D['Chuck'] == D['Chuck']:
                component_list = component_list + [D['Chuck']]
            if D['Chiller'] == D['Chiller']:
                component_list = component_list + [D['Chiller']]
            if D['Option II'] == D['Option II']:
                component_list = component_list + [D['Option II']]
        else:
            component_list = ['Chuck', 'Chiller', 'Controller', 'Other (please specify)']

        # add component list to drop down menu
        # Reset var and delete all old options
        root.component_var.set('         ')
        root.component_optmenu['menu'].delete(0, 'end')
        # Insert list of new options (tk._setit hooks them up to var)
        component_list = tuple(component_list) # convert component list to tuple
        for choice in component_list:
            root.component_optmenu['menu'].add_command(label=choice, command=tk._setit(root.component_var, choice))
        #root.component_optmenu = tk.OptionMenu(root, root.component_var, *component_list)


        global sn, name
        if dontUseActiveList == False:
            sn = D['SN Kompl.']
            name = D.name
        else:
            sn = 'NA'
            name = lauf_num

        root.component_label.grid()
        root.component_optmenu.grid()
        root.error_label.grid()
        root.error_optmenu.grid()
        root.save_button.grid()
        root.reset_button.grid()
        root.description_label.grid()
        root.description_box.grid()

        return lauf_num

    def save(lauf_num):
        save = True

        if dontUseActiveList == False:
            if active_var.get() != "Select LF" and enter_other_entry.get().strip() == "":
                lauf_num2 = active_var.get()
            elif len(enter_other_entry.get().strip()) > 3:
                lauf_num2 = enter_other_entry.get().strip()
            elif active_var.get() == "Select LF" and len(enter_other_entry.get().strip()) < 4:
                messagebox.showerror("Error", "Please select or enter a valid LF")
        else:
            if len(enter_other_entry.get().strip()) > 3:
                lauf_num2 = enter_other_entry.get().strip()
            else:
                messagebox.showerror("Error", "Please enter a valid LF")


        if lauf_num2 != lauf_num:
            messagebox.showerror("Error", "LF Number was changed. Please 'Reset' the form and then press 'Next'")
            save = False

        if dontUseActiveList == False:

            # Define paths #
            snxls_path = '//fileserver/produktion/Endtest/30_Seriennummern/Seriennummern.xlsm'  ## path of seriennummern spreadsheet

            # Import XLS #
            D = pandas.read_excel(snxls_path, sheet_name='Serien_Nummern aufsteigend',
                                  index_col=0)  # imports Seriennummern spreadsheet

            alpha = False
            i = 0
            while i < len(lauf_num2):
                if lauf_num2[i].isalpha() == True:
                    alpha = True
                i = i + 1

            if alpha == True:
                if lauf_num2 not in D.index.values:
                    messagebox.showerror('Error' , 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                    x = 1 / 0
                else:
                    D = D.loc[str(lauf_num2)]
            if alpha == False:
                if len(lauf_num2) == 6:
                    if str(
                            lauf_num2) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        messagebox.showerror('Error' , 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        x = 1 / 0
                    else:
                        D = D.loc[str(lauf_num2)]
                elif len(lauf_num2) != 6:
                    if int(
                            lauf_num2) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        messagebox.showerror('Error',  'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        x = 1 / 0
                    else:
                        D = D.loc[int(lauf_num2)]

            # cancel script if lauf number is listed more than once in the excel file
            if len(
                    D) < 28:  # we say less than 28, because if only 1 exits then it has length 28, but if 2 exist then length 2
                messagebox.showerror("Error", "Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt")
                x = 1 / 0


        component = root.component_var.get()
        error = root.error_var.get()
        description = root.description_box.get('1.0', tk.END).strip()

        error_list = pandas.read_excel('X://ERSTools/EndtestData/error_list.xlsx', sheet_name='Sheet1',
                          index_col=None, header=None)  # imports Seriennummern spreadsheet

        new = error_list.loc[0,:]

        new[0] = str(len(error_list) + 1)
        new[1] = lauf_num
        new[2] = getpass.getuser()
        new[3] = datetime.today().strftime('%Y-%m-%d')
        new[4] = component
        if dontUseActiveList == False:
            new[5] = D['SN Kompl.']
        else:
            new[5] = 'NA'
        new[6] = error
        new[7] = description

        error_list = error_list.append(new)

        if save == True:
            writer = ExcelWriter('X://ERSTools/EndtestData/error_list.xlsx')
            error_list.to_excel(writer, 'Sheet1', index=0, header=0)
            try:
                writer.save()
                messagebox.showinfo("Success", "Save Sucessful")
            except:
                pass
                messagebox.showerror("Error", "Save Not Sucessful.") # if you get this message, its probably because the error_list.xlsx is open


    def reset():
        root.save_button.grid_remove()
        root.reset_button.grid_remove()
        root.component_label.grid_remove()
        root.component_optmenu.grid_remove()
        root.error_label.grid_remove()
        root.error_optmenu.grid_remove()
        root.description_label.grid_remove()
        root.description_box.grid_remove()

except BaseException:
    import sys
    print(sys.exc_info()[0])
    import traceback
    print(traceback.format_exc())

finally:
    print("Drücken Sie zum Verlassen die Eingabetaste....")
    input()