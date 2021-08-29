try:
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

except BaseException:
    import sys
    print(sys.exc_info()[0])
    import traceback
    print(traceback.format_exc())


# Notes #
# The layout pattern used is called "grid". "grid" specifies a the place of the widget (e.g. text box, menu, button) on the page
# grid_remove() hides a widget. grid() displays it.


global hideForChuckDept
global hideForFTDept
hideForChuckDept = False
hideForFTDept = False

try:
    class SampleApp(tk.Tk):
        def __init__(self, *args, **kwargs):
            tk.Tk.__init__(self, *args, **kwargs)

            self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")
            self.label_font = tkfont.Font(family='Helvetica', size=14, weight="bold")
            self.entry_font = tkfont.Font(family='Helvetica', size=10, weight="bold")
            self.geometry("720x800+415+27")

            # the container is where we'll stack a bunch of frames
            # on top of each other, then the one we want visible
            # will be raised above the others
            container = tk.Frame(self)
            #container.pack(side="top", fill="both", expand=True)
            container.grid()
            container.grid_rowconfigure(0, weight=1)
            container.grid_columnconfigure(0, weight=1)


            # Make sure to add any new classes / pages to this list
            self.frames = {}
            for F in (StartPage, CreateNew, EditList, ModLF, DocDev, Manager, TempLow, TempMed, TempHigh, TempOthr1, TempOthr2,
                      TempOthr3, PlanLow, PlanMed, PlanHigh, PlanOthr1, PlanOthr2, PlanOthr3, Pt100):
                page_name = F.__name__
                frame = F(parent=container, controller=self)
                self.frames[page_name] = frame

                # put all of the pages in the same location;
                # the one on the top of the stacking order
                # will be the one that is visible.
                frame.grid(row=0, column=0, sticky="nsew")

            self.show_frame("StartPage")

        # this function provides a reference to a different class, so that you can access variables of any class from within a different class
        def get_page(self, page_class):
            return self.frames[page_class]

        # show a frame for a given page name
        def show_frame(self, page_name):
            frame = self.frames[page_name]
            frame.tkraise()

        # this function updates the Manger page to display the correct LF number, the correct tests, test status, and corresponding temperatures.
        def update_manager(self, lf):
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                            dtype=str, header=None)

            new_title = "LF " + lf + " EndTest Tasks"
            self.get_page("Manager").label.configure(text=new_title)
            #LF_status = manager_list[manager_list[0].str.match(lf)]

            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] == lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]


            self.lowTemp = int(LF_status[2])
            self.medTemp = int(LF_status[3])
            self.highTemp = int(LF_status[4])
            self.othr1Temp = int(LF_status[23])
            self.othr2Temp = int(LF_status[24])
            self.othr3Temp = int(LF_status[25])

            self.last4 = LF_status[20].tolist()[0]
            self.folder_path = LF_status[21].tolist()[0]

            self.type = ""
            if str((LF_status[1])) == "STD":
                self.type = "STD"
            if str((LF_status[1])) == "LN":
                self.type = "LN"
            if str((LF_status[1])) == "AM":
                self.type = "AM"
            if str((LF_status[1])) == "HTU":
                self.type = "HTU"

            if str(LF_status[2].tolist())[2:-2] != '999':
                if hideForChuckDept == False:
                    self.get_page("Manager").temp1.grid()
                    self.get_page("Manager").temp1_status.grid()
                if hideForFTDept == False:
                    self.get_page("Manager").plan1.grid()
                    self.get_page("Manager").plan1_status.grid()
                lowtemp_title = "Temp Uniformity " + str(int(LF_status[2])) + "C"
                lowplan_title = "Planarity " + str(int(LF_status[2])) + "C"
                self.get_page("Manager").temp1.configure(text=lowtemp_title)
                self.get_page("TempLow").label.configure(text=lowtemp_title)
                self.get_page("Manager").plan1.configure(text=lowplan_title)
                self.get_page("PlanLow").label.configure(text=lowplan_title)
                self.get_page("Manager").temp1_status.configure(bg="red", text="  incomplete")
                self.get_page("Manager").plan1_status.configure(bg="red", text="  incomplete")
                if int(LF_status[8]) == 1:
                    self.get_page("Manager").temp1_status.configure(bg="green", text="  complete")
                    self.get_page("TempLow").checkComplete_var.set(True)
                if int(LF_status[11]) == 1:
                    self.get_page("Manager").plan1_status.configure(bg="green", text="  complete")
                    self.get_page("PlanLow").checkComplete_var.set(True)
            else:
                self.get_page("Manager").temp1.grid_remove()
                self.get_page("Manager").temp1_status.grid_remove()
                self.get_page("Manager").plan1.grid_remove()
                self.get_page("Manager").plan1_status.grid_remove()


            if str(LF_status[3].tolist())[2:-2] != '999':
                if hideForChuckDept == False:
                    self.get_page("Manager").temp2.grid()
                    self.get_page("Manager").temp2_status.grid()
                if hideForFTDept == False:
                    self.get_page("Manager").plan2.grid()
                    self.get_page("Manager").plan2_status.grid()
                medtemp_title = "Temp Uniformity " + str(int(LF_status[3])) + "C"
                medplan_title = "Planarity " + str(int(LF_status[3])) + "C"
                self.get_page("Manager").temp2.configure(text=medtemp_title)
                self.get_page("TempMed").label.configure(text=medtemp_title)
                self.get_page("Manager").plan2.configure(text=medplan_title)
                self.get_page("PlanMed").label.configure(text=medplan_title)
                self.get_page("Manager").temp2_status.configure(bg="red", text="  incomplete")
                self.get_page("Manager").plan2_status.configure(bg="red", text="  incomplete")
                if int(LF_status[9]) == 1:
                    self.get_page("Manager").temp2_status.configure(bg="green", text="  complete")
                    self.get_page("TempMed").checkComplete_var.set(True)
                if int(LF_status[12]) == 1:
                    self.get_page("Manager").plan2_status.configure(bg="green", text="  complete")
                    self.get_page("PlanMed").checkComplete_var.set(True)
            else:
                self.get_page("Manager").temp2.grid_remove()
                self.get_page("Manager").temp2_status.grid_remove()
                self.get_page("Manager").plan2.grid_remove()
                self.get_page("Manager").plan2_status.grid_remove()


            if str(LF_status[4].tolist())[2:-2] != '999':
                if hideForChuckDept == False:
                    self.get_page("Manager").temp3.grid()
                    self.get_page("Manager").temp3_status.grid()
                if hideForFTDept == False:
                    self.get_page("Manager").plan3.grid()
                    self.get_page("Manager").plan3_status.grid()
                hightemp_title = "Temp Uniformity " + str(int(LF_status[4])) + "C"
                highplan_title = "Planarity " + str(int(LF_status[4])) + "C"
                self.get_page("Manager").temp3.configure(text=hightemp_title)
                self.get_page("TempHigh").label.configure(text=hightemp_title)
                self.get_page("Manager").plan3.configure(text=highplan_title)
                self.get_page("PlanHigh").label.configure(text=highplan_title)
                self.get_page("Manager").temp3_status.configure(bg="red", text="  incomplete")
                self.get_page("Manager").plan3_status.configure(bg="red", text="  incomplete")
                if int(LF_status[10]) == 1:
                    self.get_page("Manager").temp3_status.configure(bg="green", text="  complete")
                    self.get_page("TempHigh").checkComplete_var.set(True)
                if int(LF_status[13]) == 1:
                    self.get_page("Manager").plan3_status.configure(bg="green", text="  complete")
                    self.get_page("PlanHigh").checkComplete_var.set(True)
            else:
                self.get_page("Manager").temp3.grid_remove()
                self.get_page("Manager").temp3_status.grid_remove()
                self.get_page("Manager").plan3.grid_remove()
                self.get_page("Manager").plan3_status.grid_remove()

            ###
            if str(LF_status[23].tolist())[2:-2] != '999':
                if hideForChuckDept == False:
                    self.get_page("Manager").temp4.grid()
                    self.get_page("Manager").temp4_status.grid()
                if hideForFTDept == False:
                    self.get_page("Manager").plan4.grid()
                    self.get_page("Manager").plan4_status.grid()
                othr1temp_title = "Temp Uniformity " + str(int(LF_status[23])) + "C"
                othr1plan_title = "Planarity " + str(int(LF_status[23])) + "C"
                self.get_page("Manager").temp4.configure(text=othr1temp_title)
                self.get_page("TempOthr1").label.configure(text=othr1temp_title)
                self.get_page("Manager").plan4.configure(text=othr1plan_title)
                self.get_page("PlanOthr1").label.configure(text=othr1plan_title)
                self.get_page("Manager").temp4_status.configure(bg="red", text="  incomplete")
                self.get_page("Manager").plan4_status.configure(bg="red", text="  incomplete")
                if int(LF_status[29]) == 1:
                    self.get_page("Manager").temp4_status.configure(bg="green", text="  complete")
                    self.get_page("TempOthr1").checkComplete_var.set(True)
                if int(LF_status[32]) == 1:
                    self.get_page("Manager").plan4_status.configure(bg="green", text="  complete")
                    self.get_page("PlanOthr1").checkComplete_var.set(True)
            else:
                self.get_page("Manager").temp4.grid_remove()
                self.get_page("Manager").temp4_status.grid_remove()
                self.get_page("Manager").plan4.grid_remove()
                self.get_page("Manager").plan4_status.grid_remove()


            if str(LF_status[24].tolist())[2:-2] != '999':
                if hideForChuckDept == False:
                    self.get_page("Manager").temp5.grid()
                    self.get_page("Manager").temp5_status.grid()
                if hideForFTDept == False:
                    self.get_page("Manager").plan5.grid()
                    self.get_page("Manager").plan5_status.grid()
                othr2temp_title = "Temp Uniformity " + str(int(LF_status[24])) + "C"
                othr2plan_title = "Planarity " + str(int(LF_status[24])) + "C"
                self.get_page("Manager").temp5.configure(text=othr2temp_title)
                self.get_page("TempOthr2").label.configure(text=othr2temp_title)
                self.get_page("Manager").plan5.configure(text=othr2plan_title)
                self.get_page("PlanOthr2").label.configure(text=othr2plan_title)
                self.get_page("Manager").temp5_status.configure(bg="red", text="  incomplete")
                self.get_page("Manager").plan5_status.configure(bg="red", text="  incomplete")
                if int(LF_status[30]) == 1:
                    self.get_page("Manager").temp5_status.configure(bg="green", text="  complete")
                    self.get_page("TempOthr2").checkComplete_var.set(True)
                if int(LF_status[33]) == 1:
                    self.get_page("Manager").plan5_status.configure(bg="green", text="  complete")
                    self.get_page("PlanOthr2").checkComplete_var.set(True)
            else:
                self.get_page("Manager").temp5.grid_remove()
                self.get_page("Manager").temp5_status.grid_remove()
                self.get_page("Manager").plan5.grid_remove()
                self.get_page("Manager").plan5_status.grid_remove()

            if str(LF_status[25].tolist())[2:-2] != '999':
                if hideForChuckDept == False:
                    self.get_page("Manager").temp6.grid()
                    self.get_page("Manager").temp6_status.grid()
                if hideForFTDept == False:
                    self.get_page("Manager").plan6.grid()
                    self.get_page("Manager").plan6_status.grid()
                othr3temp_title = "Temp Uniformity " + str(int(LF_status[25])) + "C"
                othr3plan_title = "Planarity " + str(int(LF_status[25])) + "C"
                self.get_page("Manager").temp6.configure(text=othr3temp_title)
                self.get_page("TempOthr3").label.configure(text=othr3temp_title)
                self.get_page("Manager").plan6.configure(text=othr3plan_title)
                self.get_page("PlanOthr3").label.configure(text=othr3plan_title)
                self.get_page("Manager").temp6_status.configure(bg="red", text="  incomplete")
                self.get_page("Manager").plan6_status.configure(bg="red", text="  incomplete")
                if int(LF_status[31]) == 1:
                    self.get_page("Manager").temp6_status.configure(bg="green", text="  complete")
                    self.get_page("TempOthr3").checkComplete_var.set(True)
                if int(LF_status[34]) == 1:
                    self.get_page("Manager").plan6_status.configure(bg="green", text="  complete")
                    self.get_page("PlanOthr3").checkComplete_var.set(True)
            else:
                self.get_page("Manager").temp6.grid_remove()
                self.get_page("Manager").temp6_status.grid_remove()
                self.get_page("Manager").plan6.grid_remove()
                self.get_page("Manager").plan6_status.grid_remove()


            if LF_status[14].tolist()[0] == '1':
                if hideForChuckDept == False:
                    self.get_page("Manager").pt100.grid()
                    self.get_page("Manager").pt100_status.grid()
                hightemp_title = "PT100 Calibration"
                self.get_page("Manager").pt100.configure(text=hightemp_title)
                self.get_page("Manager").pt100_status.configure(bg="red", text="  incomplete")
                if int(LF_status[15]) == 1:
                    self.get_page("Manager").pt100_status.configure(bg="green", text="  complete")
            else:
                self.get_page("Manager").pt100.grid_remove()
                self.get_page("Manager").pt100_status.grid_remove()

        # this is used whenever we want to call multiple functions directly from the command of a button
        def combine_funcs(*funcs):
            def combined_func(*args, **kwargs):
                for f in funcs:
                    f(*args, **kwargs)
            return combined_func

        def return_type(self):
            return self.type

        #def focus_next_window(event):
         #   event.widget.tk_focusNext().focus()
          #  return ("break")

    class StartPage(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            label = tk.Label(self, text="Welcome", font=controller.title_font)
            label.grid(row=1, column=2, padx = 250, pady=10)

            ## Drop down list / option menu ###
            active_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1', dtype = str)
            active_list = active_list[active_list.iloc[:, 19] == '1'] # if it is active in the EditList Class (add / remove datalogs page), then keep the row
            active_list = active_list.iloc[:,0] # look at first column (index = 0), which contains the LF numbers
            active_list = active_list.values.tolist() # use .tolist() to get into array of strings
            self.choice_var = tk.StringVar()
            self.choice_var.set("Select Lauf Nummer")
            self.LF_menu = tk.OptionMenu(self,self.choice_var,*active_list) # set the options of the dropdown menu created
            self.LF_menu.grid(row=2,column=2, padx=250, pady=10)
            self.LF_menu.configure(width=20)


            continue_button = tk.Button(self, text="Continue", command=self.continue_func) # command=self.continue_func will wait for button press to call function,
                                                                                         # command=self.continue_func() will execute as soon as page is loaded (WE DON"T WANT THIS)
            continue_button.grid(row=3,column=2, padx=250, pady=10)
            continue_button.configure(width=22)

            newLF_button = tk.Button(self, text="Create New Data Log", command=lambda: controller.show_frame("CreateNew")) # if a function must be called with "()" at the end,
                                                                                                                            # then use lambda so that it is not called until a button press
            newLF_button.grid(row=4,column=2, padx=250, pady=10)
            newLF_button.configure(width=22)
            if hideForChuckDept == True or hideForFTDept == True:
                newLF_button.grid_remove()


            editList_button = tk.Button(self, text="Add/Remove Data Logs", command=lambda: controller.show_frame("EditList"))
            editList_button.grid(row=5,column=2, padx=250, pady=10)
            editList_button.configure(width=22)
            if hideForChuckDept == True or hideForFTDept == True:
                editList_button.grid_remove()


            docDev_button = tk.Button(self, text="Document Deviation",
                                       command=lambda: controller.show_frame("DocDev"))
            docDev_button.grid(row=6, column=2, padx=250, pady=10)
            docDev_button.configure(width=22)

        def continue_func(self):
            if self.choice_var.get() != "Select Lauf Nummer":
                lf = str(self.choice_var.get())                 # this returns the LF selected in the menu
                self.controller.update_manager(lf)              # call update manager function in the SampleApp class
                self.controller.show_frame("Manager")           # show manager page
            else:
                messagebox.showinfo("Error", "Please Select a Lauf Nummer")

        # this function just updates the drop down list of LFs. It is called after new LFs are created, or after LFs are modified
        def update_LF_menu(self):
            active_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                        dtype=str)
            active_list = active_list[active_list.iloc[:, 19] == '1']
            active_list = active_list.iloc[:, 0]
            active_list = active_list.values.tolist()
            self.LF_menu['menu'].delete(0, 'end')
            self.choice_var = tk.StringVar()
            self.choice_var.set("Select Lauf Nummer")
            self.LF_menu = tk.OptionMenu(self,self.choice_var,*active_list)
            self.LF_menu.grid(row=2,column=2, padx=250, pady=10)
            self.LF_menu.configure(width=20)

    class CreateNew(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="Chuck Model: ", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=7, padx = 0, pady=10)

            # this sixpoints varible is changed to True IF 6 temperature points are needed for an HTU chuck (because HTU chuck needs 6 temperatures, while other chucks only need 3)
            global sixpoints
            sixpoints = False

            newLF_label = tk.Label(self, text="New LF")
            newLF_label.grid(row=2,column=1, padx = 50, pady = 10, sticky = "E")
            self.newLF_textbox = tk.Text(self, height=1, width = 10)
            self.newLF_textbox.grid(row=2,column=2,padx=0,pady=10, sticky = "W")

            self.check_var = tk.IntVar()
            self.checkBox = tk.Checkbutton(self, text="Kein Chuck", variable = self.check_var)
            self.checkBox.grid(row=2,column=4,padx=0,pady=10, sticky = "W")

            tempLow_label = tk.Label(self, text="Test Temperature Low")
            tempLow_label.grid(row=3,column=1, padx = 5, pady = 10, sticky = "E")
            self.tempLow_textbox = tk.Text(self, height=1, width = 10)
            self.tempLow_textbox.grid(row=3,column=2,padx=0,pady=10, sticky = "W")

            tempMed_label = tk.Label(self, text="Test Temp Intermediate (25C)")
            tempMed_label.grid(row=4,column=1, padx = 5, pady = 10, sticky = "E")
            self.tempMed_textbox = tk.Text(self, height=1, width = 10)
            self.tempMed_textbox.grid(row=4,column=2,padx=0,pady=10, sticky = "W")

            tempHigh_label = tk.Label(self, text="Test Temperature High")
            tempHigh_label.grid(row=5,column=1, padx = 5, pady = 10, sticky = "E")
            self.tempHigh_textbox = tk.Text(self, height=1, width = 10)
            self.tempHigh_textbox.grid(row=5,column=2,padx=0,pady=10, sticky = "W")


            self.tempOthr1_label = tk.Label(self, text="Test Temperature Other1")
            self.tempOthr1_label.grid(row=6,column=1, padx = 5, pady = 10, sticky = "E")
            self.tempOthr1_textbox = tk.Text(self, height=1, width = 10)
            self.tempOthr1_textbox.grid(row=6,column=2,padx=0,pady=10, sticky = "W")
            self.tempOthr1_label.grid_remove()
            self.tempOthr1_textbox.grid_remove()

            self.tempOthr2_label = tk.Label(self, text="Test Temperature Other2")
            self.tempOthr2_label.grid(row=7,column=1, padx = 5, pady = 10, sticky = "E")
            self.tempOthr2_textbox = tk.Text(self, height=1, width = 10)
            self.tempOthr2_textbox.grid(row=7,column=2,padx=0,pady=10, sticky = "W")
            self.tempOthr2_label.grid_remove()
            self.tempOthr2_textbox.grid_remove()

            self.tempOthr3_label = tk.Label(self, text="Test Temperature Other3")
            self.tempOthr3_label.grid(row=8,column=1, padx = 5, pady = 10, sticky = "E")
            self.tempOthr3_textbox = tk.Text(self, height=1, width = 10)
            self.tempOthr3_textbox.grid(row=8,column=2,padx=0,pady=10, sticky = "W")
            self.tempOthr3_label.grid_remove()
            self.tempOthr3_textbox.grid_remove()

            #### Create Spec boxes / menus
            list = ["deg C", "% Temp"]          # this list of strings becomes the choices of the drop down menu
            self.choice_var_low = tk.StringVar()
            self.choice_var_low.set(list[0])

            self.choice_var_med = tk.StringVar()
            self.choice_var_med.set(list[0])

            self.choice_var_high = tk.StringVar()
            self.choice_var_high.set(list[0])

            self.specLow_label = tk.Label(self, text="Spec at ")
            self.specLow_label.grid(row=3,column=3, padx = 5, pady = 10, sticky = "E")
            self.specLow_textbox = tk.Text(self, height=1, width = 10)
            self.specLow_textbox.grid(row=3,column=4,padx=0,pady=10, sticky = "W")
            self.specLow_menu = tk.OptionMenu(self, self.choice_var_low, *list)
            self.specLow_menu.grid(row=3,column=5, padx=0, pady=0)
            self.specLow_menu.configure(width=7)

            self.specMed_label = tk.Label(self, text="Spec at ")
            self.specMed_label.grid(row=4,column=3, padx = 5, pady = 10, sticky = "E")
            self.specMed_textbox = tk.Text(self, height=1, width = 10)
            self.specMed_textbox.grid(row=4,column=4,padx=0,pady=10, sticky = "W")
            self.specMed_menu = tk.OptionMenu(self, self.choice_var_med, *list)
            self.specMed_menu.grid(row=4,column=5, padx=0, pady=0)
            self.specMed_menu.configure(width=7)

            self.specHigh_label = tk.Label(self, text="Spec at ")
            self.specHigh_label.grid(row=5,column=3, padx = 5, pady = 10, sticky = "E")
            self.specHigh_textbox = tk.Text(self, height=1, width = 10)
            self.specHigh_textbox.grid(row=5,column=4,padx=0,pady=10, sticky = "W")
            self.specHigh_menu = tk.OptionMenu(self, self.choice_var_high, *list)
            self.specHigh_menu.grid(row=5,column=5, padx=0, pady=0)
            self.specHigh_menu.configure(width=7)


            self.planLow_label = tk.Label(self, text="+/- um")
            self.planLow_label.grid(row=3,column=7, padx = 0, pady = 10, sticky = "E")
            self.planLow_textbox = tk.Text(self, height=1, width = 5)
            self.planLow_textbox.grid(row=3,column=6,padx=10,pady=10, sticky = "W")

            self.planMed_label = tk.Label(self, text="+/- um")
            self.planMed_label.grid(row=4,column=7, padx = 0, pady = 10, sticky = "E")
            self.planMed_textbox = tk.Text(self, height=1, width = 5)
            self.planMed_textbox.grid(row=4,column=6,padx=10,pady=10, sticky = "W")

            self.planHigh_label = tk.Label(self, text="+/- um")
            self.planHigh_label.grid(row=5,column=7, padx = 0, pady = 10, sticky = "E")
            self.planHigh_textbox = tk.Text(self, height=1, width = 5)
            self.planHigh_textbox.grid(row=5,column=6,padx=10,pady=10, sticky = "W")

            ### HTU ONLY
            self.choice_var_othr1 = tk.StringVar()
            self.choice_var_othr1.set(list[0])

            self.choice_var_othr2 = tk.StringVar()
            self.choice_var_othr2.set(list[0])

            self.choice_var_othr3 = tk.StringVar()
            self.choice_var_othr3.set(list[0])

            self.specOthr1_label = tk.Label(self, text="Spec at ")
            self.specOthr1_label.grid(row=6,column=3, padx = 5, pady = 10, sticky = "E")
            self.specOthr1_textbox = tk.Text(self, height=1, width = 10)
            self.specOthr1_textbox.grid(row=6,column=4,padx=0,pady=10, sticky = "W")
            self.specOthr1_menu = tk.OptionMenu(self, self.choice_var_othr1, *list)
            self.specOthr1_menu.grid(row=6,column=5, padx=0, pady=0)
            self.specOthr1_menu.configure(width=7)
            self.specOthr1_label.grid_remove()
            self.specOthr1_textbox.grid_remove()
            self.specOthr1_menu.grid_remove()

            self.specOthr2_label = tk.Label(self, text="Spec at ")
            self.specOthr2_label.grid(row=7,column=3, padx = 5, pady = 10, sticky = "E")
            self.specOthr2_textbox = tk.Text(self, height=1, width = 10)
            self.specOthr2_textbox.grid(row=7,column=4,padx=0,pady=10, sticky = "W")
            self.specOthr2_menu = tk.OptionMenu(self, self.choice_var_othr2, *list)
            self.specOthr2_menu.grid(row=7,column=5, padx=0, pady=0)
            self.specOthr2_menu.configure(width=7)
            self.specOthr2_label.grid_remove()
            self.specOthr2_textbox.grid_remove()
            self.specOthr2_menu.grid_remove()

            self.specOthr3_label = tk.Label(self, text="Spec at ")
            self.specOthr3_label.grid(row=8,column=3, padx = 5, pady = 10, sticky = "E")
            self.specOthr3_textbox = tk.Text(self, height=1, width = 10)
            self.specOthr3_textbox.grid(row=8,column=4,padx=0,pady=10, sticky = "W")
            self.specOthr3_menu = tk.OptionMenu(self, self.choice_var_othr3, *list)
            self.specOthr3_menu.grid(row=8,column=5, padx=0, pady=0)
            self.specOthr3_menu.configure(width=7)
            self.specOthr3_label.grid_remove()
            self.specOthr3_textbox.grid_remove()
            self.specOthr3_menu.grid_remove()

            self.planOthr1_label = tk.Label(self, text="+/- um")
            self.planOthr1_label.grid(row=6,column=7, padx = 0, pady = 10, sticky = "E")
            self.planOthr1_textbox = tk.Text(self, height=1, width = 5)
            self.planOthr1_textbox.grid(row=6,column=6,padx=10,pady=10, sticky = "W")
            self.planOthr1_label.grid_remove()
            self.planOthr1_textbox.grid_remove()

            self.planOthr2_label = tk.Label(self, text="+/- um")
            self.planOthr2_label.grid(row=7,column=7, padx = 0, pady = 10, sticky = "E")
            self.planOthr2_textbox = tk.Text(self, height=1, width = 5)
            self.planOthr2_textbox.grid(row=7,column=6,padx=10,pady=10, sticky = "W")
            self.planOthr2_label.grid_remove()
            self.planOthr2_textbox.grid_remove()

            self.planOthr3_label = tk.Label(self, text="+/- um")
            self.planOthr3_label.grid(row=8,column=7, padx = 0, pady = 10, sticky = "E")
            self.planOthr3_textbox = tk.Text(self, height=1, width = 5)
            self.planOthr3_textbox.grid(row=8,column=6,padx=10,pady=10, sticky = "W")
            self.planOthr3_label.grid_remove()
            self.planOthr3_textbox.grid_remove()

            ####

            HTU_button = tk.Button(self, text="Add More Temperature Steps for HTU", command=self.show_HTU)
            HTU_button.grid(row=9,column=1, columnspan=7, padx=5, pady=10)
            HTU_button.configure(width=30)

            return_button = tk.Button(self, text="Go Back", command=lambda: controller.show_frame("StartPage"))
            return_button.grid(row=10, column=2,columnspan=1, padx=10, pady=10, sticky = "E")
            return_button.configure(width=12)

            self.continue_button = tk.Button(self, text="Continue", command=lambda: self.specs(list, sixpoints))
            self.continue_button.grid(row=10,column=3, columnspan=2, padx=10, pady=10)
            self.continue_button.configure(width=12)

            self.saveLF_button = tk.Button(self, text="Save", command=lambda: self.save_LF(self.type, self.yes_pt100,
                                                                                           self.last4, sixpoints))
            self.saveLF_button.grid(row=10,column=4, columnspan=2, padx=10, pady=10)
            self.saveLF_button.grid_remove()
            self.saveLF_button.configure(width=12)

            # if "Add More Temperature Steps for HTU" button is pressed then run this function
        def show_HTU(self):
            global sixpoints
            sixpoints = True

            self.tempOthr1_label.grid()
            self.tempOthr1_textbox.grid()
            self.tempOthr2_label.grid()
            self.tempOthr2_textbox.grid()
            self.tempOthr3_label.grid()
            self.tempOthr3_textbox.grid()

            self.specOthr1_label.grid()
            self.specOthr1_textbox.grid()
            self.specOthr1_menu.grid()

            self.specOthr2_label.grid()
            self.specOthr2_textbox.grid()
            self.specOthr2_menu.grid()

            self.specOthr3_label.grid()
            self.specOthr3_textbox.grid()
            self.specOthr3_menu.grid()

            self.planOthr1_label.grid()
            self.planOthr1_textbox.grid()

            self.planOthr2_label.grid()
            self.planOthr2_textbox.grid()

            self.planOthr3_label.grid()
            self.planOthr3_textbox.grid()

            return sixpoints

        def specs(self, list, sixpoints):
            new_LF = str(self.newLF_textbox.get('1.0', tk.END)).strip()
            low_temp = str(self.tempLow_textbox.get('1.0', tk.END)).strip()  # low temp test
            med_temp = str(self.tempMed_textbox.get('1.0', tk.END)).strip()  # medium temp test
            high_temp = str(self.tempHigh_textbox.get('1.0', tk.END)).strip()  # high temp test
            othr1_temp = str(self.tempOthr1_textbox.get('1.0', tk.END)).strip() #other temp 1 for HTU
            othr2_temp = str(self.tempOthr2_textbox.get('1.0', tk.END)).strip()#other temp 2 for HTU
            othr3_temp = str(self.tempOthr3_textbox.get('1.0', tk.END)).strip()#other temp 3 for HTU

            if sixpoints == True:
                if othr1_temp == "" or othr2_temp == "" or othr3_temp == "": # check to make sure all 'other' temperatures are given if the HTU button is pressed
                    messagebox.showinfo("Missing Values", "Please fill out the remaining temperatures or restart the program")

            self.chuck, self.type, self.yes_pt100, self.last4 = self.get_type(new_LF) # call the get_type function, which reads the seriennummern xls to find the type of chuck/controller
            self.label.configure( text = "Chuck Model: " + self.chuck )
            spec_index = self.type + "_spec"
            unit_index = self.type + "_unit"

            specs = pandas.read_excel('X://ERSTools/EndtestData/specs.xlsx', sheet_name='Sheet1', index_col=0)


            if self.check_var.get() == 0: # if "no chuck" is not checked
                # get specs for temperature uniformity
                low_spec = specs.loc[spec_index][int(low_temp)]
                med_spec = specs.loc[spec_index][int(med_temp)]
                high_spec = specs.loc[spec_index][int(high_temp)]

                self.specLow_textbox.delete('1.0', tk.END)
                self.specMed_textbox.delete('1.0', tk.END)
                self.specHigh_textbox.delete('1.0', tk.END)
                self.specLow_textbox.insert('1.0', low_spec)
                self.specMed_textbox.insert('1.0', med_spec)
                self.specHigh_textbox.insert('1.0', high_spec)

                # assign default planarity spec of '10' um. This can be changed by user input
                self.planLow_textbox.delete('1.0', tk.END)
                self.planMed_textbox.delete('1.0', tk.END)
                self.planHigh_textbox.delete('1.0', tk.END)
                self.planLow_textbox.insert('1.0', '10')
                self.planMed_textbox.insert('1.0', '10')
                self.planHigh_textbox.insert('1.0', '10')

                self.specLow_label.configure(text="Spec at " + low_temp + "C")
                self.specMed_label.configure(text="Spec at " + med_temp + "C")
                self.specHigh_label.configure(text="Spec at " + high_temp + "C")

                self.choice_var_low = tk.StringVar()
                self.choice_var_med = tk.StringVar()
                self.choice_var_high = tk.StringVar()

                if sixpoints == True:
                    othr1_spec = specs.loc[spec_index][int(othr1_temp)]
                    othr2_spec = specs.loc[spec_index][int(othr2_temp)]
                    othr3_spec = specs.loc[spec_index][int(othr3_temp)]

                    self.specOthr1_textbox.delete('1.0', tk.END)
                    self.specOthr2_textbox.delete('1.0', tk.END)
                    self.specOthr3_textbox.delete('1.0', tk.END)
                    self.specOthr1_textbox.insert('1.0', othr1_spec)
                    self.specOthr2_textbox.insert('1.0', othr2_spec)
                    self.specOthr3_textbox.insert('1.0', othr3_spec)

                    self.planOthr1_textbox.delete('1.0', tk.END)
                    self.planOthr2_textbox.delete('1.0', tk.END)
                    self.planOthr3_textbox.delete('1.0', tk.END)
                    self.planOthr1_textbox.insert('1.0', '10')
                    self.planOthr2_textbox.insert('1.0', '10')
                    self.planOthr3_textbox.insert('1.0', '10')

                    self.specOthr1_label.configure(text="Spec at " + othr1_temp + "C")
                    self.specOthr2_label.configure(text="Spec at " + othr2_temp + "C")
                    self.specOthr3_label.configure(text="Spec at " + othr3_temp + "C")

                    self.choice_var_othr1 = tk.StringVar()
                    self.choice_var_othr2 = tk.StringVar()
                    self.choice_var_othr3 = tk.StringVar()

                # if spec is in "deg C", then set the menu to say "deg C" after updating the spec value in the text box
                if specs.loc[unit_index][int(low_temp)] == list[0]:
                    self.choice_var_low.set(list[0])
                    self.specLow_menu = tk.OptionMenu(self, self.choice_var_low, *list)
                    self.specLow_menu.grid(row=3, column=5, padx=0, pady=0)
                    self.specLow_menu.configure(width=7)
                else:
                    self.choice_var_low.set(list[1])
                    self.specLow_menu = tk.OptionMenu(self, self.choice_var_low, *list)
                    self.specLow_menu.grid(row=3, column=5, padx=0, pady=0)
                    self.specLow_menu.configure(width=7)

                if specs.loc[unit_index][int(med_temp)] == list[0]:
                    self.choice_var_med.set(list[0])
                    self.specMed_menu = tk.OptionMenu(self, self.choice_var_med, *list)
                    self.specMed_menu.grid(row=4, column=5, padx=0, pady=0)
                    self.specMed_menu.configure(width=7)
                else:
                    self.choice_var_med.set(list[1])
                    self.specMed_menu = tk.OptionMenu(self, self.choice_var_med, *list)
                    self.specMed_menu.grid(row=4, column=5, padx=0, pady=0)
                    self.specMed_menu.configure(width=7)

                if specs.loc[unit_index][int(high_temp)] == list[0]:
                    self.choice_var_high.set(list[0])
                    self.specHigh_menu = tk.OptionMenu(self, self.choice_var_high, *list)
                    self.specHigh_menu.grid(row=5, column=5, padx=0, pady=0)
                    self.specHigh_menu.configure(width=7)
                else:
                    self.choice_var_high.set(list[1])
                    self.specHigh_menu = tk.OptionMenu(self, self.choice_var_high, *list)
                    self.specHigh_menu.grid(row=5, column=5, padx=0, pady=0)
                    self.specHigh_menu.configure(width=7)

                # if its HTU, do these also
                if sixpoints == True:
                    if specs.loc[unit_index][int(othr1_temp)] == list[0]:
                        self.choice_var_othr1.set(list[0])
                        self.specOthr1_menu = tk.OptionMenu(self, self.choice_var_othr1, *list)
                        self.specOthr1_menu.grid(row=6, column=5, padx=0, pady=0)
                        self.specOthr1_menu.configure(width=7)
                    else:
                        self.choice_var_othr1.set(list[1])
                        self.specOthr1_menu = tk.OptionMenu(self, self.choice_var_othr1, *list)
                        self.specOthr1_menu.grid(row=6, column=5, padx=0, pady=0)
                        self.specOthr1_menu.configure(width=7)

                    if specs.loc[unit_index][int(othr2_temp)] == list[0]:
                        self.choice_var_othr2.set(list[0])
                        self.specOthr2_menu = tk.OptionMenu(self, self.choice_var_othr2, *list)
                        self.specOthr2_menu.grid(row=7, column=5, padx=0, pady=0)
                        self.specOthr2_menu.configure(width=7)
                    else:
                        self.choice_var_othr2.set(list[1])
                        self.specOthr2_menu = tk.OptionMenu(self, self.choice_var_othr2, *list)
                        self.specOthr2_menu.grid(row=7, column=5, padx=0, pady=0)
                        self.specOthr2_menu.configure(width=7)

                    if specs.loc[unit_index][int(othr3_temp)] == list[0]:
                        self.choice_var_othr3.set(list[0])
                        self.specOthr3_menu = tk.OptionMenu(self, self.choice_var_othr3, *list)
                        self.specOthr3_menu.grid(row=8, column=5, padx=0, pady=0)
                        self.specOthr3_menu.configure(width=7)
                    else:
                        self.choice_var_othr3.set(list[1])
                        self.specOthr3_menu = tk.OptionMenu(self, self.choice_var_othr3, *list)
                        self.specOthr3_menu.grid(row=8, column=5, padx=0, pady=0)
                        self.specOthr3_menu.configure(width=7)

            self.saveLF_button.grid()

        def save_LF(self, type, yes_pt100, last4, sixpoints):
            # load active list
            LF_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1', header=None)
            new_LF = self.newLF_textbox.get('1.0', tk.END)  # new LF number
            low_temp = self.tempLow_textbox.get('1.0', tk.END)  # low temp test
            med_temp = self.tempMed_textbox.get('1.0', tk.END)  # medium temp test
            high_temp = self.tempHigh_textbox.get('1.0', tk.END)  # high temp test
            othr1_temp = self.tempOthr1_textbox.get('1.0', tk.END)
            othr2_temp = self.tempOthr2_textbox.get('1.0', tk.END)
            othr3_temp = self.tempOthr3_textbox.get('1.0', tk.END)
            new = LF_list.iloc[0, :]                                # make copy of first row of active list. We will put the New LF information in it, and then add it back to the original LF_list

            i = 0
            while i < len(new): # set all values to 0 except last (because for last, 1 = on active list)
                new.loc[i] = '0'
                i=i+1
            new.loc[19] = '1'  # have new LF appear in the active list
            new.loc[2] = '999'  # set all temperature values to '999'
            new.loc[3] = '999'
            new.loc[4] = '999'
            new.loc[23] = '999'
            new.loc[24] = '999'
            new.loc[25] = '999'



                ## note: much of the next few hundred lines of code are taken from a different program I wrote. Some of the code is not functional in this program.
            if self.check_var.get() == 0:
                unit_index = type + "_unit"   # for example, these variables are not used. BE CAREFUL if you change anything. Everything so far seems to function in testing.
                spec_index = type + "_spec"
                specs = pandas.read_excel('X://ERSTools/EndtestData/specs.xlsx', sheet_name='Sheet1', index_col=0)

                if self.choice_var_high.get() == 'deg C':
                    self.high_spec = str(self.specHigh_textbox.get('1.0', tk.END)).strip()
                else:
                    self.high_spec = str(0.01*float(str(self.tempHigh_textbox.get('1.0', tk.END)).strip()) * \
                                     float(str(self.specHigh_textbox.get('1.0', tk.END)).strip()))


            if new.loc[0] != 'L':   # if it is an internal ERS order, name the directory this
                newdir = "//fileserver/produktion/Endtest/10_Dokumente/" + str(last4) + '_ERS' + str(new_LF).strip()
            else: # if its not, then name the directory this
                newdir = "//fileserver/produktion/Endtest/10_Dokumente/" + str(last4) + '_' + str(new_LF).strip()


            ### Refer to X:\ERSTools\EndtestData\Documentation\active_list_legend.xlsx for a legend to see what the different positions of the 'new' array mean
            new.loc[21] = str(newdir).strip()
            new.loc[0] = str(new_LF).strip() #for example, [0] is the position for the LF number
            new.loc[14] = yes_pt100 # [14] is the position that tells wether or not PT100 testing is needed ( if its an SP controller or TS010 chiller, we need PT100 testing)
            new.loc[20] = last4
            if self.check_var.get() == 0:
                new.loc[1] = type
                new.loc[2] = str(low_temp).strip()  # low temp test
                new.loc[3] = str(med_temp).strip()  # medium temp test
                new.loc[4] = str(high_temp).strip()  # high temp test
                new.loc[5] = str(self.specLow_textbox.get('1.0', tk.END)).strip()
                new.loc[6] = str(self.specMed_textbox.get('1.0', tk.END)).strip()
                new.loc[7] = self.high_spec # high temp spec
                new.loc[16] = str(self.planLow_textbox.get('1.0', tk.END)).strip()
                new.loc[17] = str(self.planMed_textbox.get('1.0', tk.END)).strip()
                new.loc[18] = str(self.planHigh_textbox.get('1.0', tk.END)).strip()
                new.loc[22] = str(self.controller.get_page("CreateNew").choice_var_high.get()).strip() # units of high temp measurement
                if sixpoints == True:
                    new.loc[23] = str(othr1_temp).strip()
                    new.loc[24] = str(othr2_temp).strip()
                    new.loc[25] = str(othr3_temp).strip()
                    new.loc[26] = str(self.specOthr1_textbox.get('1.0', tk.END)).strip()
                    new.loc[27] = str(self.specOthr2_textbox.get('1.0', tk.END)).strip()
                    new.loc[28] = str(self.specOthr3_textbox.get('1.0', tk.END)).strip()
                    new.loc[35] = str(self.planOthr1_textbox.get('1.0', tk.END)).strip()
                    new.loc[36] = str(self.planOthr2_textbox.get('1.0', tk.END)).strip()
                    new.loc[37] = str(self.planOthr3_textbox.get('1.0', tk.END)).strip()


            i = 1
            while i < LF_list.shape[0]:
                if len(str(LF_list.iloc[i,0])) < 4:
                    LF_list.loc[i, 0] = "delete"
                i = i+1

            LF_list = LF_list[LF_list[0].astype(str).str.contains('delete') == False]

            save = True
            if self.check_var.get() == 0:
                if len(new.loc[0]) < 4 or len(new.loc[2]) == 1 or len(new.loc[3]) == 1:   # if 2 temperature values are empty and the no chuck box is not checked
                    save = False

            exists = False
            if len(LF_list[LF_list[0].astype(str).str.contains(new.loc[0]) == True]):
                messagebox.showinfo("Error", "Entry for Lauf Number Already Exists")
                exists = True


##############
            notiinlist = False
            lauf_num = new.loc[0]
            snxls_path = '//fileserver/produktion/Endtest/30_Seriennummern/Seriennummern.xlsm'  ## path of seriennummern spreadsheet
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
                    print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                    notiinlist = True
                else:
                    D = D.loc[str(lauf_num)]
            if alpha == False:
                if len(lauf_num) == 6:
                    if str(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        notiinlist = True
                    else:
                        D = D.loc[str(lauf_num)]
                elif len(lauf_num) != 6:
                    if int(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        notiinlist = True
                    else:
                        D = D.loc[int(lauf_num)]


            loop = False
            oospec = False # out of spec
            temp = ['', '', '']
            i = 0
            while i < 3:                            # THIS IS IMPORTANT
                if self.check_var.get() == 0 and new.loc[i+16] != "": # this piece of code sets the allowable values for planarity tolerance
                    loop = True                                       # right now 8, 10, 12, and 15 are accepted values
                    if str(new.loc[i+16]).strip() == '8' or str(new.loc[i+16]).strip() == '10' \
                            or str(new.loc[i+16]).strip() == '12' or str(new.loc[i+16]).strip() == '15':
                        temp[i] = 1
                i = i + 1
            if loop == True:
                if temp[0] == 1 and temp[1] == 1 and temp[2] == 1:
                    save = True
                else:
                    save = False
                    oospec = True

            if exists == True:
                save = False

            if notiinlist == True:
                save = False
                messagebox.showinfo("Error", "LF ist nicht in die Seriennumernliste")

            saved = False
            print(save)
            if save == True:
                LF_list = LF_list.append(new)
                LF_list = LF_list.reset_index(drop=True)
                writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                LF_list.to_excel(writer, 'Sheet1', index=False, header=None)
                writer.save()
                self.controller.get_page("StartPage").update_LF_menu()
                self.make_files(new, last4, sixpoints)
                saved = True
            elif oospec == True:
                messagebox.showerror("Error",
                                     "Planarity specification not acceptable (accepted values = 8, 10, 12, and 15")
            else:
                messagebox.showinfo("Error", "Missing Values")

            if saved == True:
                messagebox.showinfo("New Message", "Save Successful")

        def make_files(self, new, last4, sixpoints):
            if new.loc[0] != 'L':
                newdir = "//fileserver/produktion/Endtest/10_Dokumente/" + str(last4) + '_ERS' + new.loc[0]
            else:
                newdir = "//fileserver/produktion/Endtest/10_Dokumente/" + str(last4) + '_' + new.loc[0]

            newdir2 = newdir + "/data/"

            try:                        # it is possible that this directory already exists because it was created by the Endtest Doc Generator Program
                os.mkdir(newdir)        # because of that, we use a try statement and pass any exceptions. It it already exists, the try statement will fail
            except:                      # but we do not care because we just the need the directory to exist
                pass
            os.mkdir(newdir2)
            writer = ExcelWriter(newdir2 + "ref.xlsx")
            new.to_excel(writer, 'Sheet1', index=False, header=None) # create ref spreadsheet
            writer.save()

            # create all excel files that will store our test data
            wb = openpyxl.Workbook()
            wb.save(newdir2 + "PT100.xlsx")
            wb.save(newdir2+"TempLow.xlsx")
            wb.save(newdir2+"TempMed.xlsx")
            wb.save(newdir2+"TempHigh.xlsx")
            wb.save(newdir2+"PlanLow.xlsx")
            wb.save(newdir2+"PlanMed.xlsx")
            wb.save(newdir2+"PlanHigh.xlsx")
            if sixpoints == True:             # if its an HTU chuck
                wb.save(newdir2 + "TempOthr1.xlsx")
                wb.save(newdir2 + "TempOthr2.xlsx")
                wb.save(newdir2 + "TempOthr3.xlsx")
                wb.save(newdir2 + "PlanOthr1.xlsx")
                wb.save(newdir2 + "PlanOthr2.xlsx")
                wb.save(newdir2 + "PlanOthr3.xlsx")


            # find out the type of chuck and type of controller
        def get_type(self, new_LF):
            dir = '//fileserver/produktion/Endtest/30_Seriennummern/Seriennummern.xlsm'
            #dir = '//fileserver/alle/Austin/old/SeriennummernValidation/Seriennummern.xlsm'
            D = pandas.read_excel(dir, sheet_name='Serien_Nummern aufsteigend',
                                  index_col=0)  # imports Seriennummern spreadsheet
            alpha = False
            i = 0
            lauf_num = new_LF
            while i < len(lauf_num):
                if lauf_num[i].isalpha() == True:
                    alpha = True
                i = i + 1

            if alpha == True:
                if lauf_num not in D.index.values:
                    messagebox.showerror("Error", "Lauf Number not in Seriennummern Spreadsheet")
                    x = 1/0
                else:
                    D = D.loc[str(lauf_num)]
            if alpha == False:
                if len(lauf_num) == 6:
                    if str(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        messagebox.showerror("Error", "Lauf Number not in Seriennummern Spreadsheet")
                        x = 1 / 0
                    else:
                        D = D.loc[str(lauf_num)]
                elif len(lauf_num) != 6:
                    if int(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        messagebox.showerror("Error", "Lauf Number not in Seriennummern Spreadsheet")
                        x = 1 / 0
                    else:
                        D = D.loc[int(lauf_num)]

            chuck = D['Chuck'].upper()
            type = "STD"   # we say the chuck is a standard chuck until something in the name tells us different
            if 'HTU' in chuck:
                type = "HTU"
            if 'AM' in chuck:
                type = "AM"
            if 'LN' in chuck:
                type = "LN"

            yes_pt100 = '0' # only SP controllers and TS010 chillers need pt100 calibration
            if D['Steuergerät'] == D['Steuergerät']:  # Is controller field empty (NaN)? NaN is not equal to NaN
                yes_pt100 = '1'
                if 'RSI' in D['Steuergerät'].upper():   # an RSI box is not a controller, even though sometimes it is put in the controller column
                    yes_pt100 = '0'

            if D['Chiller'] == D['Chiller']:
                if 'TS010' in D['Chiller'].upper(): # or is there a TS010 chiller
                    yes_pt100 = '1'

            SN = str(D['SN Kompl.'])
            last4 = SN[len(SN) - 4:len(SN)]

            return chuck, type, yes_pt100, last4

    class EditList(tk.Frame):
        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            label = tk.Label(self, text="Welcome", font=controller.title_font)
            label.grid(row=1, column=2, padx=0, pady=10, sticky="W")

            # find active LFs again (note: this code is a direct copy from elsewhere in this program. You will see repeats of code many times)
            master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1', dtype = str)
            inactive = master_list[master_list.iloc[:, 19] == '0']
            inactive = inactive.iloc[:,0]
            inactive_list = inactive.values.tolist()
            active = master_list[master_list.iloc[:, 19] == '1']
            active = active.iloc[:,0]
            active_list = active.values.tolist()

            active_label = tk.Label(self, text="Active LFs" , font=controller.label_font)
            active_label.grid(row=2, column=1)
            self.active_var = tk.StringVar()
            self.active_var.set("")
            self.active_list = tk.Listbox(self)
            self.active_list.insert(tk.END, *active_list)
            self.active_list.grid(row=3, column=1, padx=70, pady=10, sticky="E")

            inactive_label = tk.Label(self, text="Inactive LFs" , font=controller.label_font)
            inactive_label.grid(row=2, column=3)
            self.inactive_var = tk.StringVar()
            self.inactive_var.set("")
            self.inactive_list = tk.Listbox(self)
            self.inactive_list.insert(tk.END, *inactive_list)
            self.inactive_list.grid(row=3, column=3, padx=70, pady=10, sticky="E")

            self.toInactive_button = tk.Button(self, text = "Move to Inactive ====>", command=lambda: self.toInactive())
            self.toInactive_button.grid(row = 4, column = 2, pady = 5)

            self.toActive_button = tk.Button(self, text = "<==== Move to Active", command=lambda: self.toActive())
            self.toActive_button.grid(row = 5, column = 2, pady = 5)

            self.toActive_button = tk.Button(self, text = "Edit/Delete", command=lambda: self.toEdit())
            self.toActive_button.grid(row = 6, column = 2, pady = 5)


            return_button = tk.Button(self, text="Go Back",
                                      command=lambda: controller.show_frame("StartPage"))
            return_button.grid(row=6, column=1, padx=70, pady=10, sticky="E")
            return_button.configure(width=22)

        def toInactive(self):

            master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                            dtype=str)
            inactive = master_list[master_list.iloc[:, 19] == '0']
            inactive = inactive.iloc[:,0]
            inactive_list = inactive.values.tolist()
            active = master_list[master_list.iloc[:, 19] == '1']
            active = active.iloc[:,0]
            active_list = active.values.tolist()

            string = ""
            string = active_list[self.active_list.curselection()[0]]

            if string != "":
                if len(active_list) > 1:

                    i = 0
                    while i < len(master_list):
                        if master_list.iloc[i,0] == string:
                            master_list.iloc[i, 19] = '0'
                            break
                        i = i + 1

                    inactive = master_list[master_list.iloc[:, 19] == '0']
                    inactive = inactive.iloc[:,0]
                    inactive_list = inactive.values.tolist()
                    active = master_list[master_list.iloc[:, 19] == '1']
                    active = active.iloc[:,0]
                    active_list = active.values.tolist()

                    self.active_list.delete(0,tk.END)                   # update list boxes
                    self.active_list.insert(tk.END, *active_list)
                    self.inactive_list.delete(0, tk.END)
                    self.inactive_list.insert(tk.END, *inactive_list)

                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.get_page("StartPage").update_LF_menu()
                else:
                    messagebox.showinfo("Error", "You must leave atleast one LF in the active list at all times.")

        def toActive(self):
            master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                            dtype=str)

            inactive = master_list[master_list.iloc[:, 19] == '0']
            inactive = inactive.iloc[:,0]
            inactive_list = inactive.values.tolist()
            active = master_list[master_list.iloc[:, 19] == '1']
            active = active.iloc[:,0]
            active_list = active.values.tolist()

            string = inactive_list[self.inactive_list.curselection()[0]]
            i = 0
            while i < len(master_list):
                if master_list.iloc[i,0] == string:
                    master_list.iloc[i, 19] = '1'
                    break
                i = i + 1

            inactive = master_list[master_list.iloc[:, 19] == '0']
            inactive = inactive.iloc[:,0]
            inactive_list = inactive.values.tolist()
            active = master_list[master_list.iloc[:, 19] == '1']
            active = active.iloc[:,0]
            active_list = active.values.tolist()

            self.active_list.delete(0,tk.END)                   # update list boxes
            self.active_list.insert(tk.END, *active_list)
            self.inactive_list.delete(0, tk.END)
            self.inactive_list.insert(tk.END, *inactive_list)

            writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
            master_list.to_excel(writer, 'Sheet1', index=None)
            writer.save()
            self.controller.get_page("StartPage").update_LF_menu()

        def toEdit(self):
            master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                            dtype=str)
            inactive = master_list[master_list.iloc[:, 19] == '0']
            inactive = inactive.iloc[:,0]
            inactive_list = inactive.values.tolist()
            active = master_list[master_list.iloc[:, 19] == '1']
            active = active.iloc[:,0]
            active_list = active.values.tolist()

            selected_LF = ""
            try:
                selected_LF = active_list[self.active_list.curselection()[0]]
            except:
                pass
            try:
                selected_LF = inactive_list[self.inactive_list.curselection()[0]]
            except:
                pass
            self.controller.get_page("ModLF").ModLFLabel.configure(text="LF: " + selected_LF)

            selected_DF = master_list[master_list.loc[:,0] == selected_LF]

            self.controller.get_page("ModLF").tempLow_label.grid_remove()
            self.controller.get_page("ModLF").tempLow_textbox.grid_remove()
            self.controller.get_page("ModLF").tempMed_label.grid_remove()
            self.controller.get_page("ModLF").tempMed_textbox.grid_remove()
            self.controller.get_page("ModLF").tempHigh_label.grid_remove()
            self.controller.get_page("ModLF").tempHigh_textbox.grid_remove()


            ## BELOW IS COMMENTED IN ORDER TO GET RID OF ABILITY TO EDIT TEMPERATURES OF LF's AFTER CREATED

            # if selected_DF.iloc[0, 2] != '999' or selected_DF.iloc[0, 3] != '999' or selected_DF.iloc[0, 4] != '999':
            #     self.controller.get_page("ModLF").tempLow_label.grid()
            #     self.controller.get_page("ModLF").tempLow_textbox.grid()
            #     self.controller.get_page("ModLF").tempMed_label.grid()
            #     self.controller.get_page("ModLF").tempMed_textbox.grid()
            #     self.controller.get_page("ModLF").tempHigh_label.grid()
            #     self.controller.get_page("ModLF").tempHigh_textbox.grid()
            #
            #     self.controller.get_page("ModLF").tempLow_textbox.delete('1.0', tk.END)
            #     self.controller.get_page("ModLF").tempMed_textbox.delete('1.0', tk.END)
            #     self.controller.get_page("ModLF").tempHigh_textbox.delete('1.0', tk.END)
            #     if selected_DF.iloc[0, 2] != '999':
            #         self.controller.get_page("ModLF").tempLow_textbox.insert('1.0', selected_DF.iloc[0, 2])
            #     if selected_DF.iloc[0, 3] != '999':
            #         self.controller.get_page("ModLF").tempMed_textbox.insert('1.0', selected_DF.iloc[0, 3])
            #     if selected_DF.iloc[0, 4] != '999':
            #         self.controller.get_page("ModLF").tempHigh_textbox.insert('1.0', selected_DF.iloc[0, 4])


            self.controller.get_page("ModLF").tempOthr1_label.grid_remove()
            self.controller.get_page("ModLF").tempOthr1_textbox.grid_remove()
            self.controller.get_page("ModLF").tempOthr2_label.grid_remove()
            self.controller.get_page("ModLF").tempOthr2_textbox.grid_remove()
            self.controller.get_page("ModLF").tempOthr3_label.grid_remove()
            self.controller.get_page("ModLF").tempOthr3_textbox.grid_remove()

            ## BELOW IS COMMENTED IN ORDER TO GET RID OF ABILITY TO EDIT TEMPERATURES OF LF's AFTER CREATED

            # if selected_DF.iloc[0,23] != '999' or selected_DF.iloc[0,24] != '999' or selected_DF.iloc[0,25] != '999':
            #     self.controller.get_page("ModLF").tempOthr1_label.grid()
            #     self.controller.get_page("ModLF").tempOthr1_textbox.grid()
            #     self.controller.get_page("ModLF").tempOthr2_label.grid()
            #     self.controller.get_page("ModLF").tempOthr2_textbox.grid()
            #     self.controller.get_page("ModLF").tempOthr3_label.grid()
            #     self.controller.get_page("ModLF").tempOthr3_textbox.grid()
            #
            #     self.controller.get_page("ModLF").tempOthr1_textbox.delete('1.0', tk.END)
            #     self.controller.get_page("ModLF").tempOthr2_textbox.delete('1.0', tk.END)
            #     self.controller.get_page("ModLF").tempOthr3_textbox.delete('1.0', tk.END)
            #     if selected_DF.iloc[0,23] != '999':
            #         self.controller.get_page("ModLF").tempOthr1_textbox.insert('1.0', selected_DF.iloc[0, 23])
            #     if selected_DF.iloc[0, 24] != '999':
            #         self.controller.get_page("ModLF").tempOthr2_textbox.insert('1.0', selected_DF.iloc[0, 24])
            #     if selected_DF.iloc[0,25] != '999':
            #         self.controller.get_page("ModLF").tempOthr3_textbox.insert('1.0', selected_DF.iloc[0, 25])


            self.controller.show_frame("ModLF")

    class ModLF(tk.Frame):
        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.ModLFLabel = tk.Label(self, text="", font=controller.title_font)
            self.ModLFLabel.grid(row=1, column=2, columnspan=3, padx=0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=1, column=1, padx=80, pady=10)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=1, column=5, padx=80, pady=10)

            self.tempLow_label = tk.Label(self, text="Test Temperature Low")
            self.tempLow_label.grid(row=3, column=2, columnspan=2, padx=5, pady=10)
            self.tempLow_textbox = tk.Text(self, height=1, width=10)
            self.tempLow_textbox.grid(row=3, column=4, padx=0, pady=10)
            self.tempLow_label.grid_remove()
            self.tempLow_textbox.grid_remove()

            self.tempMed_label = tk.Label(self, text="Test Temp Intermediate (25C)")
            self.tempMed_label.grid(row=4, column=2, columnspan=2, padx=5, pady=10)
            self.tempMed_textbox = tk.Text(self, height=1, width=10)
            self.tempMed_textbox.grid(row=4, column=4, padx=0, pady=10)
            self.tempMed_label.grid_remove()
            self.tempMed_textbox.grid_remove()

            self.tempHigh_label = tk.Label(self, text="Test Temperature High")
            self.tempHigh_label.grid(row=5, column=2, columnspan=2, padx=5, pady=10)
            self.tempHigh_textbox = tk.Text(self, height=1, width=10)
            self.tempHigh_textbox.grid(row=5, column=4, padx=0, pady=10)
            self.tempHigh_label.grid_remove()
            self.tempHigh_textbox.grid_remove()

            self.tempOthr1_label = tk.Label(self, text="Test Temperature Other1")
            self.tempOthr1_label.grid(row=6, column=2, columnspan=2, padx=5, pady=10)
            self.tempOthr1_textbox = tk.Text(self, height=1, width=10)
            self.tempOthr1_textbox.grid(row=6, column=4, padx=0, pady=10)
            self.tempOthr1_label.grid_remove()
            self.tempOthr1_textbox.grid_remove()

            self.tempOthr2_label = tk.Label(self, text="Test Temperature Other2")
            self.tempOthr2_label.grid(row=7, column=2, columnspan=2, padx=5, pady=10)
            self.tempOthr2_textbox = tk.Text(self, height=1, width=10)
            self.tempOthr2_textbox.grid(row=7, column=4, padx=0, pady=10)
            self.tempOthr2_label.grid_remove()
            self.tempOthr2_textbox.grid_remove()

            self.tempOthr3_label = tk.Label(self, text="Test Temperature Other3")
            self.tempOthr3_label.grid(row=8, column=2, columnspan=2, padx=5, pady=10)
            self.tempOthr3_textbox = tk.Text(self, height=1, width=10)
            self.tempOthr3_textbox.grid(row=8, column=4, padx=0, pady=10)
            self.tempOthr3_label.grid_remove()
            self.tempOthr3_textbox.grid_remove()


            return_button = tk.Button(self, text="Go Back", command=lambda: controller.show_frame("EditList"))
            return_button.grid(row=9, column=2, columnspan=1, padx=5, pady=10)
            return_button.configure(width=12)

            self.save_button = tk.Button(self, text="Save", command=lambda: self.mod_Save())
            self.save_button.grid(row=9, column=3, columnspan=1, padx=5, pady=10)
            self.save_button.configure(width=12)
            self.save_button.grid_remove()

            self.delete_button = tk.Button(self, text="Delete", command=lambda: self.mod_Delete())
            self.delete_button.grid(row=9, column=4, columnspan=1, padx=5, pady=10)
            self.delete_button.configure(width=12)

        def mod_Save(self):
            active_LF = self.ModLFLabel['text']
            active_LF = active_LF.replace("LF: ", "")

            master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                            dtype=str)
            active_index = str(master_list[master_list.loc[:, 0] == active_LF].index.tolist())
            active_index = active_index[1:-1] # removes first and last characters from string.. we could have also done .tolist()[0]

            if str(self.tempLow_textbox.get('1.0', tk.END)).strip() != "":
                master_list.iloc[int(active_index), 2] = str(self.tempLow_textbox.get('1.0', tk.END)).strip()
            if str(self.tempMed_textbox.get('1.0', tk.END)).strip() != "":
                master_list.iloc[int(active_index), 3] = str(self.tempMed_textbox.get('1.0', tk.END)).strip()
            if str(self.tempHigh_textbox.get('1.0', tk.END)).strip() != "":
                master_list.iloc[int(active_index), 4] = str(self.tempHigh_textbox.get('1.0', tk.END)).strip()

            if str(self.tempOthr1_textbox.get('1.0', tk.END)).strip() != "":
                master_list.iloc[int(active_index), 23] = str(self.tempOthr1_textbox.get('1.0', tk.END)).strip()
            if str(self.tempOthr2_textbox.get('1.0', tk.END)).strip() != "":
                master_list.iloc[int(active_index), 24] = str(self.tempOthr2_textbox.get('1.0', tk.END)).strip()
            if str(self.tempOthr3_textbox.get('1.0', tk.END)).strip() != "":
                master_list.iloc[int(active_index), 25] = str(self.tempOthr3_textbox.get('1.0', tk.END)).strip()


            writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
            master_list.to_excel(writer, 'Sheet1', index=None)
            writer.save()
            self.controller.get_page("StartPage").update_LF_menu()
            messagebox.showinfo("Message", "Save Sucessful")

        def mod_Delete(self):
            delete = False
            MsgBox = messagebox.askquestion('Delete Record', 'Are you sure you want to delete this record?',
                                               icon='warning')
            if MsgBox == 'yes':
                delete = True

            if delete == True:
                active_LF = self.ModLFLabel['text']
                active_LF = active_LF.replace("LF: ", "")
                master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                dtype=str)
                master_list = master_list[master_list.loc[:,0] != active_LF]

                writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                master_list.to_excel(writer, 'Sheet1', index=None)
                writer.save()
                self.controller.get_page("StartPage").update_LF_menu()

                inactive = master_list[master_list.iloc[:, 19] == '0']
                inactive = inactive.iloc[:, 0]
                inactive_list = inactive.values.tolist()
                active = master_list[master_list.iloc[:, 19] == '1']
                active = active.iloc[:, 0]
                active_list = active.values.tolist()

                self.controller.get_page("EditList").active_list.delete(0, tk.END)  # update list boxes
                self.controller.get_page("EditList").active_list.insert(tk.END, *active_list)
                self.controller.get_page("EditList").inactive_list.delete(0, tk.END)
                self.controller.get_page("EditList").inactive_list.insert(tk.END, *inactive_list)

                messagebox.showinfo("Message", "Record Deleted")
                self.controller.show_frame("EditList")

    class DocDev(tk.Frame):
        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="Instructions: Select or Enter LF nummer. Press Next. Fill out all fields and then save.")
            self.label.grid(row=1, column=1, columnspan=7, padx=5, pady=10)

            # This program looks for active LF's in the same spreadsheet as used by the data entry app
            master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1', dtype=str)
            inactive = master_list[master_list.iloc[:, 19] == '0']
            inactive = inactive.iloc[:, 0]
            inactive_list = inactive.values.tolist()
            active = master_list[master_list.iloc[:, 19] == '1']
            active = active.iloc[:, 0]
            active_list = ["Select LF"] + active.values.tolist()

            active_label = tk.Label(self, text="Active LFs")
            active_label.grid(row=2, column=2, pady=5, padx=5)
            self.active_var = tk.StringVar()
            self.active_var.set("Select LF")
            self.active_optmenu = tk.OptionMenu(self, self.active_var, *active_list)
            self.active_optmenu.grid(row=3, column=2, padx=5, pady=10, sticky="E")
            #
            self.enter_other_label = tk.Label(self, text="Other LF:")
            self.enter_other_label.grid(row=4, column=1, pady=5, padx=5)
            self.enter_other_entry = tk.Entry(self, width=10)
            self.enter_other_entry.grid(row=4, column=2, pady=5, padx=5)

            self.component_label = tk.Label(self, text="Select Component")
            self.component_label.grid(row=2, column=3)
            self.component_var = tk.StringVar()
            self.component_var.set("         ")
            self.component_optmenu = tk.OptionMenu(self, self.component_var, [])
            self.component_optmenu.grid(row=3, column=3, padx=5, pady=10, sticky="E")
            self.component_label.grid_remove()
            self.component_optmenu.grid_remove()

            error_list = ['BOMB', 'Drawing', 'Software', 'Bad Material', 'Other']
            self.error_label = tk.Label(self, text="Select Error Type")
            self.error_label.grid(row=2, column=4)
            self.error_var = tk.StringVar()
            self.error_var.set("         ")
            self.error_optmenu = tk.OptionMenu(self, self.error_var, *error_list)
            self.error_optmenu.grid(row=3, column=4, padx=5, pady=10, sticky="E")
            self.error_label.grid_remove()
            self.error_optmenu.grid_remove()

            self.description_label = tk.Label(self, text="Error Description")
            self.description_label.grid(row=2, column=5, padx=5, pady=10)
            self.description_box = tk.Text(self)
            self.description_box.configure(width=20, height=5)
            self.description_box.grid(row=3, column=5, padx=5, pady=10)
            self.description_label.grid_remove()
            self.description_box.grid_remove()

            self.next_button = tk.Button(self, text="Next", command=lambda: next())
            self.next_button.grid(row=10, column=1, padx=5, pady=10)
            self.next_button.configure(width=10)

            self.save_button = tk.Button(self, text="Save", command=lambda: save())
            self.save_button.grid(row=10, column=3, padx=5, pady=10)
            self.save_button.configure(width=10)
            self.save_button.grid_remove()

            self.reset_button = tk.Button(self, text="Reset", command=lambda: reset())
            self.reset_button.grid(row=10, column=4, padx=5, pady=10)
            self.reset_button.configure(width=10)
            self.reset_button.grid_remove()

            def next():
                global lauf_num
                if self.active_var.get() != "Select LF" and self.enter_other_entry.get().strip() == "":
                    lauf_num = self.active_var.get()
                elif len(enter_other_entry.get().strip()) > 3:
                    lauf_num = self.enter_other_entry.get().strip()
                elif self.active_var.get() == "Select LF" and len(self.enter_other_entry.get().strip()) < 4:
                    messagebox.showerror("Error", "Please select or enter a valid LF")

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
                        messagebox.showerror('Error', 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        x = 1 / 0
                    else:
                        D = D.loc[str(lauf_num)]
                if alpha == False:
                    if len(lauf_num) == 6:
                        if str(
                                lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                            messagebox.showerror('Error', 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                            x = 1 / 0
                        else:
                            D = D.loc[str(lauf_num)]
                    elif len(lauf_num) != 6:
                        if int(
                                lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                            messagebox.showerror('Error', 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
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

                # add component list to drop down menu
                # Reset var and delete all old options
                self.component_var.set('         ')
                self.component_optmenu['menu'].delete(0, 'end')
                # Insert list of new options (tk._setit hooks them up to var)
                component_list = tuple(component_list)  # convert component list to tuple
                for choice in component_list:
                    self.component_optmenu['menu'].add_command(label=choice, command=tk._setit(self.component_var, choice))
                # self.component_optmenu = tk.OptionMenu(self, self.component_var, *component_list)

                global sn, name
                sn = D['SN Kompl.']
                name = D.name

                self.component_label.grid()
                self.component_optmenu.grid()
                self.error_label.grid()
                self.error_optmenu.grid()
                self.save_button.grid()
                self.reset_button.grid()
                self.description_label.grid()
                self.description_box.grid()

            def save():
                save = True
                if self.active_var.get() != "Select LF" and self.enter_other_entry.get().strip() == "":
                    lauf_num2 = self.active_var.get()
                elif len(enter_other_entry.get().strip()) > 3:
                    lauf_num2 = self.enter_other_entry.get().strip()
                elif self.active_var.get() == "Select LF" and len(enter_other_entry.get().strip()) < 4:
                    messagebox.showerror("Error", "Please select or enter a valid LF")

                if lauf_num2 != lauf_num:
                    messagebox.showerror("Error", "LF Number was changed. Please 'Reset' the form and then press 'Next'")
                    save = False

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
                        messagebox.showerror('Error', 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        x = 1 / 0
                    else:
                        D = D.loc[str(lauf_num2)]
                if alpha == False:
                    if len(lauf_num2) == 6:
                        if str(
                                lauf_num2) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                            messagebox.showerror('Error', 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                            x = 1 / 0
                        else:
                            D = D.loc[str(lauf_num2)]
                    elif len(lauf_num2) != 6:
                        if int(
                                lauf_num2) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                            messagebox.showerror('Error', 'Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                            x = 1 / 0
                        else:
                            D = D.loc[int(lauf_num2)]

                # cancel script if lauf number is listed more than once in the excel file
                if len(
                        D) < 28:  # we say less than 28, because if only 1 exits then it has length 28, but if 2 exist then length 2
                    messagebox.showerror("Error", "Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt")
                    x = 1 / 0

                component = self.component_var.get()
                error = self.error_var.get()
                description = self.description_box.get('1.0', tk.END).strip()

                error_list = pandas.read_excel('X://ERSTools/EndtestData/error_list.xlsx', sheet_name='Sheet1',
                                               index_col=None, header=None)  # imports Seriennummern spreadsheet

                new = error_list.loc[0, :]

                new[0] = str(len(error_list) + 1)
                new[1] = lauf_num
                new[2] = getpass.getuser()
                new[3] = datetime.today().strftime('%Y-%m-%d')
                new[4] = component
                new[5] = D['SN Kompl.']
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
                        messagebox.showerror("Error",
                                             "Save Not Sucessful.")  # if you get this message, its probably because the error_list.xlsx is open

            def reset():
                self.save_button.grid_remove()
                self.reset_button.grid_remove()
                self.component_label.grid_remove()
                self.component_optmenu.grid_remove()
                self.error_label.grid_remove()
                self.error_optmenu.grid_remove()
                self.description_label.grid_remove()
                self.description_box.grid_remove()

    class Manager(tk.Frame):
        #inst_StartPage = self.controller.get_page(StartPage)
        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller

            space = 10      # change dimensions of page depending on which version of the program it is
            if hideForChuckDept == True or hideForFTDept == True:
                space = 100

            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan =5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none.grid(row=1, column=1, padx=space, pady=5)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=1, column=5, padx=space, pady=5)

            self.temp1 = tk.Button(self, text="Temp Uniformity -40C ", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("TempLow"),
                                                                                 self.controller.get_page(
                                                                                     "TempLow").enter_TempLow()))
            self.temp1.grid(row=4, column=2, padx=5, pady=5)
            self.temp1_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.temp1_status.grid(row=4, column=3, padx=5, pady=5)
            if hideForChuckDept == True:    # we use grid_remove() to hide widgets we don't want certain people to see
                self.temp1.grid_remove()
                self.temp1_status.grid_remove()


            self.temp2 = tk.Button(self, text="Temp Uniformity +25C ", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("TempMed"),
                                                                                 self.controller.get_page(
                                                                                     "TempMed").enter_TempMed()))
            self.temp2.grid(row=5, column=2, padx=5, pady=5)
            self.temp2_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.temp2_status.grid(row=5, column=3, padx=5, pady=5)
            if hideForChuckDept == True:
                self.temp2.grid_remove()
                self.temp2_status.grid_remove()


            self.temp3 = tk.Button(self, text="Temp Uniformity +150C", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("TempHigh"),
                                                                                 self.controller.get_page(
                                                                                     "TempHigh").enter_TempHigh()))
            self.temp3.grid(row=6, column=2, padx=5, pady=5)
            self.temp3_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.temp3_status.grid(row=6, column=3, padx=5, pady=5)
            if hideForChuckDept == True:
                self.temp3.grid_remove()
                self.temp3_status.grid_remove()
            
            
            ###
            self.temp4 = tk.Button(self, text="Temp Uniformity Othr1", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("TempOthr1"),
                                                                                 self.controller.get_page(
                                                                                     "TempOthr1").enter_TempOthr1()))
            self.temp4.grid(row=7, column=2, padx=5, pady=5)
            self.temp4_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.temp4_status.grid(row=7, column=3, padx=5, pady=5)
            if hideForChuckDept == True:
                self.temp4.grid_remove()
                self.temp4_status.grid_remove()


            self.temp5 = tk.Button(self, text="Temp Uniformity Othr2", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("TempOthr2"),
                                                                                 self.controller.get_page(
                                                                                     "TempOthr2").enter_TempOthr2()))
            self.temp5.grid(row=8, column=2, padx=5, pady=5)
            self.temp5_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.temp5_status.grid(row=8, column=3, padx=5, pady=5)
            if hideForChuckDept == True:
                self.temp5.grid_remove()
                self.temp5_status.grid_remove()


            self.temp6 = tk.Button(self, text="Temp Uniformity Othr3", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("TempOthr3"),
                                                                                 self.controller.get_page(
                                                                                     "TempOthr3").enter_TempOthr3()))
            self.temp6.grid(row=9, column=2, padx=5, pady=5)
            self.temp6_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.temp6_status.grid(row=9, column=3, padx=5, pady=5)
            if hideForChuckDept == True:
                self.temp6.grid_remove()
                self.temp6_status.grid_remove()
            
            ####

            self.plan1 = tk.Button(self, text="Planarity -40C ", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("PlanLow"),
                                                                                 self.controller.get_page(
                                                                                     "PlanLow").enter_PlanLow()))
            self.plan1.grid(row=4, column=4, padx=5, pady=5)
            self.plan1_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.plan1_status.grid(row=4, column=5, padx=5, pady=5)
            if hideForFTDept == True:
                self.plan1.grid_remove()
                self.plan1_status.grid_remove()

            self.plan2 = tk.Button(self, text="Planarity +25C ", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("PlanMed"),
                                                                                 self.controller.get_page(
                                                                                     "PlanMed").enter_PlanMed()))
            self.plan2.grid(row=5, column=4, padx=5, pady=5)
            self.plan2_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.plan2_status.grid(row=5, column=5, padx=5, pady=5)
            if hideForFTDept == True:
                self.plan2.grid_remove()
                self.plan2_status.grid_remove()


            self.plan3 = tk.Button(self, text="Planarity +150C", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("PlanHigh"),
                                                                                 self.controller.get_page(
                                                                                     "PlanHigh").enter_PlanHigh()))
            self.plan3.grid(row=6, column=4, padx=5, pady=5)
            self.plan3_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.plan3_status.grid(row=6, column=5, padx=5, pady=5)
            if hideForFTDept == True:
                self.plan3.grid_remove()
                self.plan3_status.grid_remove()

            #

            self.plan4 = tk.Button(self, text="Planarity othr1 ", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("PlanOthr1"),
                                                                                 self.controller.get_page(
                                                                                     "PlanOthr1").enter_PlanOthr1()))
            self.plan4.grid(row=7, column=4, padx=5, pady=5)
            self.plan4_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.plan4_status.grid(row=7, column=5, padx=5, pady=5)
            if hideForFTDept == True:
                self.plan4.grid_remove()
                self.plan4_status.grid_remove()


            self.plan5 = tk.Button(self, text="Planarity Othr2 ", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("PlanOthr2"),
                                                                                 self.controller.get_page(
                                                                                     "PlanOthr2").enter_PlanOthr2()))
            self.plan5.grid(row=8, column=4, padx=5, pady=5)
            self.plan5_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.plan5_status.grid(row=8, column=5, padx=5, pady=5)
            if hideForFTDept == True:
                self.plan5.grid_remove()
                self.plan5_status.grid_remove()


            self.plan6 = tk.Button(self, text="Planarity Othr3", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("PlanOthr3"),
                                                                                 self.controller.get_page(
                                                                                     "PlanOthr3").enter_PlanOthr3()))
            self.plan6.grid(row=9, column=4, padx=5, pady=5)
            self.plan6_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.plan6_status.grid(row=9, column=5, padx=5, pady=5)
            if hideForFTDept == True:
                self.plan6.grid_remove()
                self.plan6_status.grid_remove()


            self.pt100 = tk.Button(self, text="PT100 Calibration", font=controller.label_font, bg="white",
                                   command=lambda: self.controller.combine_funcs(controller.show_frame("Pt100"),
                                                                                 self.controller.get_page(
                                                                                     "Pt100").enter_Pt100()))
            self.pt100.grid(row=16, column=2, columnspan=1, padx=5, pady=5, sticky="E")
            self.pt100_status = tk.Label(self, text="incomplete", bg="red", font=controller.label_font)
            self.pt100_status.grid(row=16, column=3, columnspan=1, padx=5, pady=5, sticky="W")
            if hideForChuckDept == True:
                self.pt100.grid_remove()
                self.pt100_status.grid_remove()


            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("StartPage"))
            button.grid(row=17, column=1, columnspan=5, padx=5, pady=20)
            if hideForChuckDept == False:
                button = tk.Button(self, text="Generate Files",
                               command=self.gen_Files)
                button.grid(row=18, column=1, columnspan=5, padx=5, pady=5)

        def gen_Files(self):
            # note: much of the code in this function is non functional

            self.lf = self.controller.get_page("StartPage").choice_var.get()
            lauf_num = self.lf

            # Define paths #
            snxls_path = '//fileserver/produktion/Endtest/30_Seriennummern/Seriennummern.xlsm'  ## path of seriennummern spreadsheet
            notchiller_path = '//fileserver/alle/ERSTools/Endtest Document Generator/notchiller.xls'# spreadsheet of things that are listed in the chiller column but are NOT actually chillers

            # Import XLS #
            D = pandas.read_excel(snxls_path, sheet_name='Serien_Nummern aufsteigend',
                                  index_col=0)  # imports Seriennummern spreadsheet
            notchiller_df = pandas.read_excel(notchiller_path, dtype=str)

            alpha = False
            i = 0
            while i < len(lauf_num):
                if lauf_num[i].isalpha() == True:
                    alpha = True
                i = i + 1

            if alpha == True:
                if lauf_num not in D.index.values:
                    print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                    error = 'Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste'
                    x = 1 / 0
                else:
                    D = D.loc[str(lauf_num)]
            if alpha == False:
                if len(lauf_num) == 6:
                    if str(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        error = 'Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste'
                        x = 1 / 0
                    else:
                        D = D.loc[str(lauf_num)]
                elif len(lauf_num) != 6:
                    if int(
                            lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                        print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                        error = 'Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste'
                        x = 1 / 0
                    else:
                        D = D.loc[int(lauf_num)]

            # cancel script if lauf number is listed more than once in the excel file
            if len(
                    D) < 28:  # we say less than 28, because if only 1 exits then it has length 28, but if 2 exist then length 2
                print("Error: Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt")
                error = "Error: Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt"
                x = 1 / 0

            ## is There really a chiller? ##
            yeschiller = False
            i = 0
            count = 0
            if D['Chiller'] == D['Chiller']:
                while i < len(notchiller_df):
                    if notchiller_df.at[i, 'st'] in D.at['Chiller']:    # check if a 'not chiller' string is contained in D.at['chiller']. If it is, then it is not a chiller, and thus can't have a pt100 calibration
                        count = count + 1
                    else:
                        count = count
                    i = i + 1
                if count == 0:
                    yeschiller = True

            #       Index of D:  'Serien Nr.', 'PO-Nr.' 'SN', 'SN Kompl.', 'Quartal', 'Jahr',
            #       'best. Liefertermin', 'Auslieferung', 'Kunde', 'Lieferort',
            #       'Serie gesamt', 'Steuergerät', 'Serie', 'Option I', 'Serie.1', 'Chuck',
            #      'Serie.2', 'Chiller', 'Serie.3', 'Option II', 'Serie.4',
            #       'Softwareversion', 'Bemerkungen', 'Zubehör Option', 'Serie.5',
            #      'Temperatur Bereich', 'Unnamed: 26', 'Unnamed: 27']

            SN = str(D['SN Kompl.'])
            SN_str_1 = SN[0:3]
            SN_str_2 = SN[5:12]

            ## Create new directory and add to PDF write path ##
            last4 = SN[len(SN) - 4:len(SN)]

            E = pandas.DataFrame(index=(1, 2, 3, 4, 5), columns=('st', 'or', 'sn', 'type'), data="")

            ### FIll in stuck, order number, serial number ####
            yescontroller = False
            if D['Steuergerät'] == D['Steuergerät']:  # Is controller field empty (NaN)? NaN is not equal to NaN
                yescontroller = True
                E.at[1, 'type'] = 'Controller'
                E.at[1, 'st'] = D['Steuergerät']
                E.at[1, 'or'] = '# ' + lauf_num
                if D['Serie'] == D['Serie']:
                    if len(str(int(D['Serie']))) != 2:
                        temp = '0' + str(int(D['Serie']))
                    else:
                        temp = str(int(D['Serie']))
                if D['Serie'] != D['Serie']:
                    temp = '00'
                E.at[1, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' S'

            yesopt1 = False
            if D['Option I'] == D['Option I']:
                yesopt1 = True
                E.at[2, 'type'] = 'Chiller'
                E.at[2, 'st'] = D['Option I']
                E.at[2, 'or'] = '# ' + lauf_num
                if D['Serie.1'] == D['Serie.1']:
                    if len(str(int(D['Serie.1']))) != 2:
                        temp = '0' + str(int(D['Serie.1']))
                    else:
                        temp = str(int(D['Serie.1']))
                if D['Serie.1'] != D['Serie.1']:
                    temp = '00'
                E.at[2, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' C'

            yeschuck = False
            if D['Chuck'] == D['Chuck']:
                yeschuck = True
                E.at[3, 'type'] = 'Chuck'
                E.at[3, 'st'] = D['Chuck']
                E.at[3, 'or'] = '# ' + lauf_num
                if D['Serie.2'] == D['Serie.2']:
                    if len(str(int(D['Serie.2']))) != 2:
                        temp = '0' + str(int(D['Serie.2']))
                    else:
                        temp = str(int(D['Serie.2']))
                if D['Serie.2'] != D['Serie.2']:
                    temp = '00'
                E.at[3, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' T'

            yesacb = False
            yesbooster = False
            if D['Chiller'] == D['Chiller']:  # if theres something in chiller column
                if yeschiller == True:
                    E.at[4, 'type'] = 'Chiller'
                elif D['Chiller'].__contains__('ACB'):
                    E.at[4, 'type'] = 'Air Control Box'
                    yesacb = True
                else:
                    E.at[4, 'type'] = 'Booster'
                    yesbooster = True
                E.at[4, 'st'] = D['Chiller']
                E.at[4, 'or'] = '# ' + lauf_num
                if D['Serie.3'] == D['Serie.3']:
                    if len(str(int(D['Serie.3']))) != 2:
                        temp = '0' + str(int(D['Serie.3']))
                    else:
                        temp = str(int(D['Serie.3']))
                if D['Serie.3'] != D['Serie.3']:
                    temp = '00'
                if yesbooster == False and D['Chiller'].__contains__('CH20') == False:
                    E.at[4, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' C'
                elif yesbooster == False and D['Chiller'].__contains__('CH20') == True:
                    E.at[4, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' CH'
                else:
                    E.at[4, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' B'

            yesopt2 = False
            if D['Option II'] == D['Option II']:
                yesopt2 = True
                E.at[5, 'type'] = 'Booster'
                E.at[5, 'st'] = D['Option II']
                E.at[5, 'or'] = '# ' + lauf_num
                if D['Serie.4'] == D['Serie.4']:
                    if len(str(int(D['Serie.4']))) != 2:
                        temp = '0' + str(int(D['Serie.4']))
                    else:
                        temp = str(int(D['Serie.4']))
                if D['Serie.4'] != D['Serie.4']:
                    temp = '00'
                E.at[5, 'sn'] = 'Serial No. ' + SN_str_1 + temp + SN_str_2 + ' B'

            #### Determine CHuck Size ####
            t_sz = ''
            if D['Chuck'] == D['Chuck']:
                t_sz = D['Chuck']
                if t_sz[0:4] == 'TC25':
                    t_sz = '25mm'
                else:
                    start_len = len(t_sz)  # start length of chuck name
                    t_sz = t_sz[3:15]  # get rid of first 3 characters
                    t_sz = t_sz.lstrip(
                        'aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ')  ## remove all letters from beginning
                    removed_len = start_len - len(t_sz)
                    t_sz = list(t_sz)  # convert string to list

                    if len(t_sz) != 1:
                        i = 0
                        while t_sz[i].isdigit() and i < (len(t_sz) - 1):  ## find position of first non-numeric character
                            i = i + 1
                    else:
                        i = 0

                    if t_sz[i].isdigit() == True:
                        t_sz = t_sz[0:i + 1]  ##   get rid of all characters after the number ends
                    else:
                        t_sz = t_sz[0:i]

                    t_sz = ''.join(t_sz)  ## join elements of list back together in a string

                    if t_sz == '300':
                        t_sz = '300mm'
                    elif t_sz == '200':
                        t_sz = '200mm'
                    elif t_sz == '150':
                        t_sz = '150mm'
                    elif t_sz == '25':
                        t_sz = '25mm'
                    elif t_sz == '6':
                        t_sz = '6"'
                    elif t_sz == '8':
                        t_sz = '8"'
                    elif t_sz == '12':
                        t_sz = '12"'
                    else:
                        t_sz = ''
                        print(
                            'Achtung!: Chuck größe konnte nicht bestimmt werden. Daher konnten Boxentyp und -gewicht ebenfalls nicht ermittelt werden. Bitte füllen Sie mit einem Stift aus.')

            # determine if Prime System
            prime = False
            if D['Chuck'] == D['Chuck']:
                if D['Chuck'][removed_len - 1] == 'P':
                    prime = True
            if D['Chiller'] == D['Chiller']:
                if ('20P' in D['Chiller']) or ('08P' in D['Chiller']) or ('16P' in D['Chiller']) or (
                        '10P' in D['Chiller']):  # if the chiller is a Prime series
                    prime = True

            #### creates one vs two dataframes depending on if a chiller is being shipped (1 vs 2 shipments #####
            E2 = 0  # used to determine how many times to write a PDF --- changed to dataframe if chiller is present
            if yeschiller == False:  ### NO CHILLER
                E = E[E['or'] != '']  # Removes empty rows from dataframe
                length = len(E.index.tolist())  # determines length of new dataframe
                temp_array = numpy.linspace(1, length, length)  # renumbers index of dataframe
                temp_array = temp_array.astype(int)
                E.index = temp_array.astype(str)  # changes index from 2,4 etc to 1,2,3...

                temp = 5 - len(E)  # How many rows we removed from E (original was 5)
                if temp != 0:
                    temp_array = numpy.linspace(5 - temp + 1, 5, temp)
                    temp_array = temp_array.astype(int)
                    temp_array = temp_array.astype(str)  # create index for new data array
            else:  ### YES CHILLER
                E2 = E.iloc[[3]]  ## creates new dataframe containing chiller row of E
                E2.index = numpy.linspace(1, 1, 1).astype(int).astype(str)  ## reindex E2 to be '1'

                E = E.drop(E.index[3])  # drops chiller row from E
                E = E[E['or'] != '']  # Removes empty rows from dataframe
                length = len(E.index.tolist())  # determines length of new dataframe
                temp_array = numpy.linspace(1, length, length)  # renumbers index of dataframe
                temp_array = temp_array.astype(int)
                E.index = temp_array.astype(str)  # changes index from 2,4 etc to 1,2,3...

                temp = 5 - len(E)  # How many rows we removed from E (original was 5)
                if temp != 0:
                    temp_array = numpy.linspace(5 - temp + 1, 5, temp)
                    temp_array = temp_array.astype(int)
                    temp_array = temp_array.astype(str)  # create index for new data array

            E3 = pandas.DataFrame(index=temp_array, columns=('st', 'or', 'sn', 'type'),
                                  data="")  # create new data array

            E = E.append(E3)  # append original data array with new data array to restore length = 5

            ##### Strings Used in Write Fillable PDF function (don't touch) #####
            ANNOT_KEY = '/Annots'
            ANNOT_FIELD_KEY = '/T'
            ANNOT_VAL_KEY = '/V'
            ANNOT_RECT_KEY = '/Rect'
            SUBTYPE_KEY = '/Subtype'
            WIDGET_SUBTYPE_KEY = '/Widget'

            # this is the function that writes the PDF documents
            def write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict):
                template_pdf = pdfrw.PdfReader(input_pdf_path)
                annotations = template_pdf.pages[0][ANNOT_KEY]
                for annotation in annotations:
                    if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                        if annotation[ANNOT_FIELD_KEY]:
                            key = annotation[ANNOT_FIELD_KEY][1:-1]
                            if key in data_dict.keys():
                                annotation.update(
                                    pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                                )
                pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

            # Determine title #
            i = 1
            title = ''
            while i < 6 - len(E3):
                title = title + E.at[str(i), 'st'] + ' + '
                i = i + 1
            title = title[0:len(title) - 3]
            if D['PO-Nr.'] == D['PO-Nr.']:
                title = title + "   " + D['PO-Nr.']

            sys = 'ERS® AC3:'
            if prime == True:
                sys = 'ERS® AirCool® PRIME:'


            morethanchiller = False
            if yeschuck == True or yescontroller == True or yesopt1 == True or yesopt2 == True:
                morethanchiller = True



            ref_path = '//fileserver/alle/ERSTools/Endtest Document Generator/ref_chillers.xls'
            ref_chucks_path = '//fileserver/alle/ERSTools/Endtest Document Generator/ref_chucks.xls'
            ref_controllers_path = '//fileserver/alle/ERSTools/Endtest Document Generator/ref_controllers.xls'

            # Import XLS #
            ref = pandas.read_excel(ref_path, dtype=str)
            ref = ref.set_index('index')

            ref_controllers = pandas.read_excel(ref_controllers_path, dtype=str)  # imports Seriennummern spreadsheet
            # ref_controllers = ref_controllers.set_index('index')

            ref_chucks = pandas.read_excel(ref_chucks_path, dtype=str)  # imports Seriennummern spreadsheet
            ref_chucks = ref_chucks.set_index('index')

            # find controller type #
            yesrsi = False
            yesvg5 = False
            if yescontroller == True:
                i = 0
                while i < len(ref_controllers):
                    if ref_controllers.at[i, 'st'] in D['Steuergerät']:
                        controller_index = i
                        break
                    else:
                        i = i + 1

                # is the controller actually an RSI box or VG5XX... #
                if ref_controllers.at[controller_index, 'st'] == 'RSI':
                    yesrsi = True
                if ref_controllers.at[controller_index, 'st'] == 'VG5':
                    yesvg5 = True
                    print(
                        'Achtung!:  Gewicht und Kiste #2 bleiben leer, da der Versand der VG5XX-Controller nicht bekannt war, als dieses Programm erstellt wurde. Sie müssen dies mit einem Stift ausfüllen.')

            ## Append E and E2 ##
            if yeschiller == True:
                E = E2.append(E)
                length = len(E)
                temp_array = numpy.linspace(1, length, length)
                temp_array = temp_array.astype(int)
                E.index = temp_array.astype(str)  # changes index to 1,2,3,4 etc

            ## Reformat serial numbers ##
            i = 1
            while i < len(E) + 1:
                E.at[str(i), 'sn'] = E.at[str(i), 'sn'].lstrip(
                    'aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ. ')  #
                i = i + 1


            # modifies type column text to include "- serien #"
            # makes it say PV115P or SP41P PV110 SP -serien #, etc
            i = 1
            while i < 6 and E.at[str(i), 'type'] != '':
                if E.at[str(i), 'st'].__contains__('PV'):
                    E.at[str(i), 'type'] = E.at[str(i), 'st']

                E.at[str(i), 'type'] = E.at[str(i), 'type'] + '-Serien #  '
                i = i + 1



            # find controller name
            ctl = ''
            i = 1
            pt100_pos = i
            if yescontroller == True and yesrsi == False:
                while i < len(E):
                    if E.at[str(i), 'type'] == 'Controller-Serien #  ':
                        break
                    else:
                        i = i + 1
                        pt100_pos = i

            elif yeschiller == True:
                i = 1
                while i < len(E):
                    if E.at[str(i), 'type'] == 'Chiller-Serien #  ':
                        break
                    else:
                        i = i + 1

            yests010 = False  # do we have a TS010 chiller with a display?
            if E.at[str(i), 'st'][0:5] == 'TS010':
                yests010 = True
                pt100_pos = i

            # find chuck name
            chk = ''
            eq_id = ''
            i = 1
            chuck_pos = i
            if yeschuck == True:
                while i < len(E):
                    if E.at[str(i), 'type'] == 'Chuck-Serien #  ':
                        break
                    else:
                        i = i + 1
                        chuck_pos = i

                if t_sz != '150mm' and t_sz != '6"':
                    eq_id = 'Wafer'

            # find type of system
            sys = ''
            if yescontroller == True and yeschuck == True and yesrsi == False:
                sys = 'Chuck + Controller'
            if yeschuck == True and yescontroller == False and yests010 == True:
                sys = 'Chuck + Chiller'


            ###### Write Planarity, Temperature Uniformity Protocol, and PT100 calibration protocol #####
            ################################################################################
            ################################################################################
            ################################################################################
            ################################################################################
            
            # since different directories are created for internal vs external orders, we need to check both directory locations to see if we can read in files
            # for that, we use 2 try statements for every excel file we want to read
            dir1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(lauf_num) + "/"
            dir2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(lauf_num) + "/"

            try:
                data_ref = pandas.read_excel(dir1 + '/data' + '/ref.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_ref = pandas.read_excel(dir2 + '/data' + '/ref.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            try:
                data_TempLow = pandas.read_excel(dir1 + '/data' + '/TempLow.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempLow = pandas.read_excel(dir2 + '/data' + '/TempLow.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            try:
                data_TempMed = pandas.read_excel(dir1 + '/data' + '/TempMed.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempMed = pandas.read_excel(dir2 + '/data' + '/TempMed.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            try:
                data_TempHigh = pandas.read_excel(dir1 + '/data' + '/TempHigh.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempHigh = pandas.read_excel(dir2 + '/data' + '/TempHigh.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass


            try:
                data_PlanLow = pandas.read_excel(dir1 + '/data' + '/PlanLow.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanLow = pandas.read_excel(dir2 + '/data' + '/PlanLow.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            try:
                data_PlanMed = pandas.read_excel(dir1 + '/data' + '/PlanMed.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanMed = pandas.read_excel(dir2 + '/data' + '/PlanMed.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            try:
                data_PlanHigh = pandas.read_excel(dir1 + '/data' + '/PlanHigh.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanHigh = pandas.read_excel(dir2 + '/data' + '/PlanHigh.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            try:
                data_Pt100 = pandas.read_excel(dir1 + '/data' + '/Pt100.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_Pt100 = pandas.read_excel(dir2 + '/data' + '/Pt100.xlsx', sheet_name='Sheet',
                                             dtype=str, header=None, skip_blank_lines=False)
            except:
                pass


            empty_df = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")
            data_TempLow_copy = data_TempLow
            data_TempMed_copy = data_TempMed
            data_TempHigh_copy = data_TempHigh
            data_PlanLow_copy = data_PlanLow
            data_PlanMed_copy = data_PlanMed
            data_PlanHigh_copy = data_PlanHigh
            data_Pt100_copy = data_Pt100

            # append empty df to the data frames to make them bigger, so that we can use them in bigger loops later
            data_TempLow = data_TempLow.append(empty_df)
            data_TempMed = data_TempMed.append(empty_df)
            data_TempHigh = data_TempHigh.append(empty_df)
            data_PlanLow = data_PlanLow.append(empty_df)
            data_PlanMed = data_PlanMed.append(empty_df)
            data_PlanHigh = data_PlanHigh.append(empty_df)
            data_Pt100 = data_Pt100.append(empty_df)

            # predefine these variables
            PlanLow_min = ""
            PlanLow_max = ""
            PlanMed_min = ""
            PlanMed_max = ""
            PlanHigh_min = ""
            PlanHigh_max = ""
            TempLow_min = ""
            TempLow_max = ""
            TempMed_min = ""
            TempMed_max = ""
            TempHigh_min = ""
            TempHigh_max = ""

            # find min/max of each dataset
            TempLow_missing = False
            try:
                TempLow_min = float(data_TempLow.loc[0,0].tolist()[0].replace(",","."))  # if first location in dataframe is empty, then no data has been entered. We then display an error message
                TempLow_max = float(data_TempLow.loc[0,0].tolist()[0].replace(",","."))
            except:
                messagebox.showwarning("Warning", "Low Temperature Uniformity data has not been entered")
                TempLow_missing = True
                pass

            # loop through to find max and min values
            j = 0
            while j < len(data_TempLow):
                try:
                    if float(data_TempLow.loc[j,0].tolist()[0].replace(",",".")) < TempLow_min:
                        TempLow_min = float(data_TempLow.loc[j,0].tolist()[0].replace(",","."))
                except:
                    pass
                try:
                    if float(data_TempLow.loc[j,0].tolist()[0].replace(",",".")) > TempLow_max:
                        TempLow_max = float(data_TempLow.loc[j, 0].tolist()[0].replace(",","."))
                except:
                    pass
                j = j + 1


# repeat all of this for the other temperatures
            TempMed_missing = False
            try:
                TempMed_min = float(data_TempMed.loc[0, 0].tolist()[0].replace(",", "."))
                TempMed_max = float(data_TempMed.loc[0, 0].tolist()[0].replace(",", "."))
            except:
                messagebox.showwarning("Warning", "Med Temperature Uniformity data has not been entered")
                TempMed_missing = True
                pass

            j = 0
            while j < len(data_TempMed):
                try:
                    if float(data_TempMed.loc[j,0].tolist()[0].replace(",",".")) < TempMed_min:
                        TempMed_min = float(data_TempMed.loc[j,0].tolist()[0].replace(",","."))
                except:
                    pass
                try:
                    if float(data_TempMed.loc[j,0].tolist()[0].replace(",",".")) > TempMed_max:
                        TempMed_max = float(data_TempMed.loc[j, 0].tolist()[0].replace(",","."))
                except:
                    pass
                j = j + 1

            TempHigh_missing = False
            try:
                TempHigh_min = float(data_TempHigh.loc[0, 0].tolist()[0].replace(",", "."))
                TempHigh_max = float(data_TempHigh.loc[0, 0].tolist()[0].replace(",", "."))
            except:
                messagebox.showwarning("Warning", "High Temperature Uniformity data has not been entered")
                Temphigh_missing = True
                pass

            j = 0
            while j < len(data_TempHigh):
                try:
                    if float(data_TempHigh.loc[j,0].tolist()[0].replace(",",".")) < TempHigh_min:
                        TempHigh_min = float(data_TempHigh.loc[j,0].tolist()[0].replace(",","."))
                except:
                    pass
                try:
                    if float(data_TempHigh.loc[j,0].tolist()[0].replace(",",".")) > TempHigh_max:
                        TempHigh_max = float(data_TempHigh.loc[j, 0].tolist()[0].replace(",","."))
                except:
                    pass
                j = j + 1

            PlanLow_missing = False
            try:
                PlanLow_min = float(data_PlanLow.loc[0, 0].tolist()[0].replace(",", "."))
                PlanLow_max = float(data_PlanLow.loc[0, 0].tolist()[0].replace(",", "."))
            except:
                messagebox.showwarning("Warning", "Low Temperature Planarity data has not been entered")
                PlanLow_missing = True
                pass

            j = 0
            while j < len(data_PlanLow):
                try:
                    if float(data_PlanLow.loc[j,0].tolist()[0].replace(",",".")) < PlanLow_min:
                        PlanLow_min = float(data_PlanLow.loc[j,0].tolist()[0].replace(",","."))
                except:
                    pass
                try:
                    if float(data_PlanLow.loc[j,0].tolist()[0].replace(",",".")) > PlanLow_max:
                        PlanLow_max = float(data_PlanLow.loc[j, 0].tolist()[0].replace(",","."))
                except:
                    pass
                j = j + 1

            PlanMed_missing = False
            try:
                PlanMed_min = float(data_PlanMed.loc[0, 0].tolist()[0].replace(",", "."))
                PlanMed_max = float(data_PlanMed.loc[0, 0].tolist()[0].replace(",", "."))
            except:
                messagebox.showwarning("Warning", "Med Temperature Planarity data has not been entered")
                PlanMed_missing = True
                pass

            j = 0
            while j < len(data_PlanMed):
                try:
                    if float(data_PlanMed.loc[j,0].tolist()[0].replace(",",".")) < PlanMed_min:
                        PlanMed_min = float(data_PlanMed.loc[j,0].tolist()[0].replace(",","."))
                except:
                    pass
                try:
                    if float(data_PlanMed.loc[j,0].tolist()[0].replace(",",".")) > PlanMed_max:
                        PlanMed_max = float(data_PlanMed.loc[j, 0].tolist()[0].replace(",","."))
                except:
                    pass
                j = j + 1

            PlanHigh_missing = False
            try:
                PlanHigh_min = float(data_PlanHigh.loc[0, 0].tolist()[0].replace(",", "."))
                PlanHigh_max = float(data_PlanHigh.loc[0, 0].tolist()[0].replace(",", "."))
            except:
                messagebox.showwarning("Warning", "High Temperature Planarity data has not been entered")
                PlanHigh_missing = True
                pass
            j = 0
            while j < len(data_PlanHigh):
                try:
                    if float(data_PlanHigh.loc[j,0].tolist()[0].replace(",",".")) < PlanHigh_min:
                        PlanHigh_min = float(data_PlanHigh.loc[j,0].tolist()[0].replace(",","."))
                except:
                    pass
                try:
                    if float(data_PlanHigh.loc[j,0].tolist()[0].replace(",",".")) > PlanHigh_max:
                        PlanHigh_max = float(data_PlanHigh.loc[j, 0].tolist()[0].replace(",","."))
                except:
                    pass
                j = j + 1

            if TempLow_missing == False:
                TempLow_min = round(TempLow_min - float(self.controller.lowTemp),2)
                TempLow_max = round(TempLow_max - float(self.controller.lowTemp),2)
            if TempMed_missing == False:
                TempMed_min = round(TempMed_min - float(self.controller.medTemp),2)
                TempMed_max = round(TempMed_max - float(self.controller.medTemp),2)
            if TempHigh_missing == False:
                TempHigh_min = round(TempHigh_min - float(self.controller.highTemp),2)
                TempHigh_max = round(TempHigh_max - float(self.controller.highTemp),2)


            if PlanLow_missing == False:
                PlanLow_min = round(PlanLow_min)
                PlanLow_max = round(PlanLow_max)
            if PlanMed_missing == False:
                PlanMed_min = round(PlanMed_min)
                PlanMed_max = round(PlanMed_max)
            if PlanHigh_missing == False:
                PlanHigh_min = round(PlanHigh_min)
                PlanHigh_max = round(PlanHigh_max)


            # create lists of string that we will use for the data dict in order to write the fillable PDF
            plan_low = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            i = 0
            while i < len(data_PlanLow_copy):
                plan_low[i] = data_PlanLow.loc[i,0].tolist()[0]
                i = i + 1

            plan_med = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            i = 0
            while i < len(data_PlanMed_copy):
                plan_med[i] = data_PlanMed.loc[i,0].tolist()[0]
                i = i + 1

            plan_high = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            i = 0
            while i < len(data_PlanHigh_copy):
                plan_high[i] = data_PlanHigh.loc[i,0].tolist()[0]
                i = i + 1

            data_dict4 = {
                'sys': sys,
                'chk': E.at[str(chuck_pos), 'st'],
                'chk_sn': E.at[str(chuck_pos), 'sn'],
                'lf_1': '#' + lauf_num,
                'eq_id': eq_id,
                'ht': '',  # high temperature, not complete
                'lt': '',  # low temperature, not complete

                'low_1': plan_low[0],
                'low_2': plan_low[1],
                'low_3': plan_low[2],
                'low_4': plan_low[3],
                'low_5': plan_low[4],
                'low_6': plan_low[5],
                'low_7': plan_low[6],
                'low_8': plan_low[7],
                'low_9': plan_low[8],
                'low_10': plan_low[9],
                'low_11': plan_low[10],
                'low_12': plan_low[11],
                'low_13': plan_low[12],

                'med_1': plan_med[0],
                'med_2': plan_med[1],
                'med_3': plan_med[2],
                'med_4': plan_med[3],
                'med_5': plan_med[4],
                'med_6': plan_med[5],
                'med_7': plan_med[6],
                'med_8': plan_med[7],
                'med_9': plan_med[8],
                'med_10': plan_med[9],
                'med_11': plan_med[10],
                'med_12': plan_med[11],
                'med_13': plan_med[12],

                'high_1': plan_high[0],
                'high_2': plan_high[1],
                'high_3': plan_high[2],
                'high_4': plan_high[3],
                'high_5': plan_high[4],
                'high_6': plan_high[5],
                'high_7': plan_high[6],
                'high_8': plan_high[7],
                'high_9': plan_high[8],
                'high_10': plan_high[9],
                'high_11': plan_high[10],
                'high_12': plan_high[11],
                'high_13': plan_high[12],

                'low_min': PlanLow_min,
                'low_max': PlanLow_max,
                'med_min': PlanMed_min,
                'med_max': PlanMed_max,
                'high_min': PlanHigh_min,
                'high_max': PlanHigh_max,

                'low_spec' :  data_ref.loc[16,0],
                'med_spec': data_ref.loc[17, 0],
                'high_spec': data_ref.loc[18, 0],

                'low_temp1': data_ref.loc[2, 0],
                'low_temp2': data_ref.loc[2, 0],
                'med_temp2': data_ref.loc[3, 0],
                'high_temp1': data_ref.loc[4, 0],
                'high_temp2': data_ref.loc[4, 0],
            }


            temp_low = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            i = 0
            while i < len(data_TempLow_copy):
                temp_low[i] = data_TempLow.loc[i,0].tolist()[0]
                i = i + 1

            temp_med = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            i = 0
            while i < len(data_TempMed_copy):
                temp_med[i] = data_TempMed.loc[i,0].tolist()[0]
                i = i + 1

            temp_high = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
            i = 0
            while i < len(data_TempHigh_copy):
                temp_high[i] = data_TempHigh.loc[i,0].tolist()[0]
                i = i + 1


            # note: I accidently put a p before every variable name for temperature. I meant to do this for planarity.. don't worry, the p here is correct.
            if data_ref.loc[22,0] == '% Temp':
                phigh = str(round(100*float(data_ref.loc[7,0])/float(data_ref.loc[4,0]),1)) + "    %" #if the specification is % Temp, we need to
                                                                                                    # divide the raw spec in deg C by the test temperature to get the % spec
                                                                                                    # ex. phigh = 100*(1 deg C)/(200 deg C) = 0.5 % Temperature
            else:
                phigh = str(data_ref.loc[7,0]) + "  deg C"

            data_dict5 = {
                'sys': sys,
                'chk': E.at[str(chuck_pos), 'st'],
                'chk_sn': E.at[str(chuck_pos), 'sn'],
                'lf_1': '#' + lauf_num,
                'eq_id': eq_id,
                'ht': '',  # high temperature, not complete
                'lt': '',  # low temperature, not complete
                'plow_1': temp_low[0],
                'plow_2': temp_low[1],
                'plow_3': temp_low[2],
                'plow_4': temp_low[3],
                'plow_5': temp_low[4],
                'plow_6': temp_low[5],
                'plow_7': temp_low[6],
                'plow_8': temp_low[7],
                'plow_9': temp_low[8],
                'plow_10': temp_low[9],
                'plow_11': temp_low[10],
                'plow_12': temp_low[11],
                'plow_13': temp_low[12],

                'pmed_1': temp_med[0],
                'pmed_2': temp_med[1],
                'pmed_3': temp_med[2],
                'pmed_4': temp_med[3],
                'pmed_5': temp_med[4],
                'pmed_6': temp_med[5],
                'pmed_7': temp_med[6],
                'pmed_8': temp_med[7],
                'pmed_9': temp_med[8],
                'pmed_10': temp_med[9],
                'pmed_11': temp_med[10],
                'pmed_12': temp_med[11],
                'pmed_13': temp_med[12],

                'phigh_1': temp_high[0],
                'phigh_2': temp_high[1],
                'phigh_3': temp_high[2],
                'phigh_4': temp_high[3],
                'phigh_5': temp_high[4],
                'phigh_6': temp_high[5],
                'phigh_7': temp_high[6],
                'phigh_8': temp_high[7],
                'phigh_9': temp_high[8],
                'phigh_10': temp_high[9],
                'phigh_11': temp_high[10],
                'phigh_12': temp_high[11],
                'phigh_13': temp_high[12],

                'plow_min': TempLow_min,
                'plow_max': TempLow_max,
                'pmed_min': TempMed_min,
                'pmed_max': TempMed_max,
                'phigh_min': TempHigh_min,
                'phigh_max': TempHigh_max,

                'plow_spec' : '± ' + data_ref.loc[5,0],
                'pmed_spec' : '± ' + data_ref.loc[6,0],
                'phigh_spec' : '± ' + phigh,

                'plow_temp1': data_ref.loc[2,0],
                'plow_temp2': data_ref.loc[2,0],
                'pmed_temp2': data_ref.loc[3,0],
                'phigh_temp1': data_ref.loc[4,0],
                'phigh_temp2': data_ref.loc[4,0],
            }


            pt100 = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                     ""]
            i = 0
            while i < len(data_Pt100_copy):  # I had to break the loop into 2 parts to get this to work. For some reason, the strings
                                            # are not formatted correctly when I do not do this
                if i < 13:
                    pt100[i] = data_Pt100.loc[i, 0].tolist()[0]
                if i > 12:
                    pt100[i] = data_Pt100.loc[i, 0]
                i = i + 1

            data_dict6 = {
                'sys': sys,
                'ctl': E.at[str(pt100_pos), 'st'],
                'ctl_sn': E.at[str(pt100_pos), 'sn'],
                'lf_1': '#' + lauf_num,
                'pt100_1': pt100[0],
                'pt100_2': pt100[1],
                'pt100_3': pt100[2],
                'pt100_4': pt100[3],
                'pt100_5': pt100[4],
                'pt100_6': pt100[5],
                'pt100_7': pt100[6],
                'pt100_8': pt100[7],
                'pt100_9': pt100[8],
                'pt100_10': pt100[9],
                'pt100_11': pt100[10],
                'pt100_12': pt100[11],
                'pt100_13': pt100[12],
                'pt100_14': pt100[13],
                'pt100_15': pt100[14],
                'pt100_16': pt100[15],
                'pt100_17': pt100[16],
                'pt100_18': pt100[17],
                'pt100_19': pt100[18],
                'pt100_20': pt100[19],
                'pt100_21': pt100[20],
                'pt100_22': pt100[21],
                'pt100_23': pt100[22],
                'pt100_24': pt100[23],
                'pt100_25': pt100[24],
                'pt100_26': pt100[25],

            }

            # once again main try statements because we are sure where the fire that we want is located
            template_path3 = '//fileserver/Alle/Austin/data_entry_app/planarity_digital.pdf'
            write_path3_1 = dir1 + "/Planarity_digitial_ERS"
            write_path3_2 = dir2 + "/Planarity_digitial_ERS"
            template_path4 = '//fileserver/Alle/Austin/data_entry_app/temperature_digital.pdf'
            write_path4_1 = dir1 + "/Temperature_digitial_ERS"
            write_path4_2 = dir2 + "/Temperature_digitial_ERS"
            template_path5 = '//fileserver/Alle/Austin/data_entry_app/Pt100_calibration_digital.pdf'
            write_path5_1 = dir1 + "/Pt100_digital_ERS"
            write_path5_2 = dir2 + "/Pt100_digital_ERS"

            if yeschuck == True:
                gen_plan = False
                gen_temp = False
                try:
                    write_fillable_pdf(template_path3, write_path3_1 + lauf_num + '.pdf', data_dict4)
                    gen_plan = True
                except:
                    pass
                try:
                    write_fillable_pdf(template_path3, write_path3_2 + lauf_num + '.pdf', data_dict4)
                    gen_plan = True
                except:
                    pass
                try:
                    write_fillable_pdf(template_path4, write_path4_1 + lauf_num + '.pdf', data_dict5)
                    gen_temp = True
                except:
                    pass
                try:
                    write_fillable_pdf(template_path4, write_path4_2 + lauf_num + '.pdf', data_dict5)
                    gen_temp = True
                except:
                    pass
                if gen_plan == False:
                    messagebox.showinfo("Error", "Planarity protocol not generated. Please check to see if the file is already open, or if the folder exists")
                if gen_temp == False:
                    messagebox.showinfo("Error", "Temperature Uniformity protocol not generated. Please check to see if the file is already open, or if the folder exists")

            if (yescontroller == True and yesrsi == False and yesvg5 == False) or yests010 == True:
                gen_pt100 = False
                try:
                    write_fillable_pdf(template_path5, write_path5_1 + lauf_num + '.pdf', data_dict6)
                    gen_pt100 = True
                except:
                    pass
                try:
                    write_fillable_pdf(template_path5, write_path5_2 + lauf_num + '.pdf', data_dict6)
                    gen_pt100 = True
                except:
                    pass
                if gen_pt100 == False:
                    messagebox.showinfo("Error", "Pt100 protocol not generated. Please check to see if the file is already open, or if the folder exists")


            ###############################################################################
            ################    HTU ONLY   ################################################
            ###############################################################################
            ###############################################################################

            if data_ref.loc[23,0] != '999' or data_ref.loc[24,0] != '999' or data_ref.loc[25,0] != '999': # if second set of temperatures exist

                try:
                    data_TempOthr1 = pandas.read_excel(dir1 + '/data' + '/TempOthr1.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass
                try:
                    data_TempOthr1 = pandas.read_excel(dir2 + '/data' + '/TempOthr1.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass

                try:
                    data_TempOthr2 = pandas.read_excel(dir1 + '/data' + '/TempOthr2.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass
                try:
                    data_TempOthr2 = pandas.read_excel(dir2 + '/data' + '/TempOthr2.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass

                try:
                    data_TempOthr3 = pandas.read_excel(dir1 + '/data' + '/TempOthr3.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass
                try:
                    data_TempOthr3 = pandas.read_excel(dir2 + '/data' + '/TempOthr3.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass


                try:
                    data_PlanOthr1 = pandas.read_excel(dir1 + '/data' + '/PlanOthr1.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass
                try:
                    data_PlanOthr1 = pandas.read_excel(dir2 + '/data' + '/PlanOthr1.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass

                try:
                    data_PlanOthr2 = pandas.read_excel(dir1 + '/data' + '/PlanOthr2.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass
                try:
                    data_PlanOthr2 = pandas.read_excel(dir2 + '/data' + '/PlanOthr2.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass

                try:
                    data_PlanOthr3 = pandas.read_excel(dir1 + '/data' + '/PlanOthr3.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass
                try:
                    data_PlanOthr3 = pandas.read_excel(dir2 + '/data' + '/PlanOthr3.xlsx', sheet_name='Sheet',
                                                 dtype=str, header=None, skip_blank_lines=False)
                except:
                    pass


                empty_df = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")
                data_TempOthr1_copy = data_TempOthr1
                data_TempOthr2_copy = data_TempOthr2
                data_TempOthr3_copy = data_TempOthr3
                data_PlanOthr1_copy = data_PlanOthr1
                data_PlanOthr2_copy = data_PlanOthr2
                data_PlanOthr3_copy = data_PlanOthr3

                data_TempOthr1 = data_TempOthr1.append(empty_df)
                data_TempOthr2 = data_TempOthr2.append(empty_df)
                data_TempOthr3 = data_TempOthr3.append(empty_df)
                data_PlanOthr1 = data_PlanOthr1.append(empty_df)
                data_PlanOthr2 = data_PlanOthr2.append(empty_df)
                data_PlanOthr3 = data_PlanOthr3.append(empty_df)

                PlanOthr1_min = ""
                PlanOthr1_max = ""
                PlanOthr2_min = ""
                PlanOthr2_max = ""
                PlanOthr3_min = ""
                PlanOthr3_max = ""
                TempOthr1_min = ""
                TempOthr1_max = ""
                TempOthr2_min = ""
                TempOthr2_max = ""
                TempOthr3_min = ""
                TempOthr3_max = ""
                # find min/max of each dataset
                TempOthr1_missing = False
                try:
                    TempOthr1_min = float(data_TempOthr1.loc[0,0].tolist()[0].replace(",","."))
                    TempOthr1_max = float(data_TempOthr1.loc[0,0].tolist()[0].replace(",","."))
                except:
                    messagebox.showwarning("Warning", "Othr1 Temperature Uniformity data has not been entered")
                    TempOthr1_missing = True
                    pass

                j = 0
                while j < len(data_TempOthr1):
                    try:
                        if float(data_TempOthr1.loc[j,0].tolist()[0].replace(",",".")) < TempOthr1_min:
                            TempOthr1_min = float(data_TempOthr1.loc[j,0].tolist()[0].replace(",","."))
                    except:
                        pass
                    try:
                        if float(data_TempOthr1.loc[j,0].tolist()[0].replace(",",".")) > TempOthr1_max:
                            TempOthr1_max = float(data_TempOthr1.loc[j, 0].tolist()[0].replace(",","."))
                    except:
                        pass
                    j = j + 1

                TempOthr2_missing = False
                try:
                    TempOthr2_min = float(data_TempOthr2.loc[0, 0].tolist()[0].replace(",", "."))
                    TempOthr2_max = float(data_TempOthr2.loc[0, 0].tolist()[0].replace(",", "."))
                except:
                    messagebox.showwarning("Warning", "Othr2 Temperature Uniformity data has not been entered")
                    TempOthr2_missing = True
                    pass

                j = 0
                while j < len(data_TempOthr2):
                    try:
                        if float(data_TempOthr2.loc[j,0].tolist()[0].replace(",",".")) < TempOthr2_min:
                            TempOthr2_min = float(data_TempOthr2.loc[j,0].tolist()[0].replace(",","."))
                    except:
                        pass
                    try:
                        if float(data_TempOthr2.loc[j,0].tolist()[0].replace(",",".")) > TempOthr2_max:
                            TempOthr2_max = float(data_TempOthr2.loc[j, 0].tolist()[0].replace(",","."))
                    except:
                        pass
                    j = j + 1

                TempOthr3_missing = False
                try:
                    TempOthr3_min = float(data_TempOthr3.loc[0, 0].tolist()[0].replace(",", "."))
                    TempOthr3_max = float(data_TempOthr3.loc[0, 0].tolist()[0].replace(",", "."))
                except:
                    messagebox.showwarning("Warning", "Othr3 Temperature Uniformity data has not been entered")
                    Tempothr1_missing = True
                    pass

                j = 0
                while j < len(data_TempOthr3):
                    try:
                        if float(data_TempOthr3.loc[j,0].tolist()[0].replace(",",".")) < TempOthr3_min:
                            TempOthr3_min = float(data_TempOthr3.loc[j,0].tolist()[0].replace(",","."))
                    except:
                        pass
                    try:
                        if float(data_TempOthr3.loc[j,0].tolist()[0].replace(",",".")) > TempOthr3_max:
                            TempOthr3_max = float(data_TempOthr3.loc[j, 0].tolist()[0].replace(",","."))
                    except:
                        pass
                    j = j + 1

                PlanOthr1_missing = False
                try:
                    PlanOthr1_min = float(data_PlanOthr1.loc[0, 0].tolist()[0].replace(",", "."))
                    PlanOthr1_max = float(data_PlanOthr1.loc[0, 0].tolist()[0].replace(",", "."))
                except:
                    messagebox.showwarning("Warning", "Othr1 Temperature Planarity data has not been entered")
                    PlanOthr1_missing = True
                    pass

                j = 0
                while j < len(data_PlanOthr1):
                    try:
                        if float(data_PlanOthr1.loc[j,0].tolist()[0].replace(",",".")) < PlanOthr1_min:
                            PlanOthr1_min = float(data_PlanOthr1.loc[j,0].tolist()[0].replace(",","."))
                    except:
                        pass
                    try:
                        if float(data_PlanOthr1.loc[j,0].tolist()[0].replace(",",".")) > PlanOthr1_max:
                            PlanOthr1_max = float(data_PlanOthr1.loc[j, 0].tolist()[0].replace(",","."))
                    except:
                        pass
                    j = j + 1

                PlanOthr2_missing = False
                try:
                    PlanOthr2_min = float(data_PlanOthr2.loc[0, 0].tolist()[0].replace(",", "."))
                    PlanOthr2_max = float(data_PlanOthr2.loc[0, 0].tolist()[0].replace(",", "."))
                except:
                    messagebox.showwarning("Warning", "Othr2 Temperature Planarity data has not been entered")
                    PlanOthr2_missing = True
                    pass

                j = 0
                while j < len(data_PlanOthr2):
                    try:
                        if float(data_PlanOthr2.loc[j,0].tolist()[0].replace(",",".")) < PlanOthr2_min:
                            PlanOthr2_min = float(data_PlanOthr2.loc[j,0].tolist()[0].replace(",","."))
                    except:
                        pass
                    try:
                        if float(data_PlanOthr2.loc[j,0].tolist()[0].replace(",",".")) > PlanOthr2_max:
                            PlanOthr2_max = float(data_PlanOthr2.loc[j, 0].tolist()[0].replace(",","."))
                    except:
                        pass
                    j = j + 1

                PlanOthr3_missing = False
                try:
                    PlanOthr3_min = float(data_PlanOthr3.loc[0, 0].tolist()[0].replace(",", "."))
                    PlanOthr3_max = float(data_PlanOthr3.loc[0, 0].tolist()[0].replace(",", "."))
                except:
                    messagebox.showwarning("Warning", "Othr3 Temperature Planarity data has not been entered")
                    PlanOthr3_missing = True
                    pass
                j = 0
                while j < len(data_PlanOthr3):
                    try:
                        if float(data_PlanOthr3.loc[j,0].tolist()[0].replace(",",".")) < PlanOthr3_min:
                            PlanOthr3_min = float(data_PlanOthr3.loc[j,0].tolist()[0].replace(",","."))
                    except:
                        pass
                    try:
                        if float(data_PlanOthr3.loc[j,0].tolist()[0].replace(",",".")) > PlanOthr3_max:
                            PlanOthr3_max = float(data_PlanOthr3.loc[j, 0].tolist()[0].replace(",","."))
                    except:
                        pass
                    j = j + 1

                if TempOthr1_missing == False:
                    TempOthr1_min = round(TempOthr1_min - float(self.controller.othr1Temp),2)
                    TempOthr1_max = round(TempOthr1_max - float(self.controller.othr1Temp),2)
                if TempOthr2_missing == False:
                    TempOthr2_min = round(TempOthr2_min - float(self.controller.othr2Temp),2)
                    TempOthr2_max = round(TempOthr2_max - float(self.controller.othr2Temp),2)
                if TempOthr3_missing == False:
                    TempOthr3_min = round(TempOthr3_min - float(self.controller.othr1Temp),2)
                    TempOthr3_max = round(TempOthr3_max - float(self.controller.othr1Temp),2)


                if PlanOthr1_missing == False:
                    PlanOthr1_min = round(PlanOthr1_min)
                    PlanOthr1_max = round(PlanOthr1_max)
                if PlanOthr2_missing == False:
                    PlanOthr2_min = round(PlanOthr2_min)
                    PlanOthr2_max = round(PlanOthr2_max)
                if PlanOthr3_missing == False:
                    PlanOthr3_min = round(PlanOthr3_min)
                    PlanOthr3_max = round(PlanOthr3_max)


                plan_low = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
                i = 0
                while i < len(data_PlanOthr1_copy):
                    plan_low[i] = data_PlanOthr1.loc[i,0].tolist()[0]
                    i = i + 1

                plan_med = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
                i = 0
                while i < len(data_PlanOthr2_copy):
                    plan_med[i] = data_PlanOthr2.loc[i,0].tolist()[0]
                    i = i + 1

                plan_high = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
                i = 0
                while i < len(data_PlanOthr3_copy):
                    plan_high[i] = data_PlanOthr3.loc[i,0].tolist()[0]
                    i = i + 1

                data_dict7 = {
                    'sys': sys,
                    'chk': E.at[str(chuck_pos), 'st'],
                    'chk_sn': E.at[str(chuck_pos), 'sn'],
                    'lf_1': '#' + lauf_num,
                    'eq_id': eq_id,
                    'ht': '',  # high temperature, not complete
                    'lt': '',  # low temperature, not complete

                    'low_1': plan_low[0],
                    'low_2': plan_low[1],
                    'low_3': plan_low[2],
                    'low_4': plan_low[3],
                    'low_5': plan_low[4],
                    'low_6': plan_low[5],
                    'low_7': plan_low[6],
                    'low_8': plan_low[7],
                    'low_9': plan_low[8],
                    'low_10': plan_low[9],
                    'low_11': plan_low[10],
                    'low_12': plan_low[11],
                    'low_13': plan_low[12],

                    'med_1': plan_med[0],
                    'med_2': plan_med[1],
                    'med_3': plan_med[2],
                    'med_4': plan_med[3],
                    'med_5': plan_med[4],
                    'med_6': plan_med[5],
                    'med_7': plan_med[6],
                    'med_8': plan_med[7],
                    'med_9': plan_med[8],
                    'med_10': plan_med[9],
                    'med_11': plan_med[10],
                    'med_12': plan_med[11],
                    'med_13': plan_med[12],

                    'high_1': plan_high[0],
                    'high_2': plan_high[1],
                    'high_3': plan_high[2],
                    'high_4': plan_high[3],
                    'high_5': plan_high[4],
                    'high_6': plan_high[5],
                    'high_7': plan_high[6],
                    'high_8': plan_high[7],
                    'high_9': plan_high[8],
                    'high_10': plan_high[9],
                    'high_11': plan_high[10],
                    'high_12': plan_high[11],
                    'high_13': plan_high[12],

                    'low_min': PlanOthr1_min,
                    'low_max': PlanOthr1_max,
                    'med_min': PlanOthr2_min,
                    'med_max': PlanOthr2_max,
                    'high_min': PlanOthr3_min,
                    'high_max': PlanOthr3_max,

                    'low_spec' :  data_ref.loc[35,0],
                    'med_spec': data_ref.loc[36, 0],
                    'high_spec': data_ref.loc[37, 0],

                    'othr1_temp1': data_ref.loc[23, 0],
                    'othr1_temp2': data_ref.loc[23, 0],
                    'othr2_temp1': data_ref.loc[24, 0],
                    'othr2_temp2': data_ref.loc[24, 0],
                    'othr3_temp1': data_ref.loc[25, 0],
                    'othr3_temp2': data_ref.loc[25, 0],
                }


                temp_othr1 = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
                i = 0
                while i < len(data_TempOthr1_copy):
                    temp_othr1[i] = data_TempOthr1.loc[i,0].tolist()[0]
                    i = i + 1

                temp_othr2 = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
                i = 0
                while i < len(data_TempOthr2_copy):
                    temp_othr2[i] = data_TempOthr2.loc[i,0].tolist()[0]
                    i = i + 1

                temp_othr3 = ["", "", "", "", "", "", "", "", "", "", "", "", ""]
                i = 0
                while i < len(data_TempOthr3_copy):
                    temp_othr3[i] = data_TempOthr3.loc[i,0].tolist()[0]
                    i = i + 1


              #  if data_ref.loc[22,0] == '% Temp':
               #     phigh = str(round(100*float(data_ref.loc[7,0])/float(data_ref.loc[4,0]),1)) + "    %"
              #  else:
               #     phigh = str(data_ref.loc[7,0]) + "  deg C"

                data_dict8 = {
                    'sys': sys,
                    'chk': E.at[str(chuck_pos), 'st'],
                    'chk_sn': E.at[str(chuck_pos), 'sn'],
                    'lf_1': '#' + lauf_num,
                    'eq_id': eq_id,
                    'ht': '',  # high temperature, not complete
                    'lt': '',  # low temperature, not complete
                    'plow_1': temp_othr1[0],
                    'plow_2': temp_othr1[1],
                    'plow_3': temp_othr1[2],
                    'plow_4': temp_othr1[3],
                    'plow_5': temp_othr1[4],
                    'plow_6': temp_othr1[5],
                    'plow_7': temp_othr1[6],
                    'plow_8': temp_othr1[7],
                    'plow_9': temp_othr1[8],
                    'plow_10': temp_othr1[9],
                    'plow_11': temp_othr1[10],
                    'plow_12': temp_othr1[11],
                    'plow_13': temp_othr1[12],

                    'pmed_1': temp_othr2[0],
                    'pmed_2': temp_othr2[1],
                    'pmed_3': temp_othr2[2],
                    'pmed_4': temp_othr2[3],
                    'pmed_5': temp_othr2[4],
                    'pmed_6': temp_othr2[5],
                    'pmed_7': temp_othr2[6],
                    'pmed_8': temp_othr2[7],
                    'pmed_9': temp_othr2[8],
                    'pmed_10': temp_othr2[9],
                    'pmed_11': temp_othr2[10],
                    'pmed_12': temp_othr2[11],
                    'pmed_13': temp_othr2[12],

                    'phigh_1': temp_othr3[0],
                    'phigh_2': temp_othr3[1],
                    'phigh_3': temp_othr3[2],
                    'phigh_4': temp_othr3[3],
                    'phigh_5': temp_othr3[4],
                    'phigh_6': temp_othr3[5],
                    'phigh_7': temp_othr3[6],
                    'phigh_8': temp_othr3[7],
                    'phigh_9': temp_othr3[8],
                    'phigh_10': temp_othr3[9],
                    'phigh_11': temp_othr3[10],
                    'phigh_12': temp_othr3[11],
                    'phigh_13': temp_othr3[12],

                    'plow_min': TempOthr1_min,
                    'plow_max': TempOthr1_max,
                    'pmed_min': TempOthr2_min,
                    'pmed_max': TempOthr2_max,
                    'phigh_min': TempOthr3_min,
                    'phigh_max': TempOthr3_max,

                    'plow_spec' : '± ' + data_ref.loc[26,0],
                    'pmed_spec' : '± ' + data_ref.loc[27,0],
                    'phigh_spec': '± ' + data_ref.loc[28,0],
                    #'phigh_spec' : '± ' + phigh,

                    'pothr1_temp1': data_ref.loc[23,0],
                    'pothr1_temp2': data_ref.loc[23,0],
                    'pothr2_temp1': data_ref.loc[24,0],
                    'pothr2_temp2': data_ref.loc[24,0],
                    'pothr3_temp1': data_ref.loc[25,0],
                    'pothr3_temp2': data_ref.loc[25,0],
                }



                template_path3 = '//fileserver/Alle/Austin/data_entry_app/planarity_digital_2.pdf'
                write_path3_1 = dir1 + "/Planarity_digitial_ERS"
                write_path3_2 = dir2 + "/Planarity_digitial_ERS"
                template_path4 = '//fileserver/Alle/Austin/data_entry_app/temperature_digital_2.pdf'
                write_path4_1 = dir1 + "/Temperature_digitial_ERS"
                write_path4_2 = dir2 + "/Temperature_digitial_ERS"


                if yeschuck == True:
                    gen_plan2 = False
                    gen_temp2 = False
                    try:
                        write_fillable_pdf(template_path3, write_path3_1 + lauf_num + '_2.pdf', data_dict7)
                        gen_plan2 = True
                    except:
                        pass
                    try:
                        write_fillable_pdf(template_path3, write_path3_2 + lauf_num + '_2.pdf', data_dict7)
                        gen_plan2 = True
                    except:
                        pass
                    try:
                        write_fillable_pdf(template_path4, write_path4_1 + lauf_num + '_2.pdf', data_dict8)
                        gen_temp2 = True
                    except:
                        pass
                    try:
                        write_fillable_pdf(template_path4, write_path4_2 + lauf_num + '_2.pdf', data_dict8)
                        gen_temp2 = True
                    except:
                        pass
                    if gen_plan2 == False:
                        messagebox.showinfo("Error", "Second planarity protocol not generated. Please check to see if the file is already open, or if the folder exists")
                    if gen_temp2 == False:
                        messagebox.showinfo("Error", "Second temperature Uniformity protocol not generated. Please check to see if the file is already open, or if the folder exists")

    class TempLow(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        # this function updates the displayed LF number, displayed temperatures, and loads previously saved data
        def enter_TempLow(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)

            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] == self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            # find specs to be used to verify input values
            # really we only need the specs for lowTemp, but when I wrote the code originally I found them all. Just so you know.
            self.lowTemp = LF_status[2].tolist()[0]
            self.MedTemp = LF_status[3].tolist()[0]
            self.HighTemp = LF_status[4].tolist()[0]
            self.spec_lowTemp = LF_status[5].tolist()[0]
            self.spec_MedTemp = LF_status[6].tolist()[0]
            self.spec_HighTemp = LF_status[7].tolist()[0]

            self.lf = self.controller.get_page("StartPage").choice_var.get()
            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/TempLow.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/TempLow.xlsx"
            try:
                data_TempLow = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempLow = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass


            # delete existing values from text boxes, in case there are any
            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            # depending on the size of the dataframe that we generated from the TempLow XLSX file, we will in a different amount of text boxes
            # if we didn't do this, we would get errors because we would be specifying coordinates inside the dataframe that do not exist
            if data_TempLow.empty == False:
                if type(data_TempLow.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_TempLow.loc[0].tolist()[0])

            if len(data_TempLow) > 1:
                if type(data_TempLow.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_TempLow.loc[1].tolist()[0])

            if len(data_TempLow) >2:
                if type(data_TempLow.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_TempLow.loc[2].tolist()[0])

            if len(data_TempLow) > 3:
                if type(data_TempLow.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_TempLow.loc[3].tolist()[0])

            if len(data_TempLow) > 4:
                if type(data_TempLow.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_TempLow.loc[4].tolist()[0])

            if len(data_TempLow) > 5:
                if type(data_TempLow.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_TempLow.loc[5].tolist()[0])

            if len(data_TempLow) > 6:
                if type(data_TempLow.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_TempLow.loc[6].tolist()[0])

            if len(data_TempLow) > 7:
                if type(data_TempLow.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_TempLow.loc[7].tolist()[0])

            if len(data_TempLow) > 8:
                if type(data_TempLow.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_TempLow.loc[8].tolist()[0])

            if len(data_TempLow) > 9:
                if type(data_TempLow.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_TempLow.loc[9].tolist()[0])

            if len(data_TempLow) > 10:
                if type(data_TempLow.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_TempLow.loc[10].tolist()[0])

            if len(data_TempLow) > 11:
                if type(data_TempLow.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_TempLow.loc[11].tolist()[0])

            if len(data_TempLow) > 12:
                if type(data_TempLow.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_TempLow.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")
            # data to save
            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            # check if anything is out of spec. If so, Save = False (don't allow save)
            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.lowTemp)+ float(self.spec_lowTemp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (float(self.lowTemp) - float(self.spec_lowTemp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            # if the "finished" box is checked
            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/TempLow.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/TempLow.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 8] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 8] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class TempMed(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)
            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_TempMed(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.lowTemp = LF_status[2].tolist()[0]
            self.MedTemp = LF_status[3].tolist()[0]
            self.HighTemp = LF_status[4].tolist()[0]
            self.spec_lowTemp = LF_status[5].tolist()[0]
            self.spec_MedTemp = LF_status[6].tolist()[0]
            self.spec_HighTemp = LF_status[7].tolist()[0]


            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/TempMed.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/TempMed.xlsx"
            try:
                data_TempMed = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempMed = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass


            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_TempMed.empty == False:
                if type(data_TempMed.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_TempMed.loc[0].tolist()[0])

            if len(data_TempMed) > 1:
                if type(data_TempMed.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_TempMed.loc[1].tolist()[0])

            if len(data_TempMed) >2:
                if type(data_TempMed.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_TempMed.loc[2].tolist()[0])

            if len(data_TempMed) > 3:
                if type(data_TempMed.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_TempMed.loc[3].tolist()[0])

            if len(data_TempMed) > 4:
                if type(data_TempMed.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_TempMed.loc[4].tolist()[0])

            if len(data_TempMed) > 5:
                if type(data_TempMed.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_TempMed.loc[5].tolist()[0])

            if len(data_TempMed) > 6:
                if type(data_TempMed.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_TempMed.loc[6].tolist()[0])

            if len(data_TempMed) > 7:
                if type(data_TempMed.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_TempMed.loc[7].tolist()[0])

            if len(data_TempMed) > 8:
                if type(data_TempMed.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_TempMed.loc[8].tolist()[0])

            if len(data_TempMed) > 9:
                if type(data_TempMed.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_TempMed.loc[9].tolist()[0])

            if len(data_TempMed) > 10:
                if type(data_TempMed.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_TempMed.loc[10].tolist()[0])

            if len(data_TempMed) > 11:
                if type(data_TempMed.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_TempMed.loc[11].tolist()[0])

            if len(data_TempMed) > 12:
                if type(data_TempMed.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_TempMed.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i, 0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.MedTemp) + float(self.spec_MedTemp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (float(self.MedTemp) - float(self.spec_MedTemp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/TempMed.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/TempMed.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 9] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 9] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class TempHigh(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_TempHigh(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
           # LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.lowTemp = LF_status[2].tolist()[0]
            self.HighTemp = LF_status[3].tolist()[0]
            self.HighTemp = LF_status[4].tolist()[0]
            self.spec_lowTemp = LF_status[5].tolist()[0]
            self.spec_HighTemp = LF_status[6].tolist()[0]
            self.spec_HighTemp = LF_status[7].tolist()[0]


            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/TempHigh.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/TempHigh.xlsx"
            try:
                data_TempHigh = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempHigh = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_TempHigh.empty == False:
                if type(data_TempHigh.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_TempHigh.loc[0].tolist()[0])

            if len(data_TempHigh) > 1:
                if type(data_TempHigh.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_TempHigh.loc[1].tolist()[0])

            if len(data_TempHigh) >2:
                if type(data_TempHigh.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_TempHigh.loc[2].tolist()[0])

            if len(data_TempHigh) > 3:
                if type(data_TempHigh.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_TempHigh.loc[3].tolist()[0])

            if len(data_TempHigh) > 4:
                if type(data_TempHigh.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_TempHigh.loc[4].tolist()[0])

            if len(data_TempHigh) > 5:
                if type(data_TempHigh.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_TempHigh.loc[5].tolist()[0])

            if len(data_TempHigh) > 6:
                if type(data_TempHigh.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_TempHigh.loc[6].tolist()[0])

            if len(data_TempHigh) > 7:
                if type(data_TempHigh.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_TempHigh.loc[7].tolist()[0])

            if len(data_TempHigh) > 8:
                if type(data_TempHigh.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_TempHigh.loc[8].tolist()[0])

            if len(data_TempHigh) > 9:
                if type(data_TempHigh.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_TempHigh.loc[9].tolist()[0])

            if len(data_TempHigh) > 10:
                if type(data_TempHigh.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_TempHigh.loc[10].tolist()[0])

            if len(data_TempHigh) > 11:
                if type(data_TempHigh.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_TempHigh.loc[11].tolist()[0])

            if len(data_TempHigh) > 12:
                if type(data_TempHigh.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_TempHigh.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i, 0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.HighTemp) + float(self.spec_HighTemp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (float(self.HighTemp) - float(self.spec_HighTemp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                i = i + 1


            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/TempHigh.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/TempHigh.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 10] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 10] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class TempOthr1(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_TempOthr1(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.Othr1Temp = LF_status[23].tolist()[0]
            self.spec_Othr1Temp = LF_status[26].tolist()[0]


            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/TempOthr1.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/TempOthr1.xlsx"
            try:
                data_TempOthr1 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempOthr1 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_TempOthr1.empty == False:
                if type(data_TempOthr1.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_TempOthr1.loc[0].tolist()[0])

            if len(data_TempOthr1) > 1:
                if type(data_TempOthr1.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_TempOthr1.loc[1].tolist()[0])

            if len(data_TempOthr1) >2:
                if type(data_TempOthr1.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_TempOthr1.loc[2].tolist()[0])

            if len(data_TempOthr1) > 3:
                if type(data_TempOthr1.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_TempOthr1.loc[3].tolist()[0])

            if len(data_TempOthr1) > 4:
                if type(data_TempOthr1.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_TempOthr1.loc[4].tolist()[0])

            if len(data_TempOthr1) > 5:
                if type(data_TempOthr1.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_TempOthr1.loc[5].tolist()[0])

            if len(data_TempOthr1) > 6:
                if type(data_TempOthr1.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_TempOthr1.loc[6].tolist()[0])

            if len(data_TempOthr1) > 7:
                if type(data_TempOthr1.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_TempOthr1.loc[7].tolist()[0])

            if len(data_TempOthr1) > 8:
                if type(data_TempOthr1.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_TempOthr1.loc[8].tolist()[0])

            if len(data_TempOthr1) > 9:
                if type(data_TempOthr1.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_TempOthr1.loc[9].tolist()[0])

            if len(data_TempOthr1) > 10:
                if type(data_TempOthr1.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_TempOthr1.loc[10].tolist()[0])

            if len(data_TempOthr1) > 11:
                if type(data_TempOthr1.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_TempOthr1.loc[11].tolist()[0])

            if len(data_TempOthr1) > 12:
                if type(data_TempOthr1.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_TempOthr1.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i, 0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.Othr1Temp) + float(self.spec_Othr1Temp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (float(self.Othr1Temp) - float(self.spec_Othr1Temp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                i = i + 1


            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/TempOthr1.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/TempOthr1.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 29] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 29] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class TempOthr2(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_TempOthr2(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.Othr2Temp = LF_status[24].tolist()[0]
            self.spec_Othr2Temp = LF_status[27].tolist()[0]


            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/TempOthr2.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/TempOthr2.xlsx"
            try:
                data_TempOthr2 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempOthr2 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_TempOthr2.empty == False:
                if type(data_TempOthr2.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_TempOthr2.loc[0].tolist()[0])

            if len(data_TempOthr2) > 1:
                if type(data_TempOthr2.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_TempOthr2.loc[1].tolist()[0])

            if len(data_TempOthr2) >2:
                if type(data_TempOthr2.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_TempOthr2.loc[2].tolist()[0])

            if len(data_TempOthr2) > 3:
                if type(data_TempOthr2.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_TempOthr2.loc[3].tolist()[0])

            if len(data_TempOthr2) > 4:
                if type(data_TempOthr2.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_TempOthr2.loc[4].tolist()[0])

            if len(data_TempOthr2) > 5:
                if type(data_TempOthr2.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_TempOthr2.loc[5].tolist()[0])

            if len(data_TempOthr2) > 6:
                if type(data_TempOthr2.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_TempOthr2.loc[6].tolist()[0])

            if len(data_TempOthr2) > 7:
                if type(data_TempOthr2.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_TempOthr2.loc[7].tolist()[0])

            if len(data_TempOthr2) > 8:
                if type(data_TempOthr2.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_TempOthr2.loc[8].tolist()[0])

            if len(data_TempOthr2) > 9:
                if type(data_TempOthr2.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_TempOthr2.loc[9].tolist()[0])

            if len(data_TempOthr2) > 10:
                if type(data_TempOthr2.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_TempOthr2.loc[10].tolist()[0])

            if len(data_TempOthr2) > 11:
                if type(data_TempOthr2.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_TempOthr2.loc[11].tolist()[0])

            if len(data_TempOthr2) > 12:
                if type(data_TempOthr2.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_TempOthr2.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i, 0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.Othr2Temp) + float(self.spec_Othr2Temp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (float(self.Othr2Temp) - float(self.spec_Othr2Temp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                i = i + 1


            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/TempOthr2.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/TempOthr2.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 30] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 30] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class TempOthr3(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_TempOthr3(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.Othr3Temp = LF_status[25].tolist()[0]
            self.spec_Othr3Temp = LF_status[28].tolist()[0]


            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/TempOthr3.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/TempOthr3.xlsx"
            try:
                data_TempOthr3 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_TempOthr3 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_TempOthr3.empty == False:
                if type(data_TempOthr3.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_TempOthr3.loc[0].tolist()[0])

            if len(data_TempOthr3) > 1:
                if type(data_TempOthr3.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_TempOthr3.loc[1].tolist()[0])

            if len(data_TempOthr3) >2:
                if type(data_TempOthr3.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_TempOthr3.loc[2].tolist()[0])

            if len(data_TempOthr3) > 3:
                if type(data_TempOthr3.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_TempOthr3.loc[3].tolist()[0])

            if len(data_TempOthr3) > 4:
                if type(data_TempOthr3.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_TempOthr3.loc[4].tolist()[0])

            if len(data_TempOthr3) > 5:
                if type(data_TempOthr3.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_TempOthr3.loc[5].tolist()[0])

            if len(data_TempOthr3) > 6:
                if type(data_TempOthr3.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_TempOthr3.loc[6].tolist()[0])

            if len(data_TempOthr3) > 7:
                if type(data_TempOthr3.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_TempOthr3.loc[7].tolist()[0])

            if len(data_TempOthr3) > 8:
                if type(data_TempOthr3.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_TempOthr3.loc[8].tolist()[0])

            if len(data_TempOthr3) > 9:
                if type(data_TempOthr3.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_TempOthr3.loc[9].tolist()[0])

            if len(data_TempOthr3) > 10:
                if type(data_TempOthr3.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_TempOthr3.loc[10].tolist()[0])

            if len(data_TempOthr3) > 11:
                if type(data_TempOthr3.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_TempOthr3.loc[11].tolist()[0])

            if len(data_TempOthr3) > 12:
                if type(data_TempOthr3.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_TempOthr3.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i, 0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.Othr3Temp) + float(self.spec_Othr3Temp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (float(self.Othr3Temp) - float(self.spec_Othr3Temp))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                i = i + 1


            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/TempOthr3.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/TempOthr3.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 31] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 31] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class PlanLow(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_PlanLow(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.lowPlan = LF_status[2].tolist()[0]
            self.MedPlan = LF_status[3].tolist()[0]
            self.HighPlan = LF_status[4].tolist()[0]
            self.spec_lowPlan = LF_status[16].tolist()[0]
            self.spec_MedPlan = LF_status[17].tolist()[0]
            self.spec_HighPlan = LF_status[18].tolist()[0]

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/PlanLow.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/PlanLow.xlsx"
            try:
                data_PlanLow = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanLow = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_PlanLow.empty == False:
                if type(data_PlanLow.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_PlanLow.loc[0].tolist()[0])

            if len(data_PlanLow) > 1:
                if type(data_PlanLow.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_PlanLow.loc[1].tolist()[0])

            if len(data_PlanLow) >2:
                if type(data_PlanLow.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_PlanLow.loc[2].tolist()[0])

            if len(data_PlanLow) > 3:
                if type(data_PlanLow.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_PlanLow.loc[3].tolist()[0])

            if len(data_PlanLow) > 4:
                if type(data_PlanLow.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_PlanLow.loc[4].tolist()[0])

            if len(data_PlanLow) > 5:
                if type(data_PlanLow.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_PlanLow.loc[5].tolist()[0])

            if len(data_PlanLow) > 6:
                if type(data_PlanLow.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_PlanLow.loc[6].tolist()[0])

            if len(data_PlanLow) > 7:
                if type(data_PlanLow.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_PlanLow.loc[7].tolist()[0])

            if len(data_PlanLow) > 8:
                if type(data_PlanLow.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_PlanLow.loc[8].tolist()[0])

            if len(data_PlanLow) > 9:
                if type(data_PlanLow.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_PlanLow.loc[9].tolist()[0])

            if len(data_PlanLow) > 10:
                if type(data_PlanLow.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_PlanLow.loc[10].tolist()[0])

            if len(data_PlanLow) > 11:
                if type(data_PlanLow.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_PlanLow.loc[11].tolist()[0])

            if len(data_PlanLow) > 12:
                if type(data_PlanLow.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_PlanLow.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.spec_lowPlan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (- float(self.spec_lowPlan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/PlanLow.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/PlanLow.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 11] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 11] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class PlanMed(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_PlanMed(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]


            self.lowPlan = LF_status[2].tolist()[0]
            self.MedPlan = LF_status[3].tolist()[0]
            self.HighPlan = LF_status[4].tolist()[0]
            self.spec_lowPlan = LF_status[16].tolist()[0]
            self.spec_MedPlan = LF_status[17].tolist()[0]
            self.spec_HighPlan = LF_status[18].tolist()[0]

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/PlanMed.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/PlanMed.xlsx"
            try:
                data_PlanMed = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanMed = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass


            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_PlanMed.empty == False:
                if type(data_PlanMed.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_PlanMed.loc[0].tolist()[0])

            if len(data_PlanMed) > 1:
                if type(data_PlanMed.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_PlanMed.loc[1].tolist()[0])

            if len(data_PlanMed) >2:
                if type(data_PlanMed.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_PlanMed.loc[2].tolist()[0])

            if len(data_PlanMed) > 3:
                if type(data_PlanMed.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_PlanMed.loc[3].tolist()[0])

            if len(data_PlanMed) > 4:
                if type(data_PlanMed.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_PlanMed.loc[4].tolist()[0])

            if len(data_PlanMed) > 5:
                if type(data_PlanMed.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_PlanMed.loc[5].tolist()[0])

            if len(data_PlanMed) > 6:
                if type(data_PlanMed.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_PlanMed.loc[6].tolist()[0])

            if len(data_PlanMed) > 7:
                if type(data_PlanMed.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_PlanMed.loc[7].tolist()[0])

            if len(data_PlanMed) > 8:
                if type(data_PlanMed.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_PlanMed.loc[8].tolist()[0])

            if len(data_PlanMed) > 9:
                if type(data_PlanMed.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_PlanMed.loc[9].tolist()[0])

            if len(data_PlanMed) > 10:
                if type(data_PlanMed.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_PlanMed.loc[10].tolist()[0])

            if len(data_PlanMed) > 11:
                if type(data_PlanMed.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_PlanMed.loc[11].tolist()[0])

            if len(data_PlanMed) > 12:
                if type(data_PlanMed.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_PlanMed.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.spec_MedPlan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (- float(self.spec_MedPlan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/PlanMed.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/PlanMed.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 12] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 12] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class PlanHigh(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_PlanHigh(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.lowPlan = LF_status[2].tolist()[0]
            self.MedPlan = LF_status[3].tolist()[0]
            self.HighPlan = LF_status[4].tolist()[0]
            self.spec_lowPlan = LF_status[16].tolist()[0]
            self.spec_MedPlan = LF_status[17].tolist()[0]
            self.spec_HighPlan = LF_status[18].tolist()[0]

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/PlanHigh.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/PlanHigh.xlsx"
            try:
                data_PlanHigh = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanHigh = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_PlanHigh.empty == False:
                if type(data_PlanHigh.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_PlanHigh.loc[0].tolist()[0])

            if len(data_PlanHigh) > 1:
                if type(data_PlanHigh.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_PlanHigh.loc[1].tolist()[0])

            if len(data_PlanHigh) >2:
                if type(data_PlanHigh.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_PlanHigh.loc[2].tolist()[0])

            if len(data_PlanHigh) > 3:
                if type(data_PlanHigh.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_PlanHigh.loc[3].tolist()[0])

            if len(data_PlanHigh) > 4:
                if type(data_PlanHigh.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_PlanHigh.loc[4].tolist()[0])

            if len(data_PlanHigh) > 5:
                if type(data_PlanHigh.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_PlanHigh.loc[5].tolist()[0])

            if len(data_PlanHigh) > 6:
                if type(data_PlanHigh.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_PlanHigh.loc[6].tolist()[0])

            if len(data_PlanHigh) > 7:
                if type(data_PlanHigh.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_PlanHigh.loc[7].tolist()[0])

            if len(data_PlanHigh) > 8:
                if type(data_PlanHigh.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_PlanHigh.loc[8].tolist()[0])

            if len(data_PlanHigh) > 9:
                if type(data_PlanHigh.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_PlanHigh.loc[9].tolist()[0])

            if len(data_PlanHigh) > 10:
                if type(data_PlanHigh.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_PlanHigh.loc[10].tolist()[0])

            if len(data_PlanHigh) > 11:
                if type(data_PlanHigh.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_PlanHigh.loc[11].tolist()[0])

            if len(data_PlanHigh) > 12:
                if type(data_PlanHigh.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_PlanHigh.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.spec_HighPlan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (- float(self.spec_HighPlan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/PlanHigh.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/PlanHigh.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 13] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 13] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class PlanOthr1(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_PlanOthr1(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.othr1Plan = LF_status[23].tolist()[0]
            self.spec_othr1Plan = LF_status[35].tolist()[0]

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/PlanOthr1.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/PlanOthr1.xlsx"
            try:
                data_PlanOthr1 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanOthr1 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_PlanOthr1.empty == False:
                if type(data_PlanOthr1.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_PlanOthr1.loc[0].tolist()[0])

            if len(data_PlanOthr1) > 1:
                if type(data_PlanOthr1.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_PlanOthr1.loc[1].tolist()[0])

            if len(data_PlanOthr1) >2:
                if type(data_PlanOthr1.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_PlanOthr1.loc[2].tolist()[0])

            if len(data_PlanOthr1) > 3:
                if type(data_PlanOthr1.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_PlanOthr1.loc[3].tolist()[0])

            if len(data_PlanOthr1) > 4:
                if type(data_PlanOthr1.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_PlanOthr1.loc[4].tolist()[0])

            if len(data_PlanOthr1) > 5:
                if type(data_PlanOthr1.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_PlanOthr1.loc[5].tolist()[0])

            if len(data_PlanOthr1) > 6:
                if type(data_PlanOthr1.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_PlanOthr1.loc[6].tolist()[0])

            if len(data_PlanOthr1) > 7:
                if type(data_PlanOthr1.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_PlanOthr1.loc[7].tolist()[0])

            if len(data_PlanOthr1) > 8:
                if type(data_PlanOthr1.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_PlanOthr1.loc[8].tolist()[0])

            if len(data_PlanOthr1) > 9:
                if type(data_PlanOthr1.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_PlanOthr1.loc[9].tolist()[0])

            if len(data_PlanOthr1) > 10:
                if type(data_PlanOthr1.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_PlanOthr1.loc[10].tolist()[0])

            if len(data_PlanOthr1) > 11:
                if type(data_PlanOthr1.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_PlanOthr1.loc[11].tolist()[0])

            if len(data_PlanOthr1) > 12:
                if type(data_PlanOthr1.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_PlanOthr1.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.spec_othr1Plan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (- float(self.spec_othr1Plan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/PlanOthr1.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/PlanOthr1.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 32] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 32] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class PlanOthr2(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_PlanOthr2(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.othr2Plan = LF_status[24].tolist()[0]
            self.spec_othr2Plan = LF_status[36].tolist()[0]

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/PlanOthr2.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/PlanOthr2.xlsx"
            try:
                data_PlanOthr2 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanOthr2 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_PlanOthr2.empty == False:
                if type(data_PlanOthr2.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_PlanOthr2.loc[0].tolist()[0])

            if len(data_PlanOthr2) > 1:
                if type(data_PlanOthr2.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_PlanOthr2.loc[1].tolist()[0])

            if len(data_PlanOthr2) >2:
                if type(data_PlanOthr2.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_PlanOthr2.loc[2].tolist()[0])

            if len(data_PlanOthr2) > 3:
                if type(data_PlanOthr2.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_PlanOthr2.loc[3].tolist()[0])

            if len(data_PlanOthr2) > 4:
                if type(data_PlanOthr2.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_PlanOthr2.loc[4].tolist()[0])

            if len(data_PlanOthr2) > 5:
                if type(data_PlanOthr2.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_PlanOthr2.loc[5].tolist()[0])

            if len(data_PlanOthr2) > 6:
                if type(data_PlanOthr2.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_PlanOthr2.loc[6].tolist()[0])

            if len(data_PlanOthr2) > 7:
                if type(data_PlanOthr2.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_PlanOthr2.loc[7].tolist()[0])

            if len(data_PlanOthr2) > 8:
                if type(data_PlanOthr2.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_PlanOthr2.loc[8].tolist()[0])

            if len(data_PlanOthr2) > 9:
                if type(data_PlanOthr2.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_PlanOthr2.loc[9].tolist()[0])

            if len(data_PlanOthr2) > 10:
                if type(data_PlanOthr2.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_PlanOthr2.loc[10].tolist()[0])

            if len(data_PlanOthr2) > 11:
                if type(data_PlanOthr2.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_PlanOthr2.loc[11].tolist()[0])

            if len(data_PlanOthr2) > 12:
                if type(data_PlanOthr2.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_PlanOthr2.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.spec_othr2Plan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (- float(self.spec_othr2Plan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/PlanOthr2.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/PlanOthr2.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 33] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 33] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class PlanOthr3(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=5, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=5, padx=50, pady=0)

            self.image = tk.Button(self, command=self.verify)
            self.image.grid(row=2, column=2, columnspan=3, padx=10, pady=0)
            photo_location = 'X://ERSTools/EndtestData/temp_image.png'
            self._img0 = tk.PhotoImage(file=photo_location)
            self.image.configure(image=self._img0)

            self.entry1_label = tk.Label(self, text="1", font=controller.entry_font)
            self.entry1_label.grid(row=3, column=2, padx=5, pady=5)
            self.entry1 = tk.Entry(self, width=15)
            self.entry1.grid(row=4, column=2, padx=5, pady=5)

            self.entry2_label = tk.Label(self, text="2", font=controller.entry_font)
            self.entry2_label.grid(row=3, column=3, padx=5, pady=5)
            self.entry2 = tk.Entry(self, width=15)
            self.entry2.grid(row=4, column=3, padx=5, pady=5)

            self.entry3_label = tk.Label(self, text="3", font=controller.entry_font)
            self.entry3_label.grid(row=3, column=4, padx=5, pady=5)
            self.entry3 = tk.Entry(self, width=15)
            self.entry3.grid(row=4, column=4, padx=5, pady=5)

            self.entry4_label = tk.Label(self, text="4", font=controller.entry_font)
            self.entry4_label.grid(row=5, column=2, padx=5, pady=5)
            self.entry4 = tk.Entry(self, width=15)
            self.entry4.grid(row=6, column=2, padx=5, pady=5)

            self.entry5_label = tk.Label(self, text="5", font=controller.entry_font)
            self.entry5_label.grid(row=5, column=3, padx=5, pady=5)
            self.entry5 = tk.Entry(self, width=15)
            self.entry5.grid(row=6, column=3, padx=5, pady=5)

            self.entry6_label = tk.Label(self, text="6", font=controller.entry_font)
            self.entry6_label.grid(row=5, column=4, padx=5, pady=5)
            self.entry6 = tk.Entry(self, width=15)
            self.entry6.grid(row=6, column=4, padx=5, pady=5)

            self.entry7_label = tk.Label(self, text="7", font=controller.entry_font)
            self.entry7_label.grid(row=7, column=2, padx=5, pady=5)
            self.entry7 = tk.Entry(self, width=15)
            self.entry7.grid(row=8, column=2, padx=5, pady=5)

            self.entry8_label = tk.Label(self, text="8", font=controller.entry_font)
            self.entry8_label.grid(row=7, column=3, padx=5, pady=5)
            self.entry8 = tk.Entry(self, width=15)
            self.entry8.grid(row=8, column=3, padx=5, pady=5)

            self.entry9_label = tk.Label(self, text="9", font=controller.entry_font)
            self.entry9_label.grid(row=7, column=4, padx=5, pady=5)
            self.entry9 = tk.Entry(self, width=15)
            self.entry9.grid(row=8, column=4, padx=5, pady=5)

            self.entry10_label = tk.Label(self, text="10", font=controller.entry_font)
            self.entry10_label.grid(row=9, column=2, padx=5, pady=5)
            self.entry10 = tk.Entry(self, width=15)
            self.entry10.grid(row=10, column=2, padx=5, pady=5)

            self.entry11_label = tk.Label(self, text="11", font=controller.entry_font)
            self.entry11_label.grid(row=9, column=3, padx=5, pady=5)
            self.entry11 = tk.Entry(self, width=15)
            self.entry11.grid(row=10, column=3, padx=5, pady=5)

            self.entry12_label = tk.Label(self, text="12", font=controller.entry_font)
            self.entry12_label.grid(row=9, column=4, padx=5, pady=5)
            self.entry12 = tk.Entry(self, width=15)
            self.entry12.grid(row=10, column=4, padx=5, pady=5)

            self.entry13_label = tk.Label(self, text="13", font=controller.entry_font)
            self.entry13_label.grid(row=11, column=3, padx=5, pady=5)
            self.entry13 = tk.Entry(self, width=15)
            self.entry13.grid(row=12, column=3, padx=5, pady=5)

            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=14, column=1, padx=30, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=14, column=2, padx=30, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=14, column=4, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=14, column=5, padx=0, pady=5)
            if hideForChuckDept == True or hideForFTDept == True:
                self.checkIgnore.grid_remove()

        def enter_PlanOthr3(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()
            manager_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                             dtype=str, header=None)
            #LF_status = manager_list[manager_list[0].str.match(self.lf)]
            i = 0
            while i < len(manager_list):
                if manager_list.loc[i,0] ==  self.lf:
                    index = i
                i = i + 1
            LF_status = manager_list[manager_list.index == index]

            self.othr3Plan = LF_status[25].tolist()[0]
            self.spec_othr3Plan = LF_status[37].tolist()[0]

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/PlanOthr3.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/PlanOthr3.xlsx"
            try:
                data_PlanOthr3 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_PlanOthr3 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)

            if data_PlanOthr3.empty == False:
                if type(data_PlanOthr3.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_PlanOthr3.loc[0].tolist()[0])

            if len(data_PlanOthr3) > 1:
                if type(data_PlanOthr3.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_PlanOthr3.loc[1].tolist()[0])

            if len(data_PlanOthr3) >2:
                if type(data_PlanOthr3.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_PlanOthr3.loc[2].tolist()[0])

            if len(data_PlanOthr3) > 3:
                if type(data_PlanOthr3.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_PlanOthr3.loc[3].tolist()[0])

            if len(data_PlanOthr3) > 4:
                if type(data_PlanOthr3.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_PlanOthr3.loc[4].tolist()[0])

            if len(data_PlanOthr3) > 5:
                if type(data_PlanOthr3.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_PlanOthr3.loc[5].tolist()[0])

            if len(data_PlanOthr3) > 6:
                if type(data_PlanOthr3.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_PlanOthr3.loc[6].tolist()[0])

            if len(data_PlanOthr3) > 7:
                if type(data_PlanOthr3.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_PlanOthr3.loc[7].tolist()[0])

            if len(data_PlanOthr3) > 8:
                if type(data_PlanOthr3.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_PlanOthr3.loc[8].tolist()[0])

            if len(data_PlanOthr3) > 9:
                if type(data_PlanOthr3.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_PlanOthr3.loc[9].tolist()[0])

            if len(data_PlanOthr3) > 10:
                if type(data_PlanOthr3.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_PlanOthr3.loc[10].tolist()[0])

            if len(data_PlanOthr3) > 11:
                if type(data_PlanOthr3.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_PlanOthr3.loc[11].tolist()[0])

            if len(data_PlanOthr3) > 12:
                if type(data_PlanOthr3.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_PlanOthr3.loc[12].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error", "Measurement" + str(i + 1) + " missing")
                if data.loc[i,0] != "":
                    if (float(data.loc[i, 0].replace(",",".")) > (float(self.spec_othr3Plan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i + 1) + " out of spec")
                    if (float(data.loc[i, 0].replace(",",".")) < (- float(self.spec_othr3Plan))):
                        save = False
                        messagebox.showinfo("Error", "Measurement" + str(i+1) + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/PlanOthr3.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/PlanOthr3.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 34] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 34] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x=self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    class Pt100(tk.Frame):

        def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.label = tk.Label(self, text="", font=controller.title_font)
            self.label.grid(row=1, column=1, columnspan=8, padx = 0, pady=10)

            self.label = tk.Label(self, text="PT100 Simulator Type 4503s, No 273807", font=controller.label_font)
            self.label.grid(row=2, column=1, columnspan=8, padx = 0, pady=10)

            none = tk.Label(self, text="X", fg="grey",font=controller.label_font)
            none.grid(row=2, column=1, padx=50, pady=0)
            none2 = tk.Label(self, text="X", fg="grey", font=controller.label_font)
            none2.grid(row=2, column=8, padx=50, pady=0)

            tk.Label(self, text="Nominal", font=controller.entry_font).grid(row=3, column=2, padx=5, pady=0)
            tk.Label(self, text="Value", font=controller.entry_font).grid(row=4, column=2, padx=5, pady=0)
            tk.Label(self, text="Measured", font=controller.entry_font).grid(row=3, column=3, padx=5, pady=0)
            tk.Label(self, text="Value", font=controller.entry_font).grid(row=4, column=3, padx=5, pady=0)
            tk.Label(self, text="Calibration", font=controller.entry_font).grid(row=3, column=4, padx=5, pady=0)
            tk.Label(self, text="Value", font=controller.entry_font).grid(row=4, column=4, padx=15, pady=0)

            tk.Label(self, text="-80,0", font=controller.entry_font).grid(row=5, column=2, padx=5, pady=0)
            tk.Label(self, text="-60,0", font=controller.entry_font).grid(row=6, column=2, padx=5, pady=0)
            tk.Label(self, text="-50,0", font=controller.entry_font).grid(row=7, column=2, padx=5, pady=0)
            tk.Label(self, text="-40,0", font=controller.entry_font).grid(row=8, column=2, padx=5, pady=0)
            tk.Label(self, text="-20,0", font=controller.entry_font).grid(row=9, column=2, padx=5, pady=0)
            tk.Label(self, text="-10,0", font=controller.entry_font).grid(row=10, column=2, padx=5, pady=0)
            tk.Label(self, text="0,0", font=controller.entry_font).grid(row=11, column=2, padx=5, pady=0)
            tk.Label(self, text="10,0", font=controller.entry_font).grid(row=12, column=2, padx=5, pady=0)
            tk.Label(self, text="20,0", font=controller.entry_font).grid(row=13, column=2, padx=5, pady=0)
            tk.Label(self, text="25,0", font=controller.entry_font).grid(row=14, column=2, padx=5, pady=0)
            tk.Label(self, text="30,0", font=controller.entry_font).grid(row=15, column=2, padx=5, pady=0)
            tk.Label(self, text="40,0", font=controller.entry_font).grid(row=16, column=2, padx=5, pady=0)
            tk.Label(self, text="50,0", font=controller.entry_font).grid(row=17, column=2, padx=5, pady=0)

            self.entry1 = tk.Entry(self, width=9)
            self.entry1.grid(row=5, column=3, padx=5, pady=0)
            self.entry2 = tk.Entry(self, width=9)
            self.entry2.grid(row=6, column=3, padx=5, pady=0)
            self.entry3 = tk.Entry(self, width=9)
            self.entry3.grid(row=7, column=3, padx=5, pady=0)
            self.entry4 = tk.Entry(self, width=9)
            self.entry4.grid(row=8, column=3, padx=5, pady=0)
            self.entry5 = tk.Entry(self, width=9)
            self.entry5.grid(row=9, column=3, padx=5, pady=0)
            self.entry6 = tk.Entry(self, width=9)
            self.entry6.grid(row=10, column=3, padx=5, pady=0)
            self.entry7 = tk.Entry(self, width=9)
            self.entry7.grid(row=11, column=3, padx=5, pady=0)
            self.entry8 = tk.Entry(self, width=9)
            self.entry8.grid(row=12, column=3, padx=5, pady=0)
            self.entry9 = tk.Entry(self, width=9)
            self.entry9.grid(row=13, column=3, padx=5, pady=0)
            self.entry10 = tk.Entry(self, width=9)
            self.entry10.grid(row=14, column=3, padx=5, pady=0)
            self.entry11 = tk.Entry(self, width=9)
            self.entry11.grid(row=15, column=3, padx=5, pady=0)
            self.entry12 = tk.Entry(self, width=9)
            self.entry12.grid(row=16, column=3, padx=5, pady=0)
            self.entry13 = tk.Entry(self, width=9)
            self.entry13.grid(row=17, column=3, padx=5, pady=0)

            tk.Label(self, text="-79,89", font=controller.entry_font).grid(row=5, column=4, padx=5, pady=0)
            tk.Label(self, text="-59,92", font=controller.entry_font).grid(row=6, column=4, padx=5, pady=0)
            tk.Label(self, text="-49,91", font=controller.entry_font).grid(row=7, column=4, padx=5, pady=0)
            tk.Label(self, text="-39,93", font=controller.entry_font).grid(row=8, column=4, padx=5, pady=0)
            tk.Label(self, text="-19,92", font=controller.entry_font).grid(row=9, column=4, padx=5, pady=0)
            tk.Label(self, text="-9,95", font=controller.entry_font).grid(row=10, column=4, padx=5, pady=0)
            tk.Label(self, text="0,03", font=controller.entry_font).grid(row=11, column=4, padx=5, pady=0)
            tk.Label(self, text="10,08", font=controller.entry_font).grid(row=12, column=4, padx=5, pady=0)
            tk.Label(self, text="20,08", font=controller.entry_font).grid(row=13, column=4, padx=5, pady=0)
            tk.Label(self, text="25,08", font=controller.entry_font).grid(row=14, column=4, padx=5, pady=0)
            tk.Label(self, text="30,05", font=controller.entry_font).grid(row=15, column=4, padx=5, pady=0)
            tk.Label(self, text= "40,08", font=controller.entry_font).grid(row=16, column=4, padx=5, pady=0)
            tk.Label(self, text="50,08", font=controller.entry_font).grid(row=17, column=4, padx=5, pady=0)


            tk.Label(self, text="Nominal", font=controller.entry_font).grid(row=3, column=5, padx=15, pady=0)
            tk.Label(self, text="Value", font=controller.entry_font).grid(row=4, column=5, padx=5, pady=0)
            tk.Label(self, text="Measured", font=controller.entry_font).grid(row=3, column=6, padx=5, pady=0)
            tk.Label(self, text="Value", font=controller.entry_font).grid(row=4, column=6, padx=5, pady=0)
            tk.Label(self, text="Calibration", font=controller.entry_font).grid(row=3, column=7, padx=5, pady=0)
            tk.Label(self, text="Value", font=controller.entry_font).grid(row=4, column=7, padx=5, pady=0)

            tk.Label(self, text="60,0", font=controller.entry_font).grid(row=5, column=5, padx=5, pady=0)
            tk.Label(self, text="70,0", font=controller.entry_font).grid(row=6, column=5, padx=5, pady=0)
            tk.Label(self, text="80,0", font=controller.entry_font).grid(row=7, column=5, padx=5, pady=0)
            tk.Label(self, text="90,0", font=controller.entry_font).grid(row=8, column=5, padx=5, pady=0)
            tk.Label(self, text="100,0", font=controller.entry_font).grid(row=9, column=5, padx=5, pady=0)
            tk.Label(self, text="120,0", font=controller.entry_font).grid(row=10, column=5, padx=5, pady=0)
            tk.Label(self, text="140,0", font=controller.entry_font).grid(row=11, column=5, padx=5, pady=0)
            tk.Label(self, text="150,0", font=controller.entry_font).grid(row=12, column=5, padx=5, pady=0)
            tk.Label(self, text="160,0", font=controller.entry_font).grid(row=13, column=5, padx=5, pady=0)
            tk.Label(self, text="180,0", font=controller.entry_font).grid(row=14, column=5, padx=5, pady=0)
            tk.Label(self, text="200,0", font=controller.entry_font).grid(row=15, column=5, padx=5, pady=0)
            tk.Label(self, text="250,0", font=controller.entry_font).grid(row=16, column=5, padx=5, pady=0)
            tk.Label(self, text="300,0", font=controller.entry_font).grid(row=17, column=5, padx=5, pady=0)

            self.entry14 = tk.Entry(self, width=9)
            self.entry14.grid(row=5, column=6, padx=5, pady=0)
            self.entry15 = tk.Entry(self, width=9)
            self.entry15.grid(row=6, column=6, padx=5, pady=0)
            self.entry16 = tk.Entry(self, width=9)
            self.entry16.grid(row=7, column=6, padx=5, pady=0)
            self.entry17 = tk.Entry(self, width=9)
            self.entry17.grid(row=8, column=6, padx=5, pady=0)
            self.entry18 = tk.Entry(self, width=9)
            self.entry18.grid(row=9, column=6, padx=5, pady=0)
            self.entry19 = tk.Entry(self, width=9)
            self.entry19.grid(row=10, column=6, padx=5, pady=0)
            self.entry20 = tk.Entry(self, width=9)
            self.entry20.grid(row=11, column=6, padx=5, pady=0)
            self.entry21 = tk.Entry(self, width=9)
            self.entry21.grid(row=12, column=6, padx=5, pady=0)
            self.entry22 = tk.Entry(self, width=9)
            self.entry22.grid(row=13, column=6, padx=5, pady=0)
            self.entry23 = tk.Entry(self, width=9)
            self.entry23.grid(row=14, column=6, padx=5, pady=0)
            self.entry24 = tk.Entry(self, width=9)
            self.entry24.grid(row=15, column=6, padx=5, pady=0)
            self.entry25 = tk.Entry(self, width=9)
            self.entry25.grid(row=16, column=6, padx=5, pady=0)
            self.entry26 = tk.Entry(self, width=9)
            self.entry26.grid(row=17, column=6, padx=5, pady=0)

            tk.Label(self, text="60,08", font=controller.entry_font).grid(row=5, column=7, padx=5, pady=0)
            tk.Label(self, text="70,08", font=controller.entry_font).grid(row=6, column=7, padx=5, pady=0)
            tk.Label(self, text="80,08", font=controller.entry_font).grid(row=7, column=7, padx=5, pady=0)
            tk.Label(self, text="90,07", font=controller.entry_font).grid(row=8, column=7, padx=5, pady=0)
            tk.Label(self, text="100,09", font=controller.entry_font).grid(row=9, column=7, padx=5, pady=0)
            tk.Label(self, text="120,09", font=controller.entry_font).grid(row=10, column=7, padx=5, pady=0)
            tk.Label(self, text="140,08", font=controller.entry_font).grid(row=11, column=7, padx=5, pady=0)
            tk.Label(self, text="150,06", font=controller.entry_font).grid(row=12, column=7, padx=5, pady=0)
            tk.Label(self, text="160,09", font=controller.entry_font).grid(row=13, column=7, padx=5, pady=0)
            tk.Label(self, text="180,07", font=controller.entry_font).grid(row=14, column=7, padx=5, pady=0)
            tk.Label(self, text="200,05", font=controller.entry_font).grid(row=15, column=7, padx=5, pady=0)
            tk.Label(self, text="250,02", font=controller.entry_font).grid(row=16, column=7, padx=5, pady=0)
            tk.Label(self, text="300,09", font=controller.entry_font).grid(row=17, column=7, padx=5, pady=0)


            button = tk.Button(self, text="Go Back",
                               command=lambda: controller.show_frame("Manager"))
            button.grid(row=20, column=1, padx=5, pady=20)

            save = tk.Button(self, text="Save",
                               command=self.save)
            save.grid(row=20, column=2, padx=5, pady=20)

            self.checkComplete_var = tk.IntVar()
            self.checkComplete = tk.Checkbutton(self, text="Finished", variable=self.checkComplete_var)
            self.checkComplete.grid(row=20, column=3, columnspan=2, padx=0, pady=5)
            self.checkIgnore_var = tk.IntVar()
            self.checkIgnore = tk.Checkbutton(self, text="Ignore Errors (Save Anyway)", variable=self.checkIgnore_var)
            self.checkIgnore.grid(row=20, column=5, columnspan=2, padx=0, pady=5)

        def enter_Pt100(self):
            self.lf = self.controller.get_page("StartPage").choice_var.get()

            last4 = str(self.controller.last4)
            path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(self.lf) + "/data/Pt100.xlsx"
            path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(self.lf) + "/data/Pt100.xlsx"
            try:
                data_Pt100 = pandas.read_excel(str(path1),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass
            try:
                data_Pt100 = pandas.read_excel(str(path2),dtype=str, header=None, skip_blank_lines=False)
            except:
                pass

            self.entry1.delete(0, tk.END)
            self.entry2.delete(0, tk.END)
            self.entry3.delete(0, tk.END)
            self.entry4.delete(0, tk.END)
            self.entry5.delete(0, tk.END)
            self.entry6.delete(0, tk.END)
            self.entry7.delete(0, tk.END)
            self.entry8.delete(0, tk.END)
            self.entry9.delete(0, tk.END)
            self.entry10.delete(0, tk.END)
            self.entry11.delete(0, tk.END)
            self.entry12.delete(0, tk.END)
            self.entry13.delete(0, tk.END)
            self.entry14.delete(0, tk.END)
            self.entry15.delete(0, tk.END)
            self.entry16.delete(0, tk.END)
            self.entry17.delete(0, tk.END)
            self.entry18.delete(0, tk.END)
            self.entry19.delete(0, tk.END)
            self.entry20.delete(0, tk.END)
            self.entry21.delete(0, tk.END)
            self.entry22.delete(0, tk.END)
            self.entry23.delete(0, tk.END)
            self.entry24.delete(0, tk.END)
            self.entry25.delete(0, tk.END)
            self.entry26.delete(0, tk.END)


            if data_Pt100.empty == False:
                if type(data_Pt100.loc[0].tolist()[0]) == str:
                    self.entry1.insert(0, data_Pt100.loc[0].tolist()[0])

            if len(data_Pt100) > 1:
                if type(data_Pt100.loc[1].tolist()[0]) == str:
                    self.entry2.insert(0, data_Pt100.loc[1].tolist()[0])

            if len(data_Pt100) >2:
                if type(data_Pt100.loc[2].tolist()[0]) == str:
                    self.entry3.insert(0, data_Pt100.loc[2].tolist()[0])

            if len(data_Pt100) > 3:
                if type(data_Pt100.loc[3].tolist()[0]) == str:
                    self.entry4.insert(0, data_Pt100.loc[3].tolist()[0])

            if len(data_Pt100) > 4:
                if type(data_Pt100.loc[4].tolist()[0]) == str:
                    self.entry5.insert(0, data_Pt100.loc[4].tolist()[0])

            if len(data_Pt100) > 5:
                if type(data_Pt100.loc[5].tolist()[0]) == str:
                    self.entry6.insert(0, data_Pt100.loc[5].tolist()[0])

            if len(data_Pt100) > 6:
                if type(data_Pt100.loc[6].tolist()[0]) == str:
                    self.entry7.insert(0, data_Pt100.loc[6].tolist()[0])

            if len(data_Pt100) > 7:
                if type(data_Pt100.loc[7].tolist()[0]) == str:
                    self.entry8.insert(0, data_Pt100.loc[7].tolist()[0])

            if len(data_Pt100) > 8:
                if type(data_Pt100.loc[8].tolist()[0]) == str:
                    self.entry9.insert(0, data_Pt100.loc[8].tolist()[0])

            if len(data_Pt100) > 9:
                if type(data_Pt100.loc[9].tolist()[0]) == str:
                    self.entry10.insert(0, data_Pt100.loc[9].tolist()[0])

            if len(data_Pt100) > 10:
                if type(data_Pt100.loc[10].tolist()[0]) == str:
                    self.entry11.insert(0, data_Pt100.loc[10].tolist()[0])

            if len(data_Pt100) > 11:
                if type(data_Pt100.loc[11].tolist()[0]) == str:
                    self.entry12.insert(0, data_Pt100.loc[11].tolist()[0])

            if len(data_Pt100) > 12:
                if type(data_Pt100.loc[12].tolist()[0]) == str:
                    self.entry13.insert(0, data_Pt100.loc[12].tolist()[0])

            if len(data_Pt100) > 13:
                if type(data_Pt100.loc[13].tolist()[0]) == str:
                    self.entry14.insert(0, data_Pt100.loc[13].tolist()[0])

            if len(data_Pt100) > 14:
                if type(data_Pt100.loc[14].tolist()[0]) == str:
                    self.entry15.insert(0, data_Pt100.loc[14].tolist()[0])

            if len(data_Pt100) > 15:
                if type(data_Pt100.loc[15].tolist()[0]) == str:
                    self.entry16.insert(0, data_Pt100.loc[15].tolist()[0])

            if len(data_Pt100) > 16:
                if type(data_Pt100.loc[16].tolist()[0]) == str:
                    self.entry17.insert(0, data_Pt100.loc[16].tolist()[0])

            if len(data_Pt100) > 17:
                if type(data_Pt100.loc[17].tolist()[0]) == str:
                    self.entry18.insert(0, data_Pt100.loc[17].tolist()[0])

            if len(data_Pt100) > 18:
                if type(data_Pt100.loc[18].tolist()[0]) == str:
                    self.entry19.insert(0, data_Pt100.loc[18].tolist()[0])

            if len(data_Pt100) > 19:
                if type(data_Pt100.loc[19].tolist()[0]) == str:
                    self.entry20.insert(0, data_Pt100.loc[19].tolist()[0])

            if len(data_Pt100) > 20:
                if type(data_Pt100.loc[20].tolist()[0]) == str:
                    self.entry21.insert(0, data_Pt100.loc[20].tolist()[0])

            if len(data_Pt100) > 21:
                if type(data_Pt100.loc[21].tolist()[0]) == str:
                    self.entry22.insert(0, data_Pt100.loc[21].tolist()[0])

            if len(data_Pt100) > 22:
                if type(data_Pt100.loc[22].tolist()[0]) == str:
                    self.entry23.insert(0, data_Pt100.loc[22].tolist()[0])

            if len(data_Pt100) > 23:
                if type(data_Pt100.loc[23].tolist()[0]) == str:
                    self.entry24.insert(0, data_Pt100.loc[23].tolist()[0])

            if len(data_Pt100) > 24:
                if type(data_Pt100.loc[24].tolist()[0]) == str:
                    self.entry25.insert(0, data_Pt100.loc[24].tolist()[0])

            if len(data_Pt100) > 25:
                if type(data_Pt100.loc[25].tolist()[0]) == str:
                    self.entry26.insert(0, data_Pt100.loc[25].tolist()[0])

        def save(self):
            data = pandas.DataFrame(index=(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,
                                           22, 23, 24, 25), columns=([0]), data="")

            data.at[0, 0] = self.entry1.get()
            data.at[1, 0] = self.entry2.get()
            data.at[2, 0] = self.entry3.get()
            data.at[3, 0] = self.entry4.get()
            data.at[4, 0] = self.entry5.get()
            data.at[5, 0] = self.entry6.get()
            data.at[6, 0] = self.entry7.get()
            data.at[7, 0] = self.entry8.get()
            data.at[8, 0] = self.entry9.get()
            data.at[9, 0] = self.entry10.get()
            data.at[10, 0] = self.entry11.get()
            data.at[11, 0] = self.entry12.get()
            data.at[12, 0] = self.entry13.get()
            data.at[13, 0] = self.entry14.get()
            data.at[14, 0] = self.entry15.get()
            data.at[15, 0] = self.entry16.get()
            data.at[16, 0] = self.entry17.get()
            data.at[17, 0] = self.entry18.get()
            data.at[18, 0] = self.entry19.get()
            data.at[19, 0] = self.entry20.get()
            data.at[20, 0] = self.entry21.get()
            data.at[21, 0] = self.entry22.get()
            data.at[22, 0] = self.entry23.get()
            data.at[23, 0] = self.entry24.get()
            data.at[24, 0] = self.entry25.get()
            data.at[25, 0] = self.entry26.get()

            Pt100_calibration_values = ['-79.89', '-59.92', '-49.91', '-39.93', '-19.92', '-9.95', '0.03',
                                        '10.08', '20.08', '25.08', '30.05', '40.08', '50.08', '60.08', '70.08', '80.08',
                                        '90.07',
                                        '100.09', '120.09', '140.08', '150.06', '160.09', '180.07', '200.05', '250.02',
                                        '300.09']

            i = 0
            save = True
            while i < len(data):
                if data.loc[i, 0] == "":
                    save = False
                    messagebox.showinfo("Error",
                                        "Measurement for calibration value " + Pt100_calibration_values[i] + " missing")
                if data.loc[i, 0] != "":
                    if (float(data.loc[i, 0].replace(",", ".")) > (float(Pt100_calibration_values[i]) + 0.1)):
                        save = False
                        messagebox.showinfo("Error", "Measurement for calibration value " + Pt100_calibration_values[
                            i] + " out of spec")
                    if (float(data.loc[i, 0].replace(",", ".")) < (float(Pt100_calibration_values[i]) - 0.1)):
                        save = False
                        messagebox.showinfo("Error", "Measurement for calibration value " + Pt100_calibration_values[
                            i] + " out of spec")
                i = i + 1

            yes_complete = False
            if self.checkComplete_var.get() == 1:
                yes_complete = True

            if self.checkIgnore_var.get() == 1:
                save = True

            if save == True:
                self.lf = self.controller.get_page("StartPage").choice_var.get()
                last4 = str(self.controller.last4)
                path1 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + str(
                    self.lf) + "/data/Pt100.xlsx"
                path2 = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + str(
                    self.lf) + "/data/Pt100.xlsx"

                writer1 = ExcelWriter(path1)
                writer2 = ExcelWriter(path2)
                data.to_excel(writer1, 'Sheet', index=False, header=None)
                data.to_excel(writer2, 'Sheet', index=False, header=None)
                try:
                    writer1.save()
                except:
                    pass
                try:
                    writer2.save()
                except:
                    pass

                if yes_complete == True:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 15] = '1'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                if yes_complete == False:
                    master_list = pandas.read_excel('X://ERSTools/EndtestData/active_list.xlsx', sheet_name='Sheet1',
                                                    dtype=str)
                    active_index = master_list[master_list.loc[:, 0] == self.lf].index.tolist()[0]
                    master_list.iloc[int(active_index), 15] = '0'
                    writer = ExcelWriter('X://ERSTools/EndtestData/active_list.xlsx')
                    master_list.to_excel(writer, 'Sheet1', index=None)
                    writer.save()
                    self.controller.update_manager(self.lf)
                messagebox.showinfo("Success", "Save Successful")
            else:
                messagebox.showinfo("Error", "Save Not Successful")

        def verify(self):
            x = self.controller.return_type()
            y = type(x)
            print(x)
            print(y)

    if __name__ == "__main__":
        app = SampleApp()
        app.mainloop()

except BaseException:
    import sys
    print(sys.exc_info()[0])
    import traceback
    print(traceback.format_exc())

finally:
    print("Drücken Sie zum Verlassen die Eingabetaste....")
    input()