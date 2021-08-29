import os
import pdfrw
import pandas
import numpy
import sys
import datetime

error = ''

try:
    # Prompt User for Entry #
    lauf_num = input("Laufnummer Eingeben:  ")
    name = input("Deinen Namen Eingeben:  ")

    # Define paths #
    snxls_path = '//fileserver/produktion/Endtest/30_Seriennummern/Seriennummern.xls'  ## path of seriennummern spreadsheet
    template_path = '//fileserver/alle/ERSTools/Endtest Document Generator/warranty.pdf'
    notchiller_path = '//fileserver/alle/ERSTools/Endtest Document Generator/notchiller.xls'

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
            if str(lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                error = 'Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste'
                x = 1 / 0
            else:
                D = D.loc[str(lauf_num)]
        elif len(lauf_num) != 6:
            if int(lauf_num) not in D.index.values:  # Is the Laufnummer contained in the Serien_nummern spreadsheet
                print('Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste')
                error = 'Error: Die Laufnummer ist nicht in der Seriennummer XLS-Liste'
                x = 1 / 0
            else:
                D = D.loc[int(lauf_num)]

    # cancel script if lauf number is listed more than once in the excel file
    if len(D) < 28:  # we say less than 28, because if only 1 exits then it has length 28, but if 2 exist then length 2
        print("Error: Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt")
        error = "Error: Die Laufnummer wird in der Excel-Datei doppelt oder mehrmals aufgeführt"
        x = 1 / 0

    ## is There really a chiller? ##
    yeschiller = False
    i = 0
    count = 0
    if D['Chiller'] == D['Chiller']:
        while i < len(notchiller_df):
            if notchiller_df.at[i, 'st'] in D.at['Chiller']:
                count = count + 1
            else:
                count = count
            i = i + 1
        if count == 0:
            yeschiller = True

    #       Index of D:  'Serien Nr.', 'SN', 'SN Kompl.', 'Quartal', 'Jahr',
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
    if lauf_num[0] != 'L':
        newdir = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_ERS' + lauf_num
    else:
        newdir = "//fileserver/produktion/Endtest/10_Dokumente/" + last4 + '_' + lauf_num

    os.mkdir(newdir)
    write_path = newdir + "/Warranty_Sheet_ERS"
    os.mkdir(newdir + "/Bilder Auslieferung")
    os.mkdir(newdir + "/Bilder Tempverteilung")

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
        if yesbooster == False and D['Chiller'].__contains__('CH20') == True:
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
                i = 1

            if t_sz[i].isdigit() == True:
                t_sz = t_sz[0:i+1]  ##   get rid of all characters after the number ends
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

    E3 = pandas.DataFrame(index=temp_array, columns=('st', 'or', 'sn', 'type'), data="")  # create new data array

    E = E.append(E3)  # append original data array with new data array to restore length = 5

    ##### Strings Used in Write Fillable PDF function (don't touch) #####
    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_VAL_KEY = '/V'
    ANNOT_RECT_KEY = '/Rect'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'


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

    sys = 'ERS® AC3:'
    if prime == True:
        sys = 'ERS® AirCool® PRIME:'

    ############ Writes PDF Warranty for chiller ONLY (or all if NO chiller) and opens #############

    data_dict = {  ### Assign values to all fillable PDF fields here
        'sys': sys,
        'title': title,
        'tr': ' ',  ## INCOMPLETE
        't_sz': t_sz,
        'or_0': '# ' + lauf_num,
        'or_1': E.at['1', 'or'],
    ## in ['1', 'or'], '1' is a string instead of integer because we conocated E earlier to return length = 5
        'or_2': E.at['2', 'or'],
        'or_3': E.at['3', 'or'],
        'or_4': E.at['4', 'or'],
        'or_5': E.at['5', 'or'],
        'sn_1': E.at['1', 'sn'],
        'sn_2': E.at['2', 'sn'],
        'sn_3': E.at['3', 'sn'],
        'sn_4': E.at['4', 'sn'],
        'sn_5': E.at['5', 'sn'],
        'st_1': E.at['1', 'st'],
        'st_2': E.at['2', 'st'],
        'st_3': E.at['3', 'st'],
        'st_4': E.at['4', 'st'],
        'st_5': E.at['5', 'st'],
    }

    morethanchiller = False
    if yeschuck == True or yescontroller == True or yesopt1 == True or yesopt2 == True:
        morethanchiller = True
        write_fillable_pdf(template_path, write_path + lauf_num + '.pdf', data_dict)

    # opens file in browser
    # os.startfile(write_path + lauf_num + '.pdf')

    #######   Writes 2nd Warranty PDF *IF* there is a chiller in the order #####

    if yeschiller == True:  ### if chiller exists, then E2 is changed from integer to a dataframe

        data_dict = {  # Assign values to all fillable PDF fields here
            'sys': sys,
            'title': E2.at['1', 'st'],
            'tr': ' ',  ## INCOMPLETE
            't_sz': t_sz,
            'or_0': '# ' + lauf_num,
            'or_1': E2.at['1', 'or'],
            'sn_1': E2.at['1', 'sn'],
            'st_1': E2.at['1', 'st'],
        }

        if morethanchiller == True:
            write_fillable_pdf(template_path, write_path + lauf_num + '-2.pdf', data_dict)
        else:
            write_fillable_pdf(template_path, write_path + lauf_num + '.pdf', data_dict)
        # opens file in browser
        # os.startfile(write_path + lauf_num + '-2.pdf')

    ###################################################
    ###################################################
    ###################################################
    ###################################################
    ############ Begin Versand section ##################

    template_path2 = '//fileserver/alle/ERSTools/Endtest Document Generator/versand.pdf'
    write_path2 = newdir + "/Versandanzeige_ERS"
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

    # Create Kiste reference #
    kiste = pandas.DataFrame({'index': ['3', '6', '7', '6.2', '8.1', '23', ''],
                              'kiste': ['3', '6', '7', '6.2', '8.1', '23', ''],
                              'gw': [20, 20, 40, 30, 10, 65, ''],
                              'l': [63, 74, 75, 88, 52, 90, ''],
                              'b': [52, 44, 68, 68, 52, 62, ''],
                              'h': [61, 59, 76, 74, 29, 123, ''], })
    kiste = kiste.set_index('index')

    # determine how many creates and creates list for crates#
    num_crates = 0
    if yeschiller == True:
        num_crates = 1
    if (yescontroller + yeschuck + yesopt1 + yesopt2) > 0:
        num_crates = num_crates + 1

    crates_ls = [''] * num_crates
    gewicht_ls = [''] * num_crates  # also creating this list to keep track of weight

    ####### Determines which types of crates used #######
    count = 0
    if yeschiller == True:
        count = 1
        sep = ' '
        ch_num = D['Chiller'].split(sep, 1)[0]  # splits string at space and keeps first half
        ch_num = ch_num.lstrip('aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ. ')

        i = 0
        while i < len(ch_num):
            if ch_num[i].isdigit() == False:
                break
            else:
                i = i + 1

        ch_num = ch_num[0:i]



        if next((True for item in ref.index if item == ch_num), False) == False:
            crates_ls[0] = '23'
            gewicht_ls[0] = '160'
        else:
            crates_ls[0] = ref.at[ref.loc[ch_num].name, 'verpackung']
            if crates_ls[0] == '7':
                gewicht_ls[0] = '100'

            else:
                gewicht_ls[0] = '160'

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

    ## Fill in crates_ls and gewicht_ls
    if yescontroller == True and yeschuck == False and yesvg5 == False:  ## purposefully left RSI box out of here so that it can be entered manually rather than having error
        crates_ls[count] = ref_controllers.at[controller_index, 'verpackung']
        gewicht_ls[count] = ref_controllers.at[controller_index, 'gewicht']
        count = count + 1

    if yescontroller == False and yeschuck == True:
        if t_sz != '25mm' and t_sz != '':
            temp_df = ref_chucks.loc[t_sz]
            crates_ls[count] = temp_df.at['verpackung']
            gewicht_ls[count] = temp_df.at['gewicht']
            count = count + 1
        elif t_sz == '25mm':
            crates_ls[count] = ''
            gewicht_ls[count] = ''
            count = count + 1

    if yescontroller == True and yeschuck == True and yesrsi == False and yesvg5 == False:
        if t_sz != '25mm' and t_sz != '':
            temp_df = ref_chucks.loc[t_sz]
            crates_ls[count] = temp_df.at['system verpackung']
            gewicht_ls[count] = temp_df.at['system gewicht']
            count = count + 1
        elif t_sz == '25mm':
            crates_ls[count] = ref_controllers.at[controller_index, 'verpackung']
            gewicht_ls[count] = ref_controllers.at[controller_index, 'gewicht']
            count = count + 1

    if yescontroller == True and yeschuck == True and yesrsi == True and t_sz != '25mm' and t_sz != '':
        temp_df = ref_chucks.loc[t_sz]
        crates_ls[count] = temp_df.at['system verpackung']
        gewicht_ls[count] = temp_df.at['system gewicht']

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
        E.at[str(i), 'sn'] = E.at[str(i), 'sn'].lstrip('aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ. ')  #
        i = i + 1

    ## Sum gewicht of crates ##
    i = 0
    gw = 0
    while i < num_crates and gewicht_ls[i] != '':
        if yeschuck == True or yeschiller == True or (yescontroller == True and yesrsi == False):
            gw = gw + int(gewicht_ls[i])
        i = i + 1
    if int(gw) > 50:
        gw = int(numpy.ceil(int(gw) / 5) * 5)  # round up to nearest 5

    if yesvg5 == True:  ## If it is a vg5 controller, leave weight blank
        gw = ''

    if yeschuck == False and yeschiller == False and ((
                                                              yescontroller == True and yesrsi == True) or yescontroller == False):  # if only an option (no chuck, controller, or chiller), leave blank
        gw = ''
        i = 0
        while i < len(crates_ls):
            crates_ls[i] = ''
            i = i + 1

    if yeschuck == True and t_sz == '':
        gw = ''

    # Make length of gewicht_ls and crates_ls 2 for entry into the data dictionary #
    if len(gewicht_ls) == 1:
        gewicht_ls = gewicht_ls + ['']
    if len(crates_ls) == 1:
        crates_ls = crates_ls + ['']

    # determine verpackung nummber #
    vrn = ''
    if num_crates > 1:
        vrn = crates_ls[0]
        i = 1
        while i < num_crates:
            vrn = vrn + ' / ' + crates_ls[i]
            i = i + 1
    else:
        vrn = crates_ls[0]

    # Determine date
    dt = str(D['best. Liefertermin'])[0:10]
    yyyy = dt[0:4]
    mm = dt[5:7]
    dd = dt[8:10]
    dt = dd + '/' + mm + '/' + yyyy

    # what is today? #
    today = str(datetime.datetime.today()).split()[0]
    yyyy = today[0:4]
    mm = today[5:7]
    dd = today[8:10]
    today = dd + '/' + mm + '/' + yyyy

    # modifies type column text to include "- serien #"
    # makes it say PV115P or SP41P PV110 SP -serien #, etc
    i = 1
    while i < 6 and E.at[str(i), 'type'] != '':
        if E.at[str(i), 'st'].__contains__('PV'):
            E.at[str(i), 'type'] = E.at[str(i), 'st']

        E.at[str(i), 'type'] = E.at[str(i), 'type'] + '-Serien #  '
        i = i + 1

    data_dict2 = {  ### Do we need to go up to 5 here??????? !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'lf_1': lauf_num,
        'st_1': E.at['1', 'st'],
        'st_2': E.at['2', 'st'],
        'st_3': E.at['3', 'st'],
        'st_4': E.at['4', 'st'],
        'st_5': E.at['5', 'st'],
        'dt_a': dt,
        'today': today,
        'name': name,
        'or_1': E.at['1', 'type'] + E.at['1', 'sn'],
        'or_2': E.at['2', 'type'] + E.at['2', 'sn'],
        'or_3': E.at['3', 'type'] + E.at['3', 'sn'],
        'or_4': E.at['4', 'type'] + E.at['4', 'sn'],
        'or_5': E.at['5', 'type'] + E.at['5', 'sn'],
        # 'in_1': num_crates, ## what is difference between in and an???
        'an_1': num_crates,
        'vrn': vrn,
        'gw': gw,
        'l_1': kiste.at[crates_ls[0], 'l'],
        'l_2': kiste.at[crates_ls[1], 'l'],
        'b_1': kiste.at[crates_ls[0], 'b'],
        'b_2': kiste.at[crates_ls[1], 'b'],
        'h_1': kiste.at[crates_ls[0], 'h'],
        'h_2': kiste.at[crates_ls[1], 'h'],

    }

    write_fillable_pdf(template_path2, write_path2 + lauf_num + '.pdf', data_dict2)

    # opens file in browser
    # os.startfile(write_path2 + lauf_num + '.pdf')

    ###### Write Planarity, Temperature Uniformity Protocol, and PT100 calibration protocol #####
    ################################################################################
    ################################################################################
    ################################################################################
    ################################################################################
    template_path3 = '//fileserver/alle/ERSTools/Endtest Document Generator/planarity.pdf'
    write_path3 = newdir + "/Planarity_Protocol_ERS"
    template_path4 = '//fileserver/alle/ERSTools/Endtest Document Generator/temperature.pdf'
    write_path4 = newdir + "/Temperature_Uniformity_ERS"
    template_path5 = '//fileserver/alle/ERSTools/Endtest Document Generator/Pt100abgl 273807_Kunde.pdf'
    write_path5 = newdir + "/Pt100abgl 273807_Kunde_Calibration_ERS"
    template_path6 = '//fileserver/alle/ERSTools/Endtest Document Generator/Protokoll Final_Check Chuck.pdf'
    write_path6 = newdir + "/Protokoll_Final_Check_Chuck_ERS"
    template_path7 = '//fileserver/alle/ERSTools/Endtest Document Generator/Protokoll Final_Check System.pdf'
    write_path7 = newdir + "/Protokoll_Final_Check_System_ERS"

    # find controller name
    ctl = ''
    i = 1
    if yescontroller == True:
        while i < len(E):
            if E.at[str(i), 'type'] == 'Controller-Serien #  ':
                break
            else:
                i = i + 1

    if yescontroller == False and yeschiller == True:
        i = 1
        while i < len(E):
            if E.at[str(i), 'type'] == 'Chiller-Serien #  ':
                break
            else:
                i = i + 1

    data_dict = {
        'ctl': E.at[str(i), 'st'],
        'ctl_sn': E.at[str(i), 'sn'],
        'lf_1': '#' + lauf_num,
    }

    yests010 = False  # do we have a TS010 chiller with a display?
    if E.at[str(i), 'st'][0:5] == 'TS010':
        yests010 = True

    if (
            yescontroller == True and yesrsi == False) or yests010 == True:  # if its a controller or a TS010 chiller, then make the file
        write_fillable_pdf(template_path5, write_path5 + lauf_num + '.pdf', data_dict)

    # find chuck name
    chk = ''
    eq_id = ''
    i = 1
    if yeschuck == True:
        while i < len(E):
            if E.at[str(i), 'type'] == 'Chuck-Serien #  ':
                break
            else:
                i = i + 1

        if t_sz != '150mm' and t_sz != '6"':
            eq_id = 'Wafer'

    # find type of system
    sys = ''
    if yescontroller == True and yeschuck == True:
        sys = 'Chuck + Controller'
    if yeschuck == True and yescontroller == False and yests010 == True:
        sys = 'Chuck + Chiller'

    data_dict = {
        'sys': sys,
        'chk': E.at[str(i), 'st'],
        'chk_sn': E.at[str(i), 'sn'],
        'lf_1': '#' + lauf_num,
        'eq_id': eq_id,
        'ht': '',  # high temperature, not complete
        'lt': '',  # low temperature, not complete
    }

    if yeschuck == True:
        write_fillable_pdf(template_path3, write_path3 + lauf_num + '.pdf', data_dict)

        write_fillable_pdf(template_path4, write_path4 + lauf_num + '.pdf', data_dict)

        write_fillable_pdf(template_path6, write_path6 + lauf_num + '.pdf', data_dict)

        write_fillable_pdf(template_path7, write_path7 + lauf_num + '.pdf', data_dict)

    if error == '':
        with open("//fileserver/alle/ERSTools/Endtest Document Generator/error_log.txt", "a") as myfile:
            myfile.write("#################   SUCCESS   ##############################")
            myfile.write("<" + "Lauf Number = " + str(lauf_num) + ">")
            myfile.write("<" + str(datetime.datetime.today()).split()[0] + ">")

    if error != '':
        with open("//fileserver/alle/ERSTools/Endtest Document Generator/error_log.txt", "a") as myfile:
            myfile.write("#################   SUCCESS WITH ERROR   ##############################")
            if error != '':
                myfile.write("Error = " + error)
            myfile.write("<" + "Lauf Number = " + str(lauf_num) + ">")
            myfile.write("<" + str(datetime.datetime.today()).split()[0] + ">")


except BaseException:
    import sys
    print(sys.exc_info()[0])
    import traceback
    print(traceback.format_exc())
    with open("//fileserver/alle/ERSTools/Endtest Document Generator/error_log.txt", "a") as myfile:
        myfile.write("#################   RUN FAIL ERROR   ##############################")
        if error != '':
            myfile.write("Error = " + error)
        myfile.write("<" + "Lauf Number = " + str(lauf_num) + ">")
        myfile.write("<" + str(datetime.datetime.today()).split()[0] + ">")
        myfile.write(str(sys.exc_info()[0]))
        myfile.write(str(traceback.format_exc()))

finally:
    print("Drücken Sie zum Verlassen die Eingabetaste....")
    input()