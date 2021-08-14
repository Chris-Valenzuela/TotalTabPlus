import pandas as pd
import re
from styles import bough
from datetime import datetime

# =============================================================================
# Initialize filenames, Globals, TotalTabsDictionary
# =============================================================================


name = 'R201857 ALL UNW Banner1'
filename = 'datasets/' + name + '.xlsx'
output = 'datasets/' + name + '_TotalTabsPlus' + '.xlsx'
dirpath = 'C:/Users/Chris Valenzuela/Desktop/Programming/TotalTabsPlus/'

xls = pd.ExcelFile(filename)

global start, end, firstworksheet, lastworksheet
firstworksheet = 2
lastworksheet = None

start = 5 
end = 2
TotalTabsDictionary = {}
TotalTabsDictionary['Table'] = {}
TotalTabsDictionary['Question'] = {} 
TotalTabsDictionary['Stub'] = {}
TotalTabsDictionary['StubData'] = {}
TotalTabsDictionary['RowStubData'] = {}
TotalTabsDictionary['TableLink'] = {}
TotalTabsDictionary['Banner'] = []
TotalTabsDictionary['BannerLetter'] = []



def aggr(xls):

    # =========================================================================
    # Looping through each Table and grabbing Table Link
    # Adding Dict to key "TableLink". Where key = "Table Number"
    # =========================================================================
    for indexsheet in xls.sheet_names[0:1]:
        df = pd.read_excel(filename, sheet_name = indexsheet)
        df = df.loc[4:].copy()
        for o, links in enumerate(df['Client: ']):
            o += 1

            # =================================================================
            # skipping these tables because they have numerics/duplicate rows
            # using "skip_tabs" function in bough file 
            # =================================================================
            if bough.skip_tabs(o, '56, 113, 114, 126, 127, 158, 159, 187, 188, 189, 190, 203, 204, 205, 206'):
                TotalTabsDictionary['TableLink'][o] = links
    


    # =========================================================================
    # Looping through each table and initializing TotalTabsDictionary per tab
    # =========================================================================
    for i, sheet_name in enumerate(xls.sheet_names[firstworksheet:lastworksheet]):

        # =====================================================================
        # each sheet = df
        # returns two arrays: Banner and Letter representing each banner
        # TODO what if there is no letter below the banner?
        # TODO might be better to make a dict of arrays so each tab has 
        # its own stat test ?
        # =====================================================================
        df = pd.read_excel(filename, sheet_name = sheet_name)
        TotalTabsDictionary['Banner'], TotalTabsDictionary['BannerLetter'] = bough.rowaggregator(df, start)

        # =====================================================================
        # TableNumber - search for a number in the worksheet if it exists 
        # then we consider it a table
        # =====================================================================
        TableNumber = re.search(r"\d", sheet_name)
        if TableNumber and bough.skip_tabs(sheet_name, 'T56, T113, T114, T126, T127, T158, T159, T187, T188, T189, T190, T203, T204, T205, T206'):
            
            # =================================================================
            # TableNumber - initialize, each table gets key, value
            # =================================================================
            TableNumber = sheet_name[TableNumber.start():]        
            TotalTabsDictionary['Table'][str(TableNumber)] = TableNumber

            sheetdf = df.copy()
        
            # =================================================================
            # Question - initialize, each Question gets key = table number, 
            # value = Question name 
            # =================================================================
            title = sheetdf.loc[1][0]
            title = title.split(" ")[0]
            TotalTabsDictionary['Question'][str(TableNumber)] = title
        
            # =================================================================
            # Grab first col, drop all NaN. To loop through stubs only 
            # TODO we might be able to simplify this
            # =================================================================
            columnsheetdf = sheetdf['Unnamed: 0'].copy().dropna(how='all')
            totalrowcount = len(columnsheetdf.index)
            
            
            # =================================================================
            # Initialize 'Stub'. Each table gets its own array of stub names
            # =================================================================
            TotalTabsDictionary['Stub'][str(TableNumber)] = []
            for j, row in enumerate(columnsheetdf):
                
                # =============================================================
                # pinpointing what row stub we want then appending to an array
                # TotalTabsDictionary['Stub']['TableNumber'] = []
                # =============================================================
                if j > start - 1 and j < totalrowcount - end:
                    TotalTabsDictionary['Stub'][str(TableNumber)].append(row)
            
            # StubData initialize
            # create a dict in a dict that holds an array for each column of the tableset/bannerpoint
            TotalTabsDictionary['StubData'][str(TableNumber)] = {}        
            for bannerpoint in TotalTabsDictionary['Banner']:
                TotalTabsDictionary['StubData'][str(TableNumber)][str(bannerpoint)] = []       
            
            newdf = df.loc[start:][1:].copy()
            newdf.set_index('Unnamed: 0', inplace= True)
            newdf.dropna(inplace = True, how='all')
            
            # rename the columns of the new dataframe 
            for x in range(0, len(newdf.columns)):
                newdf = newdf.rename(columns={f'Unnamed: {x + 1}': TotalTabsDictionary['Banner'][x]})

            for cols in newdf.columns:
                # list comprehension 
                #solution: if you want to get percentages then add the "shift" one row bottom in between cols and bannerpoint. to just get counts remove the shift
                # "{:.2%}".format(a_number) - this is a way to transform it into a percentage but its a problem for empty 
                
                # todo this needs to be fixed its way to complicated. not your finest moment 
    #             bannerpointdata = [ "{:.2%}".format(float(newdf[cols].shift(-1)[str(bannerpoint)])) + " " + str((newdf[cols].shift(-2)[str(bannerpoint)]) )  if str(newdf[cols].shift(-2)[str(bannerpoint)]) != "nan" and (newdf[cols].shift(-1)[str(bannerpoint)] != "-" and newdf[cols].shift(-1)[str(bannerpoint)] != "*" )  else "{:.2%}".format(float(newdf[cols].shift(-1)[str(bannerpoint)])) if (newdf[cols].shift(-1)[str(bannerpoint)] != "-" and newdf[cols].shift(-1)[str(bannerpoint)] != "*" ) else newdf[cols].shift(-1)[str(bannerpoint)] for bannerpoint in TotalTabsDictionary['Stub'][str(TableNumber)]  ]
    #             'ARCHIVE' error: line 253 335 346
                bannerpointdata = []
    #             print(str(TableNumber))
                for bannerpoint in TotalTabsDictionary['Stub'][str(TableNumber)]:
                    if str(bannerpoint) != 'Base' and str(bannerpoint) != 'Unweighted Base' and str(bannerpoint) != 'Mean':
                        if  type(bannerpoint) ==  str:
                            bannerpoint = str(bannerpoint)
                        elif type(bannerpoint) ==  int:
                            bannerpoint = bannerpoint

                        if (str(newdf[cols].shift(-2)[bannerpoint]).lower() != "nan" and str(newdf[cols].shift(-2)[bannerpoint]).isdigit() == False) and (str(newdf[cols].shift(-1)[bannerpoint]) != "*" and str(newdf[cols].shift(-1)[bannerpoint]) != "-"):
                            # fix for multiple instances of rows (i.e series)
                            # print(str(newdf[cols].shift(-2)[bannerpoint][0]))

                            bannerpointdata.append("{:.2%}".format(float(newdf[cols].shift(-1)[bannerpoint])) + " " + str((newdf[cols].shift(-2)[bannerpoint]) ) )
    #                         print(bannerpoint)
    #                         print("0 below: " + str(newdf[cols][bannerpoint]))
    #                         print("1 below: " + str(newdf[cols].shift(periods = -1)[bannerpoint]))
    #                         print("2 below: " + str(newdf[cols].shift(periods = -2)[bannerpoint]))
                        elif str(newdf[cols].shift(-2)[bannerpoint]).lower() != "nan" and (str(newdf[cols].shift(-1)[bannerpoint]) == "*" and str(newdf[cols].shift(-1)[bannerpoint]) == "-"):
                            bannerpointdata.append(str(newdf[cols].shift(-1)[bannerpoint]) + " " + str((newdf[cols].shift(-2)[bannerpoint]) ) )
    #                         print(bannerpoint)
    #                         print("0 below: " + str(newdf[cols][bannerpoint]))
    #                         print("1 below: " + str(newdf[cols].shift(periods = -1)[bannerpoint]))

                        elif (str(newdf[cols].shift(-1)[bannerpoint]) != "*" and str(newdf[cols].shift(-1)[bannerpoint]) != "-"):
                            bannerpointdata.append("{:.2%}".format(float(newdf[cols].shift(-1)[bannerpoint])) )
                            # print(bannerpoint)
                            # print("0 below: " + str(newdf[cols][bannerpoint]))
                            # print("1 below: " + str(newdf[cols].shift(periods = -1)[bannerpoint]))

                        elif (str(newdf[cols].shift(-1)[bannerpoint]) == "*" or str(newdf[cols].shift(-1)[bannerpoint]) == "-"):
                            bannerpointdata.append(newdf[cols].shift(-1)[bannerpoint])
    #                         print(bannerpoint)
    #                         print("0 below: " + str(newdf[cols][bannerpoint]))
    #                         print("1 below: " + str(newdf[cols].shift(periods = -1)[bannerpoint]))

                TotalTabsDictionary['StubData'][str(TableNumber)][str(cols)] = bannerpointdata
            
            
            # RowStubData initialize
            TotalTabsDictionary['RowStubData'][str(TableNumber)] = {}
            for j, rowineachtable in enumerate(columnsheetdf):
                if j > start - 1 and j < totalrowcount - end:
                    TotalTabsDictionary['RowStubData'][str(TableNumber)][str(rowineachtable)] = []
        
            for  rows in (newdf.index):
                if str(rows) != "nan" and str(rows) != 'Base' and str(rows) != 'Unweighted Base' and str(rows) != 'Mean':
                    for cols in newdf.columns:
                        if (newdf[cols].shift(-1)[rows] != "-" and newdf[cols].shift(-1)[rows] != "*"):
                            if str(newdf[cols].shift(-2)[rows]) != "nan":
                                
                                rowdatawithstats = str("{:.2%}".format(newdf[cols].shift(-1)[rows])) + " " + str((newdf[cols].shift(-2)[rows]) )
                                TotalTabsDictionary['RowStubData'][str(TableNumber)][str(rows)].append(rowdatawithstats)
                            else:
    #                             print(type(rows))
    #                             print(rows)
    #                             print(TotalTabsDictionary['RowStubData'][str(TableNumber)])
                                TotalTabsDictionary['RowStubData'][str(TableNumber)][str(rows)].append( "{:.2%}".format(newdf[cols].shift(-1)[rows]))                      
                        else:
                            TotalTabsDictionary['RowStubData'][str(TableNumber)][str(rows)].append(newdf[cols].shift(-1)[rows])



