import pandas as pd
import re
from styles import bough
from datetime import datetime


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


# # ================ First Table and grabbing the STAT Testing =======================
# statsdf = pd.read_excel(filename, sheet_name = 'T1')
# statsdf = statsdf.loc[4:].copy()
# stattest = []
# for statistics in statsdf['Unnamed: 0']:
#         if str(statistics).find('Statistics:') != -1:
#             stattest = statistics.split(":")[3].split(",")[:-1]
# #===================================================================================





def aggr(xls):

    # ================ Looping through each Table and grabbing Table Link =======================
    for indexsheet in xls.sheet_names[0:1]:
        df = pd.read_excel(filename, sheet_name = indexsheet)
        df = df.loc[4:].copy()
        for o, links in enumerate(df['Client: ']):
            o += 1

            # TODO 
            if str(o) != '56' and str(o) != '113' and str(o) != '114' and str(o) != '126' and str(o) != '127' and str(o) != '158' and str(o) != '159' and str(o) != '187' and str(o) != '188' and str(o) != '189' and str(o) != '190' and str(o) != '203' and str(o) != '204' and str(o) != '205' and str(o) != '206':
                TotalTabsDictionary['TableLink'][o] = links
    #==============================================================================================



    # =================== Looping through each table and initializing TotalTabsDictionary ================
    for i, sheet_name in enumerate(xls.sheet_names[firstworksheet:lastworksheet]):

        df = pd.read_excel(filename, sheet_name = sheet_name)
        TotalTabsDictionary['Banner'], TotalTabsDictionary['BannerLetter'] = bough.rowaggregator(df, start)

        ''' Table - search for a number in the worksheet if it exists then add that to the array of TotalTabsDict '''
        TableNumber = re.search(r"\d", sheet_name)
        # if TableNumber:
        # TODO 
        if TableNumber and str(sheet_name) != 'T56' and str(sheet_name) != 'T113' and str(sheet_name) != 'T114' and str(sheet_name) != 'T126' and str(sheet_name) != 'T127' and str(sheet_name) != 'T158' and str(sheet_name) != 'T159' and str(sheet_name) != 'T187' and str(sheet_name) != 'T188' and str(sheet_name) != 'T189' and str(sheet_name) != 'T190' and str(sheet_name) != 'T203' and str(sheet_name) != 'T204' and str(sheet_name) != 'T205' and str(sheet_name) != 'T206':
            
            #TableNumber initialize
            TableNumber = sheet_name[TableNumber.start():]        
            TotalTabsDictionary['Table'][str(TableNumber)] = TableNumber

            sheetdf = df.copy()
        
            # Question initialize
            title = sheetdf.loc[1][0]
            title = title.split(" ")[0]
            TotalTabsDictionary['Question'][str(TableNumber)] = title
        
        
            # grab the first column and drop all the na so we can loop through the stubs only
            columnsheetdf = sheetdf['Unnamed: 0'].copy().dropna(how='all')
            totalrowcount = len(columnsheetdf.index)
            
            # Stub intialize
            TotalTabsDictionary['Stub'][str(TableNumber)] = []
            for j, row in enumerate(columnsheetdf):
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



