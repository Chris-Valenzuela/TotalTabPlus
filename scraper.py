from aggron import *

def scraper(totaltabsdf, newcolumns, stattest):
    
    for key in TotalTabsDictionary:
        if key == 'Table' or key == 'Question' or key == 'Stub':
            totaltabsdf[str(key)] = []
            
    for key in TotalTabsDictionary['Banner']:
        totaltabsdf[str(key)] = []

        
    for key in TotalTabsDictionary:
        if key == 'TableLink':
            totaltabsdf[str(key)] = []
        

    #maybe theres a better way to do this but right now im looping through Stub because that what determines the amount of rows for that table 
    for key in TotalTabsDictionary['Stub']:
        for stub in TotalTabsDictionary['Stub'][key]:
            if str(stub) != 'Base' and str(stub) != 'Unweighted Base' and str(stub) != 'Mean':
                totaltabsdf['Table'].append(key)
                totaltabsdf['Question'].append(TotalTabsDictionary['Question'][key])
                totaltabsdf['TableLink'].append(TotalTabsDictionary['TableLink'][int(key)])
        # make sure to include invidual bases
        [totaltabsdf['Stub'].append(x) for x in TotalTabsDictionary['Stub'][key] if (str(x) != 'Base' and str(x) != 'Unweighted Base' and str(x) != 'Mean')]
        for bannerpoint in TotalTabsDictionary['Banner']:
            [ totaltabsdf[bannerpoint].append(x) for x in TotalTabsDictionary['StubData'][key][bannerpoint] ]
        
    # ================== Creating Banner point stuff =======================
        
    colordict = {}
    bannerlettername = []

    for p, bannerletter in enumerate(TotalTabsDictionary['BannerLetter']):
        colordict[bannerletter] = TotalTabsDictionary['Banner'][p]

    
    for batch in stattest:
        Letter = batch.strip().split('/')
        for q, letter in enumerate(Letter):
            Letter[q] = colordict[letter]

        bannerlettername.append(Letter)
        
    
    ''' have to rename the index so i can use the index position to match which data i want from row data '''
    for banner in bannerlettername:
        for r, bannerpoint in enumerate(banner):
            banner[r] = TotalTabsDictionary['Banner'].index(bannerpoint)


    for i in range(1, len(bannerlettername) + 1):
        totaltabsdf["Max Diff " + str(i)] = []
        
    
    for table in TotalTabsDictionary['RowStubData']:
        for stub in TotalTabsDictionary['RowStubData'][table]: 
    #         minimum = TotalTabsDictionary['RowStubData'][table][stub][banner[0]][:-1]
            for x, banner in enumerate(bannerlettername):
                
                if str(stub) != 'Base' and str(stub) != 'Unweighted Base' and str(stub) != 'Mean':
                    x += 1
                    max_value = None
                    cleancell = bough.thedeleter(str(TotalTabsDictionary['RowStubData'][table][stub][banner[0]].strip('%')))
                    if cleancell != '-' and cleancell != '*' and cleancell != '':
                        minimum = cleancell
                    else:
                        minimum = 0

                    
                    for bannerpoint in banner:

                        cleancell2 = bough.thedeleter(str(TotalTabsDictionary['RowStubData'][table][stub][bannerpoint]))
                        
                        
                        if cleancell2 != "-" and cleancell2 != "*" and cleancell2 != '':
                            
                            cleancell3 = bough.thedeleter(TotalTabsDictionary['RowStubData'][table][stub][bannerpoint].strip("%"))
                            ''' Maximum '''
                            if (max_value is None or float(cleancell3) > max_value):
                                max_value = float(cleancell3)
                                
                            ''' Minimum '''
                            if float(minimum) > float(cleancell3):
                                minimum = float(cleancell3)
                            

                    if str(max_value) != 'None':
                        statdiff = float(max_value) - float(minimum)
                        totaltabsdf["Max Diff " + str(x)].append(statdiff)
                    else:
                        # toggle this is we want empty colors (i.e no max diff)
                        totaltabsdf["Max Diff " + str(x)].append(max_value)
                        # toggle this is we dont want empty colors (i.e max diff = 0 at lowest)
                        # totaltabsdf["Max Diff " + str(x)].append(0)
                    
                
                            

    # newcolumns = ['Table', 'Question', 'Stub']
    beginnumber = len(newcolumns)

    for bannerpoint in TotalTabsDictionary['Banner']:
        newcolumns.append(bannerpoint)



    for x, statbatch in enumerate(bannerlettername):
        
        position = max(statbatch) + beginnumber + 1 + x
        newcolumns.insert(position, "Max Diff " + str(x + 1)) 


    # check if arrays are the same 
    # for z in totaltabsdf:
    #     print((z))
    #     print(len(totaltabsdf[z]))
    # print(totaltabsdf['TableLink'])
    # print(totaltabsdf['Total'])

    newcolumns.append('TableLink')
    totaltabsdf = pd.DataFrame(data = totaltabsdf, columns = newcolumns)
    totaltabsdf.set_index('Table', inplace = True)
    
    

    
    return totaltabsdf

   
   