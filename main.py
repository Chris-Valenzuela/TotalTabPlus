import pandas as pd
from aggron import *
from scraper import *
from style import *

from datetime import datetime
# from aggron import TotalTabsDictionary
# import os 
# dir_path = os.path.dirname(os.path.realpath(__file__))
# print(dir_path)

totaltabsdf = {}
newcolumns = ['Table', 'Question', 'Stub']

# =============================================================================
# Grabbing the stat testing from the first table "T1". Adding to array stattest
# =============================================================================
statsdf = pd.read_excel(filename, sheet_name = 'T1')
statsdf = statsdf.loc[4:].copy()
stattest = []
for statistics in statsdf['Unnamed: 0']:
    
    # =========================================================================
    # Looking for the Stat test row
    # =========================================================================
    if str(statistics).find('Statistics:') != -1:
        stattest = statistics.split(":")[3].split(",")[:-1]

        # =====================================================================
        # sometimes a programmer can enter ',' at the end of T_Banners and 
        # sometimes they wont. This is a work around for both case
        # =====================================================================
        if str(statistics.split(":")[3].split(",")[-1]).find('/') != -1:
            lastitersplit = statistics.split(":")[3].split(",")[-1].split(" ")
            for iter, space in enumerate(lastitersplit):
                if space.find('/') != -1:
                    stattest.append(lastitersplit[iter])
        

def main():

    # =========================================================================
    # Runs Aggron file which aggregates data from all tabs
    # =========================================================================
    print('============ Starting Data Aggron File =================')
    start_time_aggron = datetime.utcnow()
    aggr(xls)
    end_time_aggron     = datetime.utcnow()
    elapsed_time_aggron = end_time_aggron - start_time_aggron
    print("Elapsed Aggron time: " + str(elapsed_time_aggron))
    print('============ Completed Data Aggron File =================\n')

    # =========================================================================
    # Runs Scraper file which organizes data per powerpoint
    # =========================================================================
    
    print('============ Startinng Scraping File =================')
    start_time_scraper = datetime.utcnow()
    totaltabsdfnew = scraper(totaltabsdf, newcolumns, stattest)
    end_time_scraper     = datetime.utcnow()
    elapsed_time_scraper = end_time_scraper - start_time_scraper
    print("Elapsed Scraper time: " + str(elapsed_time_scraper))
    print('============ Completed Scraping File =================\n')
    
    # =========================================================================
    # Runs Styler file which does hyperlinks, openpyxl etc...
    # =========================================================================

    print('============ Starting Makeup File =================')
    start_time_makeup = datetime.utcnow()
    makeup(totaltabsdfnew, newcolumns)
    end_time_makeup     = datetime.utcnow()
    elapsed_time_makeup = end_time_makeup - start_time_makeup
    print("Elapsed Makeup time: " + str(elapsed_time_makeup))
    print('============ Completed Makeup File =================\n')

    
        
if __name__ == "__main__":
    start_time = datetime.utcnow()
    main()
    end_time     = datetime.utcnow()
    elapsed_time = end_time - start_time
    print("Elapsed Total time: " + str(elapsed_time))
