
from typing import List


#question: when i -> [] i get an a weird Use List[T] to indicate a list type or Union[T1, T2] to indicate a union type error. but it still runs and works properly
# not sure what the difference is 
def rowaggregator(df, start) -> List:
    rowarray = []
    rowarraystat = []
    row = df.loc[start:start+1].copy()

    row.dropna(inplace = True, how = 'all')
    row.drop(['Unnamed: 0'], axis = 1, inplace = True)
    for x in row.columns:
        banner = row[x][start]
        stattest = row[x][start+1]
        rowarray.append(banner)
        if str(stattest) == 'nan':
            rowarraystat.append(".")
        else:
            rowarraystat.append(stattest)
        
    
        
    return rowarray, rowarraystat


def thedeleter(string):
    string = str(string)
    delstring = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ %*-"
    for x in delstring:
        string = string.replace(x, "")

    return string


def letterfinder(sting):
    sting = str(sting)
    greenred = False
    letterstring = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ+~-*" 
    for x in letterstring:
        #exists 
        if sting.find(x) != -1:
            greenred = True
            break
        else:
            greenred = False
    return greenred

def color_positive_green(sting):
    """
    Takes a scalar and returns a string with
    the css property `'color: green'` for positive
    strings, black otherwise.
    """
    sting = letterfinder(sting)
    if sting:
        color = 'black'
    else:
        color = 'grey'
    return 'color: %s' % color


def color_font(sting):
    """
    Takes a scalar and returns a string with
    the css property `'color: green'` for positive
    strings, black otherwise.
    """
    sting = letterfinder(sting)
    if sting:
        color = '000000'
    else:
        color = '808080'
    return color


def skip_tabs(tab ,tabdelim):
    
    tabdelim_split = tabdelim.split(',')
    
    for tabs in tabdelim_split:
        if str(tab).strip() == str(tabs.strip()):
            return False
        else:
            continue
    
    return True 