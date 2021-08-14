
import re
from datetime import datetime
import numpy as np
import pandas as pd
import networkx as nx
from text_tools import *
from io_tools import *
from tableauhyperapi import HyperProcess, Telemetry, \
    Connection, CreateMode, \
    NOT_NULLABLE, NULLABLE, SqlType, TableDefinition, \
    Inserter, \
    escape_name, escape_string_literal, \
    HyperException

# ------------------------------------------------------------------------------
# Given a list of things, are there only unique values?
# ------------------------------------------------------------------------------
def validate_unique(ilist, ignore_blanks=False):
    if ignore_blanks:
        use_list = [x for x in ilist if valid_value(x)]
    else:
        use_list = ilist

    if len(use_list) == len(set(use_list)):
        return True
    else:
        return False

# ------------------------------------------------------------------------------
# Keys and values are straight reversed
# ------------------------------------------------------------------------------
def invert_dict(d1):
    d2 = {}
    for k,v in d1.items():
        if v in d2:
            print("WARNING! Inverting dictionary item " + str(k) + ", value " + str(v) + " appears more than once and is being overwritten.")
        d2[v] = k
    return d2

# ------------------------------------------------------------------------------
# Inverts a dict with the keys stored as lists by value.
# ------------------------------------------------------------------------------
def invert_dict_as_lists(d1):
    d2 = {}
    for k,v in d1.items():
        if not v in d2:
            d2[v] = []
        d2[v].append(k)
    return d2

# ------------------------------------------------------------------------------
# Get the maximum depth of a dictionary or list
# ------------------------------------------------------------------------------
def depth(x):
    if type(x) is dict and x:
        return 1 + max(depth(x[a]) for a in x)
    if type(x) is list and x:
        return 1 + max(depth(a) for a in x)
    return 0

# ------------------------------------------------------------------------------
# Make a list out of alternating keys and values of a dictionary
# TODO This can be improved by detecting types and bundling up contents
# ------------------------------------------------------------------------------
def listify_dict(input_dict):
    output_list = []
    for key in input_dict:
        output_list.append(key)
        output_list.append(input_dict[key])
    return output_list


# ------------------------------------------------------------------------------
# take any complexity of object and flatten it to a nested tuple.
# TODO This is just a placeholder for something that would be useful, i.e. for
# the above, sometime when it needs to be improved.
# ------------------------------------------------------------------------------
def bundle_structure(data):
    # --------------------------------------------------------------------------
    # --------------------------------------------------------------------------
    return True


# ------------------------------------------------------------------------------
# Many times a list of attributes/properties/etc will have a dictionary stored
# as a string like this:
# 'min=1|mintype=3|max=1|maxtype=3'
# ...this will break those back up into a dictionary
# ------------------------------------------------------------------------------
def unravel_dict_from_string(string, delim_level1, delim_level2):
    keyvals = string.split(delim_level1)

    d = {}
    for kv in keyvals:
        key, val = kv.split(delim_level2,1)
        d[key] = val

    return d

# ------------------------------------------------------------------------------
# Take a list or tuple of even number of items, where each pair of items 
# becomes key, value in a dict, and return that dict.
# ------------------------------------------------------------------------------
def dictify_alternating_list(input_list):
    output_dict = {}
    for i in range(0, len(input_list), 2):
        output_dict[input_list[i]] = input_list[i+1]

    return output_dict

# ------------------------------------------------------------------------------
# A dictionary's keys might be a number; if they are, then cast them as such
# ------------------------------------------------------------------------------
def numeric_dict_keys(input_dict):
    # --------------------------------------------------------------------------
    # start an empty output dictionary
    # --------------------------------------------------------------------------
    output_dict = {}

    # --------------------------------------------------------------------------
    # Work through each item in the dict
    # --------------------------------------------------------------------------
    for key in input_dict:
        # ----------------------------------------------------------------------
        # Get the value
        # ----------------------------------------------------------------------
        value = input_dict[key]

        # ----------------------------------------------------------------------
        # Modify the data type of the key if it's a number
        # ----------------------------------------------------------------------
        xkey = numberify(key)

        # ----------------------------------------------------------------------
        # However it worked out, add the thing to the new dict
        # ----------------------------------------------------------------------
        output_dict[xkey] = value

    # --------------------------------------------------------------------------
    # Send back the possibly-modified dictionary
    # --------------------------------------------------------------------------
    #print("#######")
    #print(input_dict)
    #print(output_dict)
    return output_dict

# ------------------------------------------------------------------------------
# Given a value that might be a number, return it as the most restrictive type
# ------------------------------------------------------------------------------
def numberify(in_item):
    try:
        out_item = int(in_item)
    except:
        try:
            out_item = float(in_item)
        except:
            out_item = in_item

    return out_item

# ------------------------------------------------------------------------------
# Take a dictionary and return a two-column matrix of the contents
# ------------------------------------------------------------------------------
def matrix_from_dict(input_dict):
    output_matrix = []
    for key in input_dict:
        line = []
        line.append(key)
        if type(input_dict[key]) is list:
            line.append(",".join(input_dict[key]))
        elif type(input_dict[key]) is tuple:
            line.append(",".join(input_dict[key]))
        elif type(input_dict[key]) is dict:
            line.append(str(input_dict[key]))
        else:
            line.append(input_dict[key])
        output_matrix.append(line)
    return output_matrix


# ------------------------------------------------------------------------------
# Takes any number of inputs and arranges them in a single matrix
# Return product will be a list of lists of primitives, no matter what is
# given as the input - any cell contents of complex type will be stringified
# ------------------------------------------------------------------------------
def matrixify(inputs, gutter=0, vertical=False):
    # --------------------------------------------------------------------------
    # Identify the type of each input and listify it if it isn't
    # --------------------------------------------------------------------------
    pass1 = []
    for item in inputs:
        if type(item) is list:
            pass1.append(item)
        elif type(item) is tuple:
            pass1.append(list(x for x in item))
        elif type(item) is dict:
            pass1.append(matrix_from_dict(item))
        else:
            pass1.append([[item]])

    # --------------------------------------------------------------------------
    # We now know that each input is a list, so we need to check that each
    # item in each of those lists is itself a list.
    # --------------------------------------------------------------------------
    pass2 = []
    for matrix in pass1:
        clean_matrix = []
        for row in matrix:
            if type(row) is list:
                clean_matrix.append(row)
            elif type(row) is tuple:
                clean_matrix.append(list(x for x in row))
            elif type(row) is dict:
                clean_matrix.append(listify_dict(row))
            else:
                clean_matrix.append([row])
        pass2.append(clean_matrix)

    # --------------------------------------------------------------------------
    # Finally, make sure that all the interior 'cells' are primitives
    # --------------------------------------------------------------------------
    matrices = []
    for matrix in pass2:
        clean_matrix = []
        for row in matrix:
            clean_line = []
            for cell in row:
                if type(cell) is list or type(cell) is tuple:
                    clean_line.append(",".join(str(cell)))
                elif type(cell) is float or type(cell) is int or type(cell) is str:
                    clean_line.append(cell)
                else:
                    clean_line.append(str(cell))
            clean_matrix.append(clean_line)
        matrices.append(clean_matrix)

    # --------------------------------------------------------------------------
    # At this point they are all matrices...
    # Presuming we want to stack these on top of each other
    # --------------------------------------------------------------------------
    if not vertical:
        # ----------------------------------------------------------------------
        # Figure out the widest row of any of them
        # ----------------------------------------------------------------------
        width = 0
        for matrix in matrices:
            # ------------------------------------------------------------------
            # Maybe it's just empty
            # ------------------------------------------------------------------
            if not matrix:
                continue
    
            # ------------------------------------------------------------------
            # check each internal row
            # ------------------------------------------------------------------
            for row in matrix:
                if len(row) > width:
                    width = len(row)
    
        # ----------------------------------------------------------------------
        # Go through each matrix line by line and write it into a final version
        # ----------------------------------------------------------------------
        final_matrix = []
        for i, matrix in enumerate(matrices):
            # ------------------------------------------------------------------
            # Maybe it's empty, so we want to make it a blank line
            # ------------------------------------------------------------------
            if not matrix:
                final_matrix.append([""]*width)
    
            # ------------------------------------------------------------------
            # Each row of the matrix, filled out with empty cells for however 
            # much to make it the common 'width'
            # ------------------------------------------------------------------
            for row in matrix:
                filler = list([""]*(width-len(row)))
                final_row = row + filler
                final_matrix.append(final_row)
    
            # ------------------------------------------------------------------
            # Gutter between matrices if necessary
            # ------------------------------------------------------------------
            if gutter and i < len(matrices)-1:
                for j in range(gutter):
                    final_matrix.append([""]*width)

    # --------------------------------------------------------------------------
    # Alternatively, if it is a vertical merge, we want to do the same thing
    # but sideways-to-slideways instead of upways-to-downways
    # --------------------------------------------------------------------------
    else:
        # ----------------------------------------------------------------------
        # Figure out the tallest column of any of them
        # ----------------------------------------------------------------------
        height = 0
        for matrix in matrices:
            if len(matrix) > height:
                height = len(matrix)

        # ----------------------------------------------------------------------
        # Go through each matrix
        # ----------------------------------------------------------------------
        for i, matrix in enumerate(matrices):
            # ------------------------------------------------------------------
            # we need to blank-fill the thing to match the tallest one
            # ------------------------------------------------------------------
            if matrix:
                width  = len(matrix[0])
            else:
                width  = 0
            filler = list([""]*width)
            for j in range(len(matrix),height):
                matrix.append(filler)

            # ------------------------------------------------------------------
            # Make the gutter with the specified number of cells wide by the
            # height that we got before
            # ------------------------------------------------------------------
            gutter_row = list([""] * gutter)
            gutter_matrix = list([gutter_row] * height)

            # ------------------------------------------------------------------
            # Now, positionally:
            # - first matrix becomes the basis for the rest
            # - any others are tacked on line-by-line
            # ------------------------------------------------------------------
            if i==0:
                final_matrix = matrix
            else:
                if gutter:
                    final_matrix = list(map(list.__add__, final_matrix, gutter_matrix))
                final_matrix = list(map(list.__add__, final_matrix, matrix))

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return final_matrix

# ------------------------------------------------------------------------------
# Any given matrix can be chunkified into a list of separate smaller matrices
# based on the content of the first column
# ------------------------------------------------------------------------------
def chunkify_matrix(matrix, key_lbl=None):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------

    # --------------------------------------------------------------------------
    # Identify the matrix as a Pandas dataframe or simple list of lists, and
    # reform the matrix into a simplified form if it's a df.
    # --------------------------------------------------------------------------
    if isinstance(matrix, pd.DataFrame):
        # ----------------------------------------------------------------------
        # Data frame gets changed to one where 'nan' becomes 'None' so we can
        # read them simply and consistently in the output...
        # ----------------------------------------------------------------------
        clean_df = matrix.where((pd.notnull(matrix)), None)
        matrix_to_read = [tuple(x) for x in clean_df.values]
    elif isinstance(matrix, list):
        if isinstance(matrix[0], [list,tuple]):
            matrix_to_read = matrix
        else:
            print("WARNING! Malformed matrix being fed into chunkify function!")
    else:
        print("WARNING! Malformed matrix being fed into chunkify function!")

    # --------------------------------------------------------------------------
    # Now go through each row
    # current_chunk starts out blank at the top of the list, and is updated
    # once a line is reached where there is something in the leftmost cell.
    # This means there could be non-blank lines with nothing in that cell that
    # are ABOVE the point where the first chunk is identified; those lines
    # are skipped and do NOT become part of the chunked structure.
    # If key_lbl is being used, this same thing applies to every line before
    # the first time the key is hit in the leftmost column.
    # --------------------------------------------------------------------------
    chunks = {}
    current_chunk = ""
    chunk_idx = 1
    for line in matrix_to_read:
        # ----------------------------------------------------------------------
        # Skip lines that are completely blank
        # ----------------------------------------------------------------------
        if all(cell is None for cell in line):
            continue

        # ----------------------------------------------------------------------
        # The chunk headings have something in the first slot: 
        # - if we've been given a key_lbl:
        #     - we're looking for that label exactly as the break point
        #     - The return matrix will be the whole section until next break
        #     - The output key is a simple integer count
        # - if there is no key_lbl:
        #     - anything in column 1 is a break point
        #     - that thing is used as the key
        # ----------------------------------------------------------------------
        if key_lbl:
            # ------------------------------------------------------------------
            # When there's a section break, start a new one and increment
            # ------------------------------------------------------------------
            if line[0] == key_lbl:
                current_chunk = chunk_idx
                chunks[current_chunk] = []
                chunk_idx += 1

            # ------------------------------------------------------------------
            # The matrix is filled in with all columns, including the first
            # ------------------------------------------------------------------
            if current_chunk:
                chunks[current_chunk].append(line)

        else:
            # ------------------------------------------------------------------
            # When there's a section break, start a new one
            # ------------------------------------------------------------------
            if line[0]:
                current_chunk = line[0]
                chunks[current_chunk] = []

            # ------------------------------------------------------------------
            # The matrix is filled in with all but that first column
            # ------------------------------------------------------------------
            chunks[current_chunk].append(line[1:])

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return chunks

# ------------------------------------------------------------------------------
# Given a 2d matrix, the first column are keys and any other columns are either
# the value or a list of values, depending on the number of entries
# Returns a simple dictionary of these results
# ------------------------------------------------------------------------------
def keyval_from_matrix(matrix, startcol=None, endcol=None, startrow=None, endrow=None, required_keys=[]):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}

    # --------------------------------------------------------------------------
    # Get all the configuration values - the key will be the leftmost item
    # --------------------------------------------------------------------------
    for row in matrix:
        # ----------------------------------------------------------------------
        # Skip empty lines
        # ----------------------------------------------------------------------
        if not row:
            continue

        # ----------------------------------------------------------------------
        # Skip lines where there is no key
        # ----------------------------------------------------------------------
        if not row[0]:
            continue

        # ----------------------------------------------------------------------
        # Make a list of whatever is to the right of the key, disregarding
        # blanks etc.
        # ----------------------------------------------------------------------
        vals = [x for x in row[1:] if valid_value(x)]

        # ----------------------------------------------------------------------
        # Key is whatever is in the first item; if there is only one left that
        # is the value, if there are more than one left the value becomes a list
        # ----------------------------------------------------------------------
        key = row[0]
        if type(key) == str:
            key = key.strip()

        if len(vals) == 0:
            data[key] = None
        elif len(vals) == 1:
            data[key] = vals[0]
        elif len(vals) > 1:
            data[key] = vals

    # --------------------------------------------------------------------------
    # Check vs expectation
    # --------------------------------------------------------------------------
    for item in required_keys:
        if not item in data:
            print("WARNING! Expected key field '" + item + " not found in matrix")

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Given a 2d matrix with field names on row 1, return a bundle of mapped-out
# properties and metadata about the contents along with the matrix itself
# ------------------------------------------------------------------------------
def data_grid_from_matrix(matrix, startcol=None, endcol=None, startrow=None, endrow=None, required_fields=[], key_fields=[]):
    # --------------------------------------------------------------------------
    # Determine if this thing is already a Pandas DataFrame; if it's not then
    # pull out the header row and recast the thing as a DataFrame
    # --------------------------------------------------------------------------
    if isinstance(matrix, pd.DataFrame):
        headers = matrix.columns
        df = matrix
    else:
        headers = matrix.pop(0)
        df = pd.DataFrame(matrix, columns=headers)

    # --------------------------------------------------------------------------
    # Check required fields
    # --------------------------------------------------------------------------
    for field in required_fields:
        if not field in headers:
            kill_program("Required field is not present in matrix - " + field)

    # --------------------------------------------------------------------------
    # Set the index if there are key fields
    # --------------------------------------------------------------------------
    if key_fields:
        df.set_index(key_fields, inplace=True, drop=False)

    # --------------------------------------------------------------------------
    # The matrix-level meta
    # --------------------------------------------------------------------------
    mtrxmeta = {}
    mtrxmeta['num_cols']   = len(headers)
    mtrxmeta['num_rows']   = len(df.index)
    mtrxmeta['var_order']  = headers

    # --------------------------------------------------------------------------
    # TODO
    # Figure out the encoding
    # - Only applies to text
    # - is there a way to do this fast?  running chardet on the whole thing wouldn't scale well
    # - Maybe it's done to text variables and aggregated after that
    # --------------------------------------------------------------------------
    #mtrxmeta['encoding']  = detect_encoding_of_matrix(matrix)

    # --------------------------------------------------------------------------
    # Variable level data
    # --------------------------------------------------------------------------
    varmeta = {}
    for var in mtrxmeta['var_order']:
        # ----------------------------------------------------------------------
        # A bundle of information for each variable
        # ----------------------------------------------------------------------
        varmeta[var] = {}

        # ----------------------------------------------------------------------
        # Get a total and a count for this variable
        # ----------------------------------------------------------------------
        total = len(df[var])
        counts = df[var].value_counts()
        items = counts.keys()

        # ----------------------------------------------------------------------
        # Stuff from the data 
        # TODO
        # This thing with the dtype - it seems like there is a difference
        # sometimes, I don't know why, this needs to be figured out.  
        # ----------------------------------------------------------------------
        try:
            if df[var].any():
                varmeta[var]['dtype']      = df.dtypes[var]
            else:
                varmeta[var]['dtype']      = "unknown"
        except:
            if var in df.dtypes:
                varmeta[var]['dtype']      = df.dtypes[var]
            else:
                varmeta[var]['dtype']      = "unknown"


        varmeta[var]['nonblank']   = df[var].count()
        varmeta[var]['blank']      = total - varmeta[var]['nonblank']
        varmeta[var]['buckets']    = len(counts)

        # ----------------------------------------------------------------------
        # Detect the data type
        # ----------------------------------------------------------------------
        varmeta[var]['dcontent'] = detect_data_type_from_list(items)

        # ----------------------------------------------------------------------
        # Get the maximum length of the data when cast as text
        # ----------------------------------------------------------------------
        varmeta[var]['text_len']   = get_longest(items)

        # ----------------------------------------------------------------------
        # Some other derived items
        # ----------------------------------------------------------------------
        if total:
            varmeta[var]['density']    = varmeta[var]['nonblank'] / total
        else:
            varmeta[var]['density']    = None

        if varmeta[var]['nonblank']:
            varmeta[var]['complexity'] = varmeta[var]['buckets'] / varmeta[var]['nonblank']
        else:
            varmeta[var]['complexity'] = None

        # ----------------------------------------------------------------------
        # Numeric items get stats stuff
        # ----------------------------------------------------------------------
        if varmeta[var]['dcontent'] == "Float":
            varmeta[var]['min']  = min([float(x) for x in items])
            varmeta[var]['max']  = max([float(x) for x in items])
        elif varmeta[var]['dcontent'] == "Integer":
            varmeta[var]['min']  = min([int(float(x)) for x in items])
            varmeta[var]['max']  = max([int(float(x)) for x in items])

        # ----------------------------------------------------------------------
        # Alpha items get alpha stuff
        # ----------------------------------------------------------------------
        if varmeta[var]['dcontent'] in ["Alpha"]:
            #varmeta[var]['encoding'] = check the encoding here
            pass

    # --------------------------------------------------------------------------
    # Stuff about the contents of the fields
    # --------------------------------------------------------------------------
    catmeta = {}

    # --------------------------------------------------------------------------
    # Stuff about the records
    # --------------------------------------------------------------------------
    recmeta = {}

    # --------------------------------------------------------------------------
    # Consolidate
    # --------------------------------------------------------------------------
    data = {}
    data['meta'] = mtrxmeta
    data['vars'] = varmeta
    data['cats'] = catmeta
    data['recs'] = recmeta
    data['data'] = df

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data



# ------------------------------------------------------------------------------
# Given a list of items, determine the type of data being stored
#   Alpha    - any item doesn't cast as a number or None or NaN
#   Date     - Special case of Alpha that parses as a date/time
#   Float    - all items evaluate as float
#   Integer  - Special case of float with only whole numbers
#   Binary   - Special case of integer with only 0 and 1
#   Blank    - No data so undetectable
# ------------------------------------------------------------------------------
def detect_data_type_from_list(items):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    dtype = ""

    # --------------------------------------------------------------------------
    # If there's nothing there, we can't tell
    # --------------------------------------------------------------------------
    if items.empty:
        dtype = "Blank"
        return dtype

    # --------------------------------------------------------------------------
    # Make sure we're looking at a minimized version of the thing by casting
    # it to a set of itself
    # --------------------------------------------------------------------------
    set_to_check = list(set(items))

    # --------------------------------------------------------------------------
    # Go through each item - maybe we can tell right away if it's alpha; if it
    # isn't, flag some stuff so we can figure the type of number
    # --------------------------------------------------------------------------
    has_decimals = False
    not_binary = False
    for item in set_to_check:
        # ----------------------------------------------------------------------
        # If it is a string and won't cast as a float, it's got alpha characters
        # ----------------------------------------------------------------------
        if type(item) is str:
            try: 
                item_as_float = float(item)
            except:
                dtype = "Alpha"
                break

        # ----------------------------------------------------------------------
        # It might be one of the datetime formats
        # ----------------------------------------------------------------------
        elif type(item) is pd._libs.tslibs.timestamps.Timestamp:
                dtype = "Date"
                break

        elif isinstance(item, datetime):
                dtype = "Date"
                break

        # ----------------------------------------------------------------------
        # If not, it has to be a number, or something that needs to be added
        # to this function.
        # ----------------------------------------------------------------------
        else:
            try:
                item_as_float = float(item)
            except:
                print("UNACCOUNTED TYPE! - " + str(type(item)))
                sys.exit()

        # ----------------------------------------------------------------------
        # If it gets this far it's at least a number...flag if there are any
        # decimals, which freeze it at the float level
        # ----------------------------------------------------------------------
        if not item_as_float.is_integer():
            has_decimals = True

        # ----------------------------------------------------------------------
        # By now it's a whole number - if it's anything other than 0 or 1 it
        # cannot be binary
        # ----------------------------------------------------------------------
        if item_as_float < 0 or item_as_float > 1:
            not_binary = True

    # --------------------------------------------------------------------------
    # If they're one of the numbers, use those flags to tell which kind
    # --------------------------------------------------------------------------
    if not dtype:
        if has_decimals:
            dtype = "Float"
        elif not_binary:
            dtype = "Integer"
        else:
            dtype = "Binary"

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return dtype

# ------------------------------------------------------------------------------
# Real quick - Does this thing evaluate as a number?
# ------------------------------------------------------------------------------
def is_number(s):
    try:
        float(str(s))
        return True
    except ValueError:
        return False

# ------------------------------------------------------------------------------
# Take a list of delimiters and split a string using all of them at once.
# Returns the list of items.
# ------------------------------------------------------------------------------
def split_by_multiple_delimiters(text, delimiters):

    if text:
        rexp = "|".join(delimiters)
        pieces = re.split(rexp, text)
    else:
        pieces = []

    return pieces



# ------------------------------------------------------------------------------
# Passing in a value from some grid...we want to pass back as true, but this
# definition differs based on the type of the input.
# ------------------------------------------------------------------------------
def valid_value(value, allow_zero=False):
    # --------------------------------------------------------------------------
    # text just isn't blank
    # --------------------------------------------------------------------------
    if type(value) is str:
        if len(value.strip()) > 0:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...integer is a nonzero number
    # --------------------------------------------------------------------------
    elif type(value) is int:
        if value > 0 or value < 0:
            return True
        elif allow_zero and value == 0:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...float is a nonzero number
    # --------------------------------------------------------------------------
    elif type(value) is float:
        if value > 0 or value < 0:
            return True
        elif allow_zero and value == 0:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...Numpy float value is not Nan
    # --------------------------------------------------------------------------
    elif isinstance(value, np.float64):
        if not np.isnan(value):
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...Numpy int value is not Nan
    # --------------------------------------------------------------------------
    elif isinstance(value, np.int64):
        if not np.isnan(value):
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...Numpy bool value is True
    # --------------------------------------------------------------------------
    elif isinstance(value, np.bool_):
        if value:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...boolean is True
    # --------------------------------------------------------------------------
    elif type(value) is bool:
        if value:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...list has 1+ items
    # --------------------------------------------------------------------------
    elif type(value) is list:
        if len(value)>0:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...tuple has 1+ items
    # --------------------------------------------------------------------------
    elif type(value) is tuple:
        if len(value)>0:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # ...dict has 1+ items
    # --------------------------------------------------------------------------
    elif type(value) is dict:
        if len(value.keys())>0:
            return True
        else:
            return False

    # --------------------------------------------------------------------------
    # Date is valid date
    # --------------------------------------------------------------------------
    elif isinstance(value, datetime):
        if value:
            return True
        else:
            print(value)
            return False

    # --------------------------------------------------------------------------
    # ...unidentified None is false
    # --------------------------------------------------------------------------
    elif not value:
        return False


    # --------------------------------------------------------------------------
    # ...or it's something we haven't thought of
    # --------------------------------------------------------------------------
    else:
        print(value)
        print("WARNING!  UNACCOUNTED TYPE IN USAGE COLUMN:",  type(value))
        return False

# ------------------------------------------------------------------------------
# Given something, check if it's a valid date.  Mighte be being passed an actual
# date object, which passes, and might be a string which is thought to be 
# parseable to a date
# ------------------------------------------------------------------------------
def valid_date(value):
    # --------------------------------------------------------------------------
    # If it's a date, it's valid
    # --------------------------------------------------------------------------
    if isinstance(value, datetime):
        return True

    # --------------------------------------------------------------------------
    # --------------------------------------------------------------------------
    try:
        dt_obj = datetime.strptime(value, '%b %d %Y %I:%M%p' )
        return True
    except:
        return False


# ------------------------------------------------------------------------------
# If this thing is valid, return it, otherwise return a blank string
# ------------------------------------------------------------------------------
def this_or_nothing(thing):
    if valid_value(thing):
        return thing
    else:
        return ""

# ------------------------------------------------------------------------------
# Given a list of things, calculate the average similarity of the items 
# Depth is a number indicating the level of thoroughness of this check...
# the number indicates how many next neighbors each entry will check against.
# a full check of each item vs all other items might get hefty.
# ------------------------------------------------------------------------------
def similarity_of_list(ilist, depth=1, debug_on=False):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    scores = []

    # --------------------------------------------------------------------------
    # Depth, if given, must be at least 1
    # --------------------------------------------------------------------------
    if depth < 1:
        depth = 1

    # --------------------------------------------------------------------------
    # Sort the list
    # --------------------------------------------------------------------------
    ilist = sorted(ilist)

    # --------------------------------------------------------------------------
    # Go through each entry and check it vs however many next-neighbors are
    # specified by depth
    # --------------------------------------------------------------------------
    if debug_on:
        print("----------------------------------")
        print(ilist)
        print("----------------------------------")

    for i, item in enumerate(ilist):
        if debug_on:
            print("-------")
            print(i, item)

        # ----------------------------------------------------------------------
        # Check this entry vs its next neighbors for depth
        # ----------------------------------------------------------------------
        for j in range(depth):
            if debug_on:
                print("pos1: ", j+1)
                print("pos2: ", i+j+1)

            if i+j+1 < len(ilist):
                string1 = item
                string2 = ilist[i+j+1]
                ngram   = 2
                jscore  = jaccard_score_of_ngrams(string1, string2, ngram)
                scores.append(jscore)

                if debug_on:
                    print("string1: ", string1)
                    print("string2: ", string2)
                    print("jscore:  ", jscore)

        # ----------------------------------------------------------------------
        # If the list is only one thing long, the measure is just 1.0
        # ----------------------------------------------------------------------
        if len(ilist) == 1:
            scores.append(1.0)


    # --------------------------------------------------------------------------
    # Take the average of those
    # --------------------------------------------------------------------------
    avg_score = average_of_list(scores)

    # --------------------------------------------------------------------------
    # For debug
    # --------------------------------------------------------------------------
    if debug_on:
        print("-----------------------------------------")
        print("Similarity of items in a list")
        print("-----------------------------------------")
        print("items:   ", ilist)
        print("depth:   ", depth)
        print("scores:  ", scores)
        print("average: ", avg_score)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return avg_score

# ------------------------------------------------------------------------------
# Average of a numeric list
# ------------------------------------------------------------------------------
def average_of_list(items):
    # --------------------------------------------------------------------------
    # There has to be some stuff in the list
    # --------------------------------------------------------------------------
    if len(items) == 0:
        return None

    # --------------------------------------------------------------------------
    # Everything in the list should be a number
    # --------------------------------------------------------------------------
    for item in items:
        if not is_number(item):
            return None

    # --------------------------------------------------------------------------
    # Make the total 
    # --------------------------------------------------------------------------
    total_sum = 0
    for item in items:
        total_sum = total_sum + item

    # --------------------------------------------------------------------------
    # Take the average
    # --------------------------------------------------------------------------
    average = total_sum / len(items)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return average


# ------------------------------------------------------------------------------
# Jaccard similarity score is the size of the intersect over size of the union
# ------------------------------------------------------------------------------
def jaccard_score(item1, item2):
    # --------------------------------------------------------------------------
    # Make sure they're reading as sets
    # --------------------------------------------------------------------------
    set1 = set(item1)
    set2 = set(item2)

    # --------------------------------------------------------------------------
    # Calculate the pieces
    # --------------------------------------------------------------------------
    intersect = len(set1.intersection(set2))
    union     = len(set1.union(set2))

    # --------------------------------------------------------------------------
    # Calculate the score
    # --------------------------------------------------------------------------
    if not union:
        score = 0
    else:
        score = intersect / union

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return score



# ------------------------------------------------------------------------------
# Jaccard similarity score of two strings, using n-grams of the given size
# as the members of the set
# ------------------------------------------------------------------------------
def jaccard_score_of_ngrams(string1, string2, n):
    # --------------------------------------------------------------------------
    # Make sure they're reading as sets
    # --------------------------------------------------------------------------
    set1 = set(ngrams_char(string1, n))
    set2 = set(ngrams_char(string2, n))

    # --------------------------------------------------------------------------
    # Calculate the pieces
    # --------------------------------------------------------------------------
    intersect = len(set1.intersection(set2))
    union     = len(set1.union(set2))

    # --------------------------------------------------------------------------
    # Calculate the score
    # --------------------------------------------------------------------------
    if not union:
        score = 0
    else:
        score = intersect / union

    # --------------------------------------------------------------------------
    # For debugging
    # --------------------------------------------------------------------------
    if False:
        print("---------------------------------------------------------")
        print("Jaccard score of n-gram")
        print("---------------------------------------------------------")
        print(string1)
        print(string2)
        print(set1)
        print(set2)
        print(intersect)
        print(union)
        print(score)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return score


# ------------------------------------------------------------------------------
# Get character-level n-grams of a string
# stole this from stack overflow 'Computing N Grams using Python',
# modified to work with strings cut into characters instead of words
# ------------------------------------------------------------------------------
def ngrams_char(istring, n):

    lstring = list(istring)

    output = {}

    for i in range(len(lstring)-n+1):
        g = ''.join(lstring[i:i+n])
        output.setdefault(g, 0)
        output[g] += 1

    return output


# ------------------------------------------------------------------------------
# Length of longest item in a list
# ------------------------------------------------------------------------------
def get_longest(tlist):
    longest = 0
    for item in tlist:
        if not valid_value(item):
            continue

        if len(str(item)) > longest:
            longest = len(str(item))

    return longest

# ------------------------------------------------------------------------------
# Maximum value found amongst collection of things that may or may not evaluate
# Return a number or null
# ------------------------------------------------------------------------------
def get_highest(tlist):
    # --------------------------------------------------------------------------
    # Init
    # --------------------------------------------------------------------------
    anything = False
    highest = 0

    # --------------------------------------------------------------------------
    # Each item on the list
    # --------------------------------------------------------------------------
    for item in tlist:
        # ----------------------------------------------------------------------
        # Skip over entries that just are nothing
        # ----------------------------------------------------------------------
        if not valid_value(item):
            continue

        # ----------------------------------------------------------------------
        # If it can work as a number, do the comparison and update if necessary
        # also flag that it has happened at all
        # ----------------------------------------------------------------------
        try:
            float(item)
            if item > highest:
                highest = item

            anything = True

        # ----------------------------------------------------------------------
        # If it doesn't evaluate as a number, skip it
        # ----------------------------------------------------------------------
        except:
            continue

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    if anything:
        return highest
    else:
        return None

# ------------------------------------------------------------------------------
# This can perform common formatting operations on blocks of lines in a list
# ------------------------------------------------------------------------------
def formatted_block(block, indent=0):
    # --------------------------------------------------------------------------
    # Just process each line in turn
    # --------------------------------------------------------------------------
    fblock = []
    for line in block:
        # ----------------------------------------------------------------------
        # Indentation
        # ----------------------------------------------------------------------
        fline = line
        if indent:
            fline = " "*indent + str(fline)

        # ----------------------------------------------------------------------
        # ?? other types of formatting?
        # ----------------------------------------------------------------------

        # ----------------------------------------------------------------------
        # Add to output list
        # ----------------------------------------------------------------------
        fblock.append(fline)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return fblock

# ------------------------------------------------------------------------------
# This takes in a dataframe and turns it into a verticalized dataframe
# ------------------------------------------------------------------------------
def verticalize(inputdf,idcolumn,questioncolumnname = 'Question',valuecolumnname = 'Value'):
    """
    Takes in the following fields:
        - inputdf (dataframe of the matrix data)
        - idoclumn (the name of the column that contains the identifying values 
        [Response_id/Respondent.id])
        - questioncolumnname (defaults to Question. This gives the name of the column describing
        the deconstructed headers)
        - valuecolumnname (defaults to Value. This gives the name of the column containing
        the value)
    """
    outputdf = pd.DataFrame()
    for column in inputdf.columns:
        if column != idcolumn:
            tempdf = pd.DataFrame()
            tempdf[idcolumn] = inputdf[idcolumn]
            tempdf[questioncolumnname] = [column for i in range(len(tempdf[idcolumn]))]
            tempdf[valuecolumnname] = inputdf[column]
            outputdf = pd.concat([outputdf,tempdf], ignore_index=True)
    return outputdf



# ------------------------------------------------------------------------------
# Parse a text file into a dictionary structure, given a delimiter
# ------------------------------------------------------------------------------
def get_dict_from_text_file(fname, delimiter):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    raw = open(fname).read().splitlines()
    data = {}
        
    # --------------------------------------------------------------------------
    # Go through each line; if there are none of the delimiter it is ignored;
    # if there is more than one only the first one is used
    # --------------------------------------------------------------------------
    for line in raw:
        if delimiter in line:
            pieces = line.split(delimiter, 1)
            data[pieces[0]] = pieces[1]

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Take in a csv file and output a Tableau hyper file
# ------------------------------------------------------------------------------
def run_create_hyper_file_from_csv(table,path_to_database,path_to_csv,delimiter=','):
    """
    Expects 3 inputs:
        - TableInfo (Class from Tableau api)
        - path to output (path_to_database)
        - path to input (path_to_csv)
        - delimiter (defaults to ',')
    """

    # --------------------------------------------------------------------------
    # Optional process parameters.
    # They are documented in the Tableau Hyper documentation, chapter "Process Settings"
    # (https://help.tableau.com/current/api/hyper_api/en-us/reference/sql/processsettings.html).
    #  - Limits the number of Hyper event log files to two.
    #  - Limits the size of Hyper event log files to 100 megabytes.
    # --------------------------------------------------------------------------
    process_parameters = {
        "log_file_max_count": "2",
        "log_file_size_limit": "100M"
    }

    # --------------------------------------------------------------------------
    # Starts the Hyper Process with telemetry enabled to send data to Tableau.
    # To opt out, simply set telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU.
    # --------------------------------------------------------------------------
    with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU, parameters=process_parameters) as hyper:

        # ----------------------------------------------------------------------
        # Optional connection parameters.
        # They are documented in the Tableau Hyper documentation, chapter "Connection Settings"
        # (https://help.tableau.com/current/api/hyper_api/en-us/reference/sql/connectionsettings.html).
        # ----------------------------------------------------------------------
        connection_parameters = {"lc_time": "en_US"}

        # ----------------------------------------------------------------------
        # Creates new Hyper file "customer.hyper".
        # Replaces file with CreateMode.CREATE_AND_REPLACE if it already exists.
        # ----------------------------------------------------------------------
        with Connection(endpoint=hyper.endpoint,
                        database=path_to_database,
                        create_mode=CreateMode.CREATE_AND_REPLACE,
                        parameters=connection_parameters) as connection:

            connection.catalog.create_table(table_definition=table)


            # ------------------------------------------------------------------
            # Load all rows into "Customers" table from the CSV file.
            # `execute_command` executes a SQL statement and returns the impacted row count.
            # ------------------------------------------------------------------
            count_in_table = connection.execute_command(
                command=f"COPY {table.table_name} from {escape_string_literal(path_to_csv)} with "
                f"(format csv, NULL '', delimiter '{delimiter}', header)")

            print(f"The number of rows in table {table.table_name} is {count_in_table}.")

        print("The connection to the Hyper file has been closed.")
    print("The Hyper process has been shut down.")

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return True

def stat_test(field1dict,field2dict,field1,field2,confidence=0.95):
    """
    Takes in the valuedicts from two fields in the same metric
    Expected available values:
        - stddev
        - value
        - ebase
        - base
        - stat (list)
        - field1/field2 : Names of the fields being stat-tested
    """
    cutoffdict = {0.99:2.57,0.95:1.96,0.90:1.645,0.85:1.44,0.80:1.282}
    cutoff = cutoffdict[confidence]
    # If there is no stddev value, use proportion testing method
    value1  = field1dict['value']
    value2  = field2dict['value']
    ebase1  = field1dict['ebase']
    ebase2  = field2dict['ebase']
    stddev1 = field1dict['stddev']
    stddev2 = field2dict['stddev']
    try:
        if stddev1 == '' and stddev2 == '':
            count1  = ebase1 * value1
            count2  = ebase2 * value2
            combined_percent = (count1 + count2)/(ebase1 + ebase2)

            z_score_sq = ((100 * count1/ebase1 - 100 * count2/ebase2)**2) / ((combined_percent*100.0)*(100.0 - combined_percent*100.0)*(1/ebase1 + 1/ebase2))
        else:
            z_score_sq = (value1 - value2)**2/(stddev1**2/ebase1 + stddev2**2/ebase2)
    except:
        return
    # 2.57 (99%)
    # 1.96 (95%)
    # 1.645 (90%)
    # 1.44 (85%)
    # 1.282 (80%)
    if z_score_sq > cutoff**2:
        if value1 > value2:
            field1dict['stat'].append(field2)
        else:
            field2dict['stat'].append(field1)
    return
    
def stat_test_df(datdf1,effdf1,datdf2,effdf2,high,low,confidence=0.95):
    """
    Do stat testing with two sets of dataframes
    """
    cutoffdict = {0.99:2.57,0.95:1.96,0.90:1.645,0.85:1.44,0.80:1.282}
    cutoff = cutoffdict[confidence]
    count1  = effdf1 * datdf1
    count2  = effdf2 * datdf2
    combined_percent = (count1 + count2)/(effdf1 + effdf2)

    zdf = ((100 * count1/effdf1 - 100 * count2/effdf2)**2) / ((combined_percent*100.0)*(100.0 - combined_percent*100.0)*(1/effdf1 + 1/effdf2))
    for idx, row in zdf.iterrows():
        for col,val in row.items():
            if val > cutoff**2:
                if datdf1.loc[idx,col] > datdf2.loc[idx,col]:
                    zdf.loc[idx,col] = high
                else:
                    zdf.loc[idx,col] = low
            else:
                zdf.loc[idx,col] = ''
    return zdf
