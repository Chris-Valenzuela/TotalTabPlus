
import os
import sys
import bs4
import sqlite3
import pandas as pd
import csv
import win32com.client


from text_tools import *
from data_tools import *
from list_tools import *
from io_tools import *

# ------------------------------------------------------------------------------
# General parse for dms program, recursive
# ------------------------------------------------------------------------------
def parse_dms_code(fname, struct):
    # --------------------------------------------------------------------------
    # At the top level, add the main pieces to the structure.  On recursive
    # calls, this won't happen.
    # --------------------------------------------------------------------------
    if not 'defines' in struct:
        struct['defines'] = {}
    if not 'program' in struct:
        struct['program'] = []

    # --------------------------------------------------------------------------
    # Get file contents as a list of lines
    # --------------------------------------------------------------------------
    fdata = get_file_contents(fname)

    # --------------------------------------------------------------------------
    # Go through line-by-line
    # --------------------------------------------------------------------------
    for line in fdata:
        # ----------------------------------------------------------------------
        # Skip blanks
        # ----------------------------------------------------------------------
        if len(line.strip()) == 0:
            continue

        # ----------------------------------------------------------------------
        # Skip full-line comments
        # ----------------------------------------------------------------------
        if line.strip()[0] == "'":
            continue

        # ----------------------------------------------------------------------
        # Defines get stored as we go
        # ----------------------------------------------------------------------
        if line.strip().lower().startswith("#define"):
            var, value = parse_define(line)
            struct['defines'][var] = value

        # ----------------------------------------------------------------------
        # Recurse includes, or store code
        # ----------------------------------------------------------------------
        if line.strip().lower().startswith("#include"):
            # ------------------------------------------------------------------
            # Call this thing recursively on the file that's getting included
            # ------------------------------------------------------------------
            subfname = parse_include_fname(line)
            struct = parse_dms_code(subfname, struct)
        else:
            # ------------------------------------------------------------------
            # Add the line
            # (replace Defines maybe?)
            # ------------------------------------------------------------------
            #for dname in struct['defines']:
            #    if dname in line:
            #        line = line.replace(dname, struct['defines'][dname])
            struct['program'].append(line)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return struct

# ------------------------------------------------------------------------------
# This cleanly parses a '#define' line in a DMS/MRS program, and returns two 
# strings - the 'var' and the 'value'
# ------------------------------------------------------------------------------
def parse_define(line):
    # --------------------------------------------------------------------------
    # Clean up the line, removing the #define itself, and converting any tabs
    # to four spaces
    # --------------------------------------------------------------------------
    line = line.strip()[8:].replace("\t", "    ")

    # --------------------------------------------------------------------------
    # Value may have in-line comments at the end...remove those
    # --------------------------------------------------------------------------
    num_quotes = 0

    for i, char in enumerate(line):
        # ----------------------------------------------------------------------
        # Increment the quotes there have been so far
        # ----------------------------------------------------------------------
        if char == '"':
            num_quotes += 1

        # ----------------------------------------------------------------------
        # If there are an even number of double quotes, then we're in bare
        # code, and if we then hit a single quote, this is a comment, so we
        # truncate the line there
        # ----------------------------------------------------------------------
        if num_quotes % 2 == 0 and char == "'":
            line = line[:i].strip()

    # --------------------------------------------------------------------------
    # split out the two parts
    # we have to account for 'false' as a value separately, because it will
    # break the double assignment
    # same thing for an empty string
    # --------------------------------------------------------------------------
    if line.strip().lower().endswith("false"):
        var   = line.split(" ")[0]
        value = "False"
    elif line.strip().endswith('""'):
        var   = line.split(" ")[0]
        value = '""'
    else:
        var, value = line.split(" ", 1)

    # --------------------------------------------------------------------------
    # The value might be a construct rather than a single thing
    # --------------------------------------------------------------------------
    value = construct_string(value)

    # --------------------------------------------------------------------------
    # finish
    # --------------------------------------------------------------------------
    return var, value


# ------------------------------------------------------------------------------
# This cleanly parses a '#include' line in a DMS/MRS program, and returns the
# name of the file in question
# ------------------------------------------------------------------------------
def parse_include_fname(line):
    # --------------------------------------------------------------------------
    # get rid of the include itself
    # --------------------------------------------------------------------------
    line = line.strip()[9:]

    # --------------------------------------------------------------------------
    # Value may have in-line comments at the end...remove those
    # --------------------------------------------------------------------------
    num_quotes = 0

    for i, char in enumerate(line):
        # ----------------------------------------------------------------------
        # Increment the quotes there have been so far
        # ----------------------------------------------------------------------
        if char == '"':
            num_quotes += 1

        # ----------------------------------------------------------------------
        # If there are an even number of double quotes, then we're in bare
        # code, and if we then hit a single quote, this is a comment, so we
        # truncate the line there
        # ----------------------------------------------------------------------
        if num_quotes % 2 == 0 and char == "'":
            line = line[:i].strip()

    # --------------------------------------------------------------------------
    # The filename will be in quotes; get rid of those
    # --------------------------------------------------------------------------
    if '"' in line:
        line = line.replace('"', "")
    if "'" in line:
        line = line.replace("'", "")

    # --------------------------------------------------------------------------
    # finish
    # --------------------------------------------------------------------------
    return line


# ------------------------------------------------------------------------------
# Parsing a whole regular MDD file and returning a useful structure
# ------------------------------------------------------------------------------
def parse_mdd(mdd_file, fdebug=False):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    print("...parsing mdd file " + mdd_file + "...")

    # --------------------------------------------------------------------------
    # Grab the whole MDD as xml
    # --------------------------------------------------------------------------
    mdd_data = bs4.BeautifulSoup(open(mdd_file, "r", encoding='utf-8'), "lxml-xml")

    # --------------------------------------------------------------------------
    # Prettified debug version
    # --------------------------------------------------------------------------
    if fdebug:
        dout = open("debug_mdd_parser.txt", "w", encoding='utf-8')
        for item in mdd_data.prettify().splitlines():
            print(item, file=dout)
        dout.close()

    # --------------------------------------------------------------------------
    # These are the contents of the mdd structure:
    # --------------------------------------------------------------------------
    raw_properties          = mdd_data.xml.metadata.properties
    raw_languages           = mdd_data.xml.metadata.languages
    raw_categorymap         = mdd_data.xml.metadata.categorymap
    raw_definition          = mdd_data.xml.metadata.definition
    raw_design              = mdd_data.xml.metadata.design
    raw_design_fields       = mdd_data.xml.metadata.design.fields
    raw_design_types        = mdd_data.xml.metadata.design.types
    raw_design_pages        = mdd_data.xml.metadata.design.pages
    raw_design_routings     = mdd_data.xml.metadata.design.routings
    raw_design_properties   = mdd_data.xml.metadata.design.properties
    raw_system              = mdd_data.xml.metadata.system

    raw_versionlist         = mdd_data.xml.metadata.versionlist
    raw_savelogs            = mdd_data.xml.metadata.savelogs

    # raw_datasources         = mdd_data.xml.metadata.datasources
    # raw_systemrouting       = mdd_data.xml.metadata.systemrouting
    # raw_mappings            = mdd_data.xml.metadata.mappings
    # raw_contexts            = mdd_data.xml.metadata.contexts
    # raw_labeltypes          = mdd_data.xml.metadata.labeltypes
    # raw_routingcontexts     = mdd_data.xml.metadata.routingcontexts
    # raw_scriptttypes        = mdd_data.xml.metadata.scripttypes
    # raw_atoms               = mdd_data.xml.metadata.atoms

    # --------------------------------------------------------------------------
    # Save history of the file
    # --------------------------------------------------------------------------

    mdd_savelogs = {}
    for node in raw_savelogs:
        # ----------------------------------------------------------------------
        # Skip over blanks
        # ----------------------------------------------------------------------
        if not isinstance(node, bs4.element.Tag):
            continue

        # ----------------------------------------------------------------------
        # Get the things
        # ----------------------------------------------------------------------
        sdate = node.attrs['date']
        sfver = node.attrs['fileversion']
        svset = node.attrs['versionset']
        suser = node.attrs['username']

        if 'count' in node.attrs:
            scount = node.attrs['count']
        else:
            scount = 1

        # ----------------------------------------------------------------------
        # Make the record
        # ----------------------------------------------------------------------
        mdd_savelogs[sdate] = {}
        mdd_savelogs[sdate]['fileversion'] = sfver
        mdd_savelogs[sdate]['versionset']  = svset
        mdd_savelogs[sdate]['username']    = suser
        mdd_savelogs[sdate]['count']       = scount

    # --------------------------------------------------------------------------
    # Metadata properties
    # --------------------------------------------------------------------------
    mdd_properties = retrieve_properties(raw_properties)

    # --------------------------------------------------------------------------
    # Languages
    # --------------------------------------------------------------------------
    mdd_languages = {}
    for node in raw_languages.children:
        if node.name == "versions":
            continue
        if node.name == "deleted":
            continue
        if type(node) == bs4.element.NavigableString:
            continue
        if not 'name' in node.attrs:
            continue
        lname = node.attrs['name']
        mdd_languages[lname] = retrieve_properties(node.properties)
        mdd_languages[lname]['langid'] = node.attrs['id']

    # --------------------------------------------------------------------------
    # Start the data structures for the category components
    # --------------------------------------------------------------------------
    mdd_categories = {}
    mdd_categories['by_cat'] = {}
    mdd_categories['shared'] = {}
    mdd_categories['by_question'] = {}

    # --------------------------------------------------------------------------
    # Compile information about categories; here just get the list of existing
    # category names and start a structure with them.
    # --------------------------------------------------------------------------
    for node in raw_categorymap:
        # ----------------------------------------------------------------------
        # 'categoryid' is what these things are called
        # ----------------------------------------------------------------------
        if not node.name == "categoryid":
            continue

        # ----------------------------------------------------------------------
        # Grab what we need
        # ----------------------------------------------------------------------
        catname = node.attrs['name'].lower()
        catval  = node.attrs['value']

        # ----------------------------------------------------------------------
        # Fill in the start of the structure for this category
        # ----------------------------------------------------------------------
        mdd_categories['by_cat'][catname] = {}
        mdd_categories['by_cat'][catname]['value'] = catval
        mdd_categories['by_cat'][catname]['catname'] = node.attrs['name']
        mdd_categories['by_cat'][catname]['usage_sl']  = []
        mdd_categories['by_cat'][catname]['usage_var'] = []


    # --------------------------------------------------------------------------
    # Go through the definitions and start structures for the shared lists
    # --------------------------------------------------------------------------
    nested_sl = []
    for node in raw_definition:
        # ----------------------------------------------------------------------
        # 'categories' indicates a shared list definition
        # ----------------------------------------------------------------------
        if not node.name == "categories":
            continue

        # ----------------------------------------------------------------------
        # Get the information about the shared list from the tree
        # ----------------------------------------------------------------------
        sl_name    = node.attrs['name'].lower()
        sl_id      = node.attrs['id']
        sl_props   = retrieve_properties(node.properties)
        sl_notes   = retrieve_properties(node.notes)
        sl_labels  = retrieve_labels(node.labels)

        # ----------------------------------------------------------------------
        # In order to parse nested shared lists, they need to already exist in
        # this data structure, so we're going to skip over any like that and
        # get them on a second pass
        # ----------------------------------------------------------------------
        if node.find_all(['categories']):
            nested_sl.append(sl_name)
            sl_cats = None
        else:
            sl_cats = retrieve_categories(node, mdd_categories['shared'])


        # ----------------------------------------------------------------------
        # Run back through the categories we just parsed and flag their
        # usage in the other section
        # ----------------------------------------------------------------------
        if sl_cats:
            # ------------------------------------------------------------------
            # Sometimes there is something referenced in a shared list that 
            # never actually gets used anywhere - it wouldn't have been in the
            # category map if that's the case, but we'll still make an entry
            # ------------------------------------------------------------------
            for catname in sl_cats['list']:
                cat_tag = catname.lower()
                if not cat_tag in mdd_categories['by_cat']:
                    mdd_categories['by_cat'][cat_tag] = {}
                    mdd_categories['by_cat'][cat_tag]['value'] = -1
                    mdd_categories['by_cat'][cat_tag]['catname'] = catname
                    mdd_categories['by_cat'][cat_tag]['usage_sl']  = []
                    mdd_categories['by_cat'][cat_tag]['usage_var'] = []

                # --------------------------------------------------------------
                # Add the shared list to the variable's usage list
                # --------------------------------------------------------------
                mdd_categories['by_cat'][cat_tag]['usage_sl'].append(sl_name)

        # ----------------------------------------------------------------------
        # Make an area in the data structure for this shared list
        # ----------------------------------------------------------------------
        mdd_categories['shared'][sl_name] = {}
        mdd_categories['shared'][sl_name]['id']     = sl_id
        mdd_categories['shared'][sl_name]['name']   = node.attrs['name']
        mdd_categories['shared'][sl_name]['props']  = sl_props
        mdd_categories['shared'][sl_name]['notes']  = sl_notes
        mdd_categories['shared'][sl_name]['labels'] = sl_labels
        mdd_categories['shared'][sl_name]['cats']   = sl_cats
        mdd_categories['shared'][sl_name]['usage_var'] = []

    # --------------------------------------------------------------------------
    # Make a second pass through the shared lists and add the actual categories
    # we can't always do this on the first pass because it is possible that
    # nested shared lists are not defined before they are used in the sequence.
    # --------------------------------------------------------------------------
    for node in raw_definition:
        # ----------------------------------------------------------------------
        # 'categories' indicates a shared list definition
        # ----------------------------------------------------------------------
        if not node.name == "categories":
            continue

        # ----------------------------------------------------------------------
        # Re-get the name, and skip past if we've already done this one
        # ----------------------------------------------------------------------
        sl_name = node.attrs['name'].lower()

        if not sl_name in nested_sl:
            continue

        # ----------------------------------------------------------------------
        # Get the category data
        # ----------------------------------------------------------------------
        sl_cats = retrieve_categories(node, mdd_categories['shared'])

        # ----------------------------------------------------------------------
        # Run back through the categories we just parsed and flag their
        # usage in the other section
        # ----------------------------------------------------------------------
        for catname in sl_cats['list']:
            # ------------------------------------------------------------------
            # Sometimes there is something referenced in a shared list that 
            # never actually gets used anywhere - it wouldn't have been in the
            # category map if that's the case, but we'll still make an entry
            # ------------------------------------------------------------------
            cat_tag = catname.lower()
            if not cat_tag in mdd_categories['by_cat']:
                mdd_categories['by_cat'][cat_tag] = {}
                mdd_categories['by_cat'][cat_tag]['value'] = -1
                mdd_categories['by_cat'][cat_tag]['catname'] = catname
                mdd_categories['by_cat'][cat_tag]['usage_sl']  = []
                mdd_categories['by_cat'][cat_tag]['usage_var'] = []

            # ------------------------------------------------------------------
            # Add the shared list to the variable's usage list
            # ------------------------------------------------------------------
            mdd_categories['by_cat'][cat_tag]['usage_sl'].append(sl_name)

        # ----------------------------------------------------------------------
        # Overwrite this info in the data for this list that was started before
        # ----------------------------------------------------------------------
        mdd_categories['shared'][sl_name]['cats'] = sl_cats

    # --------------------------------------------------------------------------
    # Go through the definitions and gather information about variables
    # --------------------------------------------------------------------------
    mdd_variables = {}
    mdd_variables['var_defs'] = {}

    for node in raw_definition:
        # ----------------------------------------------------------------------
        # 'variable' indicates an individual variable
        # ----------------------------------------------------------------------
        if not node.name == "variable":
            continue

        # ----------------------------------------------------------------------
        # Skip past nodes that are nothing
        # ----------------------------------------------------------------------
        if not node:
            continue

        # ----------------------------------------------------------------------
        # Skip past nodes that are not tags
        # ----------------------------------------------------------------------
        if not type(node) == bs4.element.Tag:
            continue

        # ----------------------------------------------------------------------
        # Get the components
        # ----------------------------------------------------------------------
        var_name   = node.attrs['name']
        var_id     = node.attrs['id']
        var_typen  = node.attrs['type']
        var_type   = metadata_type_name(var_typen)
        var_attrs  = retrieve_attributes(node, ['name','id','type'])
        var_props  = retrieve_properties(node.properties)
        var_notes  = retrieve_properties(node.notes)
        var_styles = retrieve_properties(node.styles)
        var_labels = retrieve_labels(node.labels)
        var_cats   = retrieve_categories(node.categories, mdd_categories['shared'])
        var_helper = retrieve_helpers(node.helperfields)
        var_axis   = retrieve_axis_exp(node.axis)
        var_tmplt  = retrieve_properties(node.templates)
        var_lstyle = retrieve_properties(node.labelstyles)

        if False:
            print("-------------------------------------------------")
            print(node.encode('utf-8'))
            print("    var_name:   ", var_name)
            print("    var_id:     ", var_id)
            print("    var_type:   ", var_type)
            print("    var_attrs:  ", var_attrs)
            print("    var_props:  ", var_props)
            print("    var_notes:  ", var_notes)
            print("    var_styles: ", var_styles)
            print("    var_helper: ", var_helper)
            print("    var_axis:   ", var_axis)
            print("    var_labels: ", str(var_labels).encode('utf-8'))
            print("    var_cats:   ", str(var_cats).encode('utf-8'))
            print("---------------------------------")
            something = retrieve_categories(node.categories, mdd_categories['shared'], fdebug=True)

        # ----------------------------------------------------------------------
        # Make a structure
        # ----------------------------------------------------------------------
        mdd_variables['var_defs'][var_id] = {}
        mdd_variables['var_defs'][var_id]['name']     = var_name
        mdd_variables['var_defs'][var_id]['type']     = var_type
        mdd_variables['var_defs'][var_id]['labels']   = var_labels
        if var_attrs:
            mdd_variables['var_defs'][var_id]['attrs']  = var_attrs
        if var_type == "categorical":
            mdd_variables['var_defs'][var_id]['cats']   = var_cats
            mdd_variables['var_defs'][var_id]['shared_lists'] = []
        if var_props:
            mdd_variables['var_defs'][var_id]['props']  = var_props
        if var_notes:
            mdd_variables['var_defs'][var_id]['notes']  = var_notes
        if var_styles:
            mdd_variables['var_defs'][var_id]['styles'] = var_styles
        if var_helper:
            mdd_variables['var_defs'][var_id]['helper'] = var_helper
        if var_axis:
            mdd_variables['var_defs'][var_id]['axis']   = var_axis
        if var_tmplt:
            mdd_variables['var_defs'][var_id]['templates'] = var_tmplt
        if var_lstyle:
            mdd_variables['var_defs'][var_id]['labelstyles'] = var_lstyle

        # ----------------------------------------------------------------------
        # Check if there's anything that we didn't account for
        # ----------------------------------------------------------------------
        for child in node:
            if type(child) == bs4.element.NavigableString:
                continue
            if child.name in ['versions','DIFF']:
                continue
            if child.name in ['properties','notes','labels','categories','styles','helperfields','axis','templates','labelstyles']:
                continue
            print("IGNORING UNDOCUMENTED NODE WITHIN VARIABLE      ", child.encode('utf-8'))

    # --------------------------------------------------------------------------
    # Go through and flag the pages
    # --------------------------------------------------------------------------
    mdd_pages = {}
    for node in raw_definition:
        # ----------------------------------------------------------------------
        # only page type
        # ----------------------------------------------------------------------
        if not node.name == "page":
            continue

        # ----------------------------------------------------------------------
        # THIS NEEDS TO BE DONE
        # ----------------------------------------------------------------------
        #print(node)

    # --------------------------------------------------------------------------
    # 'Other' text fields
    # --------------------------------------------------------------------------
    mdd_othervars = {}
    for node in raw_definition:
        # ----------------------------------------------------------------------
        # only other type
        # ----------------------------------------------------------------------
        if not node.name == "othervariable":
            continue

        # ----------------------------------------------------------------------
        # THIS NEEDS TO BE DONE
        # ----------------------------------------------------------------------
        #print(node)

    # --------------------------------------------------------------------------
    # What are we missing from the varible definition
    # --------------------------------------------------------------------------
    for node in raw_definition:
        if type(node) == bs4.element.NavigableString:
            continue
        if node.name in ["variable","categories","page","othervariable"]:
            continue
        print("UNDOCUMENTED NODE WITHIN DEFINITIONS      ", child.encode('utf-8'))


    # --------------------------------------------------------------------------
    # Routing scripts
    # --------------------------------------------------------------------------
    mdd_routing = {}
    for item in raw_design_routings.scripts:
        # ----------------------------------------------------------------------
        # It's nested inside a thing called scripttype
        # ----------------------------------------------------------------------
        if item.name == "scripttype":
            # ------------------------------------------------------------------
            # There are script level properties as attributes
            # ------------------------------------------------------------------
            script_props = item.attrs

            # ------------------------------------------------------------------
            # Drill down to the script nodes and parse and store each one
            # ------------------------------------------------------------------
            for script in item.children:
                if script.name == "script":
                    sname = script.attrs['name']
                    raw_script = script.contents[0][1:-1]
                    script_info = parse_routing_script(raw_script.splitlines())
                    script_dir  = compile_routing_script_directory(script_info)

                    mdd_routing[sname] = {}
                    mdd_routing[sname]['props']      = script_props
                    mdd_routing[sname]['raw_script'] = raw_script
                    mdd_routing[sname]['info']       = script_info
                    mdd_routing[sname]['directory']  = script_dir

    # --------------------------------------------------------------------------
    # Build a single structure to hold the things we have so far...we need 
    # to pass this stuff in to get some more complex parts
    # --------------------------------------------------------------------------
    mdd_info = {}
    mdd_info['properties']  = mdd_properties
    mdd_info['languages']   = mdd_languages
    mdd_info['categories']  = mdd_categories
    mdd_info['variables']   = mdd_variables
    mdd_info['routing']     = mdd_routing
    mdd_info['pages']       = mdd_pages
    mdd_info['othervars']   = mdd_othervars
    mdd_info['savelogs']    = mdd_savelogs

    # --------------------------------------------------------------------------
    # compile as a master list of variables for easy reference
    # --------------------------------------------------------------------------
    mdd_master = {}
    mdd_master['list'] = []
    mdd_master['info'] = {}

    # --------------------------------------------------------------------------
    # First pass on the master is the system-level section
    # --------------------------------------------------------------------------
    for node in raw_system:
        node_data = parse_metadata_field(mdd_info, node, 0, [], [], {}, [])
        mdd_master['list'].extend(node_data['list'])
        mdd_master['info'].update(node_data['info'])

    # --------------------------------------------------------------------------
    # Then go through the fields section and do the same thing
    # --------------------------------------------------------------------------
    for node in raw_design_fields:
        node_data = parse_metadata_field(mdd_info, node, 0, [], [], {}, [])
        mdd_master['list'].extend(node_data['list'])
        mdd_master['info'].update(node_data['info'])

    # --------------------------------------------------------------------------
    # Add a flag for response type = single/multi
    # --------------------------------------------------------------------------
    for i, item in enumerate(mdd_master['list']):
        # ----------------------------------------------------------------------
        # Default to single
        # ----------------------------------------------------------------------
        mdd_master['info'][item]['singlemulti'] = "single"

        # ----------------------------------------------------------------------
        # If it's a categorical it can be multi, change the default to that
        # ----------------------------------------------------------------------
        if mdd_master['info'][item]['type'] == "categorical":
            mdd_master['info'][item]['singlemulti'] = "multi"

        # ------------------------------------------------------------------
        # If there is an attribute section, then 'max' might be in there...
        # run through each and check if the numbers is >1; if it is, then
        # this is a multipunch
        # ------------------------------------------------------------------
        if 'attrs' in mdd_master['info'][item]:
            attrs = mdd_master['info'][item]['attrs']
            for attr in attrs:
                if attr.lower() == "max":
                    if float(attrs[attr]) == 1:
                        mdd_master['info'][item]['singlemulti'] = "single"

    # --------------------------------------------------------------------------
    # Add the compiled master to the structure
    # --------------------------------------------------------------------------
    mdd_info['master'] = mdd_master

    #print("-------")
    #for i, item in enumerate(mdd_master['list']):
    #    print(str(i).ljust(5), mdd_master['info'][item]['type'].ljust(12), item.ljust(64), str(mdd_master['info'][item]).encode('utf-8'))


    # --------------------------------------------------------------------------
    # Now that the master list is there and we have the full variable names
    # set, run back through the categories and add the variables to that
    # list that we started way back when those were initialized
    # --------------------------------------------------------------------------
    for var in mdd_master['list']:
        # ----------------------------------------------------------------------
        # Only things with categories
        # ----------------------------------------------------------------------
        if not 'cats' in mdd_master['info'][var]:
            continue

        # ----------------------------------------------------------------------
        # Go through each category
        # ----------------------------------------------------------------------
        for catname in mdd_master['info'][var]['cats']['list']:
            # ------------------------------------------------------------------
            # Get the info for this category
            # ------------------------------------------------------------------
            catinfo = mdd_master['info'][var]['cats']['info'][catname]

            # ------------------------------------------------------------------
            # Skip over subheadings
            # ------------------------------------------------------------------
            if not catinfo['cat_type'] == "leaf":
                continue

            # ------------------------------------------------------------------
            # It's possible for categories to exist in questions but not in
            # the main set, so we may need to initialize them here again
            # ------------------------------------------------------------------
            cat_tag = catinfo['catname'].lower()

            if not cat_tag in mdd_categories['by_cat']:
                #print("MISSING", var, cat_tag)
                mdd_categories['by_cat'][cat_tag] = {}
                mdd_categories['by_cat'][cat_tag]['value'] = -1
                mdd_categories['by_cat'][cat_tag]['catname'] = catname
                mdd_categories['by_cat'][cat_tag]['usage_sl']  = []
                mdd_categories['by_cat'][cat_tag]['usage_var'] = []

            # ------------------------------------------------------------------
            # Add the variable to the list
            # ------------------------------------------------------------------
            mdd_categories['by_cat'][cat_tag]['usage_var'].append(var)

            # ------------------------------------------------------------------
            # If this variable uses a shared list, add that to its info
            # ------------------------------------------------------------------
            if mdd_categories['by_cat'][cat_tag]['usage_sl']:
                for sl in mdd_categories['by_cat'][cat_tag]['usage_sl']:
                    if not sl in mdd_master['info'][var]['shared_lists']:
                        mdd_master['info'][var]['shared_lists'].append(sl)

    # --------------------------------------------------------------------------
    # We also want to track things by 'question', AKA the top of the hierarchy
    # We'll get this by trolling back through the master list
    # --------------------------------------------------------------------------
    questions = {}
    questions['list'] = []
    questions['info'] = {}

    for vname in mdd_master['list']:
        # ----------------------------------------------------------------------
        # Get the question name  or skip past if it's not a variable
        # ----------------------------------------------------------------------
        if 'question' in mdd_master['info'][vname]:
            qname = mdd_master['info'][vname]['question']
        else:
            continue

        # ----------------------------------------------------------------------
        # Initialize a new question if this is the first instance
        # ----------------------------------------------------------------------
        if not qname in questions['list']:
            questions['list'].append(qname)
            questions['info'][qname] = {}
            questions['info'][qname]['vars'] = []
            questions['info'][qname]['depth'] = 0

        # ----------------------------------------------------------------------
        # Record the maximum depth that this question gets to
        # ----------------------------------------------------------------------
        if mdd_master['info'][vname]['level'] > questions['info'][qname]['depth']:
            questions['info'][qname]['depth'] = mdd_master['info'][vname]['level']

        # ----------------------------------------------------------------------
        # Add this variable to the list
        # ----------------------------------------------------------------------
        questions['info'][qname]['vars'].append(vname)

    mdd_info['questions'] = questions

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return mdd_info

# ------------------------------------------------------------------------------
# To parse an individual 'question', which may be nested in numerous different
# configurations...so this is recursive, ultimately giving back a depth-first 
# list of variables with flattened out names, and a dictionary of info to match
# ------------------------------------------------------------------------------
def parse_metadata_field(mdd_info, node, level, namelev, vlist, vinfo, fields):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}
    data['list'] = vlist
    data['info'] = vinfo

    #if 'name' in node.attrs:
    #    print("-----------------------")
    #    print(level, namelev)
    #    print(fields)
    #    print(node.name, ":", node.attrs['name'])

    # --------------------------------------------------------------------------
    # Deleted items can just go away
    # --------------------------------------------------------------------------
    if node.name == "deleted":
        return data

    # --------------------------------------------------------------------------
    # Variables are processed directly
    # --------------------------------------------------------------------------
    elif node.name == "variable":
        # ----------------------------------------------------------------------
        # The name is constructed from the levels and the variable name
        # this will give us a unique identifier that is readable
        # ----------------------------------------------------------------------
        var_name = "".join(namelev) + node.attrs['name']

        # ----------------------------------------------------------------------
        # The 'ref' field holds the ID that we used back in the variable defs
        # so we can use it to pull over the already-organized information
        # ----------------------------------------------------------------------
        var_id = node.attrs['ref']
        var_info = mdd_info['variables']['var_defs'][var_id]
        var_info['var_id'] = var_id

        # ----------------------------------------------------------------------
        # Add the level at which this variable lies in the scheme, and the
        # names-by-level of its parents, along with the 'question', which
        # is whatever the level 0 element is (removing class/loop notation)
        # ----------------------------------------------------------------------
        var_info['level'] = level

        var_info['namelev'] = namelev

        # ----------------------------------------------------------------------
        # The question name is the root of the name, without any markers
        # ----------------------------------------------------------------------
        if level > 0:
            var_info['question'] = namelev[0].replace("[..].","").replace(".","")
        else:
            var_info['question'] = var_info['name'].replace("[..].","").replace(".","")

        # ----------------------------------------------------------------------
        # Add this to the master list
        # ----------------------------------------------------------------------
        data['list'].append(var_name)
        data['info'][var_name] = var_info
        if False:
            print(":::  ", level, var_info['type'].ljust(12), var_name.ljust(60), str(var_info).encode('utf-8'))

    # --------------------------------------------------------------------------
    # Classes (or 'blocks') are lists of variables with a level jump
    # --------------------------------------------------------------------------
    elif node.name == "class":
        # ----------------------------------------------------------------------
        # Increment the leveled naming system
        # ----------------------------------------------------------------------
        if level == 0:
            namelev = []
        if len(namelev) < level+1:
            namelev.append(node.attrs['name'] + ".")

        #elif len(namelev) > level:
        #    #namelev.pop()
        #    namelev = namelev[:-1]
        #    if len(namelev) < level+1:
        #        namelev.append(node.attrs['name'] + ".")

        #print(namelev, node)

        # ----------------------------------------------------------------------
        # Grab the list of child fields so we can check against it and know
        # where to go back to the original name level
        # ----------------------------------------------------------------------
        field_types = [f.name for f in node.fields]
        field_names = [f.attrs['name'] for f in node.fields if 'name' in f.attrs]

        fields.append(field_names)

        # ----------------------------------------------------------------------
        # Get the stuff about the class itself
        # ----------------------------------------------------------------------
        #class_name   = "".join(namelev) + node.attrs['name']
        class_name   = "".join(namelev)
        class_attrs  = retrieve_attributes(node, ['name'])
        class_props  = retrieve_properties(node.properties)
        class_notes  = retrieve_properties(node.notes)
        class_labels = retrieve_labels(node.labels)

        # ----------------------------------------------------------------------
        # Construct a variable-info-like structure for the class
        # ----------------------------------------------------------------------
        class_info = {}
        class_info['name']     = class_name
        class_info['level']    = level
        class_info['type']     = "class"
        class_info['labels']   = class_labels
        if class_attrs:
            class_info['attrs']  = class_attrs
        if class_props:
            class_info['props']  = class_props
        if class_notes:
            class_info['notes']  = class_notes

        class_info['namelev'] = namelev

        # ----------------------------------------------------------------------
        # The question name is the root of the name, without any markers
        # ----------------------------------------------------------------------
        if level > 0:
            class_info['question'] = namelev[0].replace("[..].","").replace(".","")
        else:
            class_info['question'] = class_info['name'].replace("[..].","").replace(".","")

        # ----------------------------------------------------------------------
        # Add the class as an entity to the master list
        # ----------------------------------------------------------------------
        data['list'].append(class_name)
        data['info'][class_name] = class_info
        if False:
            print("%%%  ", level, class_info['type'].ljust(12), class_name.ljust(60), str(class_info).encode('utf-8'))

        # ----------------------------------------------------------------------
        # Now recurse through the list to get the actual variables
        # ----------------------------------------------------------------------
        for fnode in node.fields:
            # ------------------------------------------------------------------
            # Call the next level down
            # ------------------------------------------------------------------
            node_data = parse_metadata_field(mdd_info, fnode, level+1, namelev, data['list'], data['info'], fields)

            # ------------------------------------------------------------------
            # Reset the name leveling system if we're done with a set of fields
            # ------------------------------------------------------------------
            while len(fields[-1]) == 0:
                fields = fields[:-1]
                namelev = namelev[:-1]
                level = len(namelev)-1

            # ------------------------------------------------------------------
            # Presuming that iteration was something real, remove it from the 
            # list.
            # ------------------------------------------------------------------
            if 'name' in fnode.attrs:
                newend = fields[-1]
                #print('cfoo', newend)
                #print('cbar', fnode.attrs['name'])
                newend.remove(fnode.attrs['name'])
        
                fields = fields[:-1]
                fields.append(newend)

    # --------------------------------------------------------------------------
    # Loops are complex data structures
    # --------------------------------------------------------------------------
    elif node.name == "loop":
        # ----------------------------------------------------------------------
        # Increment the leveled naming system
        # ----------------------------------------------------------------------
        if level == 0:
            namelev = []

        if len(namelev) < level+1:
            namelev.append(node.attrs['name'] + "[..].")

        #elif len(namelev) > level:
        #    namelev = namelev[:-1]

        #    if len(namelev) < level+1:
        #        namelev.append(node.attrs['name'] + "[..].")


        # ----------------------------------------------------------------------
        # Grab the list of child fields so we can check against it and know
        # where to go back to the original name level
        # ----------------------------------------------------------------------
        field_types = [f.name for f in node.fields]
        field_names = [f.attrs['name'] for f in node.fields if 'name' in f.attrs]

        fields.append(field_names)

        #if namelev and namelev[0].lower().startswith("rb_product"):
        #    print(field_types)
        #    print(field_names)

        # ----------------------------------------------------------------------
        # Get the stuff about the loop itself
        # ----------------------------------------------------------------------
        #loop_name   = "".join(namelev) + node.attrs['name']
        loop_name   = "".join(namelev)
        loop_attrs  = retrieve_attributes(node, ['name'])
        loop_props  = retrieve_properties(node.properties)
        loop_notes  = retrieve_properties(node.notes)
        loop_labels = retrieve_labels(node.labels)
        loop_cats   = retrieve_categories(node.categories, mdd_info['categories']['shared'])

        # ----------------------------------------------------------------------
        # Construct a variable-info-like structure for the loop
        # ----------------------------------------------------------------------
        loop_info = {}
        loop_info['name']     = loop_name
        loop_info['level']    = level
        loop_info['type']     = "loop"
        loop_info['labels']   = loop_labels
        loop_info['cats']     = loop_cats
        if loop_attrs:
            loop_info['attrs']  = loop_attrs
        if loop_props:
            loop_info['props']  = loop_props
        if loop_notes:
            loop_info['notes']  = loop_notes

        loop_info['namelev'] = namelev.copy()

        loop_info['shared_lists'] = []

        # ----------------------------------------------------------------------
        # The question name is the root of the name, without any markers
        # ----------------------------------------------------------------------
        if level > 0:
            loop_info['question'] = namelev[0].replace("[..].","").replace(".","")
        else:
            loop_info['question'] = loop_info['name'].replace("[..].","").replace(".","")

        # ----------------------------------------------------------------------
        # Add the loop as an entity to the master list
        # ----------------------------------------------------------------------
        data['list'].append(loop_name)
        data['info'][loop_name] = loop_info
        if False:
            print(">>>  ", level, loop_info['type'].ljust(12),  loop_name.ljust(60), str(loop_info).encode('utf-8'))

        # ----------------------------------------------------------------------
        # Now recurse through the list of fields to get the actual variables
        # ----------------------------------------------------------------------
        for i, fnode in enumerate(node.fields):
            # ------------------------------------------------------------------
            # Call the next level down
            # ------------------------------------------------------------------
            node_data = parse_metadata_field(mdd_info, fnode, level+1, namelev, data['list'], data['info'], fields)

            # ------------------------------------------------------------------
            # Reset the name leveling system if we're done with a set of fields
            # ------------------------------------------------------------------
            while len(fields[-1]) == 0:
                fields = fields[:-1]
                namelev = namelev[:-1]
                level = len(namelev)-1

            # ------------------------------------------------------------------
            # Presuming that iteration was something real, remove it from the 
            # list.
            # ------------------------------------------------------------------
            if 'name' in fnode.attrs:
                newend = fields[-1]
                #print('lfoo', newend)
                #print('lbar', fnode.attrs['name'])
                newend.remove(fnode.attrs['name'])
        
                fields = fields[:-1]
                fields.append(newend)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Return the working name of the type number for metadata variables
# ------------------------------------------------------------------------------
def metadata_type_name(number):
    # --------------------------------------------------------------------------
    # Determine name of Dimensions type
    # --------------------------------------------------------------------------
    if number == "0":
        name = "info"
    elif number == "1":
        name = "long"
    elif number == "2":
        name = "text"
    elif number == "3":
        name = "categorical"
    elif number == "4":
        name = "XXXXXXX"
    elif number == "5":
        name = "datetime"
    elif number == "6":
        name = "double"
    elif number == "7":
        name = "binary"
    elif number == "8":
        name = "XXXXXXX"
    else:
        name = "XXXXXXX"

    # --------------------------------------------------------------------------
    # If we have one we don't recognize...
    # --------------------------------------------------------------------------
    if name == "XXXXXXX":
        print("UNDOCUMENTED DIMENSIONS TYPE DETECTED: " + number)

    # --------------------------------------------------------------------------
    # Send back the name
    # --------------------------------------------------------------------------
    return name

# ------------------------------------------------------------------------------
# Get the attributes of a tag as a simple property list, minus some we specify
# ------------------------------------------------------------------------------
def retrieve_attributes(node, ignores):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}

    # --------------------------------------------------------------------------
    # Go through each attribute present
    # --------------------------------------------------------------------------
    for item in node.attrs:
        # ----------------------------------------------------------------------
        # Skip the ones that are marked
        # ----------------------------------------------------------------------
        if item in ignores:
            continue

        # ----------------------------------------------------------------------
        # Get the value of the property and force a blank if there isn't one
        # ----------------------------------------------------------------------
        value = node.attrs[item]
        if not value:
            value = ""
        
        # ----------------------------------------------------------------------
        # Add it to the list
        # ----------------------------------------------------------------------
        data[item] = value

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Variables may have an axis expression
# ------------------------------------------------------------------------------
def retrieve_axis_exp(root_node):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}

    # --------------------------------------------------------------------------
    # If this node doesn't exist
    # --------------------------------------------------------------------------
    if root_node == None:
        return data

    # --------------------------------------------------------------------------
    # There's an expression
    # --------------------------------------------------------------------------
    data['expression'] = root_node.attrs['expression']

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Some variables may have helper fields
# ------------------------------------------------------------------------------
def retrieve_helpers(root_node):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}

    # --------------------------------------------------------------------------
    # If this node doesn't exist
    # --------------------------------------------------------------------------
    if root_node == None:
        return data

    # --------------------------------------------------------------------------
    # 
    # --------------------------------------------------------------------------

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# 
# ------------------------------------------------------------------------------
def retrieve_categories(root_node, shared_info, fdebug=False):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}
    data['list'] = []
    data['info'] = {}


    # --------------------------------------------------------------------------
    # If this node doesn't exists gtfo
    # --------------------------------------------------------------------------
    if root_node == None:
        return data

    # --------------------------------------------------------------------------
    # Get the context of the root so we can compare levels
    # --------------------------------------------------------------------------
    root_context = get_node_context(root_node)

    #if fdebug:
    #    print("------------")
    #    print(root_node)

    #print(str(root_node).encode('utf-8'))


    # --------------------------------------------------------------------------
    # Go through the categories or lists of categories in this tree
    # --------------------------------------------------------------------------
    namelev = []
    labellev = []
    for node in root_node.find_all(['categories','category']):
        # ----------------------------------------------------------------------
        # 
        # ----------------------------------------------------------------------
        #print("------", node)

        # ----------------------------------------------------------------------
        # Flag the type of thing going on here
        # ----------------------------------------------------------------------
        if node.name == "categories":
            if 'ref_name' in node.attrs:
                cat_type = "shared"
            else:
                cat_type = "header"
        else:
            cat_type = "leaf"

        # ----------------------------------------------------------------------
        # Get this node's context and calculate the offset from the root
        # ----------------------------------------------------------------------
        node_context = get_node_context(node)
        offset = node_context['level'] - root_context['level']

        # ----------------------------------------------------------------------
        # If it's a 'categories' then it's a sublist, so we want to keep track
        # of the level-based naming.  First reset the thing if it's at the first
        # level.
        # ----------------------------------------------------------------------
        if offset == 1:
            namelev = []
            labellev = []

        # ----------------------------------------------------------------------
        # It might not have a name or ID, in which case it gets skipped (this 
        # will happen when Web Author saves versioning information for instance)
        # ----------------------------------------------------------------------
        if not 'id' in node.attrs:
            continue
        if not 'name' in node.attrs:
            continue

        # ----------------------------------------------------------------------
        # Category name is the bottommost node but not necessarily unique
        # ----------------------------------------------------------------------
        catname = node.attrs['name']

        # ----------------------------------------------------------------------
        # Full name is the name levels and the name
        # ----------------------------------------------------------------------
        fullname = ".".join(namelev)
        if fullname:
            fullname += "."
        fullname += catname

        #print("--------------")
        #print(node)
        #print(node.name)
        #print(node.attrs)
        #print(node_context)
        #print(offset)
        #print(fullname)
        #print(cat_type)

        # ----------------------------------------------------------------------
        # A shared list includes pieces from the already-defined set
        # ----------------------------------------------------------------------
        if cat_type == "shared":
            slname = node.attrs['ref_name']
            while slname.startswith("\\"):
                slname = slname[1:]
            while slname.startswith("."):
                slname = slname[1:]
            slname = slname.strip()
            sltag  = slname.lower()
            #print(slname, sltag)

            for cname in shared_info[sltag]['cats']['list']:
                cinfo = shared_info[sltag]['cats']['info'][cname]
                data['list'].append(cname)
                data['info'][cname] = cinfo

        # ----------------------------------------------------------------------
        # a category heading or a category gets added to the list
        # ----------------------------------------------------------------------
        else:
            # ------------------------------------------------------------------
            # Compile the information about the category
            # ------------------------------------------------------------------
            cat_info = {}
            cat_info['offset']   = offset
            cat_info['cat_type'] = cat_type
            cat_info['namelev']  = namelev.copy()
            cat_info['labellev'] = labellev.copy()
            cat_info['catname']  = catname
            cat_info['fullname'] = fullname
            cat_info['string']   = node.string
            cat_info['id']       = node.attrs['id']
            cat_info['attrs']    = {}
    
            for attr in node.attrs:
                if attr in ['id', 'name']:
                    continue
                cat_info['attrs'][attr] = node.attrs[attr]
    
            if node.properties:
                cat_info['props'] = retrieve_properties(node.properties)
            else:
                cat_info['props'] = {}
    
            if node.labels:
                cat_info['labels'] = retrieve_labels(node.labels)
            else:
                cat_info['labels'] = {}
    
            # ------------------------------------------------------------------
            # Store it in the running list
            # ------------------------------------------------------------------
            data['list'].append(fullname)
            data['info'][fullname] = cat_info

        # ----------------------------------------------------------------------
        # For headings, add to the leveled naming and labels for the next thing
        # ----------------------------------------------------------------------
        if cat_type == "header":
            namelev.append(node.attrs['name'])
            labellev.append(retrieve_labels(node.labels))
                
    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Store a dictionary of labels by language and context
# ------------------------------------------------------------------------------
def retrieve_labels(root_node):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    data = {}
    data['list'] = []
    data['info'] = {}

    # --------------------------------------------------------------------------
    # If this node doesn't exists
    # --------------------------------------------------------------------------
    if root_node == None:
        return data

    # --------------------------------------------------------------------------
    # Go through each text node found in this part of the tree
    # --------------------------------------------------------------------------
    for node in root_node.find_all("text"):
        # ----------------------------------------------------------------------
        # Each text piece will have a context and a language
        # ----------------------------------------------------------------------
        context = node.attrs['context']
        lang    = node.attrs['xml:lang']

        # ----------------------------------------------------------------------
        # Those properties will form a unique key for this list
        # ----------------------------------------------------------------------
        key = (context, lang)

        # ----------------------------------------------------------------------
        # Add this one to the running order list and the info
        # ----------------------------------------------------------------------
        data['list'].append(key)
        data['info'][key] = node.string

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data

# ------------------------------------------------------------------------------
# Store a dictionary of property names that might be nested...
# The structure will always start with a 'properties' node, which may
# contain individual leaf-node 'property' tags, or other lists of 'properties'
# ------------------------------------------------------------------------------
def retrieve_properties(root_node):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    props = {}

    # --------------------------------------------------------------------------
    # If this node doesn't exists
    # --------------------------------------------------------------------------
    if root_node == None:
        return props

    # --------------------------------------------------------------------------
    # Get the context of the root so we can compare levels
    # --------------------------------------------------------------------------
    root_context = get_node_context(root_node)

    # --------------------------------------------------------------------------
    # Go through every node in this tree
    # --------------------------------------------------------------------------
    namelev = []
    for node in root_node.descendants:
        # ----------------------------------------------------------------------
        # Only tags 
        # ----------------------------------------------------------------------
        if type(node) == bs4.element.NavigableString:
            continue

        # ----------------------------------------------------------------------
        # Skip old version stuff that web author puts in
        # ----------------------------------------------------------------------
        if node.name == "versions":
            continue
        if node.name == "DIFF":
            continue

        # ----------------------------------------------------------------------
        # Get this node's context and calculate the offset from the root
        # ----------------------------------------------------------------------
        node_context = get_node_context(node)
        offset = node_context['level'] - root_context['level']

        #print("--------------")
        #print(node)
        #print(node.name)
        #print(node_context)
        #print(offset)
        #print(node.attrs)

        # ----------------------------------------------------------------------
        # Take care of the 'namelev', which holds any prefixes for the name
        # ----------------------------------------------------------------------
        if offset == 1:
            namelev = []
        elif (offset % 2 == 0) and (len(namelev) < offset / 2) :
            namelev.append(node.find_parent().attrs['name'])
            continue

        #print(namelev)
                
        # ----------------------------------------------------------------------
        # Make the property name, including any prefixes
        # ----------------------------------------------------------------------
        propname = ".".join(namelev) 
        if propname:
            propname += "." 
        propname += node.attrs['name']

        # ----------------------------------------------------------------------
        # Store the property
        # ----------------------------------------------------------------------
        props[propname] = {}
        if 'context' in node.attrs:
            props[propname]['context'] = node.attrs['context']
        if 'type' in node.attrs:
            props[propname]['type']    = node.attrs['type']
        if 'value' in node.attrs:
            props[propname]['value']   = node.attrs['value']

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return props


# ------------------------------------------------------------------------------
# For any node, return some useful contextual information 
#    - list of tags leading up to this node
#    - Number of descendents
#    - Number of levels of descendents
#    - List of tags of descendents
# ------------------------------------------------------------------------------
def get_node_context(node):
    # --------------------------------------------------------------------------
    # 
    # --------------------------------------------------------------------------

    # --------------------------------------------------------------------------
    # List of tags in the hierarchy
    # --------------------------------------------------------------------------
    taglist = [p.name for p in node.parentGenerator() if p][::-1]

    # --------------------------------------------------------------------------
    # Put all of those into one thing
    # --------------------------------------------------------------------------
    node_info = {}
    node_info['taglist'] = taglist
    node_info['level']   = len(taglist)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return node_info



# ------------------------------------------------------------------------------
# Parse a routing script, already a list of lines
# ------------------------------------------------------------------------------
def parse_routing_script(script_as_list):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    rtdata = {}

    # --------------------------------------------------------------------------
    # Clean up and label the lines of script
    # --------------------------------------------------------------------------
    rtdata['sections'] = []
    rtdata['script'] = []
    rtdata['funclist'] = []
    rtdata['sublist'] = []
    rtdata['functions'] = {}
    rtdata['subs'] = {}
    rtdata['code_blocks'] = []
    rtdata['comment_heads'] = {}

    block_comment = False
    section = "BEGIN"
    code_type = "BARE"
    code_name = "MAIN"
    lvl  = 0
    conditions  = []
    sel_var = ""
    current_block = []
    prev_comment = False
    for i, line in enumerate(script_as_list):
        # ----------------------------------------------------------------------
        # Ignore white space
        # ----------------------------------------------------------------------
        if len(line.strip()) == 0:
            continue

        #print("@@ ", i, block_comment, prev_comment, line)

        # ----------------------------------------------------------------------
        # Start block comment
        # ----------------------------------------------------------------------
        if line.strip().startswith("'!"):
            if not prev_comment:
                if current_block:
                    current_block.append(line.strip())
                    rtdata['code_blocks'].append(current_block)
                current_block = []
                #print("=========================================START BLOCK")
                prev_comment = True

            current_block.append(line.strip())
            block_comment = True
            continue

        # ----------------------------------------------------------------------
        # Exit block comment
        # ----------------------------------------------------------------------
        if line.strip().endswith("!'"):
            #print("=========================================END BLOCK")
            current_block.append(line.strip())
            block_comment = False
            prev_comment = True
            continue

        # ----------------------------------------------------------------------
        # Ignore interior of block comment
        # ----------------------------------------------------------------------
        if block_comment:
            current_block.append(line.strip())
            prev_comment = True
            continue

        # ----------------------------------------------------------------------
        # Ignore whole-line comments
        # ----------------------------------------------------------------------
        if line.strip().startswith("'"):
            if not prev_comment:
                if current_block:
                    rtdata['code_blocks'].append(current_block)
                current_block = []
                #print("=========================================")

            current_block.append(line.strip())
            prev_comment = True

            continue

        # ----------------------------------------------------------------------
        # We can keep a set of just the comment headers...this will trigger
        # at the end of each comment section, right as the first line of code
        # is reached for each block, and just store what we have so far...
        # ----------------------------------------------------------------------
        if i > 0 and prev_comment:
            #print("COMMENT", i, current_block)
            #print("---------------------------------------------")
            code_block_id = len(rtdata['code_blocks'])
            rtdata['comment_heads'][code_block_id] = current_block

        # ----------------------------------------------------------------------
        # Clean up the line
        # ----------------------------------------------------------------------
        line = line.strip()
        current_block.append(line)
        prev_comment = False

        # ----------------------------------------------------------------------
        # Flag current section if it changes
        # Ignore 'FOO:', as it is not a real section but movable.
        # ----------------------------------------------------------------------
        if line.endswith(":") and not line.strip() == "FOO:":
            section = line[:-1]
            rtdata['sections'].append(section)

        # ----------------------------------------------------------------------
        # Check for the presense of single and double quotes 
        # ----------------------------------------------------------------------
        dq = line.count('"')
        sq = line.count("'") 

        # ----------------------------------------------------------------------
        # Remove any in-line comments...it might be any line with a single quote
        # ----------------------------------------------------------------------
        if sq > 0:
            # ------------------------------------------------------------------
            # Cut the line only up to the first single quote
            # ------------------------------------------------------------------
            first = line.find("'")
            subline = line[0:first].strip()

            # ------------------------------------------------------------------
            # Get the number of double quotes in that sub line
            # ------------------------------------------------------------------
            sdq = subline.count('"')

            # ------------------------------------------------------------------
            # If there are none or an even number of double quotes up to that
            # point, then the single quote is in the bare code and denotes a
            # comment...otherwise, it's part of some text string and we leave it
            # ------------------------------------------------------------------
            if sdq % 2 == 0:
                line = subline

        # ----------------------------------------------------------------------
        # Find stuff about conditional logic, and use it to set the 'level'
        # that this line is at, and a little visual indicator of movement
        # ----------------------------------------------------------------------
        new_if      = int(line.lower().startswith("if "))
        else_if     = int(line.lower().startswith("elseif"))
        or_else     = int(line.lower() == "else")
        then_end    = int(line.lower().endswith("then"))
        then_in     = int("then" in line.lower() and (not then_end))
        end_if      = int(line.lower() == "end if")
        select      = int(line.lower().startswith("select case"))
        new_case    = int(line.lower().startswith("case "))
        end_sel     = int(line.lower().startswith("end select"))
        for_start   = int(line.lower().startswith("for "))
        for_next    = int(line.lower() == "next")
        fnc_start   = int(line.lower().startswith("function "))
        fnc_end     = int(line.lower() == "end function")
        sub_start   = int(line.lower().startswith("sub "))
        sub_end     = int(line.lower() == "end sub")

        move = "-"
        if new_if and not then_in:
            move = ">"
            lvl += 1
            cond = line[3:-4].strip()
            conditions.append(cond)

        elif new_if and then_in:
            move = "><"

        elif else_if:
            move = "S"
            cond = line[7:-4].strip()
            conditions = conditions[:-1]
            conditions.append(cond)
        
        elif or_else:
            move = "S"
            cond = "else"
            conditions = conditions[:-1]
            conditions.append(cond)

        elif end_if:
            move = "<"
            lvl -= 1
            conditions = conditions[:-1]

        elif select:
            move = ">"
            lvl += 1
            conditions.append(line)
            sel_var = line[12:].strip()

        elif new_case:
            move = "S"
            conditions = conditions[:-1]
            cond = sel_var + " = " + line
            conditions.append(cond)

        elif end_sel:
            move = "<"
            lvl -= 1
            conditions = conditions[:-1]
            sel_var = ""

        elif for_start:
            move = "8"
            lvl += 1
            conditions.append(line)

        elif for_next:
            move = "6"
            lvl -= 1
            conditions = conditions[:-1]

        elif fnc_start:
            move = ">"
            lvl = 1

        elif fnc_end:
            move = "<"
            lvl -= 1

        elif sub_start:
            move = ">"
            lvl = 1

        elif sub_end:
            move = "<"
            lvl -= 1

        # ----------------------------------------------------------------------
        # Find asks and shows
        # ----------------------------------------------------------------------
        ask  = line.lower().endswith("ask()")
        show = line.lower().endswith("show()")

        # ----------------------------------------------------------------------
        # Wrap up the line data together
        # ----------------------------------------------------------------------
        this_line = {}
        this_line['line']    = line
        this_line['sect']    = section
        this_line['lvl']     = lvl
        this_line['move']    = move
        this_line['cond']    = conditions
        this_line['ask']     = ask
        this_line['show']    = show
        this_line['block']   = len(rtdata['code_blocks'])

        #print(line.ljust(80), refs, acts)

        # ----------------------------------------------------------------------
        # Figure out if we are in bare script, function, or sub, and put the
        # line data in the right place
        # ----------------------------------------------------------------------
        if   line.lower().startswith("function "):
            code_type                 = "FUNCTION"
            name                      = line.lower()
            rtdata['functions'][name] = []
            rtdata['functions'][name].append(this_line)
            rtdata['funclist'].append(name)

        elif line.lower().startswith("end function"):
            rtdata['functions'][name].append(this_line)
            code_type                 = "BARE"
            name                      = "MAIN"

        elif line.lower().startswith("sub "):
            code_type                 = "SUB"
            name                      = line.lower()
            rtdata['subs'][name]      = []
            rtdata['subs'][name].append(this_line)
            rtdata['sublist'].append(name)

        elif line.lower().startswith("end sub"):
            rtdata['subs'][name].append(this_line)
            code_type                 = "BARE"
            name                      = "MAIN"

        elif code_type == "FUNCTION":
            rtdata['functions'][name].append(this_line)
        elif code_type == "SUB":
            rtdata['subs'][name].append(this_line)
        elif code_type == "BARE":
            rtdata['script'].append(this_line)
        else:
            print("WARNING!  UNCLASSIFIED LINE OF ROUTING:", line)
            
    # --------------------------------------------------------------------------
    # Write the thing to the log if we are in debug
    # --------------------------------------------------------------------------
    #if debug and True:

    #    print("="*100, file=lout)
    #    print("ORGANIZED ROUTING SCRIPT", file=lout)
    #    print("="*100, file=lout)
    #    for i, line in enumerate(rtdata['script']):
    #        print(str(i+1).ljust(5), line['move'].ljust(3), str(line['lvl']).ljust(2), "|", "  "*line['lvl'], line['line'], file=lout)

    #    print("="*100, file=lout)
    #    print("ORGANIZED ROUTING FUNCTIONS", file=lout)
    #    print("="*100, file=lout)
    #    for func in rtdata['functions']:
    #        print(" ", file=lout)
    #        print(func, file=lout)
    #        print("--------------", file=lout)
    #        for i, line in enumerate(rtdata['functions'][func]):
    #            print(str(i+1).ljust(5), line['move'].ljust(3), str(line['lvl']).ljust(2), "|", "  "*line['lvl'], line['line'], file=lout)

    #    print("="*100, file=lout)
    #    print("ORGANIZED ROUTING SUBS", file=lout)
    #    print("="*100, file=lout)
    #    for sub in rtdata['subs']:
    #        print(" ", file=lout)
    #        print(sub, file=lout)
    #        print("--------------", file=lout)
    #        for i, line in enumerate(rtdata['subs'][sub]):
    #            print(str(i+1).ljust(5), line['move'].ljust(3), str(line['lvl']).ljust(2), "|", "  "*line['lvl'], line['line'], file=lout)

    # --------------------------------------------------------------------------
    # Take another pass through and sort out references and actions
    # --------------------------------------------------------------------------
    for i, line in enumerate(rtdata['script']):
        # ----------------------------------------------------------------------
        # Make a simple list of all the elements that are referenced at all
        # Make a dictionary with a key of the element and a value of the action
        # Make a simple list of all of the actions performed
        # ----------------------------------------------------------------------
        elements  = []
        actions   = []
        events    = {}

        # ----------------------------------------------------------------------
        # Tag lines that start sections, that'll be all there is
        # ----------------------------------------------------------------------
        if line['line'].endswith(":"):
            elements.append(line['sect'])
            actions.append("new section")
            continue

        # ----------------------------------------------------------------------
        # End If is also a single-use line
        # ----------------------------------------------------------------------
        if line['line'].lower() == "end if":
            actions.append("end-if L" + str(line['lvl']))
            continue

        # ----------------------------------------------------------------------
        # Tag lines that we know are part of conditional statements
        # ----------------------------------------------------------------------
        if line['move'] in [">","><","S","<","8","6"]:
            actions.append("conditional")

        # ----------------------------------------------------------------------
        # Parse out what all is on the line
        # ----------------------------------------------------------------------

        # ----------------------------------------------------------------------
        # 
        # ----------------------------------------------------------------------
        #print(line['line'].ljust(100), elements, actions)

        # ----------------------------------------------------------------------
        # Add these to the ongoing set
        # ----------------------------------------------------------------------
        rtdata['script'][i]['elems'] = elements
        rtdata['script'][i]['acts']  = actions
        rtdata['script'][i]['refs']  = events

    #for line in rtdata['script']:
    #    print(line['line'].ljust(100), line['refs'], line['acts'])

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return rtdata

# ------------------------------------------------------------------------------
# Given a parsed routing script, put together an index of the sections/blocks
# ------------------------------------------------------------------------------
def compile_routing_script_directory(script_info):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    rtdir = {}
    
    # --------------------------------------------------------------------------
    # Known dividers from shell
    # --------------------------------------------------------------------------
    rtdir["PREFIX"]                            = []
    rtdir["BEGINNING_OF_PROGRAM"]              = []
    rtdir["BEGIN_SCREENER"]                    = []
    rtdir["END_OF_SCREENER_SET_TIMING_VARS"]   = []
    rtdir["BEGIN_MAIN"]                        = []
    rtdir["END_MAIN_TIMING_VARS_QC_TERMS"]     = []

    # --------------------------------------------------------------------------
    # Go through each line of the script 
    # --------------------------------------------------------------------------
    big_section = "PREFIX"
    for i, line in enumerate(script_info['script']):
        # ----------------------------------------------------------------------
        # Determine which broad section they are in
        # ----------------------------------------------------------------------
        if line['sect'].upper() == "BEGINNING_OF_PROGRAM":
            big_section = line['sect']

        elif line['sect'].upper() == "BEGIN_SCREENER":
            big_section = line['sect']

        elif line['sect'].upper() == "END_OF_SCREENER_SET_TIMING_VARS":
            big_section = line['sect']

        elif line['sect'].upper() == "BEGIN_MAIN":
            big_section = line['sect']

        elif line['sect'].upper() == "END_MAIN_TIMING_VARS_QC_TERMS":
            big_section = line['sect']

        # ----------------------------------------------------------------------
        # Compiling a list of section, block, line for each instance in a big
        # section and add it to the appropriate list
        # ----------------------------------------------------------------------
        slug = (line['sect'], line['block'], i)
        rtdir[big_section].append(slug)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return rtdir





# ------------------------------------------------------------------------------
# Parse a ddf
# ------------------------------------------------------------------------------
def parse_ddf(ddf_file):
    # --------------------------------------------------------------------------
    # If the file doesn't exist, ditch
    # --------------------------------------------------------------------------
    if not os.path.isfile(ddf_file):
        return None

    # --------------------------------------------------------------------------
    # User info
    # --------------------------------------------------------------------------
    print("...parsing ddf file " + ddf_file + "...")

    # --------------------------------------------------------------------------
    # Make the db connection
    # --------------------------------------------------------------------------
    conn = sqlite3.connect(ddf_file)
    curs = conn.cursor()

    # --------------------------------------------------------------------------
    # Pull in a copy of each table as a pandas dataframe
    # --------------------------------------------------------------------------
    curs.execute("SELECT * FROM sqlite_master WHERE type='table';")

    tables = {}
    for item in curs.fetchall():
        # ----------------------------------------------------------------------
        # That gives us the name and schema
        # ----------------------------------------------------------------------
        tablename = item[1]
        tableschema = item[4]

        # ----------------------------------------------------------------------
        # Grab a sample of the data
        # ----------------------------------------------------------------------
        data500 = pd.read_sql_query("SELECT * FROM " + tablename + " LIMIT 500;", conn)

        # ----------------------------------------------------------------------
        # Get the raw row count
        # ----------------------------------------------------------------------
        tablecount = int(pd.read_sql_query("SELECT count(*) FROM " + tablename, conn).values[0][0])

        # ----------------------------------------------------------------------
        # There's always a table called 'SchemaVersion' with one row
        # ----------------------------------------------------------------------
        if tablename == "SchemaVersion":
            schema_version = data500.iloc[0].to_dict()
            continue

        # ----------------------------------------------------------------------
        # There's always a table called 'DataVersion' with one row
        # ----------------------------------------------------------------------
        if tablename == "DataVersion":
            data_version = data500.iloc[0].to_dict()
            continue

        # ----------------------------------------------------------------------
        # There's always a table called 'Levels', and we want to get the whole
        # thing.
        # ----------------------------------------------------------------------
        if tablename == "Levels":
            levels = pd.read_sql_query("SELECT * FROM " + tablename + ";", conn)
            levels = levels.set_index('TableName')
            continue

        # ----------------------------------------------------------------------
        # Parse out the schema into a useful format
        # ----------------------------------------------------------------------
        schema_info = parse_ddf_schema(tableschema)

        # ----------------------------------------------------------------------
        # Put all of this table-level information together
        # ----------------------------------------------------------------------
        tables[tablename] = {}
        tables[tablename]['schema']      = tableschema
        tables[tablename]['schema_info'] = schema_info
        tables[tablename]['data500']     = data500
        tables[tablename]['count']       = tablecount

    # --------------------------------------------------------------------------
    # Close it down
    # --------------------------------------------------------------------------
    curs.close()
    conn.close()

    # --------------------------------------------------------------------------
    # Add data from the levels table to the table info
    # --------------------------------------------------------------------------
    for table in tables:
        leveldata = levels.loc[table].to_dict()

        tables[table]['table_number'] = leveldata['TableNumber']
        tables[table]['parent_name'] = leveldata['ParentName']
        tables[table]['table_name'] = leveldata['DSCTableName']


    # --------------------------------------------------------------------------
    # Consolidate
    # --------------------------------------------------------------------------
    data = {}
    data['tables']         = tables
    data['schema_version'] = schema_version
    data['data_version']   = data_version
    data['levels']         = levels

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data


# ------------------------------------------------------------------------------
# Make a text version of the parsed metadata as a list of text fields
# ------------------------------------------------------------------------------
def make_metadata_text_listing(mdd_data):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    block = []

    # --------------------------------------------------------------------------
    # The HDATA properties
    # --------------------------------------------------------------------------
    block.append("HDATA")
    block.append("[")

    width = get_longest([x for x in mdd_data['properties']])

    for i, prop in enumerate(mdd_data['properties']):
        # ----------------------------------------------------------------------
        # Get the property value
        # ----------------------------------------------------------------------
        if 'value' in mdd_data['properties'][prop]:
            prop_val = mdd_data['properties'][prop]['value']
        else:
            prop_val = ""

        # ----------------------------------------------------------------------
        # write the text slug
        # ----------------------------------------------------------------------
        slug = prop.ljust(width) + "  = " + '"' + prop_val + '"'

        if i < len(mdd_data['properties'])-1:
            slug += ","

        # ----------------------------------------------------------------------
        # Add to the block
        # ----------------------------------------------------------------------
        block.append(slug)

    block.append("];")
    block.append("")

    # --------------------------------------------------------------------------
    # The variables from the master list
    # --------------------------------------------------------------------------
    for i, var in enumerate(mdd_data['master']['list']):
        # ----------------------------------------------------------------------
        # Get the info
        # ----------------------------------------------------------------------
        info = mdd_data['master']['info'][var]
        vname   = info['name']
        vtype   = info['type']
        vlevel  = info['level']
        vlabels = info['labels']

        if 'attrs' in info:
            vattrs = info['attrs']
        else:
            vprops = {}

        if 'props' in info:
            vprops = info['props']
        else:
            vprops = {}

        if 'cats' in info:
            vcats = info['cats']
        else:
            vcats = {}

        if 'notes' in info:
            vnotes = info['notes']
        else:
            vnotes = {}

        if 'styles' in info:
            vstyles = info['styles']
        else:
            vstyles = {}

        if 'helper' in info:
            vhelper = info['helper']
        else:
            vhelper = {}

        if 'axis' in info:
            vaxis = info['axis']
        else:
            vaxis = {}

        # ----------------------------------------------------------------------
        # The indentation is a number of spaces as a factor of the level
        # ----------------------------------------------------------------------
        indent = vlevel * 4

        # ----------------------------------------------------------------------
        # Heading
        # ----------------------------------------------------------------------
        block.append(" "*indent + "'| " + "-"*(97-indent))
        block.append(" "*indent + "'| " + vtype + ":  " + var)
        block.append(" "*indent + "'| " + "-"*(97-indent))

        # ----------------------------------------------------------------------
        # Notes, which apparently don't store with the tic mark
        # ----------------------------------------------------------------------
        if vnotes:
            block.append("")
            for item in vnotes:
                note = vnotes[item]['value']
                block.append(" "*indent + "notes|" + var + "|" + "'" + str(note))

        # ----------------------------------------------------------------------
        # Variable name
        # ----------------------------------------------------------------------
        block.append("")
        block.append(" "*indent + "vname|" + var + "|" + vname)

        # ----------------------------------------------------------------------
        # Labels
        # ----------------------------------------------------------------------
        width_context  = get_longest([x[0] for x in vlabels['list']])
        width_language = get_longest([x[1] for x in vlabels['list']])

        block.append("")
        for label in vlabels['list']:
            context  = label[0]
            language = label[1]
            text     = vlabels['info'][label]

            if not text:
                text = ""

            text = text.replace("\n","\\n")
            text = text.replace("\t","\\t")

            slug = " "*indent + "label|" + var + "|" + context.ljust(width_context) + " / " + language.ljust(width_language) + " : " + '"' + text + '"'
            block.append(slug)

        # ----------------------------------------------------------------------
        # Variable-level properties, also show attributes
        # ----------------------------------------------------------------------
        if vprops or vattrs:
            # ------------------------------------------------------------------
            # Get the widest for formatting
            # ------------------------------------------------------------------
            width_prop = max([get_longest([x for x in vprops]), get_longest([x for x in vattrs])])

            # ------------------------------------------------------------------
            # Header
            # ------------------------------------------------------------------
            block.append("")
            block.append(" "*indent + "props|" + var + "|" + "[")

            # ------------------------------------------------------------------
            # Start an overall counter
            # ------------------------------------------------------------------
            i = 0

            # ------------------------------------------------------------------
            # Each attribute
            # ------------------------------------------------------------------
            for prop in vattrs:
                value = vattrs[prop]

                if not value:
                    value = ""

                slug = " "*indent + "props|" + var + "|" + "    " + prop.ljust(width_prop) + " = " + value

                if i < len(vprops) + len(vattrs) - 1:
                    slug += ","

                block.append(slug)

                i+=1

            # ------------------------------------------------------------------
            # Each property
            # ------------------------------------------------------------------
            for prop in vprops:

                value = vprops[prop]['value']

                if not value:
                    value = ""

                slug = " "*indent + "props|" + var + "|" + "    " + prop.ljust(width_prop) + " = " + value

                if i < len(vprops) + len(vattrs) - 1:
                    slug += ","

                block.append(slug)

                i+=1

            # ------------------------------------------------------------------
            # Footer
            # ------------------------------------------------------------------
            block.append(" "*indent + "props|" + var + "|" + "]")

        # ----------------------------------------------------------------------
        # Type and constraints
        # ----------------------------------------------------------------------
        block.append("")

        if 'min' in vattrs:
            vmin = vattrs['min']
        else:
            vmin = ""

        if 'max' in vattrs:
            vmax = vattrs['max']
        else:
            vmax = ""

        slug = " "*indent + "vtype|" + var + "|" + vtype + " [" + str(vmin) + ".." + str(vmax) + "]"

        block.append(slug)

        # ----------------------------------------------------------------------
        # Categories
        # ----------------------------------------------------------------------
        if vcats:
            # ------------------------------------------------------------------
            # Header
            # ------------------------------------------------------------------
            block.append("")
            block.append(" "*indent + "vcats|" + var + "|" + "{")

            # ------------------------------------------------------------------
            # Longest category
            # ------------------------------------------------------------------
            width_cat = get_longest([x for x in vcats['list']])

            # ------------------------------------------------------------------
            # Longest full-named category, taking into account named levels
            # ------------------------------------------------------------------
            width_full = 0
            for vcat in vcats['list']:
                fullname = vcats['info'][vcat]['fullname']
                if len(fullname) > width_full:
                    width_full = len(fullname)

            # ------------------------------------------------------------------
            # Each category
            # ------------------------------------------------------------------
            for vcat in vcats['list']:
                # --------------------------------------------------------------
                # Get the stuff for this category
                # --------------------------------------------------------------
                catinfo = vcats['info'][vcat]
                fullname = catinfo['fullname']

                # --------------------------------------------------------------
                # This may take up several lines so it has its own block
                # --------------------------------------------------------------
                cblock = []

                # --------------------------------------------------------------
                # Make the indexing prefix to use on each line
                # --------------------------------------------------------------
                prefix = " "*indent + "vcats|" + var + "|" + fullname.ljust(width_full) + "|"

                # --------------------------------------------------------------
                # The first line is the name of the variable and the default
                # label (what was in the text node of the category)
                # --------------------------------------------------------------
                lblstring = catinfo['string']
                if not lblstring:
                    lblstring = ""
                lblstring = lblstring.replace("\n","\\n")
                lblstring = lblstring.replace("\t","\\t")

                if catinfo['cat_type'] == "header":
                    typetag = "sub"
                elif catinfo['cat_type'] == "leaf":
                    typetag = "cat"

                cblock.append(prefix + typetag + "|" + vcat.ljust(width_cat + 4) + '"' + lblstring + '"')

                # --------------------------------------------------------------
                # The labels, but only show if there is more than one...if it
                # is the only one, we already show it with the cat string
                # It is possible to have no labels, we just skip
                # --------------------------------------------------------------
                if catinfo['labels'] and len(catinfo['labels']['list']) > 1:
                    width_context  = get_longest([x[0] for x in catinfo['labels']['list']])
                    width_language = get_longest([x[1] for x in catinfo['labels']['list']])

                    for ltag in catinfo['labels']['list']:
                        context = ltag[0]
                        language = ltag[1]
                        text = catinfo['labels']['info'][ltag]
                        if not text:
                            text = ""
                        text = text.replace("\n","\\n")
                        text = text.replace("\t","\\t")

                        cblock.append(prefix + "lbl" + "|" + " "*(width_cat+4) + context.ljust(width_context) + " / " + language.ljust(width_language) + " : " + '"' + text + '"')

                # --------------------------------------------------------------
                # Attributes and properties
                # --------------------------------------------------------------
                if catinfo['attrs']:
                    for att in catinfo['attrs']:
                        value = catinfo['attrs'][att]
                        if not value:
                            value = ""
                        slug = att + " = " + value
                        cblock.append(prefix + "att" + "|" + " "*(width_cat+4) + "[ " + slug + " ]")

                if catinfo['props']:
                    for prp in catinfo['props']:
                        if 'value' in catinfo['props'][prp]:
                            value = catinfo['props'][prp]['value']
                        else:
                            value = ""
                        if not value:
                            value = ""
                        slug = prp + " = " + value
                        cblock.append(prefix + "prp" + "|" + " "*(width_cat+4) + "[ " + slug + " ]")

                # --------------------------------------------------------------
                # If the category was multiline, give some whitespace
                # --------------------------------------------------------------
                if (catinfo['labels'] and len(catinfo['labels']['list']) > 1) or catinfo['attrs'] or catinfo['props']:
                    cblock.append("")

                # --------------------------------------------------------------
                # --------------------------------------------------------------
                if True:
                    for item in catinfo:
                        if item in ["offset","namelev","cat_type","catname","fullname","string","id","labels","labellev","attrs","props"]:
                            continue
                        if catinfo[item]:
                            cblock.append(prefix + "    " + item.ljust(20) + str(catinfo[item]))

                # --------------------------------------------------------------
                # Tack these lines on to the running set
                # --------------------------------------------------------------
                block.extend(cblock)

            # ------------------------------------------------------------------
            # Footer
            # ------------------------------------------------------------------
            block.append(" "*indent + "vcats|" + var + "|" + "}")

        # ----------------------------------------------------------------------
        # Styles
        # ----------------------------------------------------------------------
        if vstyles:
            block.append("")
            block.append(" "*indent + "vstyles:  " + str(vstyles) )

        # ----------------------------------------------------------------------
        # Helper fields
        # ----------------------------------------------------------------------
        if vhelper:
            block.append("")
            block.append(" "*indent + "vhelper:  " + str(vhelper) )

        # ----------------------------------------------------------------------
        # Axis expression
        # ----------------------------------------------------------------------
        if vaxis:
            block.append("")
            block.append(" "*indent + "vaxis:    " + str(vaxis) )

        # ----------------------------------------------------------------------
        # whitespace at the end
        # ----------------------------------------------------------------------
        block.append("")

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return block

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
def parse_tables_from_csv(tables_file, table_delim="|"):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    print("...parsing table file '" + tables_file + "'...")

    # --------------------------------------------------------------------------
    # Get the contents of the document
    # --------------------------------------------------------------------------
    document = []
    file_raw  = open(tables_file, encoding='utf-16')
    file_csv  = csv.reader(file_raw, delimiter = '\t')
    for line in file_csv:
        document.append(line)

    # --------------------------------------------------------------------------
    # Break the thing up into its component tables...
    # --------------------------------------------------------------------------
    tables = []
    this_table = []
    for line in document:
        if len(line) == 1 and line[0] == table_delim:
            tables.append(this_table)
            this_table = []
        else:
            this_table.append(line)

    # --------------------------------------------------------------------------
    # Extract info from individual tables
    # --------------------------------------------------------------------------
    table_info = []
    for table in tables:
        table_info.append(parse_table(table))

    # for item in table_info[0]:
    #     print(item.ljust(20), table_info[0][item])

    # --------------------------------------------------------------------------
    # Make an index of different things
    # --------------------------------------------------------------------------
    index = {}

    # --------------------------------------------------------------------------
    # Consolidate info
    # --------------------------------------------------------------------------
    tabs_info = {}
    tabs_info['tables'] = table_info
    tabs_info['index']  = index

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return tabs_info


# ------------------------------------------------------------------------------
# Parse an individual table given as a list-of-lists
# ------------------------------------------------------------------------------
def parse_table(table, fdebug=True):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    table_info = {}

    # --------------------------------------------------------------------------
    # Parse the lines of the table
    # --------------------------------------------------------------------------
    lines = []
    section = "header"
    subtype = "UNKNOWN"
    for line in table:
                                                                                                     
        # ----------------------------------------------------------------------
        # Initialize and the line itself
        # ----------------------------------------------------------------------
        info = {}
        info['line'] = line
                                                                                                     
        # ----------------------------------------------------------------------
        # Clean text of html garbage and stuff
        # ----------------------------------------------------------------------
        for j, cell in enumerate(info['line']):
            if cell:
                if cell[-5:] == "<BR/>":
                    info['line'][j]=cell.replace("<BR/>", "")
                elif "<BR/>" in cell:
                    info['line'][j]=cell.replace("<BR/>", " - ")
                                                                                                     
                if "&amp;" in info['line'][j]:
                    info['line'][j] = info['line'][j].replace("&amp;", "&")
                
                                                                                                     
        # ----------------------------------------------------------------------
        # how many columns
        # ----------------------------------------------------------------------
        info['len'] = len(line)
                                                                                                     
        # ----------------------------------------------------------------------
        # We'll track the presence or absence of stuff in the first two
        # columns, which is how we parse the sections
        # first do blanks 
        # ----------------------------------------------------------------------
        if len(line) == 0:
            info['colA'] = False
            info['colB'] = False
            info['type'] = "00"
                                                                                                     
        # ----------------------------------------------------------------------
        # Next do if there's only one column
        # ----------------------------------------------------------------------
        elif len(line) == 1:
            if len(str(line[0]).strip()) > 0:
                info['colA'] = True
                info['colB'] = False
                info['type'] = "10"
            else:
                info['type'] = "00"
                info['colB'] = False
                info['colA'] = False
                                                                                                     
        # ----------------------------------------------------------------------
        # If there's more than one column we can do both A and B
        # ----------------------------------------------------------------------
        elif len(line) > 1:
            # ------------------------------------------------------------------
            # column A
            # ------------------------------------------------------------------
            if len(str(line[0]).strip()) > 0:
                info['colA'] = True
            else:
                info['colA'] = False
                                                                                                     
            # ------------------------------------------------------------------
            # column B
            # ------------------------------------------------------------------
            if len(str(line[1]).strip()) > 0:
                info['colB'] = True
            else:
                info['colB'] = False
                                                                                                     
            # ------------------------------------------------------------------
            # Set the type based on the presence or absence of stuff in
            # the first two cells
            # ------------------------------------------------------------------
            if info['colA'] == True and info['colB'] == False:
                info['type'] = "10"
            elif info['colA'] == False and info['colB'] == True:
                info['type'] = "01"
            elif info['colA'] == True and info['colB'] == True:
                info['type'] = "11"
            elif info['colA'] == False and info['colB'] == False:
                info['type'] = "00"
                                                                                                     
                                                                                                     
        # ----------------------------------------------------------------------
        # Set the section if it's changing
        # ----------------------------------------------------------------------
        if section == "header" and info['type'] == "01":
            section = "banner"
            banline = 0
        if section == "header" and info['type'] == '11':
            section = 'stubs'
        if section == "banner" and info['type'] in ["10","11"]:
            section = "stubs"
        if section == "stubs" and info['type'] == "10" and info['line'][0][0:13] == "Cell Contents":
            section = "footer"
                                                                                                     
        # ----------------------------------------------------------------------
        # Set the subtype - header
        # ----------------------------------------------------------------------
        if section == "header":
            if info['type'] == "10" and "Table:" in info['line'][0]:
                subtype="index"
            elif info['type'] == "10":
                subtype="label"
            elif info['type'] == "00":
                subtype="blank"
                                                                                                     
        # ----------------------------------------------------------------------
        # set subtype - banner
        # ----------------------------------------------------------------------
        if section == "banner":
                                                                                                     
            # ------------------------------------------------------------------
            # Flag as stats line if there are one or more labels with a 
            # single character, but none with any that are longer than one
            # ------------------------------------------------------------------
            stats_flag = False
            for elem in info['line']:
                if len(elem) == 1:
                    stats_flag = True
            for elem in info['line']:
                if len(elem) > 1:
                    stats_flag = False
                                                                                                     
            # ------------------------------------------------------------------
            # set the type
            # ------------------------------------------------------------------
            banline+=1
            if stats_flag:
                subtype="stats"
            elif info['type'] == "01":
                subtype="labels"
            elif info['type'] == "10":
                subtype="subhead"
            elif info['type'] == "00":
                subtype="blank"
                                                                                                     
        # ----------------------------------------------------------------------
        # Subtype for stubs, either base or stub
        # ----------------------------------------------------------------------
        if section == "stubs":
            if info['type'] in ["10","11"]:
                if info['line'][0].lower() in ["base", "base:","unweighted base","effective base"]:
                    subtype = "base"
                else:
                    subtype = "stub"
                                                                                                     
        # ----------------------------------------------------------------------
        # Set the stub element
        # 
        # There is now a catchall 'unknwn' tag for the line. If there are only 
        # '-' and '*', it's impossible to know if this should be a percents line
        # or a frequency line
        # ----------------------------------------------------------------------
        stubelem = ""
        if section == "stubs":
            lineitems = ''.join(info['line'][1:])
            if "%" in lineitems:
                stubelem = "prc"
            elif any(i in lineitems for i in ["0","1","2","3","4","5","6","7","8","9"]):
                stubelem = "frq"
            elif info['type'] == "01":
                stubelem = "stat"
            elif info['type'] == "00":
                stubelem = "blank"
            elif info['type'] == "10":
                stubelem = "label"
            else:
                stubelem = "unknwn"

                
        # !!!   added quick fix for stat testing
        if section == "stubs" and stubelem=="blank":
            for elem in info['line']:
                if len(elem.strip())>0:
                    stubelem = "stat"
            #print(stubelem,line)
        
        # ----------------------------------------------------------------------
        # Store the current section
        # ----------------------------------------------------------------------
        info['section'] = section
        info['subtype'] = subtype
        info['stubelem'] = stubelem
                                                                                                     
        # ----------------------------------------------------------------------
        # Add the info for this line
        # ----------------------------------------------------------------------
        lines.append(info)
                                                                                                     
    # --------------------------------------------------------------------------
    # Revise the lines of the table by collapsing multiple blank lines
    # to just one
    # --------------------------------------------------------------------------
    revised_lines = []
                                                                                                     
    for j, line in enumerate(lines):
        # ----------------------------------------------------------------------
        # Skip over a blank line if the previous line was also blank
        # ----------------------------------------------------------------------
        if j > 1 and revised_lines[-1]['stubelem'] == "blank" and line['stubelem'] == "blank":
            continue
                                                                                                     
        # ----------------------------------------------------------------------
        # Otherwise, just copy the line
        # ----------------------------------------------------------------------
        revised_lines.append(line)
                                                                                                     
    # --------------------------------------------------------------------------
    # Now add it to the structure to be saved
    # --------------------------------------------------------------------------
    table_info['lines'] = revised_lines
                                                                                                     
    # --------------------------------------------------------------------------
    # Go through the parsed lines and pull out table-level information
    # --------------------------------------------------------------------------
                                                                                                     
    # --------------------------------------------------------------------------
    # Label to use in index - first line found in header
    # 
    # Amending this to either the first or second line, will now also look for
    # project title
    # --------------------------------------------------------------------------
    table_info['label'] = "MISSING LABEL"
    table_info['projtitle'] = ""
    # --------------------------------------------------------------------------
    # Put together all the title related information
    # --------------------------------------------------------------------------
    alltitles = []
    for line in lines:
        if line['section'] == "header" and len(line['line']) > 1:
            alltitles.append(line['line'][0])
    # --------------------------------------------------------------------------
    # Store the project title as well as the table title
    # --------------------------------------------------------------------------
    table_info['label'] = alltitles[1]
    table_info['projtitle'] = alltitles[0]

    # for line in lines:
    #     if line['section'] == "header" and line['type'] == "10":
    #         table_info['label'] = line['line'][0]
    #         break
                                                                                                     
    # --------------------------------------------------------------------------
    # First base - Frequency (first cell, anyway) of the first base we 
    # come across
    # --------------------------------------------------------------------------
    table_info['firstbase'] = "NO BASE FOUND"
    for line in lines:
        if line['section'] == "stubs" and line['type'] == "11":
            if line['line'][0].lower() in ["base","base:"]:
                table_info['firstbase'] = line['line'][1]
                break
                                                                                                     
    # --------------------------------------------------------------------------
    # Number of bases found in the stubs (that's stubs that are named JUST
    # 'base' when it comes down to it, or things that obviously begin
    # that way due to punctuation
    # While we're here, grab number of stubs as well
    # --------------------------------------------------------------------------
    table_info['numbase'] = 0
    table_info['numstub'] = 0
    for line in lines:
        if line['section'] == "stubs" and line['type'] == "11":
            if line['line'][0].lower() in ["base","base:"]:
                table_info['numbase'] +=1
            else:
                table_info['numstub'] +=1
                                                                                                     
    # --------------------------------------------------------------------------
    # 'depth' of the banner...there would be two lines in the banner section
    # per level, one giving the name of the variable and one giving the
    # categories in play for that variable.
    # At the same time, we can gather the labels of the variables used
    # At the same time, we store the stat testing while excluding it
    # from the depth of the banner
    # --------------------------------------------------------------------------
    table_info['bandepth'] = 1
    table_info['banvars'] = []
    table_info['statcols'] = []
    table_info['numbps'] = 0
    ban_varstubdict = {}
    this_banvar = None
                                                                                                     
    j=0
    for line in lines:
        # ----------------------------------------------------------------------
        # Skip past things that are not the banner
        # ----------------------------------------------------------------------
        if not line['section'] == "banner":
            continue
        else:
            # ------------------------------------------------------------------
            # skip over any blank lines
            # ------------------------------------------------------------------
            if line['type'] == "00":
                continue
                                                                                                     
            # ------------------------------------------------------------------
            # If it's the line indicating stat testing, store it and skip past
            # - capital A in column B
            # TODO  A better way is a line with nothing but single letters or
            #      blanks
            # LXC MAYBE REMOVE?
            # ------------------------------------------------------------------
            # if line['line'][1] == "A":
            #     table_info['statcols'].extend(line['line'][1:-2])
            #     continue
                                                                                                     
            # ------------------------------------------------------------------
            # starting at B, count the number of non-blank things, and this
            # should be the number of banner points involved
            # ------------------------------------------------------------------
            numcols = 0
            for bp in line['line']:
                if len(str(bp).strip()) > 0:
                    numcols += 1
            if numcols > table_info['numbps']:
                table_info['numbps'] = numcols
                                                                                                     
            # ------------------------------------------------------------------
            # Add to the depth at each even number (excepting zero)
            # ------------------------------------------------------------------
            if j>0 and j %2 == 0:
                table_info['bandepth'] += 1
                                                                                                     
            # ------------------------------------------------------------------
            # At each even index STARTING with zero, we should be looking at
            # a variable name.  We're only grabbing the thing in column B.
            # ------------------------------------------------------------------
            if j %2 == 0 and line['subtype'] == 'labels':
                this_banvar = line['line'][1]
                if len(str(this_banvar).strip()) == 0:
                    this_banvar == "NO LABEL"
                ban_varstubdict[this_banvar] = []
                table_info['banvars'].append(this_banvar)

            # ------------------------------------------------------------------
            # At each odd index, we should be looking at the stubs in for this
            # banner var level
            # ------------------------------------------------------------------
            if j %2 == 1:
                this_banstub = ''
                for cell in line['line'][1:-2]:
                    if cell != '':
                        this_banstub = cell
                    ban_varstubdict[this_banvar].append(this_banstub)
            # ------------------------------------------------------------------
            # Increment
            # ------------------------------------------------------------------
            j += 1
    
    # --------------------------------------------------------------------------
    # Figure out the number of elements per stub line
    # --------------------------------------------------------------------------
    stubline = None
    for line in table_info['lines']:
        # ----------------------------------------------------------------------
        # Skip empty rows and variable rows
        # ----------------------------------------------------------------------
        if len(line['line']) <= 2:
            continue
        # ----------------------------------------------------------------------
        # Skip the line if not in the stubs
        # ----------------------------------------------------------------------
        if line['section'] != 'stubs':
            continue
        # ----------------------------------------------------------------------
        # Keep track of the element numberse
        # ----------------------------------------------------------------------
        firstcell = line['line'][0]
        if firstcell != '':
            currline = line
            currline['elemnum'] = 0
        if currline is not None:
            currline['elemnum'] += 1            
    # --------------------------------------------------------------------------
    # Grab the stat-testing (LXC CHECK PLEASE)
    # --------------------------------------------------------------------------
    for i in range(len(lines)-1,-1,-1):
        line = lines[i]
        if line['section'] == 'banner' and line['subtype'] == 'stats':
            table_info['statcols'] = line['line'][1:-2]
            break

    table_info['stubindx'] = {}
    
    # j=0
    currentbase = None
    currentunwgtbase = None
    currenteffctbase = None
    for j,line in enumerate(table_info['lines']):
        # ----------------------------------------------------------------------
        # Skip past things that are not the banner
        # ----------------------------------------------------------------------
        # if not line['section'] == "stubs":
        #     continue
        # else:
        if line['section'] == "stubs":
            # ------------------------------------------------------------------
            #  Capture the line index for the particular stub
            # ------------------------------------------------------------------
            firstcol = line['line'][0]
            if firstcol != '' and 'base' not in firstcol.lower():
                table_info['stubindx'][firstcol] = {}
                table_info['stubindx'][firstcol]['stub'] = j
                table_info['stubindx'][firstcol]['unwgt'] = currentunwgtbase
                table_info['stubindx'][firstcol]['effct'] = currenteffctbase
                table_info['stubindx'][firstcol]['base'] = currentbase
                
            if firstcol.lower()[0:min(4,len(firstcol))] == 'base' and not('unweighted' in firstcol.lower() or 'effective' in firstcol.lower()):
                currentbase = j
            if 'unweighted' in firstcol.lower():
                currentunwgtbase = j
            if 'effective' in firstcol.lower():
                currenteffctbase = j

            # if line['line'][0] != '' and 'base' not in line['line'][0].lower():
                # table_info['stubindx'][line['line'][0]] = {}
                # table_info['stubindx'][line['line'][0]]['stub'] = j
                # table_info['stubindx'][line['line'][0]]['unwgt'] = currentunwgtbase
                # table_info['stubindx'][line['line'][0]]['effct'] = currenteffctbase
                # table_info['stubindx'][line['line'][0]]['base'] = currentbase
            # if 'base' in line['line'][0].lower() and 'unweighted' not in line['line'][0].lower() and 'effective' not in line['line'][0].lower():
            #     currentbase = j
            # if 'unweighted' in line['line'][0].lower():
            #     currentunwgtbase = j
            # if 'effective' in line['line'][0].lower():
            #     currenteffctbase = j

        # ----------------------------------------------------------------------
        # Increment
        # ----------------------------------------------------------------------
        # j += 1
    
    # --------------------------------------------------------------------------
    # Clean up the banners so that all rows are the same length (extends out 
    # repeating stubs so that they will line up with the LONGEST row. This is 
    # typically the most granular row)
    # --------------------------------------------------------------------------
    stub_max = 0
    for var,stubs in ban_varstubdict.items():
        numstubs = len(stubs)
        if numstubs > stub_max:
            stub_max = numstubs

    for var in ban_varstubdict:
        while len(ban_varstubdict[var]) < stub_max:
            # print(var.ljust(20),ban_varstubdict[var])
            ban_varstubdict[var].append(ban_varstubdict[var][-1])

    # --------------------------------------------------------------------------
    # Clean up the statcols so that all rows are the same length (extends out 
    # blank stubs so that they will line up with the LONGEST row. This is 
    # typically the most granular row)
    # --------------------------------------------------------------------------
    while len(table_info['statcols']) < stub_max:
        table_info['statcols'].append('')

    # --------------------------------------------------------------------------
    # Build out the banner stubs definitions in each column 
    # (var{stub} > var{stub} ...)
    # --------------------------------------------------------------------------
    table_info['banstubs'] = []
    # for var, stubs in ban_varstubdict.items():
    #     print(var.ljust(20), str(len(stubs)))
    for var, stubs in ban_varstubdict.items():
        if len(table_info['banstubs']) == 0:
            table_info['banstubs'] = stubs
            # print(table_info['banstubs'])
            for i,stub in enumerate(stubs):
                table_info['banstubs'][i] = var + '{' + stub + '}'
        else:
            for i,stub in enumerate(stubs):
                table_info['banstubs'][i] += '>' + var + '{' + stub + '}'        

    # --------------------------------------------------------------------------
    # 
    # --------------------------------------------------------------------------
    if fdebug and False:
        print("---------------------------------------------------------------TABLE")
        for item in table_info:
            if item == "lines":
                print("-------lines")
                for line in table_info['lines']:
                    #print("  ", line)
                    print(line['section'].ljust(10), str(line['colA']).ljust(10), str(line['colB']).ljust(10), str(line['len']).ljust(5), str(line['type']).ljust(5), line['subtype'].ljust(10), line['stubelem'].ljust(10), line['line'])
            elif item == "stubindx":
                print("-------stubs")
                for thing in table_info['stubindx']:
                    print("  ", thing.ljust(100), table_info['stubindx'][thing])
            elif item == "banstubs":
                print("-------banstubs")
                for thing in table_info['banstubs']:
                    print("  ", thing)
            else:
                print(item, table_info[item])

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return table_info

    # --------------------------------------------------------------------------
    # Here's an example of the table info returned by that thing
    # --------------------------------------------------------------------------
    # lines         [
    #                {'line': [], 
    #                 'len': 0, 
    #                 'ColA': False, 
    #                 'ColB': False, 
    #                 'type': '00', 
    #                 'section': 'header', 
    #                 'subtype': 'blank', 
    #                 'stubelem': ''
    #                }, 
    #                lines                [
    #               {'line': [], 'len': 0, 'ColA': False, 'ColB': False, 'type': '00', 'section': 'header', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': [''], 'len': 1, 'type': '00', 'colA': False, 'section': 'header', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': [''], 'len': 1, 'type': '00', 'colA': False, 'section': 'header', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': [], 'len': 0, 'ColA': False, 'ColB': False, 'type': '00', 'section': 'header', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': ['', ''], 'len': 2, 'colA': False, 'colB': False, 'type': '00', 'section': 'header', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': ['DV_ScorecardLoop_Lyft_UnitedStates_Riders - Filter: source_type =* {Riders} and country_name_cat =* {UnitedStates}', ''], 'len': 2, 'colA': True, 'colB': False, 'type': '10', 'section': 'header', 'subtype': 'label', 'stubelem': ''}, 
    #               {'line': ['', ''], 'len': 2, 'colA': False, 'colB': False, 'type': '00', 'section': 'header', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': ['Table: 1 - Weighted by: Weight - Level: Top', ''], 'len': 2, 'colA': True, 'colB': False, 'type': '10', 'section': 'header', 'subtype': 'index', 'stubelem': ''}, 
    #               {'line': ['', 'Wave', '', ''], 'len': 4, 'colA': False, 'colB': True, 'type': '01', 'section': 'banner', 'subtype': 'labels', 'stubelem': ''}, 
    #               {'line': ['', 'December 2019', 'January 2020', 'February 2020', 'March 2020', 'April 2020', 'May 2020', 'June 2020', 'July 2020', '', ''], 'len': 11, 'colA': False, 'colB': True, 'type': '01', 'section': 'banner', 'subtype': 'labels', 'stubelem': ''}, 
    #               {'line': ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', '', ''], 'len': 11, 'colA': False, 'colB': True, 'type': '01', 'section': 'banner', 'subtype': 'stats', 'stubelem': ''}, 
    #               {'line': ['', ''], 'len': 2, 'colA': False, 'colB': False, 'type': '00', 'section': 'banner', 'subtype': 'blank', 'stubelem': ''}, 
    #               {'line': ['Unweighted Base', '991', '1334', '1839', '2029', '1295', '1240', '1062', '703', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '988.43288400', '1334.13519700', '1839.00208200', '2028.99984300', '1295.00059800', '1240.00009500', '1061.99893300', '703.00080600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['FavorabilityTB', '151.89870300', '232.80662300', '295.30986800', '352.26541500', '183.06006600', '190.82149100', '149.51486100', '105.74029600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '35.14121458%', '41.09910566%', '38.16638360%', '40.89415133%', '36.66988804%', '33.71041976%', '33.36686277%', '41.77792210%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['FavorabilityT2B', '381.99345000', '494.68052600', '683.76839800', '746.29693100', '429.17114600', '490.13260100', '385.80564700', '229.43061200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '88.37280062%', '87.32967708%', '88.37146942%', '86.63688892%', '85.96991259%', '86.58655601%', '86.09929469%', '90.64788541%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '880.43639100', '869.90694500', '732.63750500', '473.34712200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['StrongFutureConsideration', '*', '*', '*', '*', '225.32502900', '209.79538600', '177.47908800', '113.07739600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '25.59242568%', '24.11699173%', '24.22467957%', '23.88889480%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '988.43288400', '1334.13519700', '1839.00208200', '2028.99984300', '1295.00059800', '1240.00009500', '1061.99893300', '703.00080600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['TotalConsideration', '20.78872000', '18.07769200', '30.45607400', '22.73649000', '236.24874900', '215.19789100', '182.35782000', '120.58023000', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '2.10319996%', '1.35501200%', '1.65611960%', '1.12057623%', '18.24313822%', '17.35466730%', '17.17118674%', '17.15221789%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '427.88408300', '555.89370600', '762.40837800', '849.08507900', '486.53995400', '554.69209300', '438.64823000', '250.05665600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Preference', '179.56276300', '260.80103300', '361.26967200', '392.55244400', '222.67724200', '244.66062300', '180.96871000', '110.77387800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '41.96528222%', '46.91562977%', '47.38532294%', '46.23240400%', '45.76751409%', '44.10746540%', '41.25599914%', '44.29951187%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '427.88408300', '555.89370600', '762.40837800', '849.08507900', '486.53995400', '554.69209300', '438.64823000', '250.05665600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['AppFirst', '155.45772900', '246.73134900', '342.33326200', '355.40129900', '203.34596100', '228.89038600', '158.18187400', '104.43373200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '36.33173917%', '44.38462719%', '44.90156088%', '41.85697144%', '41.79429856%', '41.26440396%', '36.06121333%', '41.76402807%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SatisfactionTB', '192.70588000', '284.68621800', '363.91545400', '459.59256700', '252.04322000', '272.38121600', '211.37967200', '139.17525600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '44.58180713%', '50.25780111%', '47.03309411%', '53.35365660%', '50.48832801%', '48.11871597%', '47.17308006%', '54.98805303%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', 'C', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SatisfactionT2B', '403.25190200', '535.20308800', '720.70854200', '802.47015000', '469.48841200', '533.85146200', '413.35005300', '240.45717200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '93.29086647%', '94.48342999%', '93.14568071%', '93.15798359%', '94.04611218%', '94.30990598%', '92.24631185%', '95.00447209%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['TrustTB', '111.85255000', '166.06197700', '197.31004600', '235.39171600', '131.03041500', '135.71178900', '103.90639900', '75.90533600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '25.87668218%', '29.31617087%', '25.50070864%', '27.32639665%', '26.24750855%', '23.97477008%', '23.18853479%', '29.99014883%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['TrustT2B', '357.50172300', '469.30723800', '641.24017000', '718.72448000', '409.10332900', '462.76193300', '372.55426600', '216.34442600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '82.70672832%', '82.85033954%', '82.87504401%', '83.43602975%', '81.95000471%', '81.75126884%', '83.14201668%', '85.47754185%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Right_ThingTB', '104.70096500', '159.95656300', '156.26917200', '236.46262000', '107.17135100', '128.52795400', '*', '68.07819800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '24.22218890%', '28.23833618%', '20.19651156%', '27.45071686%', '21.46815266%', '22.70567774%', '0%', '26.89765170%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Right_ThingT2B', '364.98475000', '475.38271300', '640.29055400', '716.97455300', '413.48761700', '456.91851500', '*', '210.67071400', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '84.43789950%', '83.92288888%', '82.75231391%', '83.23288242%', '82.82824841%', '80.71897383%', '0%', '83.23586193%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['PositiveImpactTB', '134.03397800', '201.01623700', '213.81295400', '323.93003600', '161.51894000', '174.48796600', '148.53310600', '98.43477200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '31.00827518%', '35.48690951%', '27.63357444%', '37.60472458%', '32.35485256%', '30.82494821%', '33.14776693%', '38.89151432%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['PositiveImpactT2B', '368.23134900', '488.11620900', '657.38918000', '761.72177300', '436.26683500', '482.98729500', '399.44386800', '227.31285200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '85.18898842%', '86.17082878%', '84.96217139%', '88.42754391%', '87.39129371%', '85.32427018%', '89.14290284%', '89.81115982%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['LoveTB', '74.98751600', '124.52962800', '146.46516200', '185.04780000', '104.56554400', '114.48883100', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '17.34809013%', '21.98415266%', '18.92942351%', '21.48202013%', '20.94616742%', '20.22553398%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['LoveT2B', '296.75495600', '408.78439500', '557.42831800', '607.13658300', '365.95146300', '398.66994200', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '68.65318387%', '72.16578647%', '72.04304806%', '70.48189872%', '73.30598895%', '70.42881292%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Feel_GoodTB', '142.32899900', '230.43296300', '288.48535800', '350.09457700', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '32.92729823%', '40.68006560%', '37.28437153%', '40.64214084%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Feel_GoodT2B', '385.91430200', '508.36136300', '695.67484800', '784.33265400', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '89.27987552%', '89.74485822%', '89.91028064%', '91.05241922%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['RelyTB', '197.28837900', '286.65887300', '347.44226600', '433.14324100', '262.86675900', '256.26260300', '208.18409300', '132.98707600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '45.64195167%', '50.60604874%', '44.90406938%', '50.28317992%', '52.65645770%', '45.27121065%', '46.45993057%', '52.54310714%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', 'F', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['RelyT2B', '400.89624600', '526.35254400', '715.76906200', '795.78602600', '464.50301400', '525.86304900', '406.52013000', '240.90571800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '92.74589399%', '92.92097683%', '92.50729335%', '92.38203010%', '93.04745644%', '92.89867732%', '90.72209478%', '95.18169232%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SustainabilityTB', '*', '*', '*', '*', '*', '*', '*', '45.06715800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '17.80600478%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SustainabilityT2B', '*', '*', '*', '*', '*', '*', '*', '167.29940800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '66.09988717%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Cares_CustomersTB', '137.52972200', '186.30615300', '219.49736400', '293.00191700', '143.14383700', '160.33535900', '128.14947100', '88.49280200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '31.81700289%', '32.89002765%', '28.36823791%', '34.01430916%', '28.67402264%', '28.32475643%', '28.59880138%', '34.96344845%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Cares_CustomersT2B', '367.02129900', '492.70997200', '662.50373400', '750.67472800', '425.16025500', '482.64736800', '382.04192100', '219.98951600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '84.90904773%', '86.98180035%', '85.62318563%', '87.14510314%', '85.16646634%', '85.26421886%', '85.25935325%', '86.91771453%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Treats_DriversTB', '98.50036200', '119.01314100', '144.73227400', '209.43193600', '88.28777300', '102.54567700', '95.09617900', '53.04602800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '22.78770186%', '21.01028568%', '18.70546191%', '24.31275089%', '17.68546697%', '18.11566296%', '21.22237972%', '20.95845112%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Treats_DriversT2B', '317.67862500', '428.26637900', '589.74279000', '668.04366600', '363.50693000', '413.73195700', '322.87333900', '191.53126000', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '73.49379888%', '75.60508776%', '76.21942910%', '77.55254307%', '72.81630951%', '73.08966022%', '72.05484673%', '75.67387612%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Committed_To_SafetyTB', '130.53383100', '196.97557000', '234.98769000', '329.25852500', '144.93471000', '154.58410500', '126.72656600', '86.11062000', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '30.19852885%', '34.77358015%', '30.37023577%', '38.22330372%', '29.03276342%', '27.30874306%', '28.28125518%', '34.02224989%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Committed_To_SafetyT2B', '366.79423100', '501.11878000', '682.49871800', '772.98944500', '420.65061900', '488.45742800', '389.01151500', '228.03133400', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '84.85651637%', '88.46627052%', '88.20737367%', '89.73559706%', '84.26311341%', '86.29062087%', '86.81474035%', '90.09503159%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['feel_safeTB', '154.77853900', '242.33530700', '301.64275600', '375.84653600', '206.70645900', '208.13945100', '159.15667800', '105.83766200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '35.80745420%', '42.78127598%', '38.98485754%', '43.63166086%', '41.40664250%', '36.76980106%', '35.51860329%', '41.81639134%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['feel_safeT2B', '385.93069400', '532.38870000', '733.53196200', '794.70237000', '457.47176400', '521.90872400', '396.75723600', '234.82320400', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '89.28366775%', '93.98658489%', '94.80300280%', '92.25622952%', '91.63898349%', '92.20010844%', '88.54333380%', '92.77849500%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', 'D', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['high_safety_standardsTB', '95.66750900', '161.00925300', '198.37991800', '271.68554100', '110.94468500', '120.52943800', '97.14507900', '66.32585600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '22.13233159%', '28.42417547%', '25.63898084%', '31.53971169%', '22.22401240%', '21.29266430%', '21.67962768%', '26.20530252%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['high_safety_standardsT2B', '324.06261600', '452.65418500', '606.22017000', '710.37695100', '383.76574700', '437.14650400', '345.50657800', '211.35067600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '74.97071207%', '79.91045074%', '78.34899561%', '82.46697319%', '76.87447777%', '77.22606123%', '77.10585086%', '83.50451447%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', 'C', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Low_pricesTB', '97.87867800', '158.46907800', '187.40802200', '225.15805100', '125.78653100', '119.35151600', '99.58781200', '57.28225800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '22.64387752%', '27.97573925%', '24.22095308%', '26.13838037%', '25.19707389%', '21.08457324%', '22.22476638%', '22.63218284%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Low_pricesT2B', '308.31019900', '437.32513400', '590.05946000', '672.16362300', '381.96353300', '446.14175300', '339.99641400', '180.02143600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '71.32644747%', '77.20429797%', '76.26035611%', '78.03082490%', '76.51346520%', '78.81515697%', '75.87616115%', '71.12635215%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['BenefitsTB', '84.44963300', '132.19126700', '168.34250000', '196.77433400', '93.56488300', '86.79900700', '*', '44.64261400', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '19.53711662%', '23.33671946%', '21.75689040%', '22.84334212%', '18.74255734%', '15.33386489%', '0%', '17.63826772%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['BenefitsT2B', '270.63368000', '387.54237900', '518.06577800', '576.36130300', '307.44195000', '371.35593200', '*', '137.46920800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '62.61012131%', '68.41577348%', '66.95576191%', '66.90922623%', '61.58558844%', '65.60353492%', '0%', '54.31399457%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SmoothTB', '179.04189500', '251.66474100', '314.83312600', '410.48375400', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '41.42069370%', '44.42827119%', '40.68960491%', '47.65266199%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SmoothT2B', '392.41345300', '517.99288300', '683.58825200', '784.48690500', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '90.78343056%', '91.44518295%', '88.34818703%', '91.07032607%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['QualityDriversTB', '106.98930600', '195.31069700', '221.10823000', '302.29919700', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '24.75158830%', '34.47966759%', '28.57642916%', '35.09362141%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', 'C', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['QualityDriversT2B', '363.64959800', '502.18784800', '673.69113600', '748.84087100', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '84.12901692%', '88.65500114%', '87.06906578%', '86.93221246%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', 'A', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SupportsTB', '*', '*', '*', '*', '63.24238000', '57.06881000', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '12.66847022%', '10.08174462%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SupportsT2B', '*', '*', '*', '*', '294.93796900', '329.17238900', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '59.08083908%', '58.15141339%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SafetyOverProfitsTB', '*', '*', '*', '*', '71.00552200', '79.53077200', '58.81228900', '48.01207800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '14.22355295%', '14.04986249%', '13.12499348%', '18.96954076%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SafetyOverProfitsT2B', '*', '*', '*', '*', '303.41223600', '352.68270300', '270.34783300', '158.15172800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '60.77837164%', '62.30473255%', '60.33285911%', '62.48564476%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['HygenicTB', '*', '*', '*', '*', '111.35820700', '135.77310300', '121.62001100', '78.75281600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '22.30684754%', '23.98560178%', '27.14163789%', '31.11518632%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['HygenicT2B', '*', '*', '*', '*', '391.77042900', '472.81373300', '374.35876500', '222.28728400', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '78.47794487%', '83.52701430%', '83.54472227%', '87.82556117%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', 'E', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['ReduceCovidRiskTB', '*', '*', '*', '*', '*', '*', '110.61929700', '83.45892800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '24.68663568%', '32.97456811%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['ReduceCovidRiskT2B', '*', '*', '*', '*', '*', '*', '356.42316700', '216.73257600', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '79.54207910%', '85.63089966%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', 'G', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'stat'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SupportCommunitiesTB', '*', '*', '*', '*', '*', '*', '79.31108600', '42.20988800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '17.69965945%', '16.67709926%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '432.25228500', '566.45179800', '773.74338400', '861.40781400', '499.21086700', '566.06085700', '448.09385300', '253.10089800', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', '100.00000000%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SupportCommunitiesT2B', '*', '*', '*', '*', '*', '*', '288.46703900', '171.24080200', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '64.37647762%', '67.65712937%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['AwareofAny', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['NewsPositive', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['NewsNegative', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['NewsOutlets', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Advertising', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['OnlineAds', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['OutdoorAds', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['SocialMedia', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Base', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'base', 'stubelem': 'blank'}, 
    #               {'line': ['Email', '*', '*', '*', '*', '*', '*', '*', '*', ''], 'len': 10, 'colA': True, 'colB': True, 'type': '11', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'frq'}, 
    #               {'line': ['', '0%', '0%', '0%', '0%', '0%', '0%', '0%', '0%', ''], 'len': 10, 'colA': False, 'colB': True, 'type': '01', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'prc'}, 
    #               {'line': ['', '', '', '', '', '', '', '', '', ''], 'len': 10, 'colA': False, 'colB': False, 'type': '00', 'section': 'stubs', 'subtype': 'stub', 'stubelem': 'blank'}, 
    #               {'line': ['Cell Contents: - Count- Column Percentage- Statistical Test ResultsStatistics: - Column Proportions: \xa0\xa0\xa0Columns Tested (10%): AB,CD,EF,GH,IJ,KL,MN,OP,QR,ST,UV,WX,YZ,ab,cd,ef,gh,ij,kl,mn,op,qr,st,uv,wx,yz\xa0\xa0\xa0Minimum Base: 30 (**), Small Base: 100 (*)', ''], 'len': 2, 'colA': True, 'colB': False, 'type': '10', 'section': 'footer', 'subtype': 'stub', 'stubelem': ''}, 
    #               {'line': ['', ''], 'len': 2, 'colA': False, 'colB': False, 'type': '00', 'section': 'footer', 'subtype': 'stub', 'stubelem': ''}, 
    #               {'line': ['', ''], 'len': 2, 'colA': False, 'colB': False, 'type': '00', 'section': 'footer', 'subtype': 'stub', 'stubelem': ''}]
    # label                DV_ScorecardLoop_Lyft_UnitedStates_Riders - Filter: source_type =* {Riders} and country_name_cat =* {UnitedStates}
    # firstbase            988.43288400
    # numbase              60
    # numstub              60
    # bandepth             1
    # banvars              ['Wave']
    # statcols             ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    # numbps               8
    # stubindx             {'AidedAwareness': {'stub': 9, 'unwgt': 0, 'base': 6}, 'FavorabilityTB': {'stub': 15, 'unwgt': 0, 'base': 12}, 'FavorabilityT2B': {'stub': 21, 'unwgt': 0, 'base': 18}, 'StrongNonUserConsideration': {'stub': 27, 'unwgt': 0, 'base': 24}, 'StrongFutureConsideration': {'stub': 33, 'unwgt': 0, 'base': 30}, 'TotalConsideration': {'stub': 39, 'unwgt': 0, 'base': 36}, 'DualAppUser': {'stub': 45, 'unwgt': 0, 'base': 42}, 'Preference': {'stub': 51, 'unwgt': 0, 'base': 48}, 'AppFirst': {'stub': 57, 'unwgt': 0, 'base': 54}, 'SatisfactionTB': {'stub': 63, 'unwgt': 0, 'base': 60}, 'SatisfactionT2B': {'stub': 69, 'unwgt': 0, 'base': 66}, 'TrustTB': {'stub': 75, 'unwgt': 0, 'base': 72}, 'TrustT2B': {'stub': 81, 'unwgt': 0, 'base': 78}, 'Right_ThingTB': {'stub': 87, 'unwgt': 0, 'base': 84}, 'Right_ThingT2B': {'stub': 93, 'unwgt': 0, 'base': 90}, 'PositiveImpactTB': {'stub': 99, 'unwgt': 0, 'base': 96}, 'PositiveImpactT2B': {'stub': 105, 'unwgt': 0, 'base': 102}, 'LoveTB': {'stub': 111, 'unwgt': 0, 'base': 108}, 'LoveT2B': {'stub': 117, 'unwgt': 0, 'base': 114}, 'Feel_GoodTB': {'stub': 123, 'unwgt': 0, 'base': 120}, 'Feel_GoodT2B': {'stub': 129, 'unwgt': 0, 'base': 126}, 'RelyTB': {'stub': 135, 'unwgt': 0, 'base': 132}, 'RelyT2B': {'stub': 141, 'unwgt': 0, 'base': 138}, 'SustainabilityTB': {'stub': 147, 'unwgt': 0, 'base': 144}, 'SustainabilityT2B': {'stub': 153, 'unwgt': 0, 'base': 150}, 'Cares_CustomersTB': {'stub': 159, 'unwgt': 0, 'base': 156}, 'Cares_CustomersT2B': {'stub': 165, 'unwgt': 0, 'base': 162}, 'Treats_DriversTB': {'stub': 171, 'unwgt': 0, 'base': 168}, 'Treats_DriversT2B': {'stub': 177, 'unwgt': 0, 'base': 174}, 'Committed_To_SafetyTB': {'stub': 183, 'unwgt': 0, 'base': 180}, 'Committed_To_SafetyT2B': {'stub': 189, 'unwgt': 0, 'base': 186}, 'feel_safeTB': {'stub': 195, 'unwgt': 0, 'base': 192}, 'feel_safeT2B': {'stub': 201, 'unwgt': 0, 'base': 198}, 'high_safety_standardsTB': {'stub': 207, 'unwgt': 0, 'base': 204}, 'high_safety_standardsT2B': {'stub': 213, 'unwgt': 0, 'base': 210}, 'Low_pricesTB': {'stub': 219, 'unwgt': 0, 'base': 216}, 'Low_pricesT2B': {'stub': 225, 'unwgt': 0, 'base': 222}, 'BenefitsTB': {'stub': 231, 'unwgt': 0, 'base': 228}, 'BenefitsT2B': {'stub': 237, 'unwgt': 0, 'base': 234}, 'SmoothTB': {'stub': 243, 'unwgt': 0, 'base': 240}, 'SmoothT2B': {'stub': 249, 'unwgt': 0, 'base': 246}, 'QualityDriversTB': {'stub': 255, 'unwgt': 0, 'base': 252}, 'QualityDriversT2B': {'stub': 261, 'unwgt': 0, 'base': 258}, 'SupportsTB': {'stub': 267, 'unwgt': 0, 'base': 264}, 'SupportsT2B': {'stub': 273, 'unwgt': 0, 'base': 270}, 'SafetyOverProfitsTB': {'stub': 279, 'unwgt': 0, 'base': 276}, 'SafetyOverProfitsT2B': {'stub': 285, 'unwgt': 0, 'base': 282}, 'HygenicTB': {'stub': 291, 'unwgt': 0, 'base': 288}, 'HygenicT2B': {'stub': 297, 'unwgt': 0, 'base': 294}, 'ReduceCovidRiskTB': {'stub': 303, 'unwgt': 0, 'base': 300}, 'ReduceCovidRiskT2B': {'stub': 309, 'unwgt': 0, 'base': 306}, 'SupportCommunitiesTB': {'stub': 315, 'unwgt': 0, 'base': 312}, 'SupportCommunitiesT2B': {'stub': 321, 'unwgt': 0, 'base': 318}, 'AwareofAny': {'stub': 327, 'unwgt': 0, 'base': 324}, 'NewsPositive': {'stub': 333, 'unwgt': 0, 'base': 330}, 'NewsNegative': {'stub': 339, 'unwgt': 0, 'base': 336}, 'NewsOutlets': {'stub': 345, 'unwgt': 0, 'base': 342}, 'Advertising': {'stub': 351, 'unwgt': 0, 'base': 348}, 'OnlineAds': {'stub': 357, 'unwgt': 0, 'base': 354}, 'OutdoorAds': {'stub': 363, 'unwgt': 0, 'base': 360}, 'SocialMedia': {'stub': 369, 'unwgt': 0, 'base': 366}, 'Email': {'stub': 375, 'unwgt': 0, 'base': 372}}
    # banstubs             ['Wave{December 2019}', 'Wave{January 2020}', 'Wave{February 2020}', 'Wave{March 2020}', 'Wave{April 2020}', 'Wave{May 2020}', 'Wave{June 2020}', 'Wave{July 2020}']



# ------------------------------------------------------------------------------
# Parse out an already existing mdd map created by Hawkeye
# ------------------------------------------------------------------------------
def parse_mdd_map(map_file):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    print("...reading contents of map file " + map_file + "...")
    map_data = {}

    # --------------------------------------------------------------------------
    # Open and get the contents of the excel as raw Pandas dataframes
    # --------------------------------------------------------------------------
    wb = pd.read_excel(map_file, sheet_name=None, header=None, engine="openpyxl")

    # --------------------------------------------------------------------------
    # Summary sheet
    # --------------------------------------------------------------------------
    if 'summary' in wb:
        # ----------------------------------------------------------------------
        # Two sections - one containing the summary information and one that
        # holds the actions-to-be-taken list
        # ----------------------------------------------------------------------
        sum_head1 = wb['summary'].iloc[2 ,0:5]
        sum_data1 = wb['summary'].iloc[3:,0:5].dropna(axis=0, how='all')
        sum_data1.columns = sum_head1
        sum_data1.set_index('Metric', inplace=True)
        map_data['summary'] = sum_data1

        sum_head2 = wb['summary'].loc[2 ,6:8]
        sum_data2 = wb['summary'].loc[3:,6:8]
        sum_data2.columns = sum_head2
        map_data['actions'] = sum_data2

    # --------------------------------------------------------------------------
    # Variables sheet
    # --------------------------------------------------------------------------
    if 'variables' in wb:
        # ----------------------------------------------------------------------
        # There are three sections, but since there are a varying number of 
        # columns in some sections, we want to split it up by where things
        # lie on the title line
        # ----------------------------------------------------------------------
        title = wb['variables'].loc[:0]
        title.columns = title.iloc[0]
        section_1 = 0
        section_2 = title.columns.get_loc("Level Detail")
        section_3 = title.columns.get_loc("User Fields")

        # ----------------------------------------------------------------------
        # Grab the variable name to use as a key
        # ----------------------------------------------------------------------
        keycol = wb['variables'].iloc[3:,1:2][1].dropna().tolist()
        keycol = [x.lower() for x in keycol]

        # ----------------------------------------------------------------------
        # First section is the variable data
        # ----------------------------------------------------------------------
        var_head = wb['variables'].iloc[2 ,section_1:section_2-1]
        var_data = wb['variables'].iloc[3:,section_1:section_2-1]
        var_data.index = keycol
        var_data.columns = var_head
        map_data['variables'] = var_data

        # ----------------------------------------------------------------------
        # Second section is the level detail (keyed to the variable)
        # ----------------------------------------------------------------------
        lvl_head = wb['variables'].iloc[2 ,section_2:section_3-1]
        lvl_data = wb['variables'].iloc[3:,section_2:section_3-1]
        lvl_data.index = keycol
        lvl_data.columns = lvl_head
        map_data['level_detail'] = lvl_data

        # ----------------------------------------------------------------------
        # Third section is the user notes (keyed to the variable)
        # ----------------------------------------------------------------------
        usr_head = wb['variables'].iloc[2 ,section_3:]
        usr_data = wb['variables'].iloc[3:,section_3:]
        usr_data.index = keycol
        usr_data.columns = usr_head
        map_data['user_var'] = usr_data

    # --------------------------------------------------------------------------
    # Categories sheet
    # --------------------------------------------------------------------------
    if 'categories' in wb:
        wsc = pd.read_excel(map_file, sheet_name='categories', header=None, keep_default_na=False, engine="openpyxl")
        # ----------------------------------------------------------------------
        # There are three sections, split by the title headings
        # ----------------------------------------------------------------------
        title = wsc.loc[:0]
        title.columns = title.iloc[0]
        section_1 = 0
        section_2 = title.columns.get_loc("Categories by Variable")
        section_3 = title.columns.get_loc("User Fields")

        # ----------------------------------------------------------------------
        # First section is total categories
        # ----------------------------------------------------------------------
        cat_head = wsc.iloc[2 ,section_1:section_2-1]
        cat_data = wsc.iloc[3:,section_1:section_2-1]
        cat_data.columns = cat_head
        map_data['categories'] = cat_data

        # ----------------------------------------------------------------------
        # Grab the variable and cat name to use as a key for the others
        # ----------------------------------------------------------------------
        keycol_data = wsc.iloc[3:,9:11].values.tolist()
        keycol = []
        for rec in keycol_data:
            if valid_value(rec[0]) and valid_value(rec[1]):
                keycol.append((rec[0].lower(),rec[1].lower()))

        # ----------------------------------------------------------------------
        # Section section is categories by var (keyed to variable and cat)
        # ----------------------------------------------------------------------
        vct_head = wsc.iloc[2              ,section_2:section_3-1]
        vct_data = wsc.iloc[3:len(keycol)+3,section_2:section_3-1]
        vct_data.index = pd.MultiIndex.from_tuples(keycol)
        vct_data.columns = vct_head
        map_data['cats_by_var'] = vct_data

        # ----------------------------------------------------------------------
        # Third section is the user notes (keyed to variable and cat)
        # ----------------------------------------------------------------------
        cusr_head = wsc.iloc[2              ,section_3:]
        cusr_data = wsc.iloc[3:len(keycol)+3,section_3:]
        cusr_data.index = pd.MultiIndex.from_tuples(keycol)
        cusr_data.columns = cusr_head
        map_data['user_cat'] = cusr_data

    # --------------------------------------------------------------------------
    # routing sheet
    # --------------------------------------------------------------------------
    if 'routing' in wb:
        routing_data = wb['routing'].loc[2:,0:7]
        map_data['routing'] = routing_data

    # --------------------------------------------------------------------------
    # Go a little further with the actions section...for all the things that
    # are activated
    #  - separate into clusters; each cluster is a workflow, and each one of
    #    the workflows will be a dictionary of stuff appropriate to that workflow
    #  - any user columns mentioned - a list of column names for each sheet
    # --------------------------------------------------------------------------
    map_data['actions_to_take'] = []
    map_data['actions_user_vars'] = []
    map_data['actions_user_cats'] = []
    acts = map_data['actions'].dropna(axis=0, how='all').values.tolist()

    current_workflow = {}
    for i, line in enumerate(acts):
        # ----------------------------------------------------------------------
        # Skip past things that have no key in the first column
        # ----------------------------------------------------------------------
        if not valid_value(line[0]):
            continue

        # ----------------------------------------------------------------------
        # new workflows end with a :
        # ----------------------------------------------------------------------
        if line[0].endswith(":"):
            # ------------------------------------------------------------------
            # Add the existing one to the list
            # ------------------------------------------------------------------
            if current_workflow:
                map_data['actions_to_take'].append(current_workflow)

            # ------------------------------------------------------------------
            # Start a new workflow
            # ------------------------------------------------------------------
            current_workflow = {}
            current_workflow['type'] = line[0]
            current_workflow['active'] = False
            current_workflow['components'] = []

            # ------------------------------------------------------------------
            # Done with this line
            # ------------------------------------------------------------------
            continue

        # ----------------------------------------------------------------------
        # If the instruction doesn't contain information about the parameter,
        # it's not an instruction and can just be ignored
        # ----------------------------------------------------------------------
        use_this_line = True
        if not ("(Param:" in line[0] and line[0].endswith(")")):
            use_this_line = False

        # ----------------------------------------------------------------------
        # Lines that aren't activated are ignored
        # ----------------------------------------------------------------------
        if not evaluate_as_yes(line[1]):
            use_this_line = False

        # ----------------------------------------------------------------------
        # Everything needs to be part of a workflow, even the things that stand
        # on their own
        # ----------------------------------------------------------------------
        if not current_workflow:
            kill_program("Hey, there's a line in the actions list on the summary sheet that is missing its workflow heading!  That line needs to be there.")

        # ----------------------------------------------------------------------
        # Parse out the task and the parameter usage
        # ----------------------------------------------------------------------
        if use_this_line:
            task, param_type = line[0][:-1].split("(Param:")

            task = task.lower().strip()
            param_type = param_type.lower().strip()

            param = line[2]

        # ----------------------------------------------------------------------
        # If it is a reference to a user variable, add to those lists
        # ----------------------------------------------------------------------
        if use_this_line:
            if param_type == "variables sheet/user column":
                if not param in map_data['actions_user_vars']:
                    map_data['actions_user_vars'].append(param)

            elif param_type == "categories sheet/user column":
                if not param in map_data['actions_user_cats']:
                    map_data['actions_user_cats'].append(param)

        # ----------------------------------------------------------------------
        # If we're this far, we're going to do something so mark the workflow
        # as active
        # ----------------------------------------------------------------------
        if use_this_line:
            current_workflow['active'] = True

        # ----------------------------------------------------------------------
        # Add this component to the list 
        # ----------------------------------------------------------------------
        if use_this_line:
            current_workflow['components'].append( {'task':task, 'param':param} )

        # ----------------------------------------------------------------------
        # If this is the last line, add the workflow to the list
        # ----------------------------------------------------------------------
        if current_workflow and i == len(acts)-1:
            map_data['actions_to_take'].append(current_workflow)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return map_data

# ------------------------------------------------------------------------------
# Parse out an already existing mdd map created by Hawkeye
# REVISION TO PULL FROM NEW SHEET STYLE AS OF HAWKEYE v0.3
# ------------------------------------------------------------------------------
def parse_mdd_map2(map_file):
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    print("...reading contents of map file " + map_file + "...")
    map_data = {}

    # --------------------------------------------------------------------------
    # Open and get the contents of the excel as raw Pandas dataframes
    # --------------------------------------------------------------------------
    wb = pd.read_excel(map_file, sheet_name=None, header=None, engine="openpyxl")

    # --------------------------------------------------------------------------
    # Account for any old maps
    # --------------------------------------------------------------------------
    if not 'cats by vars' in wb:
        print("WARNING! Your existing map document is using an older style of layout, and can no longer be parsed properly.")
        print("Please set aside a copy of that map, and run a fresh one, then you may transfer any custom fields across to the appropriate place in the new map.")
        kill_program("Sorry for the inconvenience!")

    # --------------------------------------------------------------------------
    # Summary sheet
    # --------------------------------------------------------------------------
    if 'summary' in wb:
        # ----------------------------------------------------------------------
        # Two sections - one containing the summary information and one that
        # holds the actions-to-be-taken list
        # ----------------------------------------------------------------------
        sum_head1 = wb['summary'].iloc[2 ,0:5]
        sum_data1 = wb['summary'].iloc[3:,0:5].dropna(axis=0, how='all')
        sum_data1.columns = sum_head1
        sum_data1.set_index('Metric', inplace=True)
        map_data['summary'] = sum_data1

        sum_head2 = wb['summary'].loc[2 ,6:8]
        sum_data2 = wb['summary'].loc[3:,6:8]
        sum_data2.columns = sum_head2
        map_data['actions'] = sum_data2

    # --------------------------------------------------------------------------
    # Variables sheet
    # --------------------------------------------------------------------------
    if 'variables' in wb:
        # ----------------------------------------------------------------------
        # There are three sections, but since there are a varying number of 
        # columns in some sections, we want to split it up by where things
        # lie on the title line
        # ----------------------------------------------------------------------
        title = wb['variables'].loc[:0]
        title.columns = title.iloc[0]
        section_1 = 0
        section_2 = title.columns.get_loc("Level Detail")
        section_3 = title.columns.get_loc("User Fields")

        # ----------------------------------------------------------------------
        # Grab the variable name to use as a key
        # ----------------------------------------------------------------------
        keycol = wb['variables'].iloc[3:,1:2][1].dropna().tolist()
        keycol = [x.lower() for x in keycol]

        # ----------------------------------------------------------------------
        # First section is the variable data
        # ----------------------------------------------------------------------
        var_head = wb['variables'].iloc[2 ,section_1:section_2-1]
        var_data = wb['variables'].iloc[3:,section_1:section_2-1]
        var_data.index = keycol
        var_data.columns = var_head
        map_data['variables'] = var_data

        # ----------------------------------------------------------------------
        # Second section is the level detail (keyed to the variable)
        # ----------------------------------------------------------------------
        lvl_head = wb['variables'].iloc[2 ,section_2:section_3-1]
        lvl_data = wb['variables'].iloc[3:,section_2:section_3-1]
        lvl_data.index = keycol
        lvl_data.columns = lvl_head
        map_data['level_detail'] = lvl_data

        # ----------------------------------------------------------------------
        # Third section is the user notes (keyed to the variable)
        # ----------------------------------------------------------------------
        usr_head = wb['variables'].iloc[2 ,section_3:]
        usr_data = wb['variables'].iloc[3:,section_3:]
        usr_data.index = keycol
        usr_data.columns = usr_head
        map_data['user_var'] = usr_data

    # --------------------------------------------------------------------------
    # Cats by Name sheet
    # --------------------------------------------------------------------------
    if 'cats by name' in wb:
        wsc = pd.read_excel(map_file, sheet_name='cats by name', header=None, keep_default_na=False, engine="openpyxl")
        # ----------------------------------------------------------------------
        # There are two sections, split by the title headings
        # ----------------------------------------------------------------------
        title = wsc.loc[:0]
        title.columns = title.iloc[0]
        section_1 = 0
        section_2 = title.columns.get_loc("User Fields")

        # ----------------------------------------------------------------------
        # Grab the category name to use as a key
        # ----------------------------------------------------------------------
        keycol = wb['cats by name'].iloc[3:,1:2][1].dropna().tolist()
        keycol = [x.lower() for x in keycol]

        # ----------------------------------------------------------------------
        # First section is categories
        # ----------------------------------------------------------------------
        cat_head = wsc.iloc[2 ,section_1:section_2-1]
        cat_data = wsc.iloc[3:,section_1:section_2-1]
        cat_data.columns = cat_head
        map_data['cats_name_data'] = cat_data

        # ----------------------------------------------------------------------
        # Second section is the user notes (keyed to variable and cat)
        # ----------------------------------------------------------------------
        cusr_head = wsc.iloc[2              ,section_2:]
        cusr_data = wsc.iloc[3:len(keycol)+3,section_2:]
        cusr_data.index = keycol
        cusr_data.columns = cusr_head
        map_data['cats_name_user'] = cusr_data

    # --------------------------------------------------------------------------
    # Cats by List sheet
    # --------------------------------------------------------------------------
    if 'cats by list' in wb:
        # ----------------------------------------------------------------------
        # Get the raw sheet data
        # ----------------------------------------------------------------------
        wsc = pd.read_excel(map_file, sheet_name='cats by list', header=None, keep_default_na=False, engine="openpyxl")

        # ----------------------------------------------------------------------
        # There are two sections, split by the title headings
        # ----------------------------------------------------------------------
        title = wsc.loc[:0]
        title.columns = title.iloc[0]
        section_1 = 0
        section_2 = title.columns.get_loc("User Fields")

    # --------------------------------------------------------------------------
    # Cats by Vars sheet
    # --------------------------------------------------------------------------
    if 'cats by vars' in wb:
        # ----------------------------------------------------------------------
        # Get the raw sheet data
        # ----------------------------------------------------------------------
        wsc = pd.read_excel(map_file, sheet_name='cats by vars', header=None, keep_default_na=False, engine="openpyxl")

        # ----------------------------------------------------------------------
        # There are two sections, split by the title headings
        # ----------------------------------------------------------------------
        title = wsc.loc[:0]
        title.columns = title.iloc[0]
        section_1 = 0
        section_2 = title.columns.get_loc("User Fields")

        # ----------------------------------------------------------------------
        # Grab the variable and cat name to use as a key for the others
        # ----------------------------------------------------------------------
        keycol_data = wsc.iloc[3:,2:4].values.tolist()
        keycol = []
        for rec in keycol_data:
            if valid_value(rec[0]) and valid_value(rec[1]):
                keycol.append((rec[0].lower(),rec[1].lower()))

        # ----------------------------------------------------------------------
        # First section is categories by var (keyed to variable and cat)
        # ----------------------------------------------------------------------
        vct_head = wsc.iloc[2              ,section_1:section_2-1]
        vct_data = wsc.iloc[3:len(keycol)+3,section_1:section_2-1]
        vct_data.index = pd.MultiIndex.from_tuples(keycol)
        vct_data.columns = vct_head
        map_data['cats_vars_data'] = vct_data

        # ----------------------------------------------------------------------
        # Second section is the user notes (keyed to variable and cat)
        # ----------------------------------------------------------------------
        cusr_head = wsc.iloc[2              ,section_2:]
        cusr_data = wsc.iloc[3:len(keycol)+3,section_2:]
        cusr_data.index = pd.MultiIndex.from_tuples(keycol)
        cusr_data.columns = cusr_head
        map_data['cats_vars_user'] = cusr_data

    # --------------------------------------------------------------------------
    # routing sheet
    # --------------------------------------------------------------------------
    if 'routing' in wb:
        routing_data = wb['routing'].loc[2:,0:7]
        map_data['routing'] = routing_data

    # --------------------------------------------------------------------------
    # Go a little further with the actions section...for all the things that
    # are activated
    #  - separate into clusters; each cluster is a workflow, and each one of
    #    the workflows will be a dictionary of stuff appropriate to that workflow
    #  - any user columns mentioned - a list of column names for each sheet
    # --------------------------------------------------------------------------
    map_data['actions_to_take'] = []
    map_data['actions_user_vars'] = []
    map_data['actions_user_cats'] = []
    acts = map_data['actions'].dropna(axis=0, how='all').values.tolist()

    current_workflow = {}
    for i, line in enumerate(acts):
        # ----------------------------------------------------------------------
        # Skip past things that have no key in the first column
        # ----------------------------------------------------------------------
        if not valid_value(line[0]):
            continue

        # ----------------------------------------------------------------------
        # new workflows end with a :
        # ----------------------------------------------------------------------
        if line[0].endswith(":"):
            # ------------------------------------------------------------------
            # Add the existing one to the list
            # ------------------------------------------------------------------
            if current_workflow:
                map_data['actions_to_take'].append(current_workflow)

            # ------------------------------------------------------------------
            # Start a new workflow
            # ------------------------------------------------------------------
            current_workflow = {}
            current_workflow['type'] = line[0]
            current_workflow['active'] = False
            current_workflow['components'] = []

            # ------------------------------------------------------------------
            # Done with this line
            # ------------------------------------------------------------------
            continue

        # ----------------------------------------------------------------------
        # If the instruction doesn't contain information about the parameter,
        # it's not an instruction and can just be ignored
        # ----------------------------------------------------------------------
        use_this_line = True
        if not ("(Param:" in line[0] and line[0].endswith(")")):
            use_this_line = False

        # ----------------------------------------------------------------------
        # Lines that aren't activated are ignored
        # ----------------------------------------------------------------------
        if not evaluate_as_yes(line[1]):
            use_this_line = False

        # ----------------------------------------------------------------------
        # Everything needs to be part of a workflow, even the things that stand
        # on their own
        # ----------------------------------------------------------------------
        if not current_workflow:
            kill_program("Hey, there's a line in the actions list on the summary sheet that is missing its workflow heading!  That line needs to be there.")

        # ----------------------------------------------------------------------
        # Parse out the task and the parameter usage
        # ----------------------------------------------------------------------
        if use_this_line:
            task, param_type = line[0][:-1].split("(Param:")

            task = task.lower().strip()
            param_type = param_type.lower().strip()

            param = line[2]

        # ----------------------------------------------------------------------
        # If it is a reference to a user variable, add to those lists
        # ----------------------------------------------------------------------
        if use_this_line:
            if param_type == "variables sheet/user column":
                if not param in map_data['actions_user_vars']:
                    map_data['actions_user_vars'].append(param)

            elif param_type == "categories sheet/user column":
                if not param in map_data['actions_user_cats']:
                    map_data['actions_user_cats'].append(param)

        # ----------------------------------------------------------------------
        # If we're this far, we're going to do something so mark the workflow
        # as active
        # ----------------------------------------------------------------------
        if use_this_line:
            current_workflow['active'] = True

        # ----------------------------------------------------------------------
        # Add this component to the list 
        # ----------------------------------------------------------------------
        if use_this_line:
            current_workflow['components'].append( {'task':task, 'param':param} )

        # ----------------------------------------------------------------------
        # If this is the last line, add the workflow to the list
        # ----------------------------------------------------------------------
        if current_workflow and i == len(acts)-1:
            map_data['actions_to_take'].append(current_workflow)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return map_data

# ------------------------------------------------------------------------------
# Given a set of value information in the list/info format, return the best
# label to use for mapping.
# ------------------------------------------------------------------------------
def select_primary_mdd_label(vlbl):
    # --------------------------------------------------------------------------
    # Start with a default of nothing
    # --------------------------------------------------------------------------
    plbl = ""

    # --------------------------------------------------------------------------
    # If there is not even a list structure, it's blank
    # --------------------------------------------------------------------------
    if not 'list' in vlbl:
        plbl = ""

    # --------------------------------------------------------------------------
    # If there is a list, but nothing on the list, it's blank
    # --------------------------------------------------------------------------
    elif len(vlbl['list']) == 0:
        plbl = ""

    # --------------------------------------------------------------------------
    # If there's only one thing, it's that
    # --------------------------------------------------------------------------
    elif len(vlbl['list']) == 1:
        plbl = vlbl['info'][vlbl['list'][0]]

    # --------------------------------------------------------------------------
    # If there is more than one thing...
    # --------------------------------------------------------------------------
    elif len(vlbl['list']) > 1:
        # ----------------------------------------------------------------------
        # If we have a context/language for Question/English, it's that one
        # ----------------------------------------------------------------------
        for lblidx in vlbl['list']:
            context = lblidx[0].lower()
            lang    = lblidx[1].lower()
            if context == "question" and lang == "en-us":
                plbl = vlbl['info'][lblidx]

        # ----------------------------------------------------------------------
        # ...otherwise just use whatever the first one is
        # ----------------------------------------------------------------------
        if not plbl:
            plbl = vlbl['info'][vlbl['list'][0]]

    # --------------------------------------------------------------------------
    # If it was assigned as 'none', remake it as a blank
    # --------------------------------------------------------------------------
    if plbl is None:
        plbl = ""

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return plbl


# ------------------------------------------------------------------------------
# Query an MDD with a connection and return a count (assuming using sum/count)
# Declare the connection outside the function for efficiency purposes in the case of
# multiple queries in one program
# Grant notes:
#  - let's keep function names all lowercase
#  - we need a wrapper function to set up the connection that takes file info
#    and a list of queries, or build it into this one to keep it contained
#  - Please fill in the objectives of each paragraph of code
#  - The dataframe conversion can barf on datetime fields - not breaking, but
#    spits a warning for every row
#  - We need to have more freedom with the query itself - there are a bunch
#    of things that can be done
# ------------------------------------------------------------------------------
def query_MDD(varlist, sqlfilter, adoconnection):
    """
    adoconnection defined as an object with:
        adoconn = win32com.client.Dispatch(r'ADODB.Connection')
    standard connection string
        adoconn.Open("Provider=mrOleDB.Provider.2;Data Source=mrDataFileDsc;Location=R201093M.ddf;Initial Catalog=R201093M.mdd;MR Init MDM Access=0;")
    """
    # --------------------------------------------------------------------------
    # 
    # --------------------------------------------------------------------------
    objRecordset = win32com.client.Dispatch("ADODB.Recordset")    
    #print('Select ' + ','.join(varlist) + ' from HDATA WHERE ' + sqlfilter)
    objRecordset.Open('Select ' + ','.join(varlist) + ' from HDATA WHERE ' + sqlfilter, adoconnection)

    # --------------------------------------------------------------------------
    # 
    # --------------------------------------------------------------------------
    outdict = {}
    for var in varlist:
        outdict[var] = []

    # --------------------------------------------------------------------------
    # 
    # --------------------------------------------------------------------------
    while not objRecordset.EOF:
        for var in varlist:
            outdict[var].append(objRecordset.Fields[var].Value)
        objRecordset.MoveNext()

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return pd.DataFrame(outdict)

# ------------------------------------------------------------------------------
# Rewrite of that function giving more options
# ------------------------------------------------------------------------------
def query_mdd(adoconn, var="*", filt="", level="HDATA", idfield=""):
    # --------------------------------------------------------------------------
    # Put together the variable list, differently depending on what was passed
    # --------------------------------------------------------------------------
    if type(var) is str:
        varlist = [var]
    elif type(var) is list or type(var) is tuple:
        varlist = list(var)
    elif type(var) is dict:
        varlist = [x for x in var]

    # --------------------------------------------------------------------------
    # If an ID field is included, make sure it is on the variable list
    # --------------------------------------------------------------------------
    if idfield and (not varlist[0] == "*"):
        if not idfield in varlist:
            varlist.insert(0, idfield)

    # --------------------------------------------------------------------------
    # Now assemble the sql
    # --------------------------------------------------------------------------
    sql = "SELECT " + ",".join(varlist) + " FROM " + level
    if filt:
        sql = sql + " WHERE " + filt

    # --------------------------------------------------------------------------
    # Initialize the recordset for this query and open the connection
    # --------------------------------------------------------------------------
    records = win32com.client.Dispatch("ADODB.Recordset")    
    records.Open(sql, adoconn)

    # --------------------------------------------------------------------------
    # Initialize a dictionary to hold the variables
    # --------------------------------------------------------------------------
    outdict = {}
    for v in varlist:
        outdict[v] = []

    # --------------------------------------------------------------------------
    # Fill in that dictionary, record by record
    # --------------------------------------------------------------------------
    while not records.EOF:
        for v in varlist:
            outdict[v].append(records.Fields[v].Value)
        records.MoveNext()

    # --------------------------------------------------------------------------
    # Turn that whole thing into a dataframe
    # --------------------------------------------------------------------------
    data = pd.DataFrame(outdict)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return data


# ------------------------------------------------------------------------------
# Given an sqlite schema statement, break it into a usefule bundle of info
# ------------------------------------------------------------------------------
def parse_ddf_schema(schema):
    # --------------------------------------------------------------------------
    # If it's not a CREATE TABLE line, ditch
    # --------------------------------------------------------------------------
    if not schema.upper().startswith("CREATE TABLE"):
        print("WARNING! Attempting to parse a string as a schema fails if string does not start with 'CREATE TABLE'.")
        return None

    # --------------------------------------------------------------------------
    # Chop by the opening parenthesis, remove the closing parenthesis
    # --------------------------------------------------------------------------
    tdata, vdata = schema.split("(",1)
    vdata = vdata[:-1]

    # --------------------------------------------------------------------------
    # Extract the table name
    # --------------------------------------------------------------------------
    tname = re.sub("CREATE TABLE", "", tdata, flags=re.IGNORECASE).strip()

    # --------------------------------------------------------------------------
    # Separate the primary key indicator if it is there...this will detect it
    # when it is cast separately, which always happens at the end
    # --------------------------------------------------------------------------
    if ", primary key (" in vdata:
        vdata, pkey_text = vdata.split(", primary key")
        pkey = pkey_text.strip()[1:-1].strip().split(",")
        pkey = [x.strip() for x in pkey]
    else:
        pkey = ""

    #print("---------")
    #print(schema)
    #print(tname)
    #print(pkey)

    # --------------------------------------------------------------------------
    # work through the variables 
    # --------------------------------------------------------------------------
    vlist = [x.strip() for x in vdata.split(",")]

    vinfo = {}
    for i, item in enumerate(vlist):
        # ----------------------------------------------------------------------
        # Maybe it's flagged as not null
        # ----------------------------------------------------------------------
        if ' not null' in item.lower():
            notnull = True
            item = re.sub(" not null", "", item, flags=re.IGNORECASE).strip()
        else:
            notnull = False

        # ----------------------------------------------------------------------
        # Maybe it's flagged directly as a primary key
        # ----------------------------------------------------------------------
        if ' primary key' in item.lower():
            is_pkey = True
            item = re.sub(" primary key", "", item, flags=re.IGNORECASE).strip()
        else:
            is_pkey = False

        # ----------------------------------------------------------------------
        # Rest is like this:
        # [varname:dtype] vartype
        # or like this:
        # varname vartype
        # ----------------------------------------------------------------------
        if "]" in item:
            varname_dtype, vartype = item.split("]")
            varname_dtype += "]"
            vartype = vartype.strip()
            varname, dtype = varname_dtype[1:].strip().split(":")
            dtype = dtype[:-1]
        else:
            varname, vartype = item.split(" ")
            varname = varname.strip()
            vartype = vartype.strip()
            dtype = ""
            varname_dtype = varname

        # ----------------------------------------------------------------------
        # Maybe this was listed as part of the primary key set
        # ----------------------------------------------------------------------
        if varname_dtype in pkey:
            is_pkey = True

        # ----------------------------------------------------------------------
        # Put all of this together and add it to the structure
        # ----------------------------------------------------------------------
        this_vinfo = {}
        this_vinfo['varname'] = varname
        this_vinfo['not_null'] = notnull
        this_vinfo['dtype']    = dtype
        this_vinfo['vartype']  = vartype
        this_vinfo['is_pkey']  = is_pkey

        vinfo[varname_dtype] = this_vinfo

        #print(varname_dtype, vinfo[varname_dtype])

    # --------------------------------------------------------------------------
    # Consolidate
    # --------------------------------------------------------------------------
    schema_info = {}
    schema_info['tname'] = tname
    schema_info['pkey']  = pkey
    schema_info['vinfo'] = vinfo

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return schema_info


# ------------------------------------------------------------------------------
# Given a standard mdd_map data structure and the name of a variable (presumably
# something in a loop), return the whole list of variables with iterations
# included.
# ------------------------------------------------------------------------------
def get_exploded_varlist(mdata, var, sublist=False):
    # --------------------------------------------------------------------------
    # Chop the name up into its loop pieces
    # --------------------------------------------------------------------------
    pieces = var.split("[..].")

    # --------------------------------------------------------------------------
    # Work through however many levels there are
    # --------------------------------------------------------------------------
    iterations = []
    for i in range(len(pieces)):
        # ----------------------------------------------------------------------
        # Construct the varible name for this level
        # ----------------------------------------------------------------------
        vname = "[..].".join(pieces[0:i+1])

        if i < len(pieces)-1:
            vname += "[..]."

        # ----------------------------------------------------------------------
        # Grab the info for this level
        # ----------------------------------------------------------------------
        vinfo = mdata['variables'].loc[vname]

        # ----------------------------------------------------------------------
        # If this is one of the loop parts, store the iterations
        # ----------------------------------------------------------------------
        if vinfo['Type'] == "loop":
            iterations.append(vinfo['Categories'].split("|"))

    # --------------------------------------------------------------------------
    # Get the list of all permutations of the iterations
    # --------------------------------------------------------------------------
    permutations = exhaustive_permutation(iterations)

    # --------------------------------------------------------------------------
    # Now work through each permutation and build the list of variable names
    # --------------------------------------------------------------------------
    vlist = []
    for permutation in permutations:
        # ----------------------------------------------------------------------
        # Assemble the parts of each permutation as a list
        # ----------------------------------------------------------------------
        vname_parts = interleave(pieces, permutation)

        # ----------------------------------------------------------------------
        # Now build the variable name, alternating the delimiter to create
        # the proper pattern
        # ----------------------------------------------------------------------
        vname = ""
        side = 1
        for i, part in enumerate(vname_parts):
            # ------------------------------------------------------------------
            # If it's an iteration, we need to account for headers and just get
            # the category name
            # NO! 
            # ------------------------------------------------------------------
            if sublist is False and side < 0:
                part = part.split(".")[-1]

            # ------------------------------------------------------------------
            # Add the piece to the running name
            # ------------------------------------------------------------------
            vname += part

            # ------------------------------------------------------------------
            # Unless we are on the last one...
            # ------------------------------------------------------------------
            if i < len(vname_parts)-1:
                # --------------------------------------------------------------
                # We're either starting or finishing an iteration
                # --------------------------------------------------------------
                if side > 0:
                    vname += "[{"
                else:
                    vname += "}]."

                # --------------------------------------------------------------
                # Flip it to go the other direction next time
                # --------------------------------------------------------------
                side *= -1

        # ----------------------------------------------------------------------
        # Add it to the list
        # ----------------------------------------------------------------------
        vlist.append(vname)

    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    return vlist

# ------------------------------------------------------------------------------
# Given a standard tables output file, interpret into a set of dataframes
# ------------------------------------------------------------------------------
def parse_tables_xlsx(inp):
    """
    Grab index and go through each of the tables to build the output
    """
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    raw_read = pd.read_excel(inp,sheet_name = None)
    indx = get_tab_index(raw_read)
    # --------------------------------------------------------------------------
    # Get the individual tables from the file
    # --------------------------------------------------------------------------
    tabdict = {}
    for i in indx['tab'].to_list():
        tabdict[i] = get_table_df(raw_read[i])
    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    out = {}
    out['indx'] = indx
    out['tabs'] = tabdict
    return out

# ------------------------------------------------------------------------------
# Get the index
# ------------------------------------------------------------------------------
def get_tab_index(raw_read):
    """
    Get the index setup for the tables
    """
    # --------------------------------------------------------------------------
    # Get the index
    # --------------------------------------------------------------------------
    indx = raw_read['IndexSheet'].loc[4:,:].copy()
    # --------------------------------------------------------------------------
    # Reset the column so that we remove Table #
    # --------------------------------------------------------------------------
    ref_col = indx.columns.to_list()[0]
    indx[ref_col] = [i[i.find('-')+2:] for i in indx[ref_col].to_list()]
    # --------------------------------------------------------------------------
    # Put the reference tab
    # --------------------------------------------------------------------------
    indx['tab'] = [f'T{i+1}' for i in range(len(indx[ref_col]))]
    # --------------------------------------------------------------------------
    # Set the index
    # --------------------------------------------------------------------------
    indx.set_index(ref_col,inplace=True)
    return indx
# ------------------------------------------------------------------------------
# Use the standard pandas read and manipulate
# ------------------------------------------------------------------------------
def get_table_df(df):
    """
    Manipulate the table df so that it fits appropriately
    """
    # --------------------------------------------------------------------------
    # Start
    # --------------------------------------------------------------------------
    indf = df.copy()
    tabzeros = {i:'0' for i in ['-','*']}
    pctzeros = {i:'0' for i in ['*','**']}
    combzeros = {i:'0' for i in ['-','*','**']}
    # --------------------------------------------------------------------------
    # Drop the empty columns and rows
    # --------------------------------------------------------------------------
    # indf.dropna(axis=0,how='all',inplace=True)
    indf.dropna(axis=1,how='all',inplace=True)
    indf.fillna('',inplace=True)
    # --------------------------------------------------------------------------
    # Grab the first column
    # --------------------------------------------------------------------------
    cols = indf.columns.to_list()
    indxcol = indf[cols[0]].to_list()
    # --------------------------------------------------------------------------
    # Figure out where the pieces of the table are
    # --------------------------------------------------------------------------
    title  = [0,1]
    banner = []
    table = []
    inheader = True
    inbanner = False
    intable  = False
    # --------------------------------------------------------------------------
    # Go through and figure out where everything starts and stops
    # --------------------------------------------------------------------------
    for i,val in enumerate(indxcol):
        # ----------------------------------------------------------------------
        # Find the start of the information
        # ----------------------------------------------------------------------
        if inheader:
            bskip = True
            for cell in indf.iloc[i,:].values.tolist():
                if cell != '':
                    bskip = False
                    break
            if bskip:
                continue
            if val != '' or val is not np.nan:
                instart = False
                inheader = True
        # ----------------------------------------------------------------------
        # In the header
        # ----------------------------------------------------------------------
        if inheader:
            if val == '' or val is np.nan:
                inbanner = True
                inheader = False
        # ----------------------------------------------------------------------
        # In the banner
        # ----------------------------------------------------------------------
        if inbanner:
            bskip = True
            # ------------------------------------------------------------------
            # In the case that we have an entire row that is empty, skip.
            # This should handle empty rows after the header
            # ------------------------------------------------------------------
            for cell in indf.iloc[i,:].values.tolist():
                if cell != '':
                    bskip = False
                    break
            if bskip:
                continue
            if val != '':
                inbanner = False
                intable = True
            else:
                banner.append(i)
        # ----------------------------------------------------------------------
        # In the table
        # ----------------------------------------------------------------------
        if intable:
            table.append(i)
            # ------------------------------------------------------------------
            # Exit if we've reached the bottom
            # ------------------------------------------------------------------
            if 'Statistics' in str(val):
                break
            if 'Table ' in str(val):
                table = table[:-2]
    # --------------------------------------------------------------------------
    # Break the dataframe apart
    # --------------------------------------------------------------------------
    bandf = indf.iloc[banner[0]:banner[-1]+1].copy()
    tabdf = indf.iloc[table[0]:table[-2]].copy()
    # --------------------------------------------------------------------------
    # Create the column index setup
    # --------------------------------------------------------------------------
    banmat = bandf.iloc[:,1:].values.tolist()
    for row in banmat:
        for i in range(len(row)):
            if row[i] == '':
                row[i] = row[i-1]
    # --------------------------------------------------------------------------
    # Set up the tuples for the columns
    # --------------------------------------------------------------------------
    colnames = []
    for i in range(len(banmat[0])):
        colnames.append(tuple([row[i] for row in banmat]))
    multi_col = pd.MultiIndex.from_tuples(colnames)
    # --------------------------------------------------------------------------
    # Create the final using the information above
    # --------------------------------------------------------------------------
    tabindx = tabdf[cols[0]].to_list()
    for i,val in enumerate(tabindx):
        if val == '':
            tabindx[i] = currval
        else:
            currval = val
    # --------------------------------------------------------------------------
    # Set the index
    # --------------------------------------------------------------------------
    # setindx = pd.MultiIndex.from_tuples(tabindx)
    tabdf.index = [str(i) for i in tabindx]
    del tabdf[cols[0]]
    # --------------------------------------------------------------------------
    # Set the columns
    # --------------------------------------------------------------------------
    tabdf.columns = multi_col
    # --------------------------------------------------------------------------
    # Go through the rows and add a typing tag
    # --------------------------------------------------------------------------
    typecol = []
    for idx, row in tabdf.iterrows():
        # ----------------------------------------------------------------------
        # Default to a count row
        # ----------------------------------------------------------------------
        thistype = 'count'
        # ----------------------------------------------------------------------
        # If the entire row is empty, it's a stat row
        # ----------------------------------------------------------------------
        bstat = True
        for val in row:
            if val != '':
                bstat = False
                break
        if bstat:
            thistype = 'stat'
        # ----------------------------------------------------------------------
        # Go through the values and guess at the type
        # ----------------------------------------------------------------------
        for val in row:
            # ------------------------------------------------------------------
            # Skip if empty since it doesn't tell us anything
            # ------------------------------------------------------------------
            if val == '':
                continue
            # ------------------------------------------------------------------
            # If it has any of the following keys, skip
            # ------------------------------------------------------------------
            if val in tabzeros:
                continue
            # ------------------------------------------------------------------
            # If it is one of the zeros for pcts, mark
            # ------------------------------------------------------------------
            if val in pctzeros:
                thistype = 'pct'
                break
            # ------------------------------------------------------------------
            # If we are able to cast as float and the value is less than 1, pct
            # ------------------------------------------------------------------
            try:
                floatval = float(val)
                if floatval < 1:
                    thistype = 'pct'
                    break
            # ------------------------------------------------------------------
            # If it breaks, we are in a stat row
            # ------------------------------------------------------------------
            except:
                thistype = 'stat'
                break
        typecol.append(thistype)
    tabdf['type'] = typecol
    # --------------------------------------------------------------------------
    # Split the table into multiple tables (count, pct, stat)
    # --------------------------------------------------------------------------
    tmpmask = tabdf[('type')] == 'count'
    countdf = tabdf.loc[tmpmask].copy()
    tmpmask = tabdf[('type')] == 'stat'
    statdf  = tabdf.loc[tmpmask].copy()
    tmpmask = tabdf[('type')] == 'pct'
    pctdf   = tabdf.loc[tmpmask].copy()
    # --------------------------------------------------------------------------
    # Create the base, unweighted base, and effective base dataframes
    # --------------------------------------------------------------------------
    # --------------------------------------------------------------------------
    # Base df
    # --------------------------------------------------------------------------
    basedf = tabdf.copy()
    baserow = None
    for idx,row in basedf.iterrows():
        if (idx.lower()[:4] == 'base' and 'unweighted' not in idx.lower() and 'effective' not in idx.lower()) and row[('type')].to_list()[0] == 'count':
            baserow = row
        for col,val in row.items():
            if baserow is None:
                usebase = np.nan
            else:
                usebase = baserow[col]
            basedf.loc[idx,col] = usebase
    basedf = basedf[~basedf.index.duplicated(keep='first')]
    # --------------------------------------------------------------------------
    # Unweighted df
    # --------------------------------------------------------------------------
    unwbdf = tabdf.copy()
    baserow = None
    for idx,row in unwbdf.iterrows():
        if ('unweighted' in idx.lower()) and row[('type')].to_list()[0] == 'count':
            baserow = row
        for col,val in row.items():
            if baserow is None:
                usebase = np.nan
            else:
                usebase = baserow[col]
            unwbdf.loc[idx,col] = usebase
    unwbdf = unwbdf[~unwbdf.index.duplicated(keep='first')]
    # --------------------------------------------------------------------------
    # Effective df
    # --------------------------------------------------------------------------
    effbdf = tabdf.copy()
    baserow = None
    for idx,row in effbdf.iterrows():
        if ('effective' in idx.lower()) and row[('type')].to_list()[0] == 'count':
            baserow = row
        for col,val in row.items():
            if baserow is None:
                usebase = np.nan
            else:
                usebase = baserow[col]
            effbdf.loc[idx,col] = usebase
    effbdf = effbdf[~effbdf.index.duplicated(keep='first')]
    # --------------------------------------------------------------------------
    # Clean up the statdf in case there are bases
    # --------------------------------------------------------------------------
    rmlist = []
    for idx in statdf.index:
        if str(idx).lower()[:4] == 'base':
            rmlist.append(idx)
        if 'unweighted' in str(idx).lower():
            rmlist.append(idx)
        if 'effective' in str(idx).lower():
            rmlist.append(idx)
    statdf.drop(rmlist,inplace=True)
    # --------------------------------------------------------------------------
    # Drop the type column
    # --------------------------------------------------------------------------
    del countdf[('type')]
    del statdf[('type')]
    del pctdf[('type')]
    del basedf[('type')]
    del unwbdf[('type')]
    del effbdf[('type')]
    # --------------------------------------------------------------------------
    # Set the column types appropriately for the count and pct dataframes
    # --------------------------------------------------------------------------
    for col in countdf.columns:
        countdf.loc[:,col] = countdf[col].copy().replace(tabzeros).astype(float)
    for col in pctdf.columns:
        pctdf.loc[:,col] = pctdf[col].copy().replace(combzeros).astype(float)
    # --------------------------------------------------------------------------
    # If any of the dataframes are empty, replace with None
    # --------------------------------------------------------------------------
    if countdf.empty:
        countdf = None
    if statdf.empty:
        statdf  = None
    if pctdf.empty:
        pctdf   = None
    # --------------------------------------------------------------------------
    # If there is a statdf that is not none, drop the last row of the multiindex
    # --------------------------------------------------------------------------
    if statdf is not None:
        if countdf is not None:
            countdf.columns = countdf.columns.droplevel(-1)
        if pctdf is not None:
            pctdf.columns   = pctdf.columns.droplevel(-1)
        if basedf is not None:
            basedf.columns  = basedf.columns.droplevel(-1)
        if unwbdf is not None:
            unwbdf.columns  = unwbdf.columns.droplevel(-1)
        if effbdf is not None:
            effbdf.columns  = effbdf.columns.droplevel(-1)
    # --------------------------------------------------------------------------
    # Finish
    # --------------------------------------------------------------------------
    out = {}
    out['count'] = countdf
    out['pct']   = pctdf
    out['stat']  = statdf
    out['base']  = basedf
    out['unwb']  = unwbdf
    out['effb']  = effbdf
    return out
