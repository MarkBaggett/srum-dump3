import argparse
import os
import pathlib
import warnings
import openpyxl
import itertools
import sys

# Import the desired UI and DB modules
from ui_tk import show_live_system_warning, get_user_input
#from db_dissect import load_srumid_lookups, process_srum, EseDB
from db_ese import load_srumid_lookups, process_srum, EseDB
from helpers import load_interfaces, load_registry_sids
from helpers import load_template_tables, load_template_lookups



parser = argparse.ArgumentParser(description="Given an SRUM database it will create an XLS spreadsheet with analysis of the data in the database.")
parser.add_argument("--SRUM_INFILE", "-i", help="Specify the ESE (.dat) file to analyze. Provide a valid path to the file.")
parser.add_argument("--XLSX_OUTFILE", "-o", default="SRUM_DUMP_OUTPUT.xlsx", help="Full path to the XLS file that will be created.")
parser.add_argument("--XLSX_TEMPLATE", "-t", help="The Excel Template that specifies what data to extract from the srum database. You can create template_tables with ese_template.py.")
parser.add_argument("--REG_HIVE", "-r", dest="reghive", help="If SOFTWARE registry hive is provided then the names of the network profiles will be resolved.")
parser.add_argument("--quiet", "-q", help="Suppress unneeded output messages.", action="store_true")
options = parser.parse_args()

ads = itertools.cycle([
    "Did you know SANS Automating Infosec with Python SEC573 teaches you to develop Forensics and Incident Response tools?",
    "To learn how SRUM and other artifacts can enhance your forensics investigations check out SANS Windows Forensic Analysis FOR500.",
    "Yogesh Khatri's paper at https://github.com/ydkhatri/Presentations/blob/master/SRUM%20Forensics-SANS.DFIR.summit.2015.pdf was essential in the creation of this tool.",
    "By modifying the template file you have control of what ends up in the analyzed results.  Try creating an alternate template and passing it with the --XLSX_TEMPLATE option.",
    "TIP: When using a SOFTWARE registry file you can add your own SIDS to the 'lookup-Known SIDS' tab!",
    "This program was written by Twitter:@markbaggett and @donaldjwilliam5 because @ovie said so.",
    "SRUM-DUMP 2.0 will attempt to dump any ESE database! If no template defines a table it will do its best to guess."
])

if not options.SRUM_INFILE:
    get_user_input(options)

if not options.XLSX_TEMPLATE:
    options.XLSX_TEMPLATE = "SRUM_TEMPLATE2.xlsx"
if not options.XLSX_OUTFILE:
    options.XLSX_OUTFILE = "SRUM_DUMP_OUTPUT.xlsx"
if not os.path.exists(options.SRUM_INFILE):
    print("ESE File Not found: " + options.SRUM_INFILE)
    sys.exit(1)
if not os.path.exists(options.XLSX_TEMPLATE):
    print("Template File Not found: " + options.XLSX_TEMPLATE)
    sys.exit(1)
if options.reghive and not os.path.exists(options.reghive):
    print("Registry File Not found: " + options.reghive)
    sys.exit(1)

regsids = {}
if options.reghive:
    interface_table = load_interfaces(options.reghive)
    regsids = load_registry_sids(options.reghive)

try:
    warnings.simplefilter("ignore")
    ese_db = EseDB(options.SRUM_INFILE)
except Exception as e:
    print("I could not open the specified SRUM file. Check your path and file name.")
    print("Error : ", str(e))
    sys.exit(1)

try:
    template_wb = openpyxl.load_workbook(filename=options.XLSX_TEMPLATE)
except Exception as e:
    print("I could not open the specified template file %s. Check your path and file name." % (options.XLSX_TEMPLATE))
    print("Error : ", str(e))
    sys.exit(1)

skip_tables = ['MSysObjects', 'MSysObjectsShadow', 'MSysObjids', 'MSysLocales', 'SruDbIdMapTable', "SruDbCheckpointTable"]
template_tables = load_template_tables(template_wb)
template_lookups = load_template_lookups(template_wb)
if regsids:
    template_lookups.get("Known SIDS", {}).update(regsids)
id_table = load_srumid_lookups(ese_db)

target_wb = openpyxl.Workbook()
process_srum(ese_db, target_wb)

firstsheet = target_wb["Sheet"]
del target_wb[firstsheet]
print("Writing output file to disk.")
try:
    target_wb.save(options.XLSX_OUTFILE)
except Exception as e:
    print("I was unable to write the output file.  Do you have an old version open?  If not this is probably a path or permissions issue.")
    print("Error : ", str(e))
print("Done.")