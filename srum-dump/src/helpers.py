from Registry import Registry

import pathlib
import struct
import os
import tempfile
import urllib.request
import subprocess


def load_registry_sids(reg_file):
    """Given Software hive find SID usernames"""
    sids = {}
    profile_key = r"Microsoft\Windows NT\CurrentVersion\ProfileList"
    tgt_value = "ProfileImagePath"
    try:
        reg_handle = Registry.Registry(reg_file)
        key_handle = reg_handle.open(profile_key)
        for eachsid in key_handle.subkeys():
            sids_path = eachsid.value(tgt_value).value()
            sids[eachsid.name()] = sids_path.split("\\")[-1]
    except:
        return {}
    return sids

def load_interfaces(reg_file):
    """Loads the names of the wireless networks from the software registry hive"""
    try:
        reg_handle = Registry.Registry(reg_file)
    except Exception as e:
        print(r"I could not open the specified SOFTWARE registry key. It is usually located in \Windows\system32\config.  This is an optional value.  If you cant find it just dont provide one.")
        print(("WARNING : ", str(e)))
        return {}
    try:
        int_keys = reg_handle.open('Microsoft\\WlanSvc\\Interfaces')
    except Exception as e:
        print("There doesn't appear to be any wireless interfaces in this registry file.")
        print(("WARNING : ", str(e)))
        return {}
    profile_lookup = {}
    for eachinterface in int_keys.subkeys():
        if len(eachinterface.subkeys())==0:
            continue
        for eachprofile in eachinterface.subkey("Profiles").subkeys():
            profileid = [x.value() for x in list(eachprofile.values()) if x.name()=="ProfileIndex"][0]
            metadata = list(eachprofile.subkey("MetaData").values())
            for eachvalue in metadata:
                if eachvalue.name() in ["Channel Hints", "Band Channel Hints"]:
                    channelhintraw = eachvalue.value()
                    hintlength = struct.unpack("I", channelhintraw[0:4])[0]
                    name = channelhintraw[4:hintlength+4] 
                    profile_lookup[str(profileid)] = name.decode(encoding="latin1")
    return profile_lookup

def load_template_lookups(template_workbook):
    """Load any tabs named lookup-xyz form the template file for lookups of columns with the same format type"""
    template_lookups = {}
    for each_sheet in template_workbook.get_sheet_names():
        if each_sheet.lower().startswith("lookup-"):
            lookupname = each_sheet.split("-")[1]
            template_sheet = template_workbook.get_sheet_by_name(each_sheet)
            lookup_table = {}
            for eachrow in range(1,template_sheet.max_row+1):
                value = template_sheet.cell(row = eachrow, column = 1).value
                description = template_sheet.cell(row = eachrow, column = 2).value
                lookup_table[value] = description
            template_lookups[lookupname] = lookup_table
    return template_lookups
    
def load_template_tables(template_workbook):
    """Load template tabs that define the field names and formats for tables found in SRUM"""
    template_tables = {}    
    sheets = template_workbook.get_sheet_names()
    for each_sheet in sheets:
        #open the first sheet in the template
        template_sheet = template_workbook.get_sheet_by_name(each_sheet)
        #retieve the name of the ESE table to populate the sheet with from A1
        ese_template_table = template_sheet.cell(row=1,column=1).value
        #retrieve the names of the ESE table columns and cell styles from row 2 and format commands from row 3 
        template_field = {}
        #Read the first Row B & C in the template into lists so we know what data we are to extract
        for eachcolumn in range(1,template_sheet.max_column+1):
            field_name = template_sheet.cell(row = 2, column = eachcolumn).value
            if field_name == None:
                break
            template_style = template_sheet.cell(row = 4, column = eachcolumn).style
            template_format = template_sheet.cell(row = 3, column = eachcolumn).value
            template_value = template_sheet.cell(row = 4, column = eachcolumn ).value
            if not template_value:
                template_value= field_name
            template_field[field_name] = (template_style,template_format,template_value)
        template_tables[ese_template_table] = (each_sheet, template_field)
    return template_tables  

def extract_live_file():
    try:
        tmp_dir = tempfile.mkdtemp()
        fget_file = pathlib.Path(tmp_dir) / "fget.exe"
        registry_file = pathlib.Path(tmp_dir) / "SOFTWARE"
        extracted_srum = pathlib.Path(tmp_dir) / "srudb.dat"
        esentutl_path = pathlib.Path(os.environ.get("COMSPEC")).parent / "esentutl.exe"
        if esentutl_path.exists():
            cmdline = r"{} /y c:\\windows\\system32\\sru\\srudb.dat /vss /d {}".format(str(esentutl_path), str(extracted_srum))
            phandle = subprocess.Popen(cmdline, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out1, _ = phandle.communicate()
            cmdline = r"{} /y c:\\windows\\system32\\config\\SOFTWARE /vss /d {}".format(str(esentutl_path), str(registry_file))
            phandle = subprocess.Popen(cmdline, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out2, _ = phandle.communicate()
        else:
            fget_binary = urllib.request.urlopen('https://github.com/MarkBaggett/srum-dump/raw/master/FGET.exe').read()
            fget_file.write_bytes(fget_binary)
            cmdline = r"{} -extract c:\\windows\\system32\\sru\srudb.dat {}".format(str(fget_file), str(extracted_srum))
            phandle = subprocess.Popen(cmdline, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out1, _ = phandle.communicate()
            cmdline = r"{} -extract c:\\windows\\system32\\config\SOFTWARE {}".format(str(fget_file), str(registry_file))
            phandle = subprocess.Popen(cmdline, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out2, _ = phandle.communicate()
            fget_file.unlink()
    except Exception as e:
        print("Unable to automatically extract srum. {}\n{}\n{}".format(str(e), out1.decode(), out2.decode()))
        return None
    if (b"returned error" in out1 + out2) or (b"Init failed" in out1 + out2):
        print("ERROR\n SRUM Extraction: {}\n Registry Extraction {}".format(out1.decode(), out2.decode()))
    elif b"success" in out1.lower() and b"success" in out2.lower():
        return str(extracted_srum), str(registry_file)
    else:
        print("Unable to determine success or failure.", out1.decode(), "\n", out2.decode())
    return None

