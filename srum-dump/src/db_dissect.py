import pathlib

from dissect.esedb import EseDB
from dissect.esedb.tools.sru import SRU
from dissect.esedb.table import Table

sru_path = pathlib.Path("c:/users/lo127/Documents/git/srum-dump3/SRUDB.DAT")

sru_path = pathlib.Path(r"C:/Users/lo127/Documents/Tools/srum_examples/SRU/SRU/SRUDB.dat")
with sru_path.open("rb") as f:
    try:
        sru = SRU(f)
    except:
        pass
    for x in sru.entries():
        print(x)
        
    table = sru.get_table("NetworkUsage")
    for record in table.records:
        print(record)
        