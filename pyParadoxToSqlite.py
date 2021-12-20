import sys
import os
from pypxlib import Table
import xlsxwriter
import pandas as pd
from pypxlib.pxlib_ctypes import *
from sqlalchemy import create_engine

# tel de records met de vanilla PX libary, de python wrap werkte wat boggy
def countRecords(filename):
  pxdoc = PX_new()
  filename = filename.encode()
  PX_open_file(pxdoc, filename)
  num_records = PX_get_num_records(pxdoc)
  PX_close(pxdoc)
  PX_delete(pxdoc)
  return num_records

# creeer een dictionary
def add( col, row, wert, data):
  if row in data:
    data[row][col] = wert
  else:  
    data[row] = {col:wert}  
  return data

#gejat van een duitser die het wel voor elkaar had op door de records heen te loopen.
def process(path, filename):  
  if countRecords(path+filename) == 0:
    return
  
  try:
    table = Table(f"{path+filename}")      # Paradox.db einlesen
    print('fpogp')
    x = table.fields                          # Felder ermitteln
    tablenspalten = x.keys()
    liste_fuer_tablenspalten_namen = []
    for i in tablenspalten:
        liste_fuer_tablenspalten_namen.append(i)
    
    row1 = 0
    col1 = 0
    col = 1     # dannach col auf 1 setzen, um den tableninhalt erst in zeile zwei zu beginnen
    data = {}
    for rowid in range(len(table)):
        Datensatz1 = table[rowid]
        row = 0 # row auf 0 setzen, um von links anzufangen
        for i in range(len(liste_fuer_tablenspalten_namen)):  # i = zahl
            aktuelles_feld = f"{liste_fuer_tablenspalten_namen[i]}" # z.B. aktuelles feld = str "Artikel"
            try:
              wert = (getattr(Datensatz1, aktuelles_feld))
            except:
              wert = None
            finally:
              asdf=1

            if wert != None:
                data = add(liste_fuer_tablenspalten_namen[i], rowid, wert, data)
                row += 1
            else:   # nichts adden, aber ins nächste feld springen
                row += 1
        col += 1
    
    df = pd.DataFrame().from_dict(data, orient='index')
    t = filename.replace('.DB', '')
    engine = create_engine('sqlite:///gras.sqlite')
    df.to_sql(t, con=engine, if_exists='append')
    print('save to sqlite ', len(data),' rows')

  except:
    table.close()
  finally:
    table.close()    

# niet de meest elegante oplossing, maar met een try catch liep het projgramma toch vast
# vandaar iedere keer een nieuwe Python instance. Dat werkt wel
if len(sys.argv) == 3:
  process(sys.argv[1],sys.argv[2])
else:
  os.system('rm gras.sqlite')
  path = 'db/00018370/'
  dir = os.scandir(path)
  for f in dir:
    if '.DB' in f.name:
      print('python3 gras.py '+path+' '+f.name)
      os.system('python3 gras.py '+path+' '+f.name)
    else:
      print('skip ', f.name)
