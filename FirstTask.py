import pandas as pd
import datetime

file= r"C:\Users\DELL\Downloads\ExtractBrut_SQ00.xlsx"
df=pd.read_excel(file)
# set_index()
print(df)
#dropping the duplicate of colum UQ
df=df.drop(df.columns[ [8,10,13,15] ],axis=1)
print(df)
#drop last row
df.drop(df.tail(1).index,inplace=True)
#rename columns
df.rename(columns = {'Doc.inven.':'inventory_doc',

                    'Article':'material',
                    'Désignation article':'designation',
                    'TyAr':'type',
                    'UQ':'unit',
                    'Mag.':'store',
                    'Fourn.':'supplier',
                    'Quantité théorique':'theoritical-quantity',
                    'Quantité saisie':'entred_quantity',
                    'écart enregistré':'deviation',
                    'Ecart (montant)':'deviation_cost',
                    'Dev.':'dev',
                    'Sup':'delete',
                    'Dte cptage':'date_catchment',
                    'Rectifié par':'corrected_by',
                    'cpt':'catchment',
                    'Dev':'dev',
                    'Réf.inventaire':'refecrence_inventory',
                    'N° inventaire':'inventory_number',
                    'TyS':'Tys'},  inplace = True)


#Adding the year and week columns
#get current year and week
year=datetime.datetime.today().isocalendar()[0]
week=datetime.datetime.today().isocalendar()[1]
#insert year and week in first and second position
df.insert(0, 'year', year)
df.insert(1, 'week', week)


df.to_excel(r"C:\Users\DELL\Downloads\Saved_data.xlsx")


