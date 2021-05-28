import pandas as pd
import numpy as np

import io
from flask import make_response


"""
1) Charger un excel

input : path_file
output: df
"""

def load_excel(path_file):
    return pd.read_excel(path_file)


"""
2) Macro nettoyage

"""

def macro_net(df,cam,lead,cout,clics,impr,pi,u1,u2,u3,u4,u5,u6,u7,u8,u9,u10):
    if cam=="":
        cam="Campaign"
    if lead=="":
        lead="Compte API"
    if cout=="":
        cout="Cost"
    if clics=="":
        clics="Clicks"
    if impr=="":
        impr="Impr"
    if pi=="":   
        pi="Search impr share"
    if u1=="":
        u1="Autres"
    if u2=="": 
        u2="Autres"
    if u3=="":
        u3="Autres"
    if u4=="":
        u4="Autres"
    if u5=="":
        u5="Autres"
    if u6=="":   
        u6="Autres"
    if u7=="":   
        u7="Autres"
    if u8=="":  
        u8="Autres"
    if u9=="":
        u9="Autres"
    if u10=="":
        u10="Autres"
   
    ls=[cam,lead,cout,clics,impr,pi]
    df=df[df.Cost>0].reset_index(drop=True)
    df=df[ls]
   
    def Univers(x):
        if "_H_" in x:
            if "_ACQ_" in x:
                x="TURF ACQ"
            elif "APS" in x:
                x="TURF APS"
            elif "AVT" in x:
                x="TURF AVT"
            elif "PDT" in x:
                x="TURF PDT"
            else:
                x="TURF"
        elif "_M_" in x:
            x="MARQUE"
        elif "_POK_" in x:
            x="POKER"
        elif ("_S_" in x) or ("_F_" in x) or ("_FOOT_" in x) or ("Foot" in x) or ("_PAR_" in x):
            x="SPORT"
        else:
            x=""
        return x
    def UniversGen(x):
        if u1 in x:
            x=u1
        elif u2 in x:
            x=u2
        elif u3 in x:
            x=u3
        elif u4 in x:
            x=u4
        elif u5 in x:
            x=u5
        elif u6 in x:
            x=u6
        elif u7 in x:
            x=u7
        elif u8 in x:
            x=u8
        elif u9 in x:
            x=u9
        elif u10 in x:
            x=u10
        else:
            x="Autres"
        return x
    if cam=="Campaign" and lead=="Compte API":
        df["Univers"]=df.Campaign.apply(Univers)
        df=df[~(df.Univers=="TURF")].reset_index(drop=True)
    else:
        df["Univers"]=df[cam].apply(UniversGen)

    def SearchImpr(x):
        if type(x)==str:
            x=x.replace("<","").replace(">","").replace("%","")
            x=float(x)/100
        return x
    df[pi]=df[pi].apply(SearchImpr)
    # df["Requêtes"]=round(df[impr]/df[pi],2) quand il n'y a pas de nan
    # quand pi=nan impression/nan compte directement impression
    l_r=[]
    for i in range(len(df)):
        if df[pi][i]>0:
            l_r.append(round(df[impr][i]/df[pi][i],2))
        else:
            l_r.append(df[impr][i])
    df["Requêtes"]=l_r
    
    df.rename(columns={lead: "Lead",cout:"Coût",clics:"Clics",impr:"impression",pi:"PI estimé"},inplace=True)
    
    dfgu=df.groupby("Univers")[["Coût","Lead","Clics","impression","Requêtes"]].sum()
    
    dfgu["PI estimé"]=round(dfgu.impression/dfgu["Requêtes"],2)
    dfgu["CPA"]=round(dfgu.Coût/dfgu.Lead,2)
    dfgu["CPC estimé"]=round(dfgu.Coût/dfgu.Clics ,2)
    dfgu["CTR estimé"]=round(dfgu.Clics/dfgu.impression ,2)
    dfgu["TTR estimé"]=round(dfgu.Lead/dfgu.Clics ,4)
    dfgu=dfgu[["Requêtes","PI estimé","impression","Clics","CTR estimé","CPC estimé","Coût","TTR estimé","Lead","CPA"]]
    

    out = io.BytesIO()
    writer = pd.ExcelWriter(out, engine='xlsxwriter')
    dfgu.to_excel(excel_writer=writer, index=True, sheet_name='Data Search')
      
    writer.save()
    writer.close()
    
    file_name="data.xlsx"
    response = make_response(out.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=%s" % file_name
    response.headers["Content-type"] = "application/x-xls"

    return response
  

