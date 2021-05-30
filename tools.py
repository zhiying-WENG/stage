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

def macro_net(df,cam,lead,cout,clics,impr,pi,pmu,n1,n2,n3,n4,n5,n6,agg1,agg2,agg3,agg4,agg5,agg6,gs,ordre,u1,u2,u3,u4,u5,u6,u7,u8,u9,u10):
    # initialiser les valeurs par defaut si utilisateur n'est pas input
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
    # liste de colonnes on veut garder
    ls=[cam,lead,cout,clics,impr,pi]  
    # liste de colonnes pour groupby 
    lgroupby=["Univers"]
    if gs!="" and gs!=cam:
        ls+=[gs]
        if ordre=="nc":
            lgroupby=[gs]
        elif ordre=="unc":
            lgroupby+=[gs]
        elif ordre=="ncu":
            lgroupby=[gs]+lgroupby
    # liste de nouvelle colonnes ajouter
    ladd=[]
    # dictionary des nom de colonnes avec ses façon de calculer
    dictgroupby={"Coût":"sum","Lead":"sum","Clics":"sum","impression":"sum","Requêtes":"sum"}
    lnumeric=[np.float,np.float16,np.float32,np.float64,np.int,np.int16,np.int32,np.int64]
    if (n1!="") and (n1 not in ls) and (df[n1].dtype in lnumeric):
        ladd.append(n1)
        dictgroupby[n1]=agg1
    if (n2!="") and (n2 not in ls) and (df[n2].dtype in lnumeric):
        ladd.append(n2)
        dictgroupby[n2]=agg2
    if (n3!="") and (n3 not in ls) and (df[n3].dtype in lnumeric):
        ladd.append(n3)
        dictgroupby[n3]=agg3
    if (n4!="") and (n4 not in ls) and (df[n4].dtype in lnumeric):
        ladd.append(n4)
        dictgroupby[n4]=agg4
    if (n5!="") and (n2 not in ls) and (df[n5].dtype in lnumeric):
        ladd.append(n5)
        dictgroupby[n5]=agg5
    if (n6!="") and (n6 not in ls) and (df[n6].dtype in lnumeric):
        ladd.append(n6)
        dictgroupby[n6]=agg6
    ls+=ladd
    # garder que cost>0 et les colonnes dans le liste de colonnes on veut garder
    df=df[df.Cost>0].reset_index(drop=True)
    df=df[ls]
    # fonction pour traiter univers special pour pmu
    def UniversPMU(x):
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
    # fonction pour traiter univers general (hors pmu)
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
    if pmu=="ispmu":
        df["Univers"]=df.Campaign.apply(UniversPMU)
        df=df[~(df.Univers=="TURF")].reset_index(drop=True)
    else:
        df["Univers"]=df[cam].apply(UniversGen)
    # fonction pour traiter search impr share (PI)
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
    # renomer les colonnes
    df.rename(columns={lead: "Lead",cout:"Coût",clics:"Clics",impr:"impression",pi:"PI estimé"},inplace=True)
    
    # liste de colonnes final dans le excel
    lfinal=["Requêtes","PI estimé","impression","Clics","CTR estimé","CPC estimé","Coût","TTR estimé","Lead","CPA"]+ladd

    # les liste de  colonnes on veut garder quand groupby:  ["Coût","Lead","Clics","impression","Requêtes"]+ladd
    # on groupby et calculer
    dfgu=df.groupby(lgroupby).agg(dictgroupby)
    dfgu["PI estimé"]=round(dfgu.impression/dfgu["Requêtes"],2)
    dfgu["CPA"]=round(dfgu.Coût/dfgu.Lead,2)
    dfgu["CPC estimé"]=round(dfgu.Coût/dfgu.Clics ,2)
    dfgu["CTR estimé"]=round(dfgu.Clics/dfgu.impression ,2)
    dfgu["TTR estimé"]=round(dfgu.Lead/dfgu.Clics ,4)
    dfgu=dfgu[lfinal]
    
    # excel avec les donnes déjà nettoyer
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
  

