from flask import Flask,render_template,request
import tempfile
from tools import *
import config

app = Flask(__name__)

@app.route('/')
def use():
  return render_template(config.page_home)


@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
  if request.method == 'POST':
    try:
      f = request.files['file']
      # c'est le excel pour PMU ou hors PMU
      pmu=request.form['pmu']
      # occuper les nom de colonns correspend
      cam=request.form['cam']
      lead=request.form['lead']
      cout=request.form['cout']
      clics=request.form['clics']
      impr=request.form["impr"]
      pi=request.form['pi']
      # ajouter les nouvelle colonnes
      n1=request.form['n1']
      n2=request.form['n2']
      n3=request.form['n3']
      n4=request.form['n4']
      n5=request.form['n5']
      n6=request.form['n6']
      # les façons de calculer pour les nouvelle colonnes
      agg1=request.form['agg1']
      agg2=request.form['agg2']
      agg3=request.form['agg3']
      agg4=request.form['agg4']
      agg5=request.form['agg5']
      agg6=request.form['agg6']
      # occuper le nom de nouvelle colonne pour groupby
      gs=request.form['gs']
      # le ordre pour groupby
      ordre=request.form['group']
      # les elements pour Univers
      u1=request.form['u1']
      u2=request.form['u2']
      u3=request.form['u3']
      u4=request.form['u4']
      u5=request.form['u5']
      u6=request.form['u6']
      u7=request.form['u7']
      u8=request.form['u8']
      u9=request.form['u9']
      u10=request.form['u10']
      # save excel
      tempfile_path=tempfile.NamedTemporaryFile().name
      f.save(tempfile_path)
    except Exception as err:
      return render_template('home.html',err=err)
      
    # Je charge mon df
    try:
      df = load_excel(tempfile_path)
    except Exception as e:
      return render_template('home.html',e=e)

    # Je nettoyer df 
    try:
      response=macro_net(df,cam,lead,cout,clics,impr,pi,pmu,n1,n2,n3,n4,n5,n6,agg1,agg2,agg3,agg4,agg5,agg6,gs,ordre,u1,u2,u3,u4,u5,u6,u7,u8,u9,u10)
    except Exception as ex:
      return render_template('error_excel.html',ex=ex)
   
  return response

if __name__ == "__main__":
  app.run()
