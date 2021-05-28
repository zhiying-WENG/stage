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
      cam=request.form['cam']
      lead=request.form['lead']
      cout=request.form['cout']
      clics=request.form['clics']
      impr=request.form["impr"]
      pi=request.form['pi']
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
      response=macro_net(df,cam,lead,cout,clics,impr,pi,u1,u2,u3,u4,u5,u6,u7,u8,u9,u10)
    except Exception as ex:
      return render_template('error_excel.html',ex=ex)
   
  return response

if __name__ == "__main__":
  app.run()
