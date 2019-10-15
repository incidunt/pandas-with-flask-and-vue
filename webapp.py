from flask import Flask, request, render_template
from werkzeug.utils import secure_filename

import jinja2

import pandas as pd
import numpy as np
import imgkit
import os
import random
import shutil


def DataFrame_to_image(data, css, outputfile,format="png"):
    '''
    For rendering a Pandas DataFrame as an image.
    data: a pandas DataFrame
    css: a string containing rules for styling the output table. This must 
         contain both the opening an closing <style> tags.
    *outputimage: filename for saving of generated image
    *format: output format, as supported by IMGKit. Default is "png"
    '''
    fn = str(random.random()*100000000).split(".")[0] + ".html"
    
    try:
        os.remove(fn)
    except:
        None
    text_file = open(fn, "a")
    
    # write the CSS
    text_file.write(css)
    # write the HTML-ized Pandas DataFrame

    text_file.write(data.to_html(index = False))
    text_file.close()
    
    # See IMGKit options for full configuration,
    # e.g. cropping of final image
    imgkitoptions = {"format": format, 'encoding': "UTF-8"}
    imgkit.from_file(fn, outputfile, options=imgkitoptions)
    os.remove(fn)


css = """
<style type=\"text/css\">

table {
color: #333;
border-collapse:
collapse; 
border-spacing: 0;
}
tr{
text-align:left
}
td, th {
border: 1px solid ; /* No more visible border */
height: 30px;
}

th {
background: #DFDFDF; /* Darken header a bit */
font-weight: bold;
}

td {
background: #FAFAFA;
}

table tr:nth-child(odd) td{
background-color: white;
}
</style>
"""

app = Flask(__name__,
            static_url_path='', 
            static_folder='static')

UPLOAD_FOLDER = './uploads'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

app.jinja_loader = jinja2.FileSystemLoader('./template')

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':

        file = request.files['file']

        full_filename = secure_filename(file.filename)

        first_filename =full_filename.rsplit('.', 1)[0].lower() 
        outputdir=os.path.join('./output',first_filename)

        file.save(os.path.join(app.config['UPLOAD_FOLDER'], full_filename))
        os.mkdir( outputdir )
        
        df = pd.read_excel(file).replace(np.nan,'')

        staff=df.iloc[:,0].unique()

        for s in staff:
            r = df.loc[df['姓名'] == s]
            img_name=os.path.join(outputdir,s+".png")
            DataFrame_to_image(r,css,img_name)
            
        zip_file= "./static/download/"+first_filename
        shutil.make_archive(zip_file, 'zip', outputdir)
        shutil.rmtree(outputdir)

        return render_template('upload.html', download_file=zip_file+".zip")
    return render_template('upload.html')

@app.route('/static/<path:path>')
def static_file(path):
    return app.send_static_file(path)

if __name__ == '__main__':
    app.run(debug=True)