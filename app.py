from flask import Flask
from flask import render_template
import os
from flask import Flask, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from State_File_Prep import process_xlsx
UPLOAD_FOLDER = os.path.abspath("./uploads")
ALLOWED_EXTENSIONS = set(['xlsx'])

app = Flask(__name__)
## secret keys are needed for sessions
## please create one following this document
## http://flask.pocoo.org/docs/0.12/quickstart/#sessions
app.secret_key = "W\xc9\xa7\x1f!\xdb\xd3]\xf2@s\x1d\x9cL'\xfa}y\xe92R\x12\xefU"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def main():
    if request.method=="POST":
        if "file" not in request.files:
            flash("No file included")
            return redirect(request.url)
        file = request.files['file']
        if not allowed_file(file.filename):
            flash("File type not allowed")
            return redirect(request.url)
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        full_path = os.path.join(UPLOAD_FOLDER, filename)
        process_xlsx(full_path, request.form['submitter_id'],
                     request.form['auditor_id'],
                     request.form['batch_num'])
        flash("file processed successsfully")
        return redirect(request.url)
        
        
    return render_template("main.html")

if __name__ == '__main__':
    # app.run(debug=True)
    app.run(host='0.0.0.0', port=8000, debug=True)
