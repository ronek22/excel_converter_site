import os

from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, after_this_request
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from werkzeug.utils import secure_filename
from wtforms import SubmitField, FileField

from script.excel import convert

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(APP_ROOT, 'static', 'uploads')
DOWNLOAD_FOLDER = os.path.join(APP_ROOT, 'static', 'downloads')
ALLOWED_EXTENSIONS = {'xls'}

app = Flask(__name__)
app.secret_key = 'dev'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
bootstrap = Bootstrap(app)


class FileForm(FlaskForm):
    file = FileField('Arkusz do skonwertowania')
    submit = SubmitField("Konwertuj")


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    form = FileForm()
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('Nie dodałeś/aś pliku', 'error')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('Plik nie miał nazwy', 'error')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # convert
            converted_wb = convert(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            converted_wb.save(os.path.join(app.config['DOWNLOAD_FOLDER'], 'converted.xlsx'))
            return redirect(url_for('download_file', output='converted.xlsx', input_f=filename))
        else:
            flash("Tylko arkusze z rozszerzeniem .xsl", 'error')
    return render_template("index.html", form=form)


@app.route('/download/<input_f>/<output>', methods=['GET'])
def download_file(input_f, output):
    @after_this_request
    def remove_file(response):
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], input_f))
        except Exception as error:
            app.logger.error("Wystąpił problem przy usuwaniu pliku", error)
        return response
    flash("Konwertowanie zakończone pomyślnie", 'success')
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], output)



if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
