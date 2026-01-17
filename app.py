from flask import Flask, render_template, request, send_file
import os
import writer

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        form_data = {
            'fio': request.form.get('fio'),
            'address': request.form.get('address'),
            'series': request.form.get('passport_series'),
            'number': request.form.get('passport_number'),
            'issued': request.form.get('passport_issued'),
            'date': request.form.get('date_sign')
        }

        filename = "soglasie.docx"
        filepath = os.path.join('/tmp', filename)
        writer.save(form_data, filepath)

        return send_file(filepath, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)
