from flask import Flask, render_template, request, send_file
import os
import writer

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        content = request.form.get('content')
        filepath = os.path.join('/tmp', 'test.txt')
        writer.save(content, filepath)
        return send_file(filepath, as_attachment=True, download_name='test.txt', mimetype='text/plain')
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)
