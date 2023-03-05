import openpyxl
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/slice_string', methods=['POST'])
def slice_string():
    my_string = request.form['doc_id']
    fixed_width = 198

    num_chunks = (len(my_string) + fixed_width - 1) // fixed_width
    sliced_string = ""
    for i in range(0, len(my_string), fixed_width):
        chunk = my_string[i:i+fixed_width]
        new_chunk = chunk[:-1]
        sliced_string += new_chunk + "\n"

    return render_template('output.html', sliced_string=sliced_string)

if __name__ == '__main__':
    app.run(debug=True)
