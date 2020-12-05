# Flask initialization
from flask import *
import xlrd
from werkzeug.utils import secure_filename

app = Flask(__name__)

# App configs
app.config["ALLOWED_FILE_TYPES"] = ["XLS", "XLSX"]


@app.route('/')
def index():
    return render_template("index.html")


@app.route("/upload-file", methods=["GET", "POST"])
def upload_file():

    if request.method == "POST":
        if request.files:
            book = request.files["xls"]

            # upload validation
            if book.filename == "":
                print("File must have a name")
                return redirect(request.url)

            if not check_file_type(book.filename):
                print("Illegal file type.")
                return redirect(request.url)
            else:
                filename = secure_filename(book.filename)

            print("File upload successful")
            content = xlrd.open_workbook(file_contents=book.read())
            for sheet in content.sheets():
                get_student_data(sheet)
                total = count_worksheets(sheet)
                print('Total worksheets done:', total)
                total = 0
                last_workbook(sheet)
            return redirect(request.url)

    return render_template("upload_file.html")

def check_file_type(filename):
    if not "." in filename:
        return False

    extension = filename.rsplit(".", 1)[1]

    if extension.upper() in app.config["ALLOWED_FILE_TYPES"]:
        return True
    else:
        return


def get_student_data(sheet):
    # extract data from specific cells
    print('Student name:',sheet.cell(1,7).value)
    print('Subject:',sheet.cell(2,1).value)
    print('Month:',sheet.cell(2,0).value, int(sheet.cell(2,3).value))
    print()

def count_worksheets(sheet):
    total = 0
    for row in range(5,sheet.nrows):
        for col in range(4,14):
            cell_type = sheet.cell_type(row,col)

            if not cell_type == xlrd.XL_CELL_EMPTY:
                total += 1
    return total

def last_workbook(sheet):
    # last workbook done (last row of spreadsheet, don't know how many rows there are)
    # last row(starting row 6), columns 2 and 3
    print('Last worksheet done:', sheet.col_values(1,5)[-1] + ' ' + str(int(sheet.col_values(2,5)[-1])))
    


if __name__ == '__main__':
    app.run(debug=True)
