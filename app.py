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
            spreadsheet = request.files["xls"]

            # upload validation
            if spreadsheet.filename == "":
                print("File must have a name")
                return redirect(request.url)

            if not check_file_type(spreadsheet.filename):
                print("Illegal file type.")
                return redirect(request.url)
            else:
                filename = secure_filename(spreadsheet.filename)

            print("File upload successful")
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


def get_student_data(file):
    # extract data from specific cells
    worksheet = file.sheet_by_index(0) # maybe we can make it so that the app goes through all sheets in a xls/xlsx file?
    print('Student name:',worksheet.cell(1,7).value)
    print('Subject:',worksheet.cell(2,1).value)
    print('Month:',worksheet.cell(2,0).value, int(worksheet.cell(2,3).value))
    print()

def count_worksheets(file):
    # total worksheets done (starting row 6, columns 5-14)
    worksheet = file.sheet_by_index(0) # maybe we can make it so that the app goes through all sheets in a xls/xlsx file?
    total = 0
    for row in range(5,worksheet.nrows):
        for col in range(4,14):
            cell_type = worksheet.cell_type(row,col)

            if not cell_type == xlrd.XL_CELL_EMPTY:
                total += 1
    return total

def last_workbook(file):
    # last workbook done (last row of spreadsheet, don't know how many rows there are)
    # last row(starting row 6), columns 2 and 3
    worksheet = file.sheet_by_index(0) # maybe we can make it so that the app goes through all sheets in a xls/xlsx file?
    return worksheet.col_values(1,5)[-1] + str(int(worksheet.col_values(2,5)[-1]))
    


if __name__ == '__main__':
    app.run(debug=True)
