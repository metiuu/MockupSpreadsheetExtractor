# Flask initialization
from flask import *
from werkzeug.utils import secure_filename

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


app.config["ALLOWED_FILE_TYPES"] = ["XLS", "XLSX"]


def check_file_type(filename):
    if not "." in filename:
        return False

    extension = filename.rsplit(".", 1)[1]

    if extension.upper() in app.config["ALLOWED_FILE_TYPES"]:
        return True
    else:
        return


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


if __name__ == '__main__':
    app.run(debug=True)
