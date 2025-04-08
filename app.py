
from flask import Flask, request, send_file, render_template_string
import tempfile, zipfile, os
from SAPXMLTool import process_xmls

app = Flask(__name__)

HTML = """
<!doctype html>
<title>XML Processor</title>
<h1>Upload Excel and ZIP of XMLs</h1>
<form method=post enctype=multipart/form-data>
  Excel File: <input type=file name=excel><br><br>
  XML Zip File: <input type=file name=xmls><br><br>
  <input type=submit value=Process>
</form>
"""

@app.route("/test")
def test():
    return "It works!"

@app.route("/", methods=["GET", "POST"])
def upload_files():
    if request.method == "POST":
        excel = request.files["excel"]
        xmls_zip = request.files["xmls"]

        with tempfile.TemporaryDirectory() as tmpdir:
            xmls_path = os.path.join(tmpdir, "xmls")
            os.mkdir(xmls_path)

            xml_zip_path = os.path.join(tmpdir, "input.zip")
            excel_path = os.path.join(tmpdir, "input.xlsm")
            output_zip = os.path.join(tmpdir, "output.zip")

            xmls_zip.save(xml_zip_path)
            excel.save(excel_path)

            with zipfile.ZipFile(xml_zip_path, 'r') as zip_ref:
                zip_ref.extractall(xmls_path)

            output_folder = os.path.join(tmpdir, "output_xmls")
            os.mkdir(output_folder)

            process_xmls(excel_path, xmls_path, output_folder)

            with zipfile.ZipFile(output_zip, 'w') as zipf:
                for file in os.listdir(output_folder):
                    zipf.write(os.path.join(output_folder, file), file)

            return send_file(output_zip, as_attachment=True, download_name="processed_xmls.zip")

    return render_template_string(HTML)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
