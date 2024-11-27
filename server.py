from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import io

app = Flask(__name__)

@app.route("/")
def index():
    # Serve the HTML page with the file upload form
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process_excel():
    try:
        # Check if a file is included in the request
        if "file" not in request.files:
            return jsonify({"error": "No file part in the request"}), 400

        file = request.files["file"]

        # Ensure the file has been uploaded
        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400

        # Read the Excel file
        try:
            df = pd.read_excel(file)
        except Exception as e:
            return jsonify({"error": f"Failed to read Excel file: {str(e)}"}), 400

        # Sample Processing (Modify as needed)
        if 'Client_REF_Number' in df.columns:
            # Remove spaces from 'Client_REF_Number'
            df['Client_REF_Number'] = df['Client_REF_Number'].astype(str).str.replace(' ', '')

        # Adding missing columns with default values
        for col in ['Vendor_Name', 'Bank_Account_Number']:
            if col not in df.columns:
                df[col] = None

        # Create an in-memory byte stream
        output = io.BytesIO()

        # Write the processed DataFrame to the byte stream
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Processed Data")

        # Seek to the beginning of the byte stream
        output.seek(0)

        # Send the file as a download
        return send_file(output, as_attachment=True, download_name="processed_file.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host='0.0.0.0', debug=False)