from flask import Flask, render_template, request, send_file
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# List of valid reasons
valid_reasons = [
    "No activities monitor blong",
    "No activities dispencer issue",
    "Safe lock not open",
    "No loading power failure",
    "Loading not done software issue",
    "Loading not power failure",
    "Software issue",
    "Access denied",
    "Road block so no loading return to office",
    "Loading skipped power off",
    "Loading not done power failure",
    "No loading card reader issue switch not updated",
    "No loading dispenser issue",
    "Dispenser issue already call log",
    "Captive site maintenance work permission denied",
    "Admin card not working",
    "Tuesday only loading",
    "Admin card not working",
    "EOD not done rp issue",
    "EOD not done card reader issue",
    "Loading not done link issue",
    "Loading not done RP issue",
    "Loading not done road block",
    "EOD not done. Rp issue.",
    "EOD not done. Ystd done.",
    "NO CLEARING. BRANCH CLOSED",
    "Vault not open",
    "Dispenser fault",
    "Vault battery down"
]

# Function to update LIABILITY and REMARKS
def update_liability_and_remarks(row):
    activity_status = str(row['ACTIVITY STATUS']).strip().upper()
    reason = str(row['REASON']).strip()
    
    if activity_status == 'YES':
        return 'ACTIVITY DONE', 'ACTIVITY DONE'
    elif activity_status == 'NO':
        if reason in valid_reasons:
            return 'MSP', reason
        else:
            return 'CMS', 'Skipped'
    else:
        return row['LIABILITY'], row['REMARKS']

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            # Load the Excel file
            df = pd.read_excel(filepath)

            # Fix ATM ID Column if exists
            if 'ATM ID' in df.columns:
                df['ATM ID'] = df['ATM ID'].apply(lambda x: str(x).replace('.0', '') if isinstance(x, float) else str(x))

            # Apply the function to update LIABILITY and REMARKS
            df[['LIABILITY', 'REMARKS']] = df.apply(lambda row: pd.Series(update_liability_and_remarks(row)), axis=1)

            # Save the updated file
            output_path = os.path.join(UPLOAD_FOLDER, 'updated_' + file.filename)
            df.to_excel(output_path, index=False)

            return send_file(output_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(environ.get("PORT", 5000)), debug=False)
