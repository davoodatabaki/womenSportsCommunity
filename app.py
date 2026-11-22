from flask import Flask,url_for,redirect,flash,render_template,request
from openpyxl import Workbook,load_workbook
import os
from pathlib import Path
from datetime import datetime
import json
import gspread
from google.oauth2.service_account import Credentials


app=Flask(__name__)
#app.secret_key = "dev-temp-key-replace-before-deploy"
app.secret_key = os.environ.get("SECRET_KEY", "dev-temp-key-replace-before-deploy")

base_dir=Path(__file__).resolve().parent

google_credentials_json =os.environ.get("GOOGLE_CREDENTIALS_JSON")
if not google_credentials_json:
    creds_path = base_dir / "credentials.json"
    if creds_path.exists():
        google_credentials_json = creds_path.read_text()

spreadsheet_name="womenSportsCommunityDB"
worksheet_name="Sheet1"

def get_gsheet_worksheet():
    """Create a gspread worksheet object using service account credentials."""
    if not google_credentials_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable is not set.")
    
    creds_dict=json.loads(google_credentials_json)
    scopes=["https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"]
    creds=Credentials.from_service_account_info(creds_dict,scopes=scopes)
    client=gspread.authorize(creds)

    # Open the spreadsheet and worksheet
    sheet=client.open(spreadsheet_name).worksheet(worksheet_name)
    return sheet
def save_to_sheet(row_data):
    """Append a new row to Google Sheets and return the new Id."""
    sheet=get_gsheet_worksheet()
    
    # Get existing Ids to calculate the new Id (assuming Id is in column A)
    existing_ids = sheet.col_values(1)[1:]  # Skip header row
    new_id = 1
    if existing_ids:
        numeric_ids = [int(i) for i in existing_ids if i.isdigit()]
        if numeric_ids:
            new_id = max(numeric_ids) + 1
            
    sheet.append_row([new_id] + row_data)
    return new_id
"""
base_dir=Path(__file__).resolve().parent
excel_file=base_dir/"database.xlsx"
sheet_name="Sheet1"
"""
Headers=["Id","Naam", "Email", "Leeftijd", "Speler", "Bestuur", "Sponsor", "Tijdstempel"]

"""
def init_excel():
    #create the excel file with the headers if it doesn't exist.
    if not os.path.exists(excel_file):
        wb=Workbook()
        ws=wb.active
        ws.title=sheet_name
        ws.append(Headers)
        wb.save(excel_file)

def get_new_id(ws):
    #return the next auto incerement Id based on existing rows.
    max_id=0
    for row in ws.iter_rows(min_row=2,max_col=1,values_only=True):
        value=row[0]
        if isinstance(value,int) and value>max_id:
            max_id=value
    return max_id+1
    
def save_to_excel(date):
    #append a new row with auto incerement Id.
    wb=load_workbook(excel_file)
    ws=wb.active
    new_id=get_new_id(ws)
    ws.append([new_id]+date)
    wb.save(excel_file)
    return new_id
"""




@app.route("/",methods=["GET","POST"])
def form():
    # Defaults — always defined, whether GET or POST
    naam = ""
    email = ""
    leeftijd = ""
    speler = False
    bestuur = False
    sponsor = False

    if request.method=="POST":
        #get text inputs
        naam= request.form.get("naam","").strip()
        email= request.form.get("email","").strip()
        leeftijd= request.form.get("leeftijd","").strip()

        #get check boxes : Ja/Nee
        speler= "Ja" if request.form.get("speler") else "Nee"
        bestuur= "Ja" if request.form.get("bestuur") else "Nee"
        sponsor= "Ja" if request.form.get("sponsor") else "Nee"

        #validation
        if not naam or not email or not leeftijd:
            flash("Naam, Emali, en Leeftijd zijn verplicht!","error")
            return render_template("form.html",
                                   naam=naam, email=email, leeftijd=leeftijd,
                                   speler=(speler == "Ja"),
                                    bestuur=(bestuur == "Ja"),
                                    sponsor=(sponsor == "Ja"),
                                    )

        try:
            leeftijd_int= int(leeftijd)
            if (leeftijd_int<1 or leeftijd_int>120):
                raise ValueError
        except ValueError:
            flash("Leeftijd moet een getal zijn tussen 1 en 120.","error")
            return render_template("form.html",
                                   naam=naam, email=email, leeftijd=leeftijd,
                                   speler=(speler == "Ja"),
                                    bestuur=(bestuur == "Ja"),
                                    sponsor=(sponsor == "Ja"),
                                    )

        if not any([speler=="Ja", bestuur=="Ja", sponsor=="Ja"]):
            flash("Selecteer minstens één rol.","error")
            return render_template("form.html",
                                   naam=naam, email=email, leeftijd=leeftijd,
                                   speler=(speler == "Ja"),
                                    bestuur=(bestuur == "Ja"),
                                    sponsor=(sponsor == "Ja"),
                                    )

        #save to excel file
        tijdstempel=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        #new_id=save_to_excel([naam,email,leeftijd_int,speler,bestuur,sponsor,tijdstempel])
        new_id=save_to_sheet([naam,email,leeftijd_int,speler,bestuur,sponsor,tijdstempel])

        flash(f"Bedankt {naam}! Je gegevens zijn opgeslagen.","success")
        return redirect(url_for("form"))

    return render_template("form.html",
                           naam=naam, email=email, leeftijd=leeftijd,
                           speler=(speler == "Ja"),
                            bestuur=(bestuur == "Ja"),
                            sponsor=(sponsor == "Ja"),
                            )
"""
@app.route("/view")
def view():
    wb= load_workbook(excel_file)
    ws= wb[sheet_name]
    rows= list(ws.iter_rows(values_only=True))
    if not rows:
        return render_template("view.html", headers=Headers, data=[])
    return render_template("view.html", headers=rows[0], data=rows[1:])
"""
if __name__=="__main__":
    #init_excel()
    app.run()
    