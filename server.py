import os
import zipfile
from flask import Flask, request, redirect, send_file, flash, render_template
from werkzeug.utils import secure_filename
from multiprocessing import Process
import pandas as pd
import io

UPLOAD_FOLDER = "./static/data/input-files/"
OUTPUT_FOLDER = "./static/data/output-file/"
OUTPUT_FILE_PATH = os.path.join(OUTPUT_FOLDER, "output.xlsx")
ALLOWED_EXTENSIONS = set(['zip'])

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html', inputFiles=getInputFiles(), outputFile=getOutputFiles())

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def getInputFiles():
    inputFiles = [f for f in os.listdir(UPLOAD_FOLDER) if os.path.isfile(os.path.join(UPLOAD_FOLDER, f))]
    return inputFiles

def getOutputFiles():
    outputFile = [f for f in os.listdir(OUTPUT_FOLDER) if os.path.isfile(os.path.join(OUTPUT_FOLDER, f))]
    return outputFile

@app.route('/uploadZipfile', methods=['GET', 'POST'])
def upload_zipfile():
    if request.method == 'POST':
        if os.path.isfile(OUTPUT_FILE_PATH):
            os.remove(OUTPUT_FILE_PATH)

        if 'zipfile' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['zipfile']

        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(UPLOAD_FOLDER, filename))
            zip_ref = zipfile.ZipFile(os.path.join(UPLOAD_FOLDER, filename), 'r')
            zip_ref.extractall(UPLOAD_FOLDER)
            zip_ref.close()
            os.remove(os.path.join(UPLOAD_FOLDER, filename))
            inputFiles = getInputFiles()
        
    return render_template('index.html', inputFiles=inputFiles)

def create_outputFile(checkFiles):
    fixedColumns = ['Email', 'First Name', 'Last Name', 'Year', 'Number of Events']
    if os.path.isfile(OUTPUT_FILE_PATH):
        output_df = pd.read_excel(OUTPUT_FILE_PATH)
    else:
        output_df = pd.DataFrame(columns=fixedColumns)

    for fileIdx in range(len(checkFiles)):
        filename = checkFiles[fileIdx]
        input_file_path = os.path.join(UPLOAD_FOLDER, filename)
        df_input = pd.read_excel(input_file_path)

        for index, input_row in df_input.iterrows():
            output_row = {'Email': '', 'First Name': '', 'Last Name': '', 'Year': '', 'Number of Events': ''}

            if input_row['Email'] in output_df['Email'].values:
                old_output_row = output_df.loc[output_df['Email'] == input_row['Email']]
                output_df = output_df[output_df['Email'] != input_row['Email']]

                for col in old_output_row.keys():
                    if col not in input_row.keys() or input_row[col] == '':
                        if old_output_row[col].values[0] == '':
                            output_row[col] = ''
                        else:
                            output_row[col] = old_output_row[col].values[0]
                    else:
                        if old_output_row[col].values[0] == '':
                            output_row[col] = input_row[col]
                        else:
                            output_row[col] = old_output_row[col].values[0]

            else:     
                if input_row['Email'] != '':
                    for col in fixedColumns:
                        if col not in input_row.keys():
                            output_row[col] = ''
                        else: 
                            output_row[col] = input_row[col]

            output_row[filename] = 1   
            output_df = pd.concat([output_df, pd.DataFrame([output_row])], ignore_index = True)

        os.remove(input_file_path)
    
    output_df.fillna(0, inplace=True)
    eventColumns = [c for c in output_df.columns if c not in fixedColumns]
    output_df['Number of Events'] = output_df[eventColumns].sum(axis=1)
    output_df = output_df.sort_values(by=['Number of Events'], ascending=False)

    output_df.to_excel(OUTPUT_FILE_PATH, sheet_name="All Attendance", index=False) 
    return

def filter_outputFile():
    output_df = pd.read_excel(OUTPUT_FILE_PATH)

    output_df.fillna(0, inplace=True) # fill empty cells with 0s
    output_df = output_df[output_df['Email'] != 0] # remove cells with 0 value as 'Email'

    # remove empty spaces and turn all strings to lowercase
    for col in output_df.columns: 
        output_df[col] = output_df[col].apply(lambda s: s.lower() or s.replace(" ", "") if type(s) == str else s)

    fixedColumns = ['Email', 'First Name', 'Last Name', 'Year']

    # get a list of emails that are duplicated
    dupEmailList = output_df['Email'].loc[output_df['Email'].duplicated()].tolist()

    # combine values of events where emails are duplicates
    for i in range(len(dupEmailList)):
        dupEmailRows = output_df[output_df['Email'] == dupEmailList[i]]
        output_df = output_df[output_df['Email'] != dupEmailList[i]]

        new_email_row = {}
        for col in dupEmailRows.head(1).columns:
            if col in fixedColumns:
                new_email_row[col] = dupEmailRows.head(1)[col].values[0]
            else:
                new_email_row[col] = dupEmailRows[col].sum()            
        output_df = pd.concat([output_df, pd.DataFrame([new_email_row])], ignore_index = True)

    # get a list of names that are duplicated
    dupNameList = output_df.loc[output_df['Last Name'].duplicated(keep=False)] 
    dupFirstNameList = dupNameList['First Name'].loc[dupNameList['First Name'].duplicated()].tolist() 
    dupLastNameList = dupNameList['Last Name'].loc[dupNameList['First Name'].duplicated()].tolist() 

    # combine values of events where emails are duplicates
    for i in range(len(dupFirstNameList)):
        dupNameRows = output_df[(output_df['First Name'] == dupFirstNameList[i]) & (output_df['Last Name'] == dupLastNameList[i])]
        output_df = output_df[~((output_df['First Name'] == dupFirstNameList[i]) & (output_df['Last Name'] == dupLastNameList[i]))]
        sbuemail = dupNameRows.loc[dupNameRows['Email'].str.contains('@stonybrook.edu')]

        new_name_row = {}
        for col in sbuemail.columns:
            if col in fixedColumns:
                new_name_row[col] = sbuemail[col].values[0]
            else:
                new_name_row[col] = dupNameRows[col].sum()               
        output_df = pd.concat([output_df, pd.DataFrame([new_name_row])], ignore_index = True)

    output_df = output_df.sort_values(by=['Number of Events'], ascending=False)   
    output_df.to_excel(OUTPUT_FILE_PATH, sheet_name="All Attendance", index=False) 
    return

@app.route('/checkedFiles', methods=['GET', 'POST'])
def submit_checkedFiles():
    if request.method == 'POST':
        checkFiles = request.form.getlist('checkedFiles')
        inputFiles = getInputFiles()
        inputFiles = [f for f in inputFiles if f not in checkFiles]
        create_outputFile(checkFiles)
        filter_outputFile()
        outputFile = getOutputFiles()
    return render_template('index.html', inputFiles=inputFiles, outputFile=outputFile)

if __name__ == '__main__':
   app.run(debug = True)