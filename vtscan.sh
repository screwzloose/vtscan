import os
import requests
import xlwt
import hashlib

# VirusTotal API key
API_KEY = "API_KEY"

# server directory to scan
dir_to_scan = "/your/directory"

# create an excel workbook to store the results
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Scan Results")

# write the headers to the first row of the sheet
sheet.write(0, 0, "File Name")
sheet.write(0, 1, "VirusTotal Result")

# create an excel workbook to store the results
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Scan Results")

# write the headers to the first row of the sheet
sheet.write(0, 0, "File Name")
sheet.write(0, 1, "VirusTotal Result")

# iterate through the files in the directory
i = 1
for dirpath, dirnames, filenames in os.walk(dir_to_scan):
    for file_name in filenames:
        # construct the file's path
        file_path = os.path.join(dirpath, file_name)
        print(f'Scanning {file_path}')

        # check if the file is a hidden file
        if file_name.startswith("."):
            continue

        # get the file's hash
        with open(file_path, "rb") as f:
            file_data = f.read()
        file_hash = hashlib.md5(file_data).hexdigest()

        # check the file on VirusTotal
        params = {"apikey": API_KEY, "resource": file_hash}
        response = requests.get("https://www.virustotal.com/vtapi/v2/file/report", params=params)

        # check if the API call was successful
        if response.status_code != 200:
            print("Error: {}".format(response.text))
            sheet.write(i + 1, 0, file_path)
            sheet.write(i + 1, 1, "Error")
            i += 1
            continue
        data = response.json()

        # check if the key "positives" is present in the data dictionary
        if "positives" in data:
            sheet.write(i + 1, 0, file_path)
            sheet.write(i + 1, 1, data["positives"])
        else:
            sheet.write(i + 1, 0, file_path)
            sheet.write(i + 1, 1, "Error")
        i += 1
    print(f'Scanned {i} files')

# save the workbook
workbook.save("scan_results.xls")
