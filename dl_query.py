import sys  
import json
import csv
import logging
import param # From directory

import requests
import msal
from requests.structures import CaseInsensitiveDict

# Enter parameters here ---------------------------------------------------------------------------------------------
par_authority = 'https://login.microsoftonline.com/<your-tenant>/'
par_client_cred = '<secret>'
par_client_id = '<client-id-or-app-id>'
par_odata = '<odata-link>'

par_outfile_prefix = 'query_data_from_py_' # Prefix for output file
par_outfile_ext = '.csv' # Extension for output file

# Main script starts -------------------------------------------------------------------------------------------------

app = msal.ConfidentialClientApplication(
    client_id = par_client_id,
    authority = par_authority,
    client_credential = par_client_cred,
    )

result = app.acquire_token_for_client("https://workplaceanalytics.office.com/.default")

if "access_token" in result:
    headers = CaseInsensitiveDict()
    headers["Accept"] = "application/json"
    headers["Authorization"] = "Bearer "+result['access_token']    
    print(result['access_token'])
    
    file_counter = 0
    while par_odata != None: 
        r = requests.get(  
            url = par_odata,
            headers = headers
        )
    
        print("Status code: ", r.status_code)

        ## TOGGLE FOR DEBUGGING ONLY
        # file = open("resp_text.txt", "w", encoding = "utf-8")
        # file.write(r.text)

        data = json.loads(r.text)
        person_data = data['value']
        
        # Temporary file name based on prefix and extension parameters
        tmp_filename = par_outfile_prefix + str(file_counter).zfill(4) + par_outfile_ext
        data_file = open(tmp_filename, 'w', newline = '')
        
        data_file = open(par_outfile, 'w', newline = '')
        csv_writer = csv.writer(data_file)
        count = 0
        
        for p in person_data:
            if count == 0:
                # Writing headers of CSV file
                header = p.keys()
                csv_writer.writerow(header)
                count += 1
        
            # Writing data of CSV file
            csv_writer.writerow(p.values())

        data_file.close()

        print("Query data has been saved to file.")
        
        if data['@odata.nextLink']:
            print("More data exists")
            par_odata=data['@odata.nextLink']
            file_counter += 1
        else:
            par_odata = None

else:
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id")) 

        
