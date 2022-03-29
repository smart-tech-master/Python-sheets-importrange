# Python-sheets-importrange

Prerequisites: Ensure you have Python3.8+ installed, along with pip as well.


Run this to install required modules, if you don't already have them.

pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

Instructions
1) Download the python script. 

2) Create a GCP Project. turn on Sheets API. Set up Desktop Credentials, download the json it generates. Keep the default name of 'credentials.json'.

3) Save the 'credentials.json' in the same directory as the script.

4) Change the file permissions to an executable using chmod if necessary. Run the script with 'python3 path-to-file-goes-here' or whichever method you prefer. 
If step 1 and 2 were done correctly you will be asked to complete a OAuth flow. Be sure the account you choose has permissions to access the sheets you plan to read from and write to. This will generate a 'token.pickle' file. Delete this if you want to sign-in again.

5) Open the python script and edit/add/remove Sheets and ranges in the 'Sheet Configurations' section. Ensure that your spelling is exact. Add corresponding calls to importRange in main.

6) run the script with 'python3 path-to-file-goes-here' or whichever method you prefer.
