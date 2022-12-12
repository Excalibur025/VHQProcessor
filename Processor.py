# importing required modules 
import PyPDF2
import csv
import requests

# Step 1: Locate the automated Email from VHQ and "Print as PDF".
# Step 2: Move that file to the same directory as this script.
# Step 3: rename the vairable "VHQData" (below) with the name of the email you've exported as PDF.
# Step 4: Run the script.

# How to automate this script: We need to configure the SMTP email client (Outlook) to automatically export and save the VHQ email and save it to a directory.
# To my knowledge, only Outlook on Windows has this feature.

VHQData = 'VHQ_Email_Example.pdf' # Hypothetically by changing the path of this file, you can get it to run another place other than local.

#locate and open the PDF. Currently it searches in the same directory the py script is located.
pdfFileObj = open(VHQData, 'rb')
#convert to object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pageObj = pdfReader.getPage(0)
#It's only one page so we're good here (array starts at zero!)
text=pageObj.extractText()

#Extract ONLY the URL by splitting the strings on either side.
vhq_url=text.split("DownloadURL:") [1].split("Log in to") [0]
print("Sending request to URL:")
print(vhq_url) #Letmee make sure we got the right URL.

#because we can't curl with python, use requests. AND it takes a second.
print("Processing, please wait...")
with requests.Session() as s:
    download = s.get(vhq_url)

    decoded_content = download.content.decode('utf-8') #use standard unicode. Last output was all gibberish.

#Idk found a solution in StackOverflow.
    cr = csv.reader(decoded_content.splitlines(), delimiter=',')
    my_list = list(cr)
    for row in my_list:
        print(row) #make sure to print (row) NOT (my_list). It crashed the VSCode console lmao.
    #Additionally, you can now use the vairable "row" to export / send it whever it needs to be.

#Now that we've got the data, send it to a CSV for New Relic intake.
