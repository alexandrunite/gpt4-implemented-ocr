This script is designed to automate the process of extracting information from invoices, analyzing images of invoices using GPT-4 Vision API, and storing the extracted data in an Excel spreadsheet.

Prerequisites
Python installed on your system
Required libraries: openai, openpyxl, smtplib, pdf2image, string, requests, datetime, base64, email.mime, json, os, random, time
Configuration
Before running the script, make sure to update the following variables in the code:

MY_ADDRESS: Your email address for sending emails
PASSWORD: Your email password
email: Recipient email address
excel_path: Path to the Excel file where the extracted data will be stored
folder_path: Path to the folder containing PDF files to convert
api_key: Your OpenAI API key
image_path: Path to the image file to be analyzed by GPT-4 Vision API
image_path_2: Path to the second image file (unused in the current script)
Functions
scan_for_files(): Scans the specified folder for new files
is_pdf_file(file_path): Checks if a file is a PDF
conversie_pdf(): Converts PDF files to images
encode_image(image_path): Encodes an image to base64 format
datetime_generation(): Generates current datetime in a specific format
delete_file(file_path): Deletes a specified file
message_generation(): Generates email message and attachment
send_mail(): Sends an email with the generated message and attachment
data_scraper(): Extracts data from the image using GPT-4 Vision API
deschidere_adaugare(): Opens Excel file, extracts data, and adds it to the spreadsheet
Running the Script
To run the script:

Ensure all prerequisites are met and configurations are updated.
Execute the script, which will extract data from the provided image, analyze it, and add the extracted data to the Excel spreadsheet.
Feel free to reach out if you have any questions or need further assistance!