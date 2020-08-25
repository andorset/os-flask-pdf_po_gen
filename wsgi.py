from __future__ import print_function

import smtplib
import socket
import smartsheet
import logging
import os.path
import requests
import pypdftk

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from flask import Flask, render_template, request

TASKORDERS = [
    "001",
    "002",
]

SMARTSHEET_ACCESS_TOKEN = "1wu79j0a8jy95n5cgw9xmhmit7"
SMARTSHEET_SHEET_ID = {
    '001': 4952568381106052,
    '002': 2490346643974020
}

_dir = os.path.dirname(os.path.abspath(__file__))
PO_TEMPLATE = {
    '001': 'po_template_001.pdf',
    '002': 'po_template_002.pdf'
}

RNSD = {
    "001": "19997650",
    "002": "20544388"
}

SMARTACCOUNT = {
    "001": "ea-001.cisco.com",
    "002": "ea-002.cisco.com"
}

# Account Manager to Email Table
ACCTMGRTOEMAIL = {
    'Maury Walker': "mauwalke@cisco.com",
    'Mark Yarish': "mauwalke@cisco.com",
    'Jamie Mazziotta': "mauwalke@cisco.com",
    'Mark Looser': "cjamar@cisco.com",
    'Steve Mullet': "cjamar@cisco.com",
    'John Nester': "cjamar@cisco.com"
}

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
column_map = {}

application = Flask(__name__, template_folder='.')


@application.route("/")
def root():
    return render_template('root.html', len=len(TASKORDERS), taskorders=TASKORDERS)


@application.route("/execute", methods=['POST', 'GET'])
def execute():
    output = ""

    if request.method == 'POST':
        taskOrder = request.form.get('TaskOrder')

        if taskOrder == '001':
        # Execute for Task Order 001
            output += "<P ALIGN=\"CENTER\"><H1>Task Order " + taskOrder + "</H1></P>"
            output += executePOCreationbyTaskOrder("001")
        elif taskOrder == '002':
        # Execute for Task Order 002
            output += "<P ALIGN=\"CENTER\"><H1>Task Order " + taskOrder + "</H1></P>"
            output += executePOCreationbyTaskOrder("002")
        elif taskOrder == 'ALL':
            output = "<P ALIGN=\"CENTER\"><H1>Task Order " + "001" + "</H1></P>"
            output += executePOCreationbyTaskOrder("001")
            output += "<P ALIGN=\"CENTER\"><H1>Task Order " + "002" + "</H1></P>"
            output += executePOCreationbyTaskOrder("002")
        else:
            return "ERROR: Invalid Task Order Provided"

    return output
#        return render_template("execute.html", result=result)


@application.route("/result", methods=['POST', 'GET'])
def result():
    if request.method == 'POST':
        result = request.form
        return render_template("result.html", result=result)


# Send Email + Attachments
def send_email(data_dict):
    subject = data_dict["PONumber"]
    body = "PO#: " + data_dict["PONumber"] + "\r\r\n"
    body += "Deal ID (RNSD): " + data_dict["DealID"] + "\r\r\n"
    body += "\r\r\n"
    body += "eDelivery Destination: " + data_dict["eDeliveryDestination"] + " , " + ACCTMGRTOEMAIL[data_dict["AccountManager"]] + "\r\r\n"
    body += "Smart Account: " + data_dict["SmartAccount"] + "\r\r\n"
    body += "Virtual Account: " + data_dict["VirtualAccount"] + "\r\r\n"

    sender_email = "andorset@cisco.com"
    receiver_email = ACCTMGRTOEMAIL[data_dict["AccountManager"]]

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    #message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    #
    # Attach PO to Email
    #
    POFileName = data_dict["PONumber"] + " - PO.pdf"

    # Open PDF file in binary mode
    with open(POFileName, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        "attachment; filename="+POFileName,
    )
    # Add attachment to message
    message.attach(part)

    #
    # Attach Estimate to Email
    #
    EstimateFileName = data_dict["PONumber"] + " - Estimate.xls"

    # Open PDF file in binary mode
    with open(EstimateFileName, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        "attachment; filename="+EstimateFileName,
    )

    # Add attachment to message
    message.attach(part)

    # Convert message to string
    text = message.as_string()

    # Log in to server using secure context and send email
    try:
        server = smtplib.SMTP("rcdn-mx-01.cisco.com")
        #server.set_debuglevel(1)
        server.ehlo()
        logging.info("Sending to: %s", receiver_email)
        print ("Sending to: %s" % receiver_email)
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
    except socket.error as e:
        logging.error("Could Not Connect to Mail Server: %s", data_dict["PONumber"])
        print("Could Not Connect to Mail Server %s" % data_dict["PONumber"])
        return False
    except smtplib.SMTPServerDisconnected as e:
        logging.error("Mail Server Disconnected: %s", data_dict["PONumber"])
        print("ERROR: Mail Server Disconnected: %s" % data_dict["PONumber"])
        return False

    return True


# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)


# TODO: Replace the body of this function with your code
# This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
#
# Return a new Row with updated cell values, else None to leave unchanged
def evaluate_row_and_build_updates(taskOrder, source_row, smart_session, sheet_id, row_id):
    output = ""

    # Set the PO Template
    po_template = PO_TEMPLATE[taskOrder]

    # Find the cell and value we want to evaluate
    status_cell = get_cell_by_column_name(source_row, "Package Created")
    status_value = status_cell.value

    if ((status_value is None) or (status_value is False)):
        # Retrieve Fields for PO Form
        po_cell = get_cell_by_column_name(source_row, "PO Number")
        po_value = po_cell.display_value

        quoteid_cell = get_cell_by_column_name(source_row, "Quote ID")
        quoteid_value = quoteid_cell.display_value

        email_cell = get_cell_by_column_name(source_row, "Contact Email")
        email_value = email_cell.display_value

        virtualacct_cell = get_cell_by_column_name(source_row, "Virtual Account")
        virtualacct_value = virtualacct_cell.display_value

        acctmanager_cell = get_cell_by_column_name(source_row, "Account Manager")
        acctmanager_value = acctmanager_cell.display_value

        # Skip if we are missing information
        if ((po_value is None) or (quoteid_value is None) or (email_value is None) or (virtualacct_value is None) or (acctmanager_value is None)):
            logging.warning("WARNING: Incomplete Information - Skipping PO: %s", po_value)
            output += ("WARNING: Incomplete Information - Skipping PO: %s" % po_value)
            output += "<BR>"
        else:
            found_estimate = False
            response = smart_session.Attachments.list_row_attachments(sheet_id, row_id, include_all=True)
            attachments = response.data

            EstimateFileName = po_value + " - Estimate.xls"

            for attachment in attachments:
                if ((attachment.mime_type == "application/vnd.ms-excel") or (attachment.mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")):
                    attachmentobj = smart_session.Attachments.get_attachment(sheet_id, attachment.id)
                    response = requests.get(attachmentobj.url)
                    open(EstimateFileName, 'wb').write(response.content)
                    if (response.status_code == 200):
                        found_estimate = True

            if (found_estimate == False):
                logging.warning("WARNING: Incomplete Information - Skipping PO: %s", po_value)
                output += ("WARNING: Incomplete Information - Skipping PO: %s" % po_value)
                output += "<BR>"
                return None, None, output

            logging.info("PROCESSING: PO: %s", po_value)
            output += ("PROCESSING: PO: %s" % po_value)
            output += "<BR>"

            data_dict = {
                'DealID': RNSD[taskOrder],
                'EstimateID': quoteid_value,
                'PONumber': po_value,
                'eDeliveryDestination': email_value,
                'SmartAccount': SMARTACCOUNT[taskOrder],
                'VirtualAccount': virtualacct_value,
                'SpecialInstructions': '',
                'AccountManager': acctmanager_value
            }

            # Filename is in the format of "CISCO-EA-001-000000001 - PO.docx"
            filename = (po_value + "\ -\ PO.pdf")

            pypdftk.fill_form(po_template, datas=data_dict, out_file=filename, flatten=True)

            if (send_email(data_dict) is True):
                # True when sent successfully
                new_cell = smart_session.models.Cell()
                new_cell.column_id = column_map["Package Created"]
                new_cell.value = True

                new_row = smart_session.models.Row()
                new_row.id = source_row.id
                new_row.cells.append(new_cell)

                return new_row, po_value, output

    return None, None, output


def executePOCreationbyTaskOrder(taskOrder):
    output = ""

    # Set the Sheet ID based on Task Order
    sheet_id = SMARTSHEET_SHEET_ID[taskOrder]

    # Initialize client
    smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)

    # Load entire sheet
    sheet = smart.Sheets.get_sheet(sheet_id)

    logging.info("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)
    output += ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)
    output += "<BR>"

    # Build column map for later reference - translates column names to column id
    for column in sheet.columns:
        column_map[column.title] = column.id

    # Accumulate rows needing update here
    rowsToUpdate = []

    for row in sheet.rows:
        outputbuffer = ""
        (rowToUpdate, poNumber, outputbuffer) = evaluate_row_and_build_updates(taskOrder, row, smart, sheet_id, row.id)
        output += outputbuffer
        if rowToUpdate is not None:
            rowsToUpdate.append(rowToUpdate)

    # Finally, write updated cells back to Smartsheet
    if rowsToUpdate:
        logging.info("Writing " + str(len(rowsToUpdate)) + " rows back to sheet id " + str(sheet.id))
        output += ("Writing " + str(len(rowsToUpdate)) + " rows back to sheet id " + str(sheet.id))
        output += "<BR>"
        result = smart.Sheets.update_rows(sheet_id, rowsToUpdate)
    else:
        logging.info("No updates required")
        output += "No updates required"
        output += "<BR>"

    return output


if __name__ == "__main__":
    # Log all calls
    logging.basicConfig(filename='pdf_po_gen.log', level=logging.INFO,
                        format='%(asctime)s %(levelname)-8s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    application.run(host="0.0.0.0", port="80")
