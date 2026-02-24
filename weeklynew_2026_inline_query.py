#!/usr/bin/env python3

"""Create and email a list of new items

Author: Jeremy Goldstein 
Modification to script originally written by Gem Stone-Logan
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import xlsxwriter
import configparser
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# execute a Sierra SQL query and return the results
def run_query(query):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("config.ini")

    # Connecting to Sierra PostgreSQL database
    try:
        conn = psycopg2.connect(config["db"]["connection_string"])
    except psycopg2.Error as e:
        print("Unable to connect to database: " + str(e))

    # Opening a session and querying the database
    cursor = conn.cursor()
    cursor.execute(query)
    # For now, just storing the data in a variable. We'll use it later.
    rows = cursor.fetchall()
    conn.close()
    return rows

#take a set of results from a sql query and write them to an Excel file, returning the file
def write_excel(query_results):
    # Name of Excel File
    excel_file = "WeeklyNewItem.xlsx"
    
    # Creating the Excel file for staff
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    # Formatting our Excel worksheet
    worksheet.set_landscape()
    worksheet.hide_gridlines(0)

    # Formatting Cells
    eformat = workbook.add_format({"text_wrap": True, "valign": "top"})
    eformatlabel = workbook.add_format(
        {"text_wrap": True, "valign": "top", "bold": True}
    )

    # Setting the column widths
    worksheet.set_column("A:A", 10.29)
    worksheet.set_column("B:B", 6.29)
    worksheet.set_column("C:C", 12.71)
    worksheet.set_column("D:D", 16.57)
    worksheet.set_column("E:E", 24.71)
    worksheet.set_column("F:F", 11.14)
    worksheet.set_column("G:G", 18.47)
    worksheet.set_column("H:J", 4.5)

    # Inserting a header
    worksheet.set_header("&CWeekly New List")

    # Adding column labels
    worksheet.write("A1", "Bib Record#", eformatlabel)
    worksheet.write("B1", "Location", eformatlabel)
    worksheet.write("C1", "Call#", eformatlabel)
    worksheet.write("D1", "Author", eformatlabel)
    worksheet.write("E1", "Title", eformatlabel)
    worksheet.write("F1", "Barcode", eformatlabel)
    worksheet.write("G1", "Series", eformatlabel)
    worksheet.write("H1", "Item Count", eformatlabel)
    worksheet.write("I1", "Order Count", eformatlabel)
    worksheet.write("J1", "Hold Count", eformatlabel)

    # Writing the report for staff to the Excel worksheet
    for rownum, row in enumerate(query_results):
        worksheet.write(rownum + 1, 0, row[0], eformat)
        worksheet.write(rownum + 1, 1, row[1], eformat)
        worksheet.write(rownum + 1, 2, row[2], eformat)
        worksheet.write(rownum + 1, 3, row[3], eformat)
        worksheet.write(rownum + 1, 4, row[4], eformat)
        worksheet.write(rownum + 1, 5, row[5], eformat)
        worksheet.write(rownum + 1, 6, row[6], eformat)
        worksheet.write(rownum + 1, 7, row[7], eformat)
        worksheet.write(rownum + 1, 8, row[8], eformat)
        worksheet.write(rownum + 1, 9, row[9], eformat)

    workbook.close()
    return excel_file

#Send an email with an attachment passed to the function
def send_email(attachment):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("config.ini")

    # These are variables for the email that will be sent.
    # Make sure to use your own library's email server (emailhost)
    emailhost = config["email"]["host"]
    # user and pw was not in the original script, necessary for Minuteman's Gmail accounts
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = "25"
    emailsubject = "Weekly New Report"
    emailmessage = """***This is an automated email***
    
    
    The weekly new report has been attached. Please take a look and let the Technology Librarian know if there are any questions about it."""

    # Enter your own email information
    emailfrom = "jgoldstein@minlib.net"
    emailto = ["jgoldstein@minlib.net"]

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(emailto) is list:
        msg["To"] = ", ".join(emailto)
    else:
        msg["To"] = emailto
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = emailsubject
    msg.attach(MIMEText(emailmessage))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment; filename=%s" % attachment)
    msg.attach(part)

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    # for Gmail connection used within Minuteman
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, emailto, msg.as_string())
    smtp.quit()

#main function take a parameter for item location, which will be incorporated into sql query
def main(location):
    query = """
      /* Weekly New Report
      The report will retrieve all items, not in the specifically excluded statuses
      ('m','n','z','t','s','$','d','8','w','y'), that were created in the last 10 days.
      It will also include a count of items per bib, count of orders with a status of "o",
      and a count of bib-level holds.
       */

      SELECT 
        distinct 'b'|| rmb.record_num || 'a' AS "Bib Record Num",
        i.location_code, 
        CASE
          WHEN pei.index_entry IS NULL THEN UPPER(peb.index_entry) 
          ELSE UPPER(pei.index_entry)
        END AS "Call#",
        brp.best_author AS "Author",
        brp.best_title AS "Title",
        string_agg(distinct i.barcode, ' ') AS "Barcode", 
        string_agg(distinct pes.index_entry, ' | ') AS "Series Info",
        count(distinct ic.id) AS "Item Count",  
        count(distinct o.id) AS "Order Count",
        count(distinct h.id) AS "Hold Count"
      FROM sierra_view.item_view i
      JOIN sierra_view.bib_record_item_record_link bri
        ON i.id = bri.item_record_id
      JOIN sierra_view.record_metadata rmb
        ON bri.bib_record_id = rmb.id
        AND rmb.record_type_code='b'
      JOIN sierra_view.phrase_entry peb
        ON i.id = peb.record_id
        AND peb.index_tag='c'
      JOIN sierra_view.bib_record_property brp
        ON brp.bib_record_id = bri.bib_record_id
      LEFT JOIN sierra_view.phrase_entry pei
        ON pei.record_id=bri.item_record_id
        AND pei.index_tag='c'
      LEFT JOIN sierra_view.phrase_entry pes
        ON pes.record_id=bri.bib_record_id
        AND pes.index_tag='t'
        AND pes.varfield_type_code='s'
      LEFT JOIN sierra_view.item_record ic
        ON ic.id=i.id
        AND ic.item_status_code not in ('m','n','z','t','s','$','d','8','w','y')
      LEFT JOIN sierra_view.hold h
        ON (bri.bib_record_id = h.record_id OR i.id = h.record_id)
      LEFT JOIN sierra_view.bib_record_order_record_link bro
        ON bri.bib_record_id = bro.bib_record_id
      LEFT JOIN sierra_view.order_record o
        ON bro.order_record_id = o.id
        AND o.order_status_code ='o'

      WHERE i.record_creation_date_gmt::date>=NOW()::DATE-EXTRACT(DOW FROM NOW())::INTEGER-10 
        AND i.location_code = '""" + location + """'  
      GROUP BY rmb.record_num, i.location_code, "Call#", brp.best_author, brp.best_title
      ORDER BY i.location_code, "Call#", brp.best_author, brp.best_title
    """
    
    query_results = run_query(query)
    local_file = write_excel(query_results)
    send_email(local_file)
    
    #delete local_file
    os.remove(local_file)

#call main function, passing different item locations that will feed into sql query
main('adfic')
main('jfic')
