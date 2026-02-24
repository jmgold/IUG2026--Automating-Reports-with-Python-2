#!/usr/bin/env python3

"""Create and email a list of new items

Author: Jeremy Goldstein 
Modification to script originally written by Gem Stone-Logan
Contact Info: jgoldstein@minlib.net
"""

import psycopg2
import csv
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
    # Gather column headers, which are not included in the cursor.fetchall() 
    columns = [i[0] for i in cursor.description]
    conn.close()
    return rows, columns


#takes results of a sql query and a list of the results column headers and write them to a .csv file
def write_csv(query_results, headers):
    
    csvfile = "WeeklyNewItem.csv"
    
    #open file in write mode and write all rows from query results to the file
    #set encoding to utf-8 to match Sierra characters and newline to '' to avoid excess line breaks
    with open(csvfile,'w', encoding='utf-8', newline='') as tempFile:
        myFile = csv.writer(tempFile, delimiter=',')
        #write header row then write full query results
        myFile.writerow(headers)
        myFile.writerows(query_results)
    tempFile.close()
    
    return csvfile

#Send an email with an attachment passed to the function
def send_email(attachment, subject, message, recipient):
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

    # Enter your own email information
    emailfrom = "jgoldstein@minlib.net"

    # Creating the email message
    msg = MIMEMultipart()
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(message))
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
    smtp.sendmail(emailfrom, recipient, msg.as_string())
    smtp.quit()


def main():
    query = """
      /*
      Jeremy Goldstein
      Minuteman Library Network
      Retrives top 50 titles based on recently placed holds
      */

      WITH holds_count AS (
        SELECT
          t.bib_record_id,
          COUNT(t.bib_record_id) AS holds_on_title

        FROM (
          SELECT
            CASE
	          WHEN r.record_type_code = 'i' THEN (
		        SELECT
		          l.bib_record_id
		        FROM sierra_view.bib_record_item_record_link as l
		        WHERE l.item_record_id = h.record_id
		        LIMIT 1)
    
              WHEN r.record_type_code = 'b' THEN h.record_id
              ELSE NULL
            END AS bib_record_id

          FROM sierra_view.hold h
          JOIN sierra_view.record_metadata as r
            ON r.id = h.record_id

          WHERE h.placed_gmt::DATE > (CURRENT_DATE - INTERVAL '7 days') 
        ) t

        GROUP BY t.bib_record_id
        HAVING COUNT(t.bib_record_id) > 1
        ORDER BY holds_on_title
        )

      SELECT
        ROW_NUMBER() OVER (ORDER BY hc.holds_on_title DESC) AS rank,
        rm.record_type_code||rm.record_num||'a' AS bib_number,
        best_title as title,
        b.best_author AS author,
        hc.holds_on_title

      FROM holds_count AS hc
      JOIN sierra_view.bib_record_property b
        ON hc.bib_record_id = b.bib_record_id
      JOIN sierra_view.record_metadata rm
        ON hc.bib_record_id = rm.id

      GROUP BY 2,3,4,5
      ORDER BY hc.holds_on_title DESC
      LIMIT 50
    """
    
    email_subject = "Weekly trending titles"
    email_message = """***This is an automated email***
        
    The weekly trending titles report has been attached. 
    Please take a look and let the Technology Librarian know if there are any questions about it.
    """
    emailto = ["jgoldstein@minlib.net"]
    
    query_results, headers = run_query(query)
    local_file = write_csv(query_results, headers)
    send_email(local_file, email_subject, email_message, emailto)
    
    #delete local_file
    os.remove(local_file)


main()
