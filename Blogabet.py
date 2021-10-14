# packages to install: requests, lxml, imaplib2 and psutil

# Enable IMAP by logging in to the Gmail account and going to Settings -> Forwarding and POP/IMAP and enabling IMAP access

# Make sure to allow 'less secure apps' to be able to access the gmail account
# Turn on the option here: https://myaccount.google.com/security?pli=1#connectedapps
# And also here: https://myaccount.google.com/lesssecureapps

# Make sure to add a filter to the email account that sends the emails about picks to a specific label or folder. 
# The default is 'Blogabet', but this can be changed by editing the MAIL_LABEL variable.

import smtplib
import email
import os
import re
import time
import getpass
import csv
import json
import traceback
from datetime import datetime, timedelta
from threading import *
from stat import S_IREAD, S_IRGRP, S_IROTH, S_IWUSR
import requests # used to retrieve html page
from lxml import html # used to parse html
from imaplib2 import imaplib2 # used to connect using IMAP and use IDLE
from psutil import process_iter # used to close connection left open by previous iterations of this script
from signal import SIGTERM

# Settings for gmail accounts
MAIL_ADDRESS  = "norbiyoda@gmail.com"
MAIL_PASSWORD = "Oliver2019!"
SMTP_SERVER   = "imap.gmail.com" # The url of the imap server
MAIL_LABEL    = "Blogabet" # The label or folder of which the mails should be processed. Note that spaces in the label name might not be supported

# Settings for off-ice accounts
#MAIL_ADDRESS  = "john.mueller@off-ice.com"
#MAIL_PASSWORD = "gY8@19bh"
#SMTP_SERVER   = "mail.off-ice.com" # The url of the imap server
#MAIL_LABEL    = "INBOX.Blogabet" # The label or folder of which the mails should be processed. Note that spaces in the label name might not be supported

SMTP_PORT     = 993 # The port of the imap server
MAIL_CHECK_INTERVAL = 60 # The amount of time in seconds before the bot checks if new mails have arrived
CSV_FILE_NAME = "Blogabet.csv"
CSV_COMBO_FILE_NAME = "Blogabet_Combo.csv"
CSV_DELIMITER = ","
WEBHOOK_URL   = "https://hook.integromat.com/kixji72iskavnlgxmr5v1p85wqjqepex" # Address to send the scraped JSON object
WEBHOOK_LOG_URL = "https://hook.integromat.com/3hkisren2ezuxuj3mrd2ykd3ph3v6a29" # Address to send log messages
WEBHOOK_COMBO_URL = "https://hook.integromat.com/u5fwuezmv17r19r4t9tyife4hmn2j79e" # Address to send scraped combo tips
CAPTCHA_API_KEY = "61a16aec9f4224cffea018ff6ca7153b" # The 2captcha API key
SLEEP_TIME = 3 # Time in seconds between processing emails

def log(title, exc):
    e = Exception("test")
    print(title + ": " + str(exc))
    print(traceback.format_exc())

    try:
        response = requests.post(
            WEBHOOK_LOG_URL, data=json.dumps({"title": title, "message": str(exc)}),
            headers={'Content-Type': 'application/json'}
        )
        if response.status_code != 200:
            print(
                'Request to logging webhook returned an error %s, the response is:\n%s'
                % (response.status_code, response.text)
            )
    except Exception as e:
        print("Could not write to logging webhook: " + str(e))

# Credits: https://blog.timstoop.nl/posts/2009/03/11/python-imap-idle-with-imaplib2.html
class Idler(object):
    def __init__(self, conn):
        self.thread = Thread(target=self.idle)
        self.M = conn
        self.event = Event()
        self.needsReset = Event()
        self.needsResetExc = None
        self.idling = Event()
        self.timeSinceMailCheck = datetime.now()

    def start(self):
        self.thread.start()

    def stop(self):
        # This is a neat trick to make thread end. Took me a 
        # while to figure that one out!
        self.event.set()

    def join(self):
        self.thread.join()

    def idle(self):
        self.dosync()
        # Starting an unending loop here
        while True:
            # This is part of the trick to make the loop stop 
            # when the stop() command is given
            if self.event.isSet():
                return
            self.needsync = False
            # A callback method that gets called when a new 
            # email arrives. Very basic, but that's good.
            def callback(args):
                if not self.event.isSet():
                    self.needsync = True
                    self.event.set()
            # Do the actual idle call. This returns immediately, 
            # since it's asynchronous.
            if not self.needsReset.isSet():
                self.idling.set()
                try:
                    typ = self.M.idle(10*60, callback=callback)
                except (imaplib2.IMAP4.abort, imaplib2.IMAP4.error) as e:
                    self.needsReset.set()
                    self.needsResetExc = e
            # This waits until the event is set. The event is 
            # set by the callback, when the server 'answers' 
            # the idle call and the callback function gets 
            # called.
            self.event.wait()
            self.idling.clear()
            # Because the function sets the needsync variable,
            # this helps escape the loop without doing 
            # anything if the stop() is called. Kinda neat 
            # solution.
            if self.needsync:
                try:
                    self.dosync()
                except (imaplib2.IMAP4.abort, imaplib2.IMAP4.error) as e:
                    self.needsReset.set()
                    self.needsResetExc = e
                self.event.clear()

    # The method that gets called when a new email arrives. 
    # Replace it with something better.
    def dosync(self):
        self.timeSinceMailCheck = datetime.now()
        ids = get_mail_ids()
        while len(ids) > 0:
            read_email_from_server(ids)
            time.sleep(5)
            self.timeSinceMailCheck = datetime.now()
            ids = get_mail_ids()

jsonfields = {}

def getStrippedText(text):
    if text:
        strippedtext = re.sub(r"\s{2,}", "", str.strip(text))
        if strippedtext:
            return strippedtext
    return ''

def parseChildren(node):
    global jsonfields
    if node.tag == 'h3':
        jsonfields["Teams"] = getStrippedText(node.text)
    if node.tag == 'div' and 'class' in node.attrib and node.attrib['class'] == 'pick-line':
        jsonfields["Pick"] = getStrippedText(re.sub(r"\s{2}", " ", re.sub("@", "", node.text)))
        for child in node:
            if child.tag == 'span' and 'class' in child.attrib and child.attrib['class'] == 'feed-odd':
                jsonfields["Odds"] = getStrippedText(child.text)
            if child.tag == 'small' and jsonfields["Pick"] == '':
                jsonfields["Pick"] = getStrippedText(re.sub(r"\s{2}", " ", re.sub("@", "", child.tail)))
    if node.tag == 'div' and 'class' in node.attrib and node.attrib['class'] == 'labels':
        for child in node:
            if child.tag == 'span' and 'class' in child.attrib and child.attrib['class'] == 'label label-default':
                jsonfields["Stake"] = getStrippedText(child.text)
            if child.tag == 'a' and 'href' in child.attrib and child.attrib['href'].find('bookie') >= 0:
                bookie = getStrippedText(child.text)
                if 'Bookie' in jsonfields:
                    jsonfields["Bookie"] += ';' + bookie
                else:
                    jsonfields["Bookie"] = bookie
    if node.tag == 'div' and 'class' in node.attrib and  node.attrib['class'] == 'sport-line':
        for child in node:
            if child.tag == 'small':
                counter = 0
                for subchild in child:
                    if subchild.tag == 'span' and counter == 0:
                        jsonfields["Sport"] = getStrippedText(re.sub(r'/', "-", subchild.text)) + " " + getStrippedText(re.sub(r'/',  "", subchild.tail))
                        jsonfields["Group"] = "Sprt" + getStrippedText(re.sub(r'\s*/\s*', "", subchild.text))
                    if subchild.tag == 'span' and counter != 0:
                        jsonfields["Date"] = getStrippedText(subchild.tail)
                    counter += 1

mySession = requests.Session()

# See https://github.com/2captcha/2captcha-api-examples/blob/master/ReCaptcha%20v2%20API%20Examples/Python%20Example/2captcha_python_api_example.py for adding Proxy support
def solveCaptcha(url, site_key):
    s = mySession

    # here we post site key to 2captcha to get captcha ID (and we parse it here too)
    print("solving ref captcha...")
    captcha_id = s.post("http://2captcha.com/in.php?key={}&method=userrecaptcha&googlekey={}&pageurl={}".format(CAPTCHA_API_KEY, site_key, url)).text.split('|')[1]
    time.sleep(30)
    # then we parse gresponse from 2captcha response
    recaptcha_answer = s.get("http://2captcha.com/res.php?key={}&action=get&id={}".format(CAPTCHA_API_KEY, captcha_id)).text
    while 'CAPCHA_NOT_READY' in recaptcha_answer:
        time.sleep(5)
        recaptcha_answer = s.get("http://2captcha.com/res.php?key={}&action=get&id={}".format(CAPTCHA_API_KEY, captcha_id)).text
    recaptcha_answer = recaptcha_answer.split('|')[1]

    postHeaders= {
        'X-Requested-With': 'XMLHttpRequest'
    }
    # then send the get request to the url
    pickPage = s.get(url=url + '?g-recaptcha-response=' + recaptcha_answer,headers=postHeaders)
    return pickPage

def tryCaptcha(url, captchaNode, tries):
    for key_node in captchaNode:
        site_key = key_node.text.split("var recaptchaKey = '")[1].split("';")[0]
    try:
        page = solveCaptcha(url, site_key)
    except Exception as e:
        log("Error while solving captcha", e)
        if tries >= 2:
            raise
        print("Retrying...")
        time.sleep(5)
        return tryCaptcha(url, captchaNode, tries + 1)

    tree = html.fromstring(page.content)
    path = tree.xpath("//div[contains(concat(' ', normalize-space(@class), ' '), ' feed-pick-title ')]/div[1]")
    if len(path) == 0:
        if tries >= 2:
            raise Exception("Failed to complete captcha")
        log("Failed to solve captcha", "")
        print("Retrying...")
        time.sleep(5)
        page = mySession.get(url)
        tree = html.fromstring(page.content)
        captchaNode = tree.xpath("//script[contains(text(), 'recaptchaKey = ')]")
        return tryCaptcha(url, captchaNode, tries + 1)
    else:
        return tree

combofields = {}
def parseCombo(node, index):
    global combofields
    if index == 0:
        combofields["Pick Type"] = "Combo Pick"
        combofields["Odds"] = jsonfields["Odds"]
        combofields["Stake"] = jsonfields["Stake"]
        if 'Bookie' in jsonfields:
            combofields["Bookie"] = jsonfields["Bookie"]
    if index > 5:
        raise Exception("Error: cannot process more than 5 picks in one combo pick")
    i = str(index)
    tdIter = iter(node)
    child = next(tdIter)
    if child.tag != 'td':
        return
    combofields["Sport" + i] = next(iter(child)).attrib["title"]
    combofields["Group" + i] = "Sprt" + combofields["Sport" + i].split(' ')[0]
    combofields["Teams" + i] = next(tdIter).text
    combofields["Pick" + i] = next(tdIter).text
    combofields["Odds" + i] = next(tdIter).text

def printComboCSV():
    if not "Sport1" in combofields:
        raise Exception("Could not read combo picks")
    os.chmod(CSV_COMBO_FILE_NAME, S_IWUSR)
    csv_f = open(CSV_COMBO_FILE_NAME, "a+")
    csv_writer = csv.writer(csv_f, delimiter=CSV_DELIMITER, lineterminator="\n")
    csv_writer.writerow([
        combofields["Pick Type"],
        combofields["Date"],
        combofields["Odds"],
        combofields["Stake"],
        combofields["Bookie"] if "Bookie" in combofields else "",
        combofields["Sport1"] if "Sport1" in combofields else "",
        combofields["Group1"] if "Group1" in combofields else "",
        combofields["Teams1"] if "Teams1" in combofields else "",
        combofields["Pick1"] if "Pick1" in combofields else "",
        combofields["Odds1"] if "Odds1" in combofields else "",
        combofields["Sport2"] if "Sport2" in combofields else "",
        combofields["Group2"] if "Group2" in combofields else "",
        combofields["Teams2"] if "Teams2" in combofields else "",
        combofields["Pick2"] if "Pick2" in combofields else "",
        combofields["Odds2"] if "Odds2" in combofields else "",
        combofields["Sport3"] if "Sport3" in combofields else "",
        combofields["Group3"] if "Group3" in combofields else "",
        combofields["Teams3"] if "Teams3" in combofields else "",
        combofields["Pick3"] if "Pick3" in combofields else "",
        combofields["Odds3"] if "Odds3" in combofields else "",
        combofields["Sport4"] if "Sport4" in combofields else "",
        combofields["Group4"] if "Group4" in combofields else "",
        combofields["Teams4"] if "Teams4" in combofields else "",
        combofields["Pick4"] if "Pick4" in combofields else "",
        combofields["Odds4"] if "Odds4" in combofields else "",
        combofields["Sport5"] if "Sport5" in combofields else "",
        combofields["Group5"] if "Group5" in combofields else "",
        combofields["Teams5"] if "Teams5" in combofields else "",
        combofields["Pick5"] if "Pick5" in combofields else "",
        combofields["Odds5"] if "Odds5" in combofields else "",
        ])
    csv_f.close()
    os.chmod(CSV_COMBO_FILE_NAME, S_IREAD|S_IRGRP|S_IROTH) # Make csv read-only


def parsePage(url):
    global jsonfields
    print(url)
    page = mySession.get(url)
    tree = html.fromstring(page.content)
    captcha = tree.xpath("//script[contains(text(), 'recaptchaKey = ')]")
    path = tree.xpath("//div[contains(concat(' ', normalize-space(@class), ' '), ' feed-pick-title ')]/div[1]")
    if len(path) == 0 and len(captcha) > 0:
        tree = tryCaptcha(url, captcha, 0)
        path = tree.xpath("//div[contains(concat(' ', normalize-space(@class), ' '), ' feed-pick-title ')]/div[1]")
    combo = tree.xpath("//table[contains(@class, 'combo-table')]")
    for node in path:
        for child in node:
            parseChildren(child)
    if len(combo) > 0:
        global combofields
        combofields = {}
        comboIndex = 0
        for tableBody in combo:
            for tr in tableBody:
                parseCombo(tr, comboIndex)
                comboIndex += 1
        date = tree.xpath("//small[contains(@class, 'bet-age')]")
        combofields["Date"] = date[0].text
        printComboCSV()
        try:
            response = requests.post(
                WEBHOOK_COMBO_URL, data=json.dumps(combofields),
                headers={'Content-Type': 'application/json'}
            )
            if response.status_code != 200:
                raise ValueError(
                    'Request to combo webhook returned an error %s, the response is:\n%s'
                    % (response.status_code, response.text)
                )
        except Exception as e:
            print("Could not write to combo webhook: " + str(e))
        print(json.dumps(combofields))
    else:
        if len(jsonfields) > 1:
            print(json.dumps(jsonfields))
            # Write csv to file
            os.chmod(CSV_FILE_NAME, S_IWUSR)
            csv_f = open(CSV_FILE_NAME, "a+")
            csv_writer = csv.writer(csv_f, delimiter=CSV_DELIMITER, lineterminator="\n")
            csv_writer.writerow([
                jsonfields["Type"],
                jsonfields["Date"],
                jsonfields["Sport"],
                jsonfields["Teams"],
                jsonfields["Pick"],
                "'" + jsonfields["Stake"] + "'",
                jsonfields["Odds"],
                jsonfields["Bookie"],
                jsonfields["Group"]
                ])
            csv_f.close()
            os.chmod(CSV_FILE_NAME, S_IREAD|S_IRGRP|S_IROTH) # Make csv read-only

            try:
                response = requests.post(
                    WEBHOOK_URL, data=json.dumps(jsonfields),
                    headers={'Content-Type': 'application/json'}
                )
                if response.status_code != 200:
                    raise ValueError(
                        'Request to webhook returned an error %s, the response is:\n%s'
                        % (response.status_code, response.text)
                    )
            except Exception as e:
                print("Could not write to json webhook: " + str(e))
        else:
            raise Exception("Something went wrong scraping the page")

def parseEmail(htmlstring):
    tree = html.fromstring(str(htmlstring))
    path = tree.xpath('//a/@href')
    for url in path:
        if str.find(url, 'pick') != -1:
            parsePage(url)
            return

def get_mail_ids():
    print("Checking for unread mails")
    try:
        typ, data = mail.search(None, '(UNSEEN)')
        mail_ids = data[0]
        id_list = mail_ids.split()
        print(str(len(id_list)) + " unread mails found")
        return id_list
    except Exception as e:
        log("Error while retrieving mail ids", e)
        raise

def read_email_from_server(id_list):
    global jsonfields
    mail_counter = len(id_list)
    print('Mails to process: ' + str(mail_counter))
    for i in id_list:
        i = int(i)
        try:
            typ, data = mail.fetch(str(i), '(RFC822)' )
        except Exception as e:
            log("Error while retrieving mail with id: " + str(i), e)
        for response_part in data:
            try:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    jsonfields["Type"] = ' '.join(str(msg["SUBJECT"]).split(' ')[0:2]) # Grab the first two words of the subject as type
                    if msg.is_multipart():
                        parseEmail(msg.get_payload(1).get_payload(decode=True))
                    else:
                        parseEmail(msg.get_payload(decode=True))
                    mail_counter -= 1
                    if mail_counter != 0:
                        time.sleep(SLEEP_TIME)
            except Exception as e:
                log("Error while processing mail with id: " + str(i), e)
        jsonfields = {}


if os.path.exists("ports"):
    fPorts = open("ports", "r")
    if fPorts.mode == "r":
        port = fPorts.read()
        if port.isdigit():
            try:
                for proc in process_iter():
                    for conns in proc.connections(kind='inet'):
                        if conns.laddr.port == int(port):
                            proc.send_signal(SIGTERM)
                            print("Closed socket on port " + port)
            except Exception as e:
                print("Could not close socket running on port: " + port)
                print(str(e))
    else:
        print("Could not open ports file")

if not os.path.exists(CSV_FILE_NAME):
    csv_file = open(CSV_FILE_NAME, "w+")
    csv.writer(csv_file, delimiter=CSV_DELIMITER, lineterminator="\n").writerow(["Type", "Date", "Sports", "Teams", "Pick", "Stake", "Odds", "Bookie", "Group field"])
    csv_file.close()
    os.chmod(CSV_FILE_NAME, S_IREAD|S_IRGRP|S_IROTH) # Set to read only so Excel doesn't lock the file

if not os.path.exists(CSV_COMBO_FILE_NAME):
    csv_file = open(CSV_COMBO_FILE_NAME, "w+")
    csv.writer(csv_file, delimiter=CSV_DELIMITER, lineterminator="\n").writerow(
        ["Type", "Date", "Odds", "Stake", "Bookie", 
         "Sport1", "Team1", "Pick1", "Odds1", 
         "Sport2", "Team2", "Pick2", "Odds2",
         "Sport3", "Team3", "Pick3", "Odds3",
         "Sport4", "Team4", "Pick4", "Odds4",
         "Sport5", "Team5", "Pick5", "Odds5",])
    csv_file.close()
    os.chmod(CSV_COMBO_FILE_NAME, S_IREAD|S_IRGRP|S_IROTH) # Set to read only so Excel doesn't lock the file

mail = None
idler = None
manual_close = False
while not manual_close:
    mailtries = 0
    logintries = 0
    print("Connecting")
    while mailtries < 2:
        try:
            mail = imaplib2.IMAP4_SSL(SMTP_SERVER, SMTP_PORT)
            break
        except Exception as e:
            log("Error while connecting to server", e)
            if mailtries >= 2:
                raise
            mailtries += 1
            print("Retrying connection...")
            time.sleep(5)

    print("Connected")

    if not os.path.exists("ports"):
        fPorts = open("ports", "w+")
        fPorts.write("")
        fPorts.close()
    fPorts = open("ports", "a+")
    fPorts.write(str(mail.socket().getsockname()[1]))
    fPorts.close()

    print("Logging in")
    while logintries < 2:
        try:
            mail.login(MAIL_ADDRESS,MAIL_PASSWORD)
            break
        except Exception as e:
            log("Error while logging in to server", e)
            if logintries >= 2:
                raise
            logintries += 1
            print("Retrying login...")
            time.sleep(5)

    print("Logged in")

    try:
        # debug lines to find the correct value for MAIL_LABEL
        #res, data = mail.list()
        #print(str(data))
        #raise Exception("debug")

        mail.select(MAIL_LABEL)

        # Start the Idler thread
        idler = Idler(mail)
        idler.start()
        # Keep running forever
        while not idler.needsReset.isSet():
            time.sleep(3)
            if (datetime.now() - idler.timeSinceMailCheck).total_seconds() >= MAIL_CHECK_INTERVAL:
                idler.needsync = True
                idler.event.set()
            elif idler.idling.isSet():
                print("Waiting for new messages...")
        print("Idler thread was interrupted:")
        raise idler.needsResetExc
    except KeyboardInterrupt as e:
        print("Closing down")
        manual_close = True
    except imaplib2.IMAP4.abort as e:
        log("Connection error", e)
        print("Retrying connection...")
        time.sleep(5)
    except Exception as e:
        log("An uncaught exception occurred", e)
        raise
    finally:
        # Clean up.
        if idler:
            idler.needsReset.clear()
            idler.stop()
            idler.join()
        if mail:
            try:
                mail.close()
            except Exception as e:
                log("Error while closing the mailbox", e)
            # This is important!
            mail.logout()
        os.remove("ports")
        print("Connection closed")
