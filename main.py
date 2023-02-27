import smtplib
import numpy as np
import pandas as pd
import random
import glob
import csv
import os
from random import randrange
from datetime import timedelta, datetime
from xlwings import Book

SEND = False # When True, emails will be sent
REFRESH_QUESTIONS = True # When True, question parameters will be randomized, for demonstration purposes


if __name__ == "__main__":
    Book("RFI_TRC_Automated_Tool.xlsm").set_mock_caller()

if REFRESH_QUESTIONS:
    aa_questions = ['what color to paint?', 'what color baffles?', 'how many baffles?', 'where to put baffles?',
                    'knockout panel?', 'what is AA?', 'where is the door?', 'where to put bench?', "hand rail?",
                    "stainless steel?", "exit corridor?", "public flooring?", "hang a picture?",
                    "is the station pretty?",
                    ]
    me_questions = ['what pipe to use?', 'what circuit to connect?', 'which conduit?', 'where to put pipe?',
                    'where to put AC?', 'how to connect pipe?', 'where drain?', 'why drain?', 'what is electricity?',
                    'what is plumbing?', 'what is HVAC?', 'where is the pipe?', 'where is the cable?',
                    "what's a cable tray?", "what's MEP?", 'how cold hvac?', "fresh air in the station?",
                    "how much fresh air intake?", "what should the spi pressure be?", "how many hvac grilles?",
                    "how many conduits?", "how many pipes?", "hdpe or steel?", "can unistrut be used for hvac?",
                    "which opening in GN1212 to use?"]
    sd_questions = ['what is concrete?', "what's b-30 concrete?", 'where to put rebar?', 'what concrete to use?',
                    'how to demolish?', 'can we demolish this?', 'where to put concrete?', 'which rebar?',
                    "what diameter rebar?", "king pile diameter?", "should we demolish king piles?",
                    "what about queen piles?", "what about jack piles?"]
    STATIONS = ["DP01", "DP02", "DP03", "DP04", "DP05", "DP06", "DP07", "DP08", "DP09", "DP10"]
    question_count = (len(aa_questions) + len(sd_questions) + len(me_questions))
    for i in range(1, question_count + 1):
        name = f"RFI {str(i).zfill(5)}.CSV"
        with open(name, "w", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["Name", "Date Document Received", "Date of Required Response",
                             "Discipline", "Location", "LPMC Input"])

            delta_date = timedelta(10)
            today_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            start_date = today_date - delta_date
            random_days = randrange(delta_date.days)
            end_date = start_date + timedelta(random_days)
            random_start = start_date + timedelta(random_days)

            # Select Question:
            discipline_choice = random.choice(["aa", "me", "sd"])
            if discipline_choice == "aa":
                try:
                    LPMC_input = random.choice(aa_questions)
                except IndexError:
                    discipline_choice = "me"
                else:
                    aa_questions.pop(aa_questions.index(LPMC_input))
                    Discipline = "AA"

            elif discipline_choice == "me":
                try:
                    LPMC_input = random.choice(me_questions)
                except IndexError:
                    discipline_choice = "sd"
                else:
                    me_questions.pop(me_questions.index(LPMC_input))
                    Discipline = "ME"
            elif discipline_choice == "sd":
                try:
                    LPMC_input = random.choice(sd_questions)
                except IndexError:
                    discipline_choice = "aa"
                else:
                    sd_questions.pop(sd_questions.index(LPMC_input))
                    Discipline = "SD"

            writer.writerow([name[:9], random_start, random_start + timedelta(random.choice([7, 14])), Discipline,
                             random.choice(STATIONS), LPMC_input])


def book_caller():
    global wb
    wb = Book.caller()


book_caller()
sheet = wb.sheets("Teams")
aa_manager = sheet.range("I2").value
me_manager = sheet.range("J2").value
sd_manager = sheet.range("K2").value
all_managers = [aa_manager, me_manager, sd_manager]
aa_staff = list(sheet.range("I3").value.split(","))
me_staff = list(sheet.range("J3").value.split(","))
sd_staff = list(sheet.range("K3").value.split(","))
all_staff = aa_staff + me_staff + sd_staff

rfitrc_list = []
files = glob.glob(r"""C:\Users\akiva\Documents\Python\RFI TRC Tool\RFI_TRC_Automated_Tool/*.csv""")
for file in files:
    df = pd.read_csv(file)
    rfitrc_file = {}
    for column in df.columns:
        rfitrc_file[column] = df[column][0]
    rfitrc_list.append(rfitrc_file)

aa_rfitrc = []
me_rfitrc = []
sd_rfitrc = []
rfitrc_dict = {}
for i in rfitrc_list:
    rfitrc_dict[i["Name"]] = i

for i in rfitrc_list:
    if i["Discipline"] == "AA":
        aa_rfitrc.append(i)
    elif i["Discipline"] == "ME":
        me_rfitrc.append(i)
    elif i["Discipline"] == "SD":
        sd_rfitrc.append(i)


def get_staff():
    wb = Book.caller()
    sht = wb.sheets("Stations")
    df = sht.range("A1:E12").options(pd.DataFrame).value
    return df


staff_df = get_staff()

for key, value in rfitrc_dict.items():
    value["Responsible"] = \
        staff_df[staff_df["Station"] == value["Location"]].iloc[:,
        staff_df.columns.get_loc(value["Discipline"])].values[0]

responsible_dict = {}
for i in all_staff:
    responsible_dict[i] = []

for person in all_staff:
    for key, value in rfitrc_dict.items():
        if value["Responsible"] == person:
            responsible_dict[person].append(key)

# Calculate discipline totals
disciplines = ["aa", "me", "sd"]
me_totals = {}
aa_totals = {}
sd_totals = {}
me_totals["late"] = []
me_totals["today"] = []
me_totals["future"] = []
me_totals["this week"] = []
aa_totals["late"] = []
aa_totals["today"] = []
aa_totals["future"] = []
aa_totals["this week"] = []
sd_totals["late"] = []
sd_totals["today"] = []
sd_totals["future"] = []
sd_totals["this week"] = []

for key, value in rfitrc_dict.items():
    if value["Discipline"] == "ME":
        if datetime.strptime(value["Date of Required Response"], "%Y-%m-%d %H:%M:%S") < (datetime.today().replace(
                hour=0, minute=0, second=0, microsecond=0
        )):
            me_totals["late"].append(key)
        elif datetime.strptime(value["Date of Required Response"], "%Y-%m-%d %H:%M:%S") > datetime.today().replace(
                hour=0, minute=0, second=0, microsecond=0
        ) + \
                timedelta(days=7):
            me_totals["future"].append(key)
        else:
            me_totals["this week"].append(key)
    if value["Discipline"] == "AA":
        if datetime.strptime(value["Date of Required Response"], "%Y-%m-%d %H:%M:%S") < (datetime.today().replace(
                hour=0, minute=0, second=0, microsecond=0
        )):
            aa_totals["late"].append(key)
        elif datetime.strptime(value["Date of Required Response"], "%Y-%m-%d %H:%M:%S") > datetime.today().replace(
                hour=0, minute=0, second=0, microsecond=0
        ) + timedelta(days=7):
            aa_totals["future"].append(key)
        else:
            aa_totals["this week"].append(key)
    if value["Discipline"] == "SD":
        if datetime.strptime(value["Date of Required Response"], "%Y-%m-%d %H:%M:%S") < (datetime.today().replace(
                hour=0, minute=0, second=0, microsecond=0
        )):
            sd_totals["late"].append(key)
        elif datetime.strptime(value["Date of Required Response"], "%Y-%m-%d %H:%M:%S") > datetime.today().replace(
                hour=0, minute=0, second=0, microsecond=0
        ) + timedelta(days=7):
            sd_totals["future"].append(key)
        else:
            sd_totals["this week"].append(key)
rfitrc_df = pd.DataFrame.from_dict(rfitrc_dict)
rfitrc_df = pd.DataFrame.transpose(rfitrc_df)

# add status late/today/future
rfitrc_df["Date of Required Response"] = pd.to_datetime(rfitrc_df["Date of Required Response"])
rfitrc_df["Date of Required Response"] = rfitrc_df["Date of Required Response"]
rfitrc_df["Today"] = datetime.today()
rfitrc_df["Today Date"] = rfitrc_df["Today"].dt.date
rfitrc_df["Date of Required Response Date"] = rfitrc_df["Date of Required Response"].dt.date
rfitrc_df["Days Late TD"] = (rfitrc_df["Today Date"] - rfitrc_df["Date of Required Response Date"])
rfitrc_df["Days Late Int"] = rfitrc_df["Days Late TD"].dt.days
conditions = [
    (rfitrc_df["Days Late Int"] > 0),
    (rfitrc_df["Days Late Int"] == 0),
    (rfitrc_df["Days Late Int"] < -7),
    True
]
values = ["Late", "Today", "Future", "This Week"]
rfitrc_df["Status"] = np.select(conditions, values)

wb = Book.caller()
sheet = wb.sheets("Backend")
sheet.range("C3:E5").clear_contents()
sheet.range("C3").value = len(aa_totals["late"])
sheet.range("C4").value = len(me_totals["late"])
sheet.range("C5").value = len(sd_totals["late"])
sheet.range("D3").value = len(aa_totals["this week"])
sheet.range("D4").value = len(me_totals["this week"])
sheet.range("D5").value = len(sd_totals["this week"])
sheet.range("E3").value = len(aa_totals["future"])
sheet.range("E4").value = len(me_totals["future"])
sheet.range("E5").value = len(sd_totals["future"])

# Save photo of discipline Backend table
sht = wb.sheets("Pictures")
discipline_table = sht.pictures[0]
discipline_table.api.Copy()
from PIL import ImageGrab

img = ImageGrab.grabclipboard()
img.save("discipline_table.png")

discipline_table = sht.pictures[1]
discipline_table.api.Copy()
img = ImageGrab.grabclipboard()
img.save("discipline_chart.png")

# Dictionary member - RFITRC
cols = ["AA", "ME", "SD"]
members_stations_dict = {}
members_rfitrc_dict = {}
for member in all_staff:
    members_stations_dict[member] = list((staff_df["Station"][(staff_df[cols] == member).any(axis="columns")].values))
    for key, value in rfitrc_dict.items():
        if value["Responsible"] == member:
            if member in members_rfitrc_dict:
                members_rfitrc_dict[member].append(key)
            else:
                members_rfitrc_dict[member] = list(key.split(","))

# Email Message - Staff
my_email = "" # Deleted for privacy
connection = smtplib.SMTP("smtp.gmail.com")
connection.starttls()
connection.login(user=my_email, password="") # Deleted for privacy

for member in all_staff:
    try:
        member_message = f"""Good morning {member},\n\nYou have {len(members_rfitrc_dict[member])} open RFI/TRC's:\n"""
    except KeyError:
        pass
    else:
        for i in members_rfitrc_dict[member]:
            member_message += f"""{i} ({rfitrc_df[rfitrc_df["Name"] == i]["Location"].item()}) - {rfitrc_df[rfitrc_df["Name"] == i]["Status"].
            item()} - due in {0 - rfitrc_df[rfitrc_df["Name"] == i]
            ["Days Late Int"].item()} day(s).\n"""
        member_message += "\nGood luck and have a great day.\n\nYours truly,\nThe RFI/TRC Robot"
        if SEND:
            connection.sendmail(from_addr=my_email,
                                to_addrs="", # Deleted for privacy
                                msg=f"Subject:{member}'s Daily RFI/TRC Status\n\n{member_message}")

from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

# Report list, sorted by discipline, showing name and days late
report_list = f"""AA: {len(rfitrc_df[rfitrc_df["Discipline"] == "AA"])}, \
{len(rfitrc_df[rfitrc_df["Days Late Int"] > 0][rfitrc_df["Discipline"] == "AA"])} late\n"""
for rfitrc in aa_rfitrc:
    report_list += f"""{rfitrc["Name"]} ({rfitrc["Location"]}) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Status"].item()}:\
 {0 - rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Days Late Int"].item()} day(s)\n"""

report_list += f"""\nMEP: {len(rfitrc_df[rfitrc_df["Discipline"] == "ME"])}, \
{len(rfitrc_df[rfitrc_df["Days Late Int"] > 0][rfitrc_df["Discipline"] == "ME"])} late\n"""
for rfitrc in me_rfitrc:
    report_list += f"""{rfitrc["Name"]} ({rfitrc["Location"]}) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Status"].item()}:\
 {0 - rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Days Late Int"].item()} day(s)\n"""

report_list += f"""\nSD: {len(rfitrc_df[rfitrc_df["Discipline"] == "SD"])}, \
{len(rfitrc_df[rfitrc_df["Days Late Int"] > 0][rfitrc_df["Discipline"] == "SD"])} late\n"""
for rfitrc in sd_rfitrc:
    report_list += f"""{rfitrc["Name"]} ({rfitrc["Location"]}) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Status"].item()}:\
 {0 - rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Days Late Int"].item()} day(s)\n"""

# Import complete lists into excel
sheet = wb.sheets("Backend")
sheet.range("H1").clear_contents()
sheet.range("H1").value = report_list

# Manager emails
mep_manager_message = f"Good morning Teddy,\n\nHere are your team's open RFI/TRC's:\n"
for rfitrc in me_rfitrc:
    mep_manager_message += f"""{rfitrc["Name"]} ({rfitrc["Location"]}) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Status"].item()}:\
 {0 - rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Days Late Int"].item()} day(s) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Responsible"].item()}\n"""
mep_manager_message += "\nGood luck and have a great day.\n\nYours truly,\nThe RFI/TRC Robot"

aa_manager_message = f"Good morning Chaya,\n\nHere are your team's open RFI/TRC's:\n"
for rfitrc in aa_rfitrc:
    aa_manager_message += f"""{rfitrc["Name"]} ({rfitrc["Location"]}) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Status"].item()}:\
 {0 - rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Days Late Int"].item()} day(s) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Responsible"].item()}\n"""
aa_manager_message += "\nGood luck and have a great day.\n\nYours truly,\nThe RFI/TRC Robot"

sd_manager_message = f"Good morning Thabet,\n\nHere are your team's open RFI/TRC's:\n"
for rfitrc in sd_rfitrc:
    sd_manager_message += f"""{rfitrc["Name"]} ({rfitrc["Location"]}) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Status"].item()}:\
 {0 - rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Days Late Int"].item()} day(s) - \
{rfitrc_df[rfitrc_df["Name"] == rfitrc["Name"]]["Responsible"].item()}\n"""
sd_manager_message += "\nGood luck and have a great day.\n\nYours truly,\nThe RFI/TRC Robot"

# send manager messages
manager_messages = {"MEP": mep_manager_message, "AA": aa_manager_message, "SD": sd_manager_message}

if SEND:
    for key, value in manager_messages.items():
        connection.sendmail(from_addr=my_email,
                            to_addrs="", # Deleted for privacy
                            msg=f"Subject:{key} Team's Daily RFI/TRC Status\n\n{value}")

# Overall status email
overall_message1 = "Good morning everyone. Here is the current status of all RFI/TRC's. Please see charts attached.\n\n"
overall_message2 = f"{report_list}\n\n"
overall_message2 += "Good luck and have a great day.\n\nYours truly,\nThe RFI/TRC Robot"

if SEND:

    image_names = ["discipline_chart.png", "discipline_table.png"]

    msg = MIMEMultipart()
    msg['Subject'] = "Overall Daily RFI/TRC Status"
    msg['From'] = "" # Deleted for privacy
    msg['To'] = "" # Deleted for privacy

    text = MIMEText(overall_message1)
    msg.attach(text)
    for i in image_names:
        with open(i, 'rb') as f:
            img_data = f.read()
        image = MIMEImage(img_data, name=os.path.basename(i))
        msg.attach(image)
    text = MIMEText(overall_message2)
    msg.attach(text)
    s = smtplib.SMTP("smtp.gmail.com")
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login(user="", password="") # Deleted for privacy
    s.sendmail(from_addr="", to_addrs="", msg=msg.as_string()) # Deleted for privacy
    s.quit()
