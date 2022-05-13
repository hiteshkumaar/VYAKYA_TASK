import json
import re
import xlwt
from xlwt import Workbook


def init_sheets(ab):
    withdrawsh = ab.add_sheet("WITHDRAWALS")
    deposh = ab.add_sheet("DEPOSITS")
    insightsh = ab.add_sheet("INSIGHTS")
    withdrawsh.write(0, 0, "DATE")
    withdrawsh.write(0, 1, "DESCRIPTION")
    withdrawsh.write(0, 2, "AMOUNT")
    withdrawsh.write(0, 3, "DAY")
    withdrawsh.write(0, 4, "MONTH")
    withdrawsh.write(0, 5, "YEAR")
    deposh.write(0, 0, "DATE")
    deposh.write(0, 1, "DESCRIPTION")
    deposh.write(0, 2, "AMOUNT")
    deposh.write(0, 3, "DAY")
    deposh.write(0, 4, "MONTH")
    deposh.write(0, 5, "YEAR")
    insightsh.write(0, 0, "key")
    insightsh.write(0, 1, "value")
    return withdrawsh, deposh, insightsh

with open("E://task_input_list.json", "r") as f:
    data = json.load(f)

site_pattern = r"(http(s)?://www.)?([A-Za-z])+([\w])*((\.com)|(\.in))"
email_pattern = r"[a-zA-Z]+[\w]*@(([A-Za-z])+\.(\w)+)"
amount_pattern = r"(-)?(\$)?((\d){1,2},)?(\d)+(\.)(\d)+"
phone_pattern = r"(((\(\d{3}\) )|(\d-\d{3}-))\d{3}-\d{4})|(\+(\d{1,2,3} \d{10}))"
date_pattern = r"\d{2}/\d{2}/\d{2}"

sites = set()
emails = set()
amounts = list()
phones = set()

ab = Workbook()

withdrawsh, deposh, insightsh = init_sheets(ab)

def save_data(ab: str, date: str, des: str, amount: str, row: int):
    mon, day, year = date.split("/")
    ab.write(row, 0, date)
    ab.write(row, 1, des)
    ab.write(row, 2, amount)
    ab.write(row, 3, day)
    ab.write(row, 4, mon)
    ab.write(row, 5, "20" + year)

ds_row, ws_row = 1, 1

i = 0
while True:
    try:
        item = data[i]
    except:
       
        break

    site_match = re.search(site_pattern, item)
    email_match = re.search(email_pattern, item)
    phone_match = re.search(phone_pattern, item)
    date_match = re.search(date_pattern, item)

 
    if date_match:
        date = date_match.group()
        i += 1
        description = ""
        while True:
            try:
                item = data[i]
            except:
                break

            amount_match = re.search(amount_pattern, item)
            if not amount_match is None:
                amount = amount_match.group()

                if amount[0] != "$" and amount[1] != "$":
                    # string to float
                    amount = float(amount.replace(",", ""))
                    amounts.append(amount)
                    if amount >= 0.0:
                        save_data(ab=deposh, date=date, des=description, amount=amount, row=ds_row)
                        ds_row += 1
                    else:
                        save_data(ab=withdrawsh, date=date, des=description, amount=amount, row=ws_row)
                        ws_row += 1
                    break
            else:
                description += item

            i += 1

        pass

    if not site_match is None:
        sites.add(site_match.group())

    if not email_match is None:
        emails.add(email_match.group())
    # phone numbers
    if not phone_match is None:
        phones.add(phone_match.group())

    i += 1
insightsh.write(2, 0, "email")
if not len(emails):
    insightsh.write(2, 1, "NA")
else:
    email_data = ""
    for email in emails:
        email_data += ", " + email
    insightsh.write(2, 1, email_data.strip(","))

insightsh.write(3, 0, "phone_numbers")
if not len(phones):
    insightsh.write(3, 1, "NA")
else:
    phone_data = ""
    for phone in phones:
        phone_data += ", " + phone
    insightsh.write(3, 1, phone_data.strip(","))

insightsh.write(1, 0, "website")
if not len(sites):
    insightsh.write(1, 1, "NA")
else:
    site_data = ""
    for site in sites:
        site_data += ", " + site
    insightsh.write(1, 1, site_data.strip(","))

insightsh.write(4, 0, "max amount")
insightsh.write(5, 0, "min amount")
if len(amounts):
    insightsh.write(4, 1, max(amounts))
    insightsh.write(5, 1, min(amounts))
else:
    insightsh.write(4, 1, "NA")
    insightsh.write(5, 1, "NA")
print(ab)

ab.save("C:\\Users\\user\\Desktop\\task_output1.xls")
