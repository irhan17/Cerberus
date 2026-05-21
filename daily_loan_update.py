#Description: Scrapes web for news articles regarding specified loan properties and emails daily update of findings

#Imports libraries
import pandas as pd
import requests
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import win32com.client as win32
import os

#List of loan properties
properties = [
    "Augusta Mall",
    "Shops at Palm Desert",
    "Courtyard Basking Ridge",
    "Courtyard Newark Silicon Valley",
    "Townplace Suites Manhattan Beach",
    "Residence Inn Las Vegas",
    "Residence Inn Newark",
    "SpringHill Suites Manhattan Beach",
    "Residence Inn Phoenix",
    "Courtyard Scottsdale",
    "SpringHill Suites Plymouth Meeting",
    "90 Hudson Jersey City",
    "Eastview Mall",
    "Prince Building New York",
    "Cottonwood Mall",
    "West County Center",
    "One South Broad",
    "805 Third Avenue New York",
    "Hughes Center Las Vegas",
    "Fox River Mall",
    "555 11th Street NW"
]

#Defines number of hours to look back and output file name
LOOKBACK_HOURS = 24
output_path = "daily_loan_news.xlsx"

#Function that pulls news from Google RSS
def get_news(property_name):
    query = property_name.replace(" ", "+")
    #Fixed URL
    url = f"https://news.google.com/rss/search?q={query}&hl=en-US&gl=US&ceid=US:en"
    response = requests.get(url)
    root = ET.fromstring(response.content)
    articles = []
    for item in root.findall(".//item"):
        title = item.find("title").text
        link = item.find("link").text
        pub_date = item.find("pubDate").text
        pub_date = pd.to_datetime(pub_date)
        if pub_date >= datetime.now() - timedelta(hours=LOOKBACK_HOURS):
            articles.append({
                "Property": property_name,
                "Title": title,
                "Link": link,
                "Published": pub_date
            })
    return articles

#Main loop
all_articles = []
for prop in properties:
    results = get_news(prop)
    all_articles.extend(results)
df = pd.DataFrame(all_articles)

#Exports to Excel
if not df.empty:
    df.sort_values("Published", ascending=False, inplace=True)
    df.to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active
    #Auto-fits columns A and B
    for col in ["A", "B"]:
        max_length = 0
        for cell in ws[col]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col].width = max_length + 2
    #Hyperlinks the article links in Column C
    for row in range(2, ws.max_row + 1):
        cell = ws[f"C{row}"]
        if cell.value:
            cell.hyperlink = cell.value
            cell.style = "Hyperlink"
    wb.save(output_path)
    print(f"{len(df)} new articles found.")
    print(f"File saved as: {output_path}")
else:
    print("No new articles found.")

#Sends email via Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
date_string = datetime.now().strftime("%#m/%#d")
mail.Subject = f"{date_string} Daily Loan News Update"
mail.To = "consult-iiftikar@cerberus.com; mwaldenberg@cerberus.com"

#Puts table from Excel file in a table within the email
if not df.empty:
    df_html = df.copy()
    df_html["Link"] = df_html["Link"].apply(lambda x: f'<a href="{x}">Open Article</a>' if pd.notnull(x) else "")
    html_table = df_html.to_html(escape=False, index=False)
else:
    html_table = "<p>No new articles found.</p>"
mail.HTMLBody = f"""
<p>{len(df)} new articles found in past 24 hours. See attached file (attached only if new articles were found).</p>
{html_table}"""

#Sends email
if os.path.exists(output_path):
    mail.Attachments.Add(os.path.abspath(output_path))
mail.Send()
print("Email sent via Outlook.")