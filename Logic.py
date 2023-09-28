import os
import requests
import openpyxl
import shutil
import datetime
import csv

from bs4 import BeautifulSoup

# Constants
BASE_URL = "http://181.48.43.68:90/crm/677976/admin/tickets/ticket/"
COOKIE_STRING = "autologin=a%3A2%3A%7Bs%3A7%3A%22user_id%22%3Bs%3A3%3A%22120%22%3Bs%3A3%3A%22key%22%3Bs%3A16%3A%229716885475f092b3%22%3B%7D"
DOWNLOAD_DIRECTORY = "Download/CRM"

# Parse the cookie string and convert it to a dictionary
def parse_cookie_string(cookie_string):
    cookies = {}
    parts = cookie_string.split(';')
    for part in parts:
        key, value = part.strip().split('=')
        cookies[key] = value
    return cookies

# Send an HTTP GET request and return the BeautifulSoup object
def send_get_request(url, cookies):
    response = requests.get(url, cookies=cookies)
    if response.status_code == 200:
        return BeautifulSoup(response.text, 'html.parser')
    else:
        return None

# Get the description from the page
def getDescription(soup):
    tc_content_div = soup.find('div', {'class': 'tc-content'})
    if tc_content_div:
        return tc_content_div.text.strip()
    else:
        return ""

# Get the subject from the page
def getSubject(soup):
    subject_span = soup.find('span', id='ticket_subject')
    if subject_span:
        return subject_span.text.strip()
    else:
        return "N/A"

# Extract the project name from the page
def extractProjectName(soup):
    def filter_small_element(tag):
        return tag.name == 'small' and "Este ticket está enlazado con el proyecto:" in tag.get_text()

    small_element = soup.find(filter_small_element)

    if small_element:
        a_element = small_element.find('a')
        if a_element:
            return a_element.text.strip()
    return ""

# Download a file and return the file path
def getTalla(soup, cookies, TICKET_NUM):
    links = soup.find_all('a', href=True, class_='display-block mbot5')
    # last_link =
    os.makedirs(DOWNLOAD_DIRECTORY, exist_ok=True)

    file_path = os.path.join(DOWNLOAD_DIRECTORY, "CRM-" + TICKET_NUM + ".xlsx")
    response = requests.get(links[len(links) - 1]['href'], cookies=cookies)
    with open(file_path, 'wb') as file:
      file.write(response.content)

    workbook = openpyxl.load_workbook(file_path, data_only=True)
    worksheet = workbook["Tallaje"]
    cell_value = worksheet["G6"].value
    workbook.close()
    return cell_value

# Get the value of the Talla from the downloaded Excel file
def getTallaValue(file_path):
    if not file_path:
        return None

    workbook = openpyxl.load_workbook(file_path, data_only=True)
    worksheet = workbook.active

    # Assuming the value is in cell G6
    cell_value = worksheet['G6'].value

    workbook.close()
    return cell_value

# Get the value for a given size string
def get_value_for_size(input_str):
  try:
    size_to_value = {
        "XS": "0.5",
        "S" : "0.7",
        "M" : "1",
        "L" : "1.7",
        "XL": "2"
    }
    input_str = input_str.upper()

    # Check if the input string is in the dictionary
    if input_str in size_to_value:
        return size_to_value[input_str]
    else:
        return "0"  # Return None for invalid input
  except:
    return "0"

def getServicioValue(soup):
    servicio_span = soup.find_all('span', {'class': 'ticket-label label label-default inline-block'})[1]
    if servicio_span:
        # Extract the text after "Servicio: "
        text = servicio_span.get_text(strip=True).replace('Servicio:', '')
        return text
    else:
        return "N/A"


def getData(TICKET_NUM):
    try:

      Start_date = datetime.date.today()
      ticket_url = BASE_URL + TICKET_NUM
      cookies = parse_cookie_string(COOKIE_STRING)
      soup = send_get_request(ticket_url, cookies)

      if soup:
          print(f"Titulo: {getSubject(soup)}")
          print(f"Descripcion: {getDescription(soup)}")

          tpIncidente = getServicioValue(soup).strip()
          print(f"Tipo Incidente: {tpIncidente}")
          if tpIncidente != "incidente":
              talla = getTalla(soup, cookies, TICKET_NUM)
              print(f"Talla: {talla}")
              print(f"Puntos: {get_value_for_size(talla)}")
          else:
              print("Talla: XS")
              print("Puntos: 0.5")

          print(f"ClientName: {extractProjectName(soup)}")
          print("Responsable: N/A" )
          print("Formatos Requeridos: N/A")
          print("Start date: " , Start_date)
          print("Fecha de vencimiento: " , Start_date + datetime.timedelta(days=6))
          print("Sprint: N/A  ")
          print(f"Url Ticket: {ticket_url}")

      else:
          print(f"Failed to retrieve the desired page. URL: {ticket_url}")
    except:
        return "0"

def getData_2(TICKET_NUM):
  try:
      Start_date = datetime.date.today()
      ticket_url = BASE_URL + TICKET_NUM
      cookies = parse_cookie_string(COOKIE_STRING)
      soup = send_get_request(ticket_url, cookies)

      if soup:
          tpIncidente = getServicioValue(soup).strip()
          if tpIncidente != "incidente":
              talla = getTalla(soup, cookies, TICKET_NUM)
          else:
              talla = "XS"

          # Append data to the CSV file
          with open('ticket_data.csv', 'a', newline='', encoding='utf-8-sig') as csvfile:
              csvwriter = csv.writer(csvfile, delimiter=';')

              csvwriter.writerow([
                  TICKET_NUM, getSubject(soup), tpIncidente, talla,
                  get_value_for_size(talla), extractProjectName(soup), "N/A",
                  "N/A", Start_date, Start_date + datetime.timedelta(days=6),
                  "N/A", ticket_url
              ])

  except Exception as e:
      print(f"Error processing ticket {TICKET_NUM}: {str(e)}")

def dropFolder():
  try:
    # Use shutil.rmtree to remove the folder and its contents
    shutil.rmtree("Download")
    print(f"Folder '{DOWNLOAD_DIRECTORY}' and its contents have been deleted.")
  except Exception as e:
      print(f"An error occurred while deleting the folder: {str(e)}")

def createFinalFile():
    with open('ticket_data.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
      csvwriter = csv.writer(csvfile, delimiter=';')
      # Write the header row
      csvwriter.writerow([
          'CRM-Num', 'Resumen', 'Tipo Incidente', 'Talla', 'Story point estimate',
          'Nombre Cliente', 'Responsable', 'Formatos Requeridos', 'Start date',
          'Fecha de vencimiento', 'Sprint', 'Url Ticket'
      ])

def init_APP(TicketsToSearch):
    createFinalFile()
    for ticket in TicketsToSearch.split(','):
      getData_2(ticket)
    dropFolder()
