import anvil.server
from urllib.request import urlopen, urlretrieve
from bs4 import BeautifulSoup
import json
import datetime
import locale
import re
import time
import pyodbc

from openpyxl import Workbook
from epub_conversion.utils import open_book, convert_epub_to_lines


def create_workbook():
  book = open_book("mwb.epub")
  lines = convert_epub_to_lines(book)

  wb = Workbook()
  ws = wb.active

  with open('mwb.html', 'w', encoding="utf-8") as fp:
      for row in lines:
          fp.write(row)

  with open('mwb.html', 'r', encoding="utf-8") as fp:
      soup = BeautifulSoup(fp, 'html.parser')

  numbers = re.compile(r"\d+(?:\.\d+)?")

  # Set of preselected part titles
  part_titles = {"Permata Rohani",
                "Pembacaan Alkitab",
                "Memulai Percakapan",
                "Melakukan Kunjungan Kembali",
                "Menjadikan Murid",
                "Menjelaskan Kepercayaan Saudara",
                "Khotbah",
                "Kebutuhan Setempat",
                "Pelajaran Alkitab Sidang"}

  # find the year
  cover_title = soup.find("h1", class_="coverTtl", id="p1")
  year_string = numbers.findall(cover_title.text)[0]

  all_sections = soup.find_all("div", class_="bodyTxt")
  all_weeks = [w
              for w in all_sections
              if w.find_previous_sibling("header") is not None and "HARTA" in w.text]

  # Temporarily set location to Indonesian
  locale.setlocale(locale.LC_ALL, "IND")

  for week in all_weeks:
      date_string = week.find_previous_sibling("header").find("h1").text
      dash_position = date_string.find("-")
      if dash_position == -1:
          dash_position = date_string.find(chr(8211))
      space_position = date_string.find(" ")
      if space_position < 3:
          day_month = date_string[0:dash_position]
      else:
          day_month = date_string[0:dash_position] + " " + date_string[space_position:]
      monday_date = time.strptime(year_string + " " + day_month, "%Y %d %B")

      # find if local need week
      local_need_living = [living.text for living in week.find_all("h3", class_="du-color--maroon-600") if
                          "Kebutuhan Setempat" in living.text]
      is_local_need_week = len(local_need_living) > 0
      student = ""
      if is_local_need_week:
          student = "Local Need"

      # Assign the Indonesian chairman for the week
      if not is_local_need_week:
          excel_row = [datetime.datetime.strptime(str(monday_date.tm_mon) + "/"
                                                  + str(monday_date.tm_mday) + "/"
                                                  + str(monday_date.tm_year), "%m/%d/%Y"),
                      student, "", "", "", "Ketua", "", "", "Ketua Perhimpunan Kehidupan Pelayanan"]
          ws.append(excel_row)

      # Student Parts in Apply Yourself to Ministry
      parts = week.find_all("h3", class_="du-color--gold-700")
      for part in parts:
          part_title_strong = part.find("strong")
          if part_title_strong:
              part_title = part_title_strong.text.replace(":", "")  # 4. Pikirkan Minat Mereka—Apa yang Yesus Lakukan?
          else:
              part_title = ""

          part_title_only = part_title.split(".")[1].strip()
          if part_title_only not in part_titles:
              part_title = "Pembahasan"

          part_desc_all = (week.find("strong", string=part_title_strong.text).find_parent().find_next_sibling("div")
                          .find("p"))
          part_duration = part_desc_all.text[:part_desc_all.text.find(")") + 1]
          part_teaching = part_desc_all.text[part_desc_all.text.rfind("("):part_desc_all.text.rfind(")") + 1]
          if part_teaching == part_duration:
              part_teaching = ""
          part_desc_only = part_desc_all.text.replace(part_duration, "").replace(part_teaching, "").strip()
          part_type = part_desc_only.split(".")[0]

          excel_row = [datetime.datetime.strptime(str(monday_date.tm_mon) + "/"
                                                  + str(monday_date.tm_mday) + "/"
                                                  + str(monday_date.tm_year), "%m/%d/%Y"),
                      student, "", "", "",
                      part_title_only, part_teaching.replace("(", "").replace(")", ""), "", part_desc_all.text]
          ws.append(excel_row)

      # Parts in Living As Christians
      parts_lac = week.find_all("h3", class_="du-color--maroon-600")
      for part in parts_lac:
          part_title_strong = part.find("strong")
          if part_title_strong:
              part_title = part_title_strong.text.split(".")[1].strip()  # 7. Apakah Saudara Mau Memberikan Kesaksian Tidak Resmi?
          else:
              part_title = ""

          if len(part_title) == 0:
              continue

          if part_title not in part_titles:
              part_title = "Pembahasan/Khotbah"

          if part_title == "Kebutuhan Setempat":
              continue

          part_desc_all = part_title_strong.find_parent().find_next_sibling("div").find("p")

          part_duration = part_desc_all.text[:part_desc_all.text.find(")") + 1]
          part_desc_only = part_desc_all.text.strip()
          part_desc_only += " " + part_title_strong.text.split(".")[1].strip()
          part_type = part_desc_only.split(".")[0]

          excel_row = [datetime.datetime.strptime(str(monday_date.tm_mon) + "/"
                                                  + str(monday_date.tm_mday) + "/"
                                                  + str(monday_date.tm_year), "%m/%d/%Y"),
                      "", "", "", "", part_title, "", "", part_desc_only]

          if is_local_need_week:
              # only insert the PAS
              if part_title == "Pelajaran Alkitab Sidang":
                  ws.append(excel_row)
          else:
              ws.append(excel_row)

  locale.setlocale(locale.LC_ALL, "")

  excel_file_name = "mwb_" + str(monday_date.tm_year) + f"{monday_date.tm_mon:02}" + ".xlsx"

  wb.save(excel_file_name)

  media = anvil.media.from_file(excel_file_name)

  return media


def push_to_azure():
    # Send directly to Azure
  server = f'publisherscheduler1.database.windows.net'
  database = 'scheduler'
  username = 'indonesianphiladelphiapausa'
  password = '2033Ellsworth.'
  connstr = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
  cnxn = pyodbc.connect(connstr)
  cursor = cnxn.cursor()

  task_type = {"Pembacaan Alkitab": 1,
              "Pengabaran": 2,
              "Kunjungan Kembali": 3,
              "Pelajaran Alkitab": 4,
              "Khotbah": 5,
              "Video": 6,
              "Pembahasan": 7,
              "Khotbah KK": 8,
              "Pelajaran Alkitab Sidang": 9,
              "Ketua": 11,
              "Permata Rohani": 12,
              "Memulai Percakapan": 13,
              "Melakukan Kunjungan Kembali": 14,
              "Menjadikan Murid": 15,
              "Menjelaskan Kepercayaan Saudara": 16,
              "Kebutuhan Setempat": 17}

  for row in ws.rows:
      training_id_value = "NULL"
      if len(row[6].value) > 0:
          training_id_value = row[6].value.split()[2]
      task_id_value = task_type.get(row[5].value)
      if task_id_value is None:
          task_id_value = 0

      sqlcmd1 = (
          f"IF NOT EXISTS (select id from slots where datediff(day, begintime, '{row[0].value.strftime('%Y-%m-%d')}')=0 "
          f" AND tasktypeid = {str(task_id_value)} ")
      if training_id_value == "NULL":
          sqlcmd2 = (f" AND trainingid IS NULL) ")
      else:
          sqlcmd2 = (f" AND trainingid = {training_id_value}) ")

      sqlcmd3 = (f"BEGIN "
                f"insert into slots (projectid, begintime, endtime, isactive, description, tasktypeid, trainingid) "
                f"values (1"
                f", '{row[0].value.strftime('%m/%d/%Y')}'"
                f", '{row[0].value.strftime('%m/%d/%Y')}'"
                f", 1 "
                f", '{row[8].value}'"
                f", {str(task_id_value)}"
                f", {training_id_value}) "
                f"END"
                )
      cursor.execute(sqlcmd1 + sqlcmd2 + sqlcmd3)
      cnxn.commit()

  cursor.close()
  cnxn.close()

  return


@anvil.server.callable
def get_epub_issues(site):
  if not site:
    site = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"

  soup = BeautifulSoup(urlopen(site).read(), "html.parser")

  mwb_epub_issues = [i.get("data-issuedate") for i in soup.find_all("a", {"data-preselect": "epub"})]

  return mwb_epub_issues

@anvil.server.callable
def get_epub(site, issue):
  soup = BeautifulSoup(urlopen(site).read(), "html.parser")

  mwb_epub_issues = [
    {
      "issue": i.get("data-issuedate"), 
      "url": i.get("data-jsonurl")
    } for i in soup.find_all("a", {"data-preselect": "epub", "data-issuedate": f"{issue}"})]
  
  for i in mwb_epub_issues:
    download_json = json.load(urlopen(i.get('url')))
  
  if download_json:
    file_url = download_json.get('files').get('IN').get('EPUB')[0].get('file').get('url')
  
  if file_url:
    file_name = "mwb.epub"
    urlretrieve(file_url, file_name)
    excel_file = create_workbook()


