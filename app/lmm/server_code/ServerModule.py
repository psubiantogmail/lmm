import anvil.server
from urllib.request import urlopen


@anvil.server.callable
def get_epub(site):
  if not site:
    site = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"

  from bs4 import BeautifulSoup
  from selenium.webdriver.support.ui import WebDriverWait as wait
  soup = BeautifulSoup(urlopen(site).read(), "html.parser")

  mwb_epub_issues = [i.get("data-issuedate") for i in soup.find_all("a", {"data-preselect": "epub"})]

  return mwb_epub_issues

