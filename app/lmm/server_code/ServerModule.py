import anvil.server


@anvil.server.callable
def get_epub(site):
  if not site:
    site = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"

  from bs4 import BeautifulSoup
  from selenium.webdriver.support.ui import WebDriverWait as wait
  soup = BeautifulSoup(urlopen(self.link_epub.url).read(), "html.parser")

    
  

  
  return "2. No Error"

