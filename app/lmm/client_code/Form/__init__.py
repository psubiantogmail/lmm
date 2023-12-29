from ._anvil_designer import FormTemplate
from anvil import *
import anvil.server

class Form(FormTemplate):
  def __init__(self, **properties):
    # Set Form properties and Data Bindings.
    self.init_components(**properties)

    # Any code you write here will run before the form opens.
    language_list = ['Indonesian', 'English']
    self.drop_down_language.items = language_list
    self.link_epub.url = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"
    self.link_epub.text = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"


  def drop_down_language_change(self, **event_args):
    """This method is called when an item is selected"""
    if self.drop_down_language.selected_value == 'Indonesian':
      self.link_epub.url = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"
      self.link_epub.text = "https://www.jw.org/id/perpustakaan/jw-lembar-pelajaran/"
    else:
      self.link_epub.url = "https://www.jw.org/en/library/jw-meeting-workbook/"
      self.link_epub.text = "https://www.jw.org/en/library/jw-meeting-workbook/"

  def link_jworg_click(self, **event_args):
    """This method is called when the link is clicked"""
    pass

  def button_download_page_click(self, **event_args):
    """This method is called when the button is clicked"""
    issues = anvil.server.call('get_epub_issues', self.link_epub.text)
    self.drop_down_issues.items = issues

  def button_process_click(self, **event_args):
    """This method is called when the button is clicked"""
    excel_file = anvil.server.call('get_epub', self.link_epub.text, self.drop_down_issues.selected_value)
    anvil.media.download(excel_file)
