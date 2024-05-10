from ._anvil_designer import BaseTemplate
from anvil import *
import anvil.server
import anvil.google.auth, anvil.google.drive
from anvil.google.drive import app_files
import anvil.users
import stripe.checkout
import anvil.tables as tables
import anvil.tables.query as q
from anvil.tables import app_tables
from ..Home import Home
from ..HHAapp import HHAapp

class Base(BaseTemplate):
  def __init__(self, **properties):
    # Set Form properties and Data Bindings.
    self.init_components(**properties)
    self.change_sign_in_text()
    self.go_to_home()
    
  def go_to_home(self):
    self.content_panel.clear()
    self.content_panel.add_component(Home())

  def change_sign_in_text(self):
    user = anvil.users.get_user()
    if user: 
      email = user["email"]
      self.sign_in.text = email
    else:
      self.sign_in.text = 'Sign In'
  
  def launch_hha_app_click(self, **event_args):
    """This method is called when the link is clicked"""
    user = anvil.users.get_user()
    if user:
      hha_app_instance = HHAapp()  # Create an instance of HHAapp
      # Remove hha_app_panel from its parent container in HHAapp
      hha_app_instance.hha_app_panel.remove_from_parent()
      # Clear existing content in content_panel (if necessary)
      self.content_panel.clear()
      # Add hha_app_panel to the content_panel
      self.content_panel.add_component(hha_app_instance.hha_app_panel)
    else: 
      alert("Please sign in to access the app")
  
  def sign_in_click(self, **event_args):
    """This method is called when the link is clicked"""    
    user = anvil.users.get_user()
    if user: 
      logout = confirm("Would you like to logout?")
      if logout:
        anvil.users.logout()
        self.go_to_home()  
    else:
      anvil.users.login_with_form()
    self.change_sign_in_text()

  def title_click(self, **event_args):
    """This method is called when the link is clicked"""
    self.go_to_home()


