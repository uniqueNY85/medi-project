from ._anvil_designer import HHAappTemplate
from anvil import *
import anvil.google.auth, anvil.google.drive
from anvil.google.drive import app_files
import anvil.tables as tables
import anvil.tables.query as q
from anvil.tables import app_tables
import anvil.users
import stripe.checkout
import anvil.server

class HHAapp(HHAappTemplate):
  def __init__(self, **properties):
    self.HHA_data=[]
    # Set Form properties and Data Bindings.
    self.init_components(**properties)
    self.drop_down_year.set_event_handler('change', self.drop_down_year_change)
    
    # Any code you write above will run before the form opens.

  def log_out_click(self, **event_args):
    """This method is called when the button is clicked"""
    anvil.users.logout()
    anvil.open_form('MainPage')
    #Function to be called when the "Search" button is clicked or Enter key is pressed
  def perform_search(self):
    try:
      self.label_status.text = ''
      user_input = self.HHA_search_bar.text
      result = anvil.server.call('get_user_input_and_search', user_input)
      if result['status'] == 'No matches':
        self.label_status.text = "No matches found. Please try again."
        self.label_results_count.text = ""
        self.data_grid_1.visible = False #Hide the grid
        self.HHA_select_bar.visible = False  # Hide the selection bar
        self.selection_button.visible = False  # Hide the selection button
      else:
        self.HHA_data = result['results']
        self.data_grid_1.visible = True #show the grid
        self.repeating_panel_1.items = self.HHA_data
        self.repeating_panel_1.visible = True
        self.label_results_count.text = f"Displaying {min(50, len(result['results']))} of {len(result['results'])} results"        
        self.HHA_select_bar.visible = True  # Show the selection bar
        self.selection_button.visible = True  # Show the selection button
    except Exception as e:
      print("An error occurred:", e)
      return

  # handles user selecting a provider
  def selection_button_click(self, **event_args):
  # Get the user input from HHA_select_bar
    user_selection_id = self.HHA_select_bar.text
    try:
    # Call the server function to get the selected data
      result = anvil.server.call('get_selected_data', self.HHA_data, user_selection_id)    
      if result['status'] == 'Success':
      # Prepare the display message
        selected_data = result['selected_data']                                               
        display_message = f"Confirm your selection:\nProvider Number: {selected_data['provider_number']}\nHHA Name: {selected_data['hha_name']}\nAddress: {selected_data['address']}"
      # Show the data in a confirmation dialog
        if anvil.confirm(display_message):
        # User confirmed, proceed to next step
          self.provider_number = selected_data['provider_number']
          self.hha_name = selected_data['hha_name']
          self.address = selected_data['address']
          self.send_data_to_server(self.provider_number) 
        else:
        # User cancelled, do nothing
          pass
      else:
      # Show an error message
        anvil.alert(result['message'])
    except Exception as e:
      anvil.alert(f"An error occurred: {str(e)}")

  def send_data_to_server(self, provider_number):
    try:
    # Call the server function to get available years
      available_years = anvil.server.call('get_available_years', provider_number)
      if available_years['status'] == 'success':
      # Populate the dropdown with available years 
        self.drop_down_year.items = [("Please select a fiscal year", None)] + [(str(year), year) for year in available_years['available_years']]
        self.year_identifiers = available_years['identifiers']#storing the identifiers
        self.year_second_identifiers = available_years['second_identifiers']#storing potential second identifiers
        self.additional_data = available_years['additional_data']  # Storing additional data for the first identifier
        self.additional_second_data = available_years['additional_second_data']  # Storing additional data for the second identifier

        self.drop_down_year.visible = True  # Make the dropdown visible 
        self.drop_down_instr.visible = True 
        self.current_selection.visible = True
        self.current_selection.text = f"Current Provider Selected: {self.provider_number}, {self.hha_name}, {self.address}"
      else:
        anvil.alert("No available years found for the selected provider.")   
    except Exception as e:
      anvil.alert(f"An error occurred: {str(e)}")
  
    # Called when the "Search" button is clicked
  def search_button_click(self, **event_args):
    self.perform_search()
        
    # Called when a key is pressed in the HHA_search_bar
  def HHA_search_bar_pressed_enter(self, **event_args):
    self.perform_search()
  
  def HHA_select_bar_pressed_enter(self, **event_args):
    """This method is called when the user presses Enter in this text box"""
    self.selection_button_click()

  # Triggered when the user selects a year from the dropdown
  def drop_down_year_change(self, **event_args):
    selected_year = self.drop_down_year.selected_value
    identifier = self.year_identifiers.get(selected_year)
    # Extracting additional data for the first identifier
    first_additional = self.additional_data.get(selected_year)
    first_rpt_status_cd = first_additional['rpt_status_code']
    first_beg_rpt_period = first_additional['beg_rpt_period']
    first_end_rpt_period = first_additional['end_rpt_period']   
    second_identifier = self.year_second_identifiers.get(selected_year)
    # Extracting additional data for the second identifier, if it exists
    second_additional = self.additional_second_data.get(selected_year, {})
    second_rpt_status_cd = second_additional.get('rpt_status_code', None)
    second_beg_rpt_period = second_additional.get('beg_rpt_period', None)
    second_end_rpt_period = second_additional.get('end_rpt_period', None)
    user_email = anvil.users.get_user()['email']  #fetching the email of the currently logged-in user
    self.session_id=anvil.server.call('update_session_info', identifier, first_rpt_status_cd, second_identifier, first_beg_rpt_period, first_end_rpt_period, second_rpt_status_cd, second_beg_rpt_period, second_end_rpt_period, selected_year, self.provider_number, self.hha_name, self.address, user_email) 
    #Consider adding a confirmation here
    if identifier:
      try:
        # Get the Stripe token
        token, info = stripe.checkout.get_token(amount=5000, currency="usd", description=f"Report for {selected_year}")        
        # Initiate Stripe payment on the server
        charge_response = anvil.server.call('charge_user', token, info['email'], selected_year, identifier, self.session_id, second_identifier)      
        # Check the payment status
        if charge_response['payment_status'] == 'succeeded':
            anvil.alert("Payment Successful. A receipt will be emailed to you.")
            # Check if report generation also succeeded
            if charge_response['report_gen_status'] == 'succeeded':
                anvil.alert("Report generated successfully.")
            elif charge_response['report_gen_status'] == 'report_error':
                anvil.alert("An error occurred while generating the report. Please try again.")
        elif charge_response['payment_status'] == 'charge_failed':
            anvil.alert("Payment failed. Please try again.")
        elif charge_response['payment_status'] == 'charge_error':
            anvil.alert("An error occurred during the payment process. Please try again.")
        else:
            anvil.alert("An unexpected error occurred. Please try again.")
      except Exception as e:
        if str(e) == "Stripe checkout cancelled":
            self.reset_form()
            anvil.alert("Payment was cancelled.")
        else:
            anvil.alert(f"An error occurred: {str(e)}")

  
   
  def reset_form(self):   
    # Reset variables
    self.provider_number = None
    self.hha_name = None
    self.address = None
    
    # Reset UI components
    self.current_selection.text = ''
    self.drop_down_year.items = [("Please select a fiscal year", None)]
    self.HHA_select_bar.text = None  
    self.HHA_search_bar.text = None

    # Hide optional components
    self.drop_down_year.visible = False
    self.drop_down_instr.visible = False
    self.current_selection.visible = False
    self.selection_button.visible = False 
    self.HHA_select_bar.visible = False 
    self.repeating_panel_1.visible = False 
    self.data_grid_1.visible = False
    self.label_results_count.visible = False 

  def reset_form_button_click(self, **event_args):
    """This method is called when the button is clicked"""
    self.reset_form()
    


  




 


  




 





   
    




