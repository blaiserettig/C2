#pyinstaller --noconsole --add-data "D:\C2\assets\64.ico;assets" --add-data "D:\C2\assets\256.ico;assets" --add-data "D:\C2\assets\ZedMonoNerdFont-Light.ttf;assets" --add-data "C:\Users\Blais\.wdm;wdm" --collect-all selenium --icon="D:\C2\assets\64.ico" main.py

import dearpygui.dearpygui as dpg
from dearpygui_ext.themes import create_theme_imgui_dark
from datetime import datetime
import os
import sys
import json
import threading
import sched
import time
import traceback
import dateparser
import ctypes
import keyring
from typing import Tuple

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    selenium_available = True
except ImportError as e:
    print(f"Selenium import error: {e}")
    selenium_available = False

dpg.create_context()

# --- GLOBAL VARIABLES
scheduled_event = None
is_scheduled = False
is_auth_complete = False
has_auth_failed = False
is_invalid_email = False
credentials = ["", ""]
progress_index = 0
s = sched.scheduler(time.time, time.sleep)
# ---

# --- CONFIG SETTINGS
CONFIG_FILE = "preferences.json"
CLOCK_24 = False
DEFAULT_LOCATIONS = {
    "Chrome": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ],
    "Firefox": [
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        os.path.expanduser(r"~\AppData\Local\Mozilla Firefox\firefox.exe"),
    ],
}

valid_buildings = [
    "Chapman Court",
    "Chapman Grand",
    "Davis Apartments",
    "Glass Hall",
    "Harris Apartments",
    "Henley Hall",
    "Morlan Hall",
    "Panther Village Apartments",
    "Pralle-Sodaro Hall",
    "Sandhu Residence Center",
    "The K",
]
# ---


class Logger:
    """Rediriected I/O for in-house terminal"""

    def __init__(self):
        self.buffer = []
        self.terminal_visible = False

    def write(self, message):
        self.buffer.append(message)
        print(message, end="")  # also print to console for debugging
        if (
            hasattr(self, "terminal_window")
            and self.terminal_visible
            and dpg.does_item_exist("terminal_text")
        ):
            try:
                dpg.add_text(message, parent="terminal_text")
                dpg.set_y_scroll("terminal_text", -1.0)
            except Exception as e:
                print(f"Logger error: {e}")

    def flush(self):
        pass


logger = Logger()


def load_preferences():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading preferences: {e}")
            return {"dev_mode": False}
    return {"dev_mode": False}


def save_preferences(preferences):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(preferences, f, indent=4)
    except Exception as e:
        print(f"Error saving preferences: {e}")


def save_cred_to_keyring(email: str, password: str):
    keyring.set_password("roompact", "email", email)
    keyring.set_password("roompact", "password", password)


def load_email_from_keyring() -> str:
    return keyring.get_password("roompact", "email").strip()


def load_password_from_keyring() -> str:
    return keyring.get_password("roompact", "password").strip()


def format_time(time_tuple: Tuple[int, int, str]) -> str:
    """Format a time tuple (hour, minute, am/pm) into a string"""
    hour, minute, period = time_tuple
    if len(str(minute)) < 2:
        minute_str = f"0{minute}"
    else:
        minute_str = str(minute)

    return f"{hour}:{minute_str} {period}"


def dpg_to_reg_time(date_dict: any) -> str:
    date_obj = datetime(
        year=date_dict["year"] + 1900,  # adjust by +1900 for years in dpg date pickers
        month=date_dict["month"]
        + 1,  # adjust by +1 as months in dpg date pickers are 0-indexed
        day=date_dict["month_day"],  # tm_mday
    )
    date_str = date_obj.strftime("%Y-%m-%d")
    return date_str


def set_date(tag_button, tag_popup, tag_picker):
    try:
        # Get the date picker's tm-struct-like dict
        date_dict = dpg.get_value(tag_picker)
        print(f"Date picker value: {date_dict}")

        # Update the label on the button
        dpg.configure_item(tag_button, label=dpg_to_reg_time(date_dict))
        dpg.configure_item(tag_popup, show=False)
    except Exception as e:
        print(f"Error in set_date: {e}")


def on_cancel_or_execute(string):
    """Called when the form submission process completes, either through successful submission or cancellation"""
    global is_scheduled

    is_scheduled = False
    is_auth_complete = True

    # Reset UI elements
    dpg.configure_item("date_schedule", enabled=True)
    dpg.configure_item("schedule_hour", enabled=True)
    dpg.configure_item("schedule_minute", enabled=True)
    dpg.configure_item("schedule_period", enabled=True)
    dpg.configure_item("schedule_button", label="SCHEDULE")
    dpg.configure_item("schedule_button", callback=on_schedule_button)

    if string == "execute":
        current_time = datetime.now()
        dpg.configure_item("modal_title", default_value="Success")
        dpg.configure_item(
            "modal_message",
            default_value=f"Auto-submitted as scheduled on {current_time}",
        )
        dpg.configure_item("modal_dialog", show=True)


def get_chromedriver(headless=False):
    options = Options()
    if headless:
        options.add_argument("--headless")
    else:
        options.add_argument("--start-maximized")

    try:

        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        chromedriver_path = os.path.join(
            base_path, "wdm", "chromedriver", "chromedriver.exe"
        )

        try:
            if not os.path.exists(chromedriver_path):
                # Download chromedriver dynamically if not bundled
                chromedriver_path = ChromeDriverManager().install()
        except Exception as e:
            logger.write(f"{datetime.now()}: Failed to install ChromeDriver: {e}\n")
            dpg.configure_item("modal_title", default_value="Error")
            dpg.configure_item(
                "modal_message", default_value=f"Failed to install ChromeDriver: {e}"
            )
            dpg.configure_item("modal_dialog", show=True)
            return

        logger.write(f"{datetime.now()}: Using chromedriver at: {chromedriver_path}\n")

        service = Service(ChromeDriverManager().install())

        return webdriver.Chrome(service=service, options=options)

    except Exception as e:
        logger.write(f"Error in execute: {str(e)}\n")
        dpg.configure_item("modal_title", default_value="Execution Error")
        dpg.configure_item("modal_message", default_value=f"Error: {str(e)}")
        dpg.configure_item("modal_dialog", show=True)
        return


def throw_modal_error(err: TimeoutException, driver):
    dpg.configure_item("modal_title", default_value="Execution Error")
    dpg.configure_item(
        "modal_message",
        default_value=f"Script could not execute: {type(err).__name__} – {err}",
    )
    dpg.configure_item("modal_dialog", show=True)
    driver.quit()
    return False


def login_flow(driver):


    try:  # Wait until we arrive at the Roompact login page
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "login-input"))
        )
    except TimeoutException as err:
        throw_modal_error(err, driver)

    # Continue with login process
    rp_input_email = driver.find_element(By.ID, "login-input")
    rp_input_email.send_keys(load_email_from_keyring())
    rp_input_email.send_keys(Keys.RETURN)

    try:  # See if we are at the Microsoft SSO page
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "i0116"))
        )
    except TimeoutException:
        try:  # If not, see if we are still on Roompact being prompted for the password
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "password-input"))
            )

            rp_input_password = driver.find_element(By.ID, "password-input")
            rp_input_password.send_keys(load_password_from_keyring())
            rp_input_password.send_keys(Keys.RETURN)

            try:  # If that succeeds, wait once again to be redirected to Microsoft SSO
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "i0116"))
                )
            except TimeoutException as err:
                throw_modal_error(err, driver)
        except TimeoutException as err:
            throw_modal_error(err, driver)

    msft_input_email = driver.find_element(By.ID, "i0116")
    msft_input_email.send_keys(load_email_from_keyring())
    msft_input_email.send_keys(Keys.RETURN)

    time.sleep(1)

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "i0118"))
        )
    except TimeoutException as err:
        throw_modal_error(err, driver)

    msft_input_password = driver.find_element(By.ID, "i0118")
    msft_input_password.send_keys(load_password_from_keyring())
    msft_input_password.send_keys(Keys.RETURN)

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "idSIButton9"))
        )
    except TimeoutException as err:
        throw_modal_error(err, driver)

    time.sleep(1)

    msft_ssi = driver.find_element(By.ID, "idSIButton9")
    msft_ssi.send_keys(Keys.RETURN)

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "main_search_bar"))
        )
        return True
    except TimeoutException as err:
        throw_modal_error(err, driver)


def execute(cred: list):
    """Execute the form submission using Selenium"""
    global logger

    if not selenium_available:
        dpg.configure_item("modal_title", default_value="Error")
        dpg.configure_item(
            "modal_message",
            default_value="Selenium is not available. Please install required packages.",
        )
        dpg.configure_item("modal_dialog", show=True)
        return

    driver = get_chromedriver()

    driver.get("https://roompact.com/login")

    if (not login_flow(driver)):
        return

    driver.get("https://roompact.com/forms/#/form/7r3gX9")

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "elm-datepicker--input"))
        )
    except TimeoutException as err:
        throw_modal_error(err, driver)

    general_date = dpg_to_reg_time(dpg.get_value("picker_date_general"))
    year, month, day = general_date.split("-")
    name = dpg.get_value("name_entry")
    building = dpg.get_value("building_combo")

    date_field = driver.find_element(By.CLASS_NAME, "elm-datepicker--input")
    date_field.send_keys(f"{month}/{day}/{year}")

    who = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter text for Which RAs were on duty?']",
    )
    who.send_keys(name)

    input_hall = driver.find_element(
        By.CSS_SELECTOR, "input[aria-label='Tag Buildings']"
    )
    input_hall.send_keys(building)

    # Wait for the dropdown to appear
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, "forms-subscriptions-search-result-row")
            )
        )
    except TimeoutException as err:
        throw_modal_error(err, driver)

    # Find the dropdown option and click it
    select_hall = driver.find_element(
        By.CLASS_NAME, "forms-subscriptions-search-result-row"
    )
    time.sleep(1)
    select_hall.click()

    # Walk 1
    w1_time = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter time for Time of First Community Walk']",
    )
    hour1 = dpg.get_value("walk1_hour")
    minute1 = dpg.get_value("walk1_minute")
    period1 = dpg.get_value("walk1_period")
    w1_time.send_keys(format_time((hour1, minute1, period1)))

    w1_event = driver.find_element(By.ID, "desc_resp_sub_Zdl2oy")
    w1_event.send_keys(dpg.get_value("walk1_text"))

    # Walk 2
    w2_time = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter time for Time of Second Community Walk']",
    )
    hour2 = dpg.get_value("walk2_hour")
    minute2 = dpg.get_value("walk2_minute")
    period2 = dpg.get_value("walk2_period")
    w2_time.send_keys(format_time((hour2, minute2, period2)))

    w2_event = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter text for Events of Second Community Walk']",
    )
    w2_event.send_keys(dpg.get_value("walk2_text"))

    # Walk 3 (weekend)
    w3_time = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter time for Time of Third Community Walk (weekends only)']",
    )
    # is_weekend = dpg.get_value("weekend_checkbox")

    # if is_weekend:
    hour3 = dpg.get_value("walk3_hour")
    minute3 = dpg.get_value("walk3_minute")
    period3 = dpg.get_value("walk3_period")
    w3_time.send_keys(format_time((hour3, minute3, period3)))

    w3_event = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter text for Events of Third Community Walk (weekends only)']",
    )
    # if is_weekend:
    w3_event.send_keys(dpg.get_value("walk3_text"))

    # Details
    call = driver.find_element(By.ID, "desc_resp_sub_N1dlMj")
    call.send_keys(dpg.get_value("calls_text"))

    inter = driver.find_element(By.ID, "desc_resp_sub_Zdl2o4")
    inter.send_keys(dpg.get_value("interactions_text"))

    loc = driver.find_element(
        By.CSS_SELECTOR,
        "input[aria-label='Enter text for Where did you see residents hanging out in the community?']",
    )
    loc.send_keys(dpg.get_value("areas_text"))

    incident = driver.find_element(By.ID, "desc_resp_sub_RovZrr")
    incident.send_keys(dpg.get_value("incidents_text"))

    wr = driver.find_element(By.ID, "desc_resp_sub_Vrmag2")
    wr.send_keys(dpg.get_value("workrequests_text"))

    add = driver.find_element(By.ID, "desc_resp_sub_aljG2k")
    add.send_keys(dpg.get_value("notes_text"))
    add.send_keys("\nC2 Automated Submission")

    time.sleep(2)

    try:
        submit = driver.find_element(
            By.XPATH, '//button[normalize-space()="Submit"]'
        )
        time.sleep(1)
        if dpg.get_value("dev_mode"):
            logger.write(
                f"{datetime.now()}: Submission aborted due to Developer Mode\n"
            )
        else:
            submit.click()
            logger.write(f"{datetime.now()}: Successful Submission\n")
            time.sleep(5)
    except TimeoutException as err:
        throw_modal_error(err, driver)

    driver.quit()
    on_cancel_or_execute("execute")

    return


def auth():
    """Open authentication dialog for email/password input"""
    global credentials, is_auth_complete

    dpg.configure_item("auth_dialog", show=True)

    dpg.configure_item("auth_progress", default_value="")


def attempt_login():
    global is_auth_complete
    """Try to log in with the given credentials"""
    if not selenium_available:
        return False

    global logger

    is_auth_complete = False

    driver = get_chromedriver(True)

    driver.get("https://roompact.com/login")

    return login_flow(driver)


def validate_email():
    """Validate email format"""
    return len(load_email_from_keyring()) >= 13 and load_email_from_keyring().endswith("@chapman.edu")


def show_progress():
    """Show progress indicator during authentication"""
    global progress_index, is_auth_complete, logger

    if (not is_auth_complete):
        dots = [".", "..", "...", ""]
        dpg.configure_item(
            "auth_progress",
            default_value=f"Attempting Authentication with Roompact{dots[progress_index]}",
        )
        progress_index = (progress_index + 1) % len(dots)
        # Schedule the next update in 500ms
        dpg.set_frame_callback(
            dpg.get_frame_count() + 30, show_progress
        )  # 30 frames ~ 500ms at 60 FPS
    else:
        dpg.configure_item("auth_progress", default_value="")


def submit_credentials():
    """Handle credential submission in auth dialog"""
    global credentials, is_auth_complete, progress_index, logger, is_invalid_email

    is_auth_complete = False

    save_cred_to_keyring(dpg.get_value("email_input").strip(), dpg.get_value("password_input").strip())

    # Validate email format
    if not validate_email():

        is_invalid_email = True
        dpg.configure_item("auth_dialog", show=False)
        time.sleep(0.5)

        dpg.configure_item("modal_title", default_value="Invalid Email")
        dpg.configure_item(
            "modal_message",
            default_value="The email must be 13 characters or more and end with '@chapman.edu'.",
        )
        dpg.configure_item("modal_dialog", show=True)
        return

    # Check for empty password
    if not load_password_from_keyring():
        dpg.configure_item("modal_title", default_value="Empty Password")
        dpg.configure_item("modal_message", default_value="Password cannot be empty.")
        dpg.configure_item("modal_dialog", show=True)
        return

    # Reset and start progress indicator
    progress_index = 0
    show_progress()

    # Login attempt in a separate thread
    def login_task():
        global is_auth_complete, credentials, has_auth_failed

        if attempt_login():  # authenticated, credentials are valid
            dpg.configure_item("auth_dialog", show=False)
            time.sleep(0.5)

            is_auth_complete = True
            complete_scheduling()
        else:
            has_auth_failed = True
            dpg.configure_item("auth_dialog", show=False)
            time.sleep(0.5)

            dpg.configure_item("modal_title", default_value="Authentication Failed")
            dpg.configure_item(
                "modal_message", default_value="Invalid email or password."
            )
            dpg.configure_item("modal_dialog", show=True)
            is_auth_complete = True
            dpg.configure_item("auth_progress", default_value="")

    threading.Thread(target=login_task, daemon=True).start()


def combobox_mismatch():
    """Check if selected building is valid"""
    building = dpg.get_value("building_combo")
    return building not in valid_buildings


def missing_entry_exists():
    """Check if required fields are filled"""
    if dpg.get_value("name_entry") == "":
        return True
    if dpg.get_value("building_combo") == "":
        return True
    if dpg.get_value("walk1_text") == "":
        return True
    return False


def complete_scheduling():
    """Completes the scheduling process after authentication"""
    global scheduled_event, is_scheduled

    date_str = dpg_to_reg_time(dpg.get_value("picker_date_schedule"))
    hour = dpg.get_value("schedule_hour")
    minute = dpg.get_value("schedule_minute")
    period = dpg.get_value("schedule_period")

    formatted_time = format_time((hour, minute, period))
    parse_time = dateparser.parse(f"{date_str} {formatted_time}")

    current_time = datetime.now()

    if parse_time <= current_time:
        dpg.configure_item("modal_title", default_value="Error")
        dpg.configure_item(
            "modal_message",
            default_value="Scheduled time has become earlier than current time. Scheduled time must be in the future.",
        )
        dpg.configure_item("modal_dialog", show=True)
        return

    dpg.configure_item("modal_title", default_value="Success")
    dpg.configure_item(
        "modal_message", default_value=f"Scheduled to submit on {parse_time}"
    )
    dpg.configure_item("modal_dialog", show=True)

    scheduled_event = s.enterabs(
        parse_time.timestamp(), 999, execute, argument=(credentials,)
    )
    threading.Thread(target=s.run, daemon=True).start()
    is_scheduled = True

    dpg.configure_item("date_schedule", enabled=False)
    dpg.configure_item("schedule_hour", enabled=False)
    dpg.configure_item("schedule_minute", enabled=False)
    dpg.configure_item("schedule_period", enabled=False)
    dpg.configure_item("schedule_button", label="CANCEL")
    dpg.configure_item("schedule_button", callback=on_cancel_button)


def on_schedule_button():
    """Handle schedule button click"""
    global scheduled_event, is_scheduled, logger

    if combobox_mismatch():
        dpg.configure_item("modal_title", default_value="Error")
        dpg.configure_item(
            "modal_message",
            default_value="Selected building does not exist. Ensure your entry is one of the provided drop-down options.",
        )
        dpg.configure_item("modal_dialog", show=True)
        return

    if missing_entry_exists():
        dpg.configure_item("modal_title", default_value="Error")
        dpg.configure_item(
            "modal_message", default_value="A required question is blank."
        )
        dpg.configure_item("modal_dialog", show=True)
        return

    date_str = dpg_to_reg_time(dpg.get_value("picker_date_schedule"))
    hour = dpg.get_value("schedule_hour")
    minute = dpg.get_value("schedule_minute")
    period = dpg.get_value("schedule_period")

    logger.write(
        f"{datetime.now()}: Scheduled for {date_str} at {hour}:{minute} {period}\n"
    )

    formatted_time = format_time((hour, minute, period))
    parse_time = dateparser.parse(f"{date_str} {formatted_time}")

    current_time = datetime.now()

    if parse_time <= current_time:
        dpg.configure_item("modal_title", default_value="Error")
        dpg.configure_item(
            "modal_message",
            default_value="Invalid time. Scheduled time must be in the future.",
        )
        dpg.configure_item("modal_dialog", show=True)
    else:  # good to go, start authentication
        auth()


def on_cancel_button():
    """Handle cancel button click"""
    global scheduled_event, is_scheduled

    try:
        s.cancel(scheduled_event)
        dpg.configure_item("modal_title", default_value="Cancelled")
        dpg.configure_item("modal_message", default_value="Auto-submission cancelled.")
        dpg.configure_item("modal_dialog", show=True)
    except ValueError:
        dpg.configure_item("modal_title", default_value="Error")
        dpg.configure_item(
            "modal_message", default_value="No scheduled event to cancel."
        )
        dpg.configure_item("modal_dialog", show=True)

    on_cancel_or_execute("cancel")


def dev_mode_warn():
    """Show warning when dev mode is activated"""
    dpg.configure_item("dev_mode_warning", show=dpg.get_value("dev_mode"))


def show_terminal():
    """Show the terminal window"""
    global logger

    logger.terminal_visible = True
    dpg.configure_item("terminal_window", show=True)


def hide_terminal():
    """Hide the terminal window"""
    global logger

    logger.terminal_visible = False
    dpg.configure_item("terminal_window", show=False)


# OBSOLETE. toggle_weekend() is removed to reduce complexity and clutter.

# def toggle_weekend():
# """Toggle weekend option to enable/disable Walk 3"""
# is_weekend = dpg.get_value("weekend_checkbox")

# Enable or disable Walk 3 inputs
# dpg.configure_item("walk3_text", enabled=is_weekend)
# dpg.configure_item("walk3_hour", enabled=is_weekend)
# dpg.configure_item("walk3_minute", enabled=is_weekend)
# dpg.configure_item("walk3_period", enabled=is_weekend)


def on_preferences_button():
    """Show preferences dialog"""
    dpg.configure_item("preferences_dialog", show=True)


def save_preferences_callback():
    """Save preferences and close dialog"""
    preferences = load_preferences()
    preferences["dev_mode"] = dpg.get_value("dev_mode")
    if dpg.get_value("dev_mode"):
        show_terminal()
    else:
        hide_terminal()
    save_preferences(preferences)

    dpg.configure_item("modal_title", default_value="Preferences Saved")
    dpg.configure_item(
        "modal_message", default_value="Your preferences have been saved."
    )
    dpg.configure_item("modal_dialog", show=True)

    dpg.configure_item("preferences_dialog", show=False)


def on_quit():
    """Handle quit action with confirmation"""
    dpg.configure_item("quit_dialog", show=True)


def confirm_quit():
    """Confirm and execute application quit"""
    dpg.stop_dearpygui()


def modal_callback():
    global is_invalid_email, has_auth_failed
    dpg.configure_item("modal_dialog", show=False)
    if (is_invalid_email or has_auth_failed):
        time.sleep(0.5)
        dpg.configure_item("auth_dialog", show=True)
        is_invalid_email = False
        has_auth_failed = False


def get_dpi_scale():
    dpi = ctypes.windll.user32.GetDpiForSystem()
    scale = dpi / 96  # 96 is the std DPI for 1.0 scaling
    return scale


def scale(value, dpi_scale):
    return int(value * dpi_scale)


def setup_ui():
    """Create the UI layout and widgets"""

    preferences = load_preferences()
    dpi_scale = get_dpi_scale()

    # Set up theme, kinda obsolete now that the default dpg dark theme is in use
    # but I will keep this around as maybe I want to go back someday, or have a
    # preferences option to change it
    with dpg.theme() as global_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (50, 50, 50))
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg, (30, 30, 30))
            dpg.add_theme_color(dpg.mvThemeCol_TitleBg, (35, 35, 35))
            dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, (45, 45, 45))
            dpg.add_theme_color(dpg.mvThemeCol_Button, (60, 60, 60))
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding, scale(8, dpi_scale), scale(8, dpi_scale))
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, scale(8, dpi_scale), scale(8, dpi_scale))

    # Create button styles
    with dpg.theme() as button_theme:
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button, (80, 80, 80))
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (100, 100, 100))
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (70, 70, 70))
            dpg.add_theme_color(dpg.mvThemeCol_Text, (255, 255, 255))

    # Create input field styles
    with dpg.theme() as input_theme:
        with dpg.theme_component(dpg.mvInputText):
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, (60, 60, 60))
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, (70, 70, 70))

    # Create combo box styles
    with dpg.theme() as combo_theme:
        with dpg.theme_component(dpg.mvCombo):
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, (60, 60, 60))
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, (70, 70, 70))

    # Modal dialogs
    with dpg.window(
        label="Message",
        modal=True,
        show=False,
        tag="modal_dialog",
        width=scale(400, dpi_scale),
        height=scale(150, dpi_scale),
    ):
        dpg.add_text("", tag="modal_title")
        dpg.add_separator()
        dpg.add_text("", tag="modal_message", wrap=scale(380, dpi_scale))
        dpg.add_separator()
        dpg.add_button(
            label="OK",
            width=scale(75, dpi_scale),
            callback=modal_callback)
        dpg.bind_font(nerd_mono)


    # Authentication dialog
    with dpg.window(
        label="Authentication",
        modal=True,
        show=False,
        tag="auth_dialog",
        width=scale(400, dpi_scale),
        height=scale(250, dpi_scale),
    ):
        dpg.add_text("Please enter your Roompact credentials:")
        dpg.add_text("Email:")
        dpg.add_input_text(tag="email_input", width=-1)
        dpg.add_text("Password:")
        dpg.add_input_text(tag="password_input", width=-1, password=True)
        dpg.add_text("", tag="auth_progress")
        dpg.add_button(label="Submit", width=scale(100, dpi_scale), callback=submit_credentials)

    # Preferences dialog
    with dpg.window(
        label="Preferences",
        modal=True,
        show=False,
        tag="preferences_dialog",
        width=scale(400, dpi_scale),
        height=scale(150, dpi_scale),
    ):
        dpg.add_checkbox(
            label="Developer Mode",
            tag="dev_mode",
            default_value=preferences["dev_mode"],
            callback=dev_mode_warn,
        )
        dpg.add_text(
            "Warning: Developer mode disables form submission!",
            tag="dev_mode_warning",
            show=False,
            color=(211, 65, 34),
        )
        dpg.add_separator()
        dpg.add_button(label="Save", width=scale(100, dpi_scale), callback=save_preferences_callback)

    # Quit confirmation dialog
    with dpg.window(
        label="Confirm Exit",
        modal=True,
        show=False,
        tag="quit_dialog",
        width=scale(400, dpi_scale),
        height=scale(150, dpi_scale),
    ):
        dpg.add_text("Do you want to quit? All progress will be lost.")
        dpg.add_separator()
        with dpg.group(horizontal=True):
            dpg.add_button(label="Yes", width=scale(75, dpi_scale), callback=confirm_quit)
            dpg.add_button(
                label="No",
                width=scale(75, dpi_scale),
                callback=lambda: dpg.configure_item("quit_dialog", show=False),
            )

    # Terminal window
    with dpg.window(
        label="Integrated Terminal",
        show=False,
        tag="terminal_window",
        width=scale(800, dpi_scale),
        height=scale(400, dpi_scale),
    ):
        with dpg.child_window(tag="terminal_text", width=-1, height=-1):
            pass
        logger.terminal_window = dpg.last_item()

    # Main Application Window
    with dpg.window(tag="main_window", label="C2", width=scale(1400, dpi_scale), height=scale(600, dpi_scale)):
        # Menu Bar
        with dpg.menu_bar():
            with dpg.menu(label="File"):
                dpg.add_menu_item(label="Quit", callback=on_quit)
            with dpg.menu(label="Edit"):
                dpg.add_menu_item(label="Preferences", callback=on_preferences_button)

        # Layout with splitting
        with dpg.group(horizontal=True):
            # Left Panel (~40% total screen sapce)
            with dpg.child_window(width=scale(500, dpi_scale), height=-1, tag="left_panel"):
                # General Information Section
                with dpg.collapsing_header(
                    label="General Information", default_open=True
                ):
                    with dpg.group():
                        with dpg.group(horizontal=True):
                            dpg.add_text("DATE *", color=(211, 65, 34))

                            with dpg.group(horizontal=True):
                                dpg.add_button(
                                    label=datetime.now().strftime("%Y-%m-%d"),
                                    tag="date_general",
                                    width=-1,
                                    callback=lambda: dpg.configure_item(
                                        "popup_date_general", show=True
                                    ),
                                )
                                with dpg.popup(
                                    parent="date_general",
                                    modal=True,
                                    tag="popup_date_general",
                                    mousebutton=dpg.mvMouseButton_Left,
                                ):
                                    dpg.add_date_picker(
                                        default_value={
                                            "month": datetime.now().month
                                            - 1,  # 0-based month (0-11)
                                            "month_day": datetime.now().day,  # 1-based day (1-31)
                                            "year": datetime.now().year
                                            - 1900,  # Year offset for dpg
                                        },
                                        tag="picker_date_general",
                                    )
                                    dpg.add_button(
                                        label="OK",
                                        callback=lambda: set_date(
                                            "date_general",
                                            "popup_date_general",
                                            "picker_date_general",
                                        ),
                                    )

                        with dpg.group(horizontal=True):
                            dpg.add_text("NAME *", color=(211, 65, 34))
                            dpg.add_input_text(tag="name_entry", width=-1)

                        with dpg.group(horizontal=True):
                            dpg.add_text("BUILDING *", color=(211, 65, 34))
                            dpg.add_combo(
                                valid_buildings, tag="building_combo", width=-1
                            )

                # Walks Section
                with dpg.collapsing_header(label="Walks", default_open=True):
                    # Walk 1
                    dpg.add_text("WALK 1 *", color=(211, 65, 34))
                    dpg.add_input_text(
                        tag="walk1_text", multiline=True, width=-1, height=scale(100, dpi_scale)
                    )

                    with dpg.group(horizontal=True):
                        dpg.add_text("TIME *", color=(211, 65, 34))
                        dpg.add_combo(
                            items=list(range(1, 13)),
                            tag="walk1_hour",
                            default_value=8,
                            width=scale(50, dpi_scale),
                        )
                        dpg.add_text(":")
                        dpg.add_combo(
                            items=list(range(0, 60)),
                            tag="walk1_minute",
                            default_value=0,
                            width=scale(50, dpi_scale),
                        )
                        dpg.add_combo(
                            items=["am", "pm"],
                            tag="walk1_period",
                            default_value="pm",
                            width=scale(50, dpi_scale),
                        )

                    # Walk 2
                    dpg.add_separator()
                    dpg.add_text("WALK 2")
                    dpg.add_input_text(
                        tag="walk2_text", multiline=True, width=-1, height=scale(100, dpi_scale)
                    )

                    with dpg.group(horizontal=True):
                        dpg.add_text("TIME")
                        dpg.add_combo(
                            items=list(range(1, 13)),
                            tag="walk2_hour",
                            default_value=10,
                            width=scale(50, dpi_scale),
                        )
                        dpg.add_text(":")
                        dpg.add_combo(
                            items=list(range(0, 60)),
                            tag="walk2_minute",
                            default_value=0,
                            width=scale(50, dpi_scale),
                        )
                        dpg.add_combo(
                            items=["am", "pm"],
                            tag="walk2_period",
                            default_value="pm",
                            width=scale(50, dpi_scale),
                        )

                    # --- OBSOLETE

                    # Weekend checkbox
                    # dpg.add_separator()
                    # dpg.add_checkbox(label="WEEKEND?", tag="weekend_checkbox", callback=toggle_weekend)

                    # ---

                    # Walk 3 (Weekend only)
                    dpg.add_text("WALK 3 (Weekend Only)")
                    dpg.add_input_text(
                        tag="walk3_text",
                        multiline=True,
                        width=-1,
                        height=scale(100, dpi_scale),
                        enabled=True,
                    )

                    with dpg.group(horizontal=True):
                        dpg.add_text("TIME")
                        dpg.add_combo(
                            items=list(range(1, 13)),
                            tag="walk3_hour",
                            default_value=12,
                            width=scale(50, dpi_scale),
                            enabled=True,
                        )
                        dpg.add_text(":")
                        dpg.add_combo(
                            items=list(range(0, 60)),
                            tag="walk3_minute",
                            default_value=0,
                            width=scale(50, dpi_scale),
                            enabled=True,
                        )
                        dpg.add_combo(
                            items=["am", "pm"],
                            tag="walk3_period",
                            default_value="am",
                            width=scale(50, dpi_scale),
                            enabled=True,
                        )

                # Scheduling Section
                with dpg.collapsing_header(label="Schedule", default_open=True):
                    with dpg.group(horizontal=True):
                        dpg.add_text("DATE")

                        with dpg.group(horizontal=True):
                            # Button with current date as label
                            dpg.add_button(
                                label=datetime.now().strftime("%Y-%m-%d"),
                                tag="date_schedule",
                                width=scale(100, dpi_scale),
                                callback=lambda: dpg.configure_item(
                                    "popup_date_schedule", show=True
                                ),
                            )
                            with dpg.popup(
                                parent="date_schedule",
                                modal=True,
                                tag="popup_date_schedule",
                                mousebutton=dpg.mvMouseButton_Left,
                            ):
                                dpg.add_date_picker(
                                    default_value={
                                        "month": datetime.now().month
                                        - 1,  # 0-based month (0-11)
                                        "month_day": datetime.now().day,  # 1-based day (1-31)
                                        "year": datetime.now().year
                                        - 1900,  # Year offset for dpg
                                    },
                                    tag="picker_date_schedule",
                                )
                                dpg.add_button(
                                    label="OK",
                                    callback=lambda: set_date(
                                        "date_schedule",
                                        "popup_date_schedule",
                                        "picker_date_schedule",
                                    ),
                                )

                        dpg.add_text("TIME")
                        dpg.add_combo(
                            items=list(range(1, 13)),
                            tag="schedule_hour",
                            default_value=8,
                            width=scale(50, dpi_scale),
                        )
                        dpg.add_text(":")
                        dpg.add_combo(
                            items=list(range(0, 60)),
                            tag="schedule_minute",
                            default_value=0,
                            width=scale(50, dpi_scale),
                        )
                        dpg.add_combo(
                            items=["am", "pm"],
                            tag="schedule_period",
                            default_value="am",
                            width=scale(50, dpi_scale),
                        )

                        dpg.add_button(
                            label="SCHEDULE",
                            tag="schedule_button",
                            callback=on_schedule_button,
                            width=scale(100, dpi_scale),
                        )

            # Right Panel (~60% total screen space)
            with dpg.child_window(width=scale(800, dpi_scale), height=-1, tag="right_panel"):
                with dpg.collapsing_header(label="Details", default_open=True):
                    dpg.add_text("CALLS RECEIVED")
                    dpg.add_input_text(
                        tag="calls_text", multiline=True, width=-1, height=scale(100, dpi_scale)
                    )
                    dpg.add_separator()

                    dpg.add_text("RESIDENT INTERACTIONS")
                    dpg.add_input_text(
                        tag="interactions_text", multiline=True, width=-1, height=scale(75, dpi_scale)
                    )
                    dpg.add_separator()

                    dpg.add_text("AREAS RESIDENTS WERE SEEN")
                    dpg.add_input_text(
                        tag="areas_text", multiline=True, width=-1, height=scale(75, dpi_scale)
                    )
                    dpg.add_separator()

                    dpg.add_text("INCIDENTS")
                    dpg.add_input_text(
                        tag="incidents_text", multiline=True, width=-1, height=scale(75, dpi_scale)
                    )
                    dpg.add_separator()

                    dpg.add_text("WORK REQUESTS SUBMITTED")
                    dpg.add_input_text(
                        tag="workrequests_text", multiline=True, width=-1, height=scale(75, dpi_scale)
                    )
                    dpg.add_separator()

                    dpg.add_text("ADDITIONAL NOTES")
                    dpg.add_input_text(
                        tag="notes_text", multiline=True, width=-1, height=scale(75, dpi_scale)
                    )

    dpg.bind_theme(global_theme)
    dpg.bind_item_theme("schedule_button", button_theme)

    if preferences.get("dev_mode", False):
        show_terminal()

    def update_layout(sender, app_data):
        width = dpg.get_viewport_width()
        dpg.configure_item("left_panel", width=int(width * 0.4))
        dpg.configure_item("right_panel", width=int(width * 0.6))

    dpg.set_viewport_resize_callback(update_layout)

    dpg.set_primary_window("main_window", True)

    dev_mode_warn()


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if getattr(sys, 'frozen', False): 
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    try:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except:
            ctypes.windll.user32.SetProcessDPIAware()

        screensize = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)

        w = int(screensize[0] / 1.2)
        h = int(screensize[1] / 1.35)

        dpg.create_context()
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        small_icon_path = resource_path(os.path.join('assets', '64.ico'))
        large_icon_path = resource_path(os.path.join('assets', '256.ico'))
        font_path = resource_path(os.path.join('assets', 'ZedMonoNerdFont-Light.ttf'))

        dpi_scale = get_dpi_scale()
        font_scale = int(w / 106)
        with dpg.font_registry():
            nerd_mono = dpg.add_font(font_path, font_scale * dpi_scale)

        dpg.set_global_font_scale(1.0 / dpi_scale)

        setup_ui()

        dpg.create_viewport(title="C2", width=w, height=h, small_icon=small_icon_path, large_icon=large_icon_path)
        #dpg.configure_viewport(0, dpi_aware=True)

        # --- Theming
        dark_theme = create_theme_imgui_dark()
        dpg.bind_theme(dark_theme)
        # ---

        dpg.setup_dearpygui()

        dpg.set_viewport_resizable(True)
        #dpg.maximize_viewport()

        dpg.show_viewport()
        dpg.start_dearpygui()
        dpg.destroy_context()

    except Exception as e:
        print(f"Application failed to start: {e}")
        traceback.print_exc()
