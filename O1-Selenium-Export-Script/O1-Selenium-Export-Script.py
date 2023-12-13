import json
import logging
import os
import pickle
import sys
import time
import warnings
from os import environ as env
from zipfile import ZipFile

import colorlog
from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    ElementNotVisibleException,
    NoSuchAttributeException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.remote_connection import LOGGER
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

warnings.filterwarnings("ignore")

LOGGER.setLevel(logging.WARNING)

#   Load environment variables  #
load_dotenv()
os.environ["WDM_LOG"] = "false"
logging.getLogger("WDM").setLevel(logging.NOTSET)

#   GLOBALS CONSTANTS/VARIABLES   #
COOKIES_FILE = env.get("COOKIES_FILE")
PROJ_URLS_FILE = env.get("PROJ_URLS_FILE")
DELAY = 1  # delay in seconds
SHORT_DELAY = 0.5  # short delay
ELEMENT_WAIT_TIME = 5  # wait time in sec for getting elements
MENU_WAIT_TIME = 10  # menu wait time in sec
DEFAULT_WEBDRIVER_WAIT_TIME = 30  # default wait time in sec for webdriver
SHORT_WEBDRIVER_WAIT_TIME = 15  # short wait time in seconds for webdriver
REFRESH_WAIT_TIME = 10  # refresh wait time in sec
MAX_RETRY = 5  # max number of retries for export
# Build absolute path to output folder in current directory
OUTPUT_PATH = os.path.join(os.getcwd(), env.get("OUTPUT_DIR_NAME"))
# Build absolute path to temp download folder in current directory
TEMP_OUTPUT_PATH = os.path.join(OUTPUT_PATH, env.get("TEMP_DIR_NAME"))
# Build export log directory path
EXPORT_LOG_DIR = os.path.join(os.getcwd(), env.get("EXPORT_LOG_DIR_NAME"))
# Build export log messages file for skipped/failed downloads path
# Add the current date and time of script start in the file name
export_log_file = "{}_{}".format(
    time.strftime("%d-%m-%Y_%H-%M-%S", time.localtime()), env.get("EXPORT_LOG_NAME")
)
EXPORT_LOG_FILE = os.path.join(EXPORT_LOG_DIR, export_log_file)
# global EXPORT_LOG dict object
EXPORT_LOG = {}

#   CREDENTIALS  #
USERNAME = env.get("O1_email")
PASSWORD = env.get("O1_password")

#   URLS   #
BENTLEY_LOGIN_URL = env.get("BENTLEY_LOGIN_URL")
ALL_SYNCHRO_URL = env.get("ALL_SYNCHRO_URL")
PROJ_URL_PLACEHOLDER = env.get("PROJ_URL_PLACEHOLDER")


#   Selenium webdriver options  #
def get_options(dev=False):
    options = webdriver.ChromeOptions()
    # run in headless mode if dev is False, else run foreground mode
    if dev is False:
        options.add_argument("--headless")
    # overcome limited resource problems
    options.add_argument("--disable-dev-shm-usage")
    # set window size
    options.add_argument("--window-size=1920,1200")
    # avoid granting access to camera and microphone
    options.add_argument("--use-fake-ui-for-media-stream")
    # disable notifications
    options.add_argument("--disable-notifications")
    # disable popup blocking
    options.add_argument("--disable-popup-blocking")
    # disable infobars
    options.add_argument("--disable-infobars")
    # disable extensions
    options.add_argument("--disable-extensions")
    # disable crash reporting
    options.add_argument("--disable-crash-reporter")
    # disable crash uploading list
    options.add_argument("--disable-crash-upload")
    # disable oopr debug crash dump
    options.add_argument("--disable-oopr-debug-crash-dump")
    # disable background timer throttling
    options.add_argument("--disable-background-timer-throttling")
    # disable backgrounding occluded windows
    options.add_argument("--disable-backgrounding-occluded-windows")
    # disable renderer backgrounding
    options.add_argument("--disable-renderer-backgrounding")
    # disable logging switch
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    # add prefs to change default download directory
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": TEMP_OUTPUT_PATH,
            "profile.default_content_settings.popups": 0,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        },
    )
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")
    return options


# Get chrome webdriver options
if env.get("dev_mode").lower() == "true":
    OPTIONS = get_options(dev=True)
else:
    OPTIONS = get_options()


#   Custom Exceptions   #
class CustomException(Exception):
    "Base class for other exceptions"

    def __init__(self, message):
        super().__init__(message)


class NavigationError(CustomException):
    "Raised when there is an error navigating to a page"

    def __init__(self, message):
        super().__init__(message)


class CookiesInvalidError(CustomException):
    "Raised when cookies are invalid"

    def __init__(self, message):
        super().__init__(message)


class CookiesFileNotFoundError(CustomException):
    "Raised when cookies file is not found"

    def __init__(self, message):
        super().__init__(message)


class LoginError(CustomException):
    "Raised when login fails"

    def __init__(self, message):
        super().__init__(message)


class ProjectUrlsFileNotFoundError(CustomException):
    "Raised when project urls file is not found"

    def __init__(self, message):
        super().__init__(message)


# All Selenium Exceptions to catch
SELENIUM_ERROR = (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementNotVisibleException,
)
# Ignored exceptions for WebDriverWait
IGNORED_EXCEPTIONS = (
    NoSuchElementException,
    StaleElementReferenceException,
    ElementNotVisibleException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoSuchAttributeException,
)
# All Exceptions Classes
ALL_ERRORS = SELENIUM_ERROR + (Exception, OSError)

# Get xpaths and css selectors for elements
table_row = env.get("table_row")
table_row_title = env.get("table_row_title")
three_dots = env.get("three_dots")
export_to_pdf = env.get("export_to_pdf")
tippy_box = env.get("tippy_box")
archive_export_pdf = env.get("archive_export_pdf")
select_all_box = env.get("select_all_box")
export_btn_xpath = env.get("export_btn_xpath")
archive_export_excel = env.get("archive_export_excel")
export_excel = env.get("export_excel")
total_forms_item = env.get("total_forms_item")
table_rows = env.get("table_rows")
page_item = env.get("page_item")
active_page_item = env.get("active_page_item")
next_page_item = env.get("next_page_item")
export_modal = env.get("export_modal")
checkboxes_xpath_dict = {
    "Comments": env.get("comments_box"),
    "Audit trail": env.get("audit_trail_box"),
    "Images": env.get("images_box"),
    "Export attachments": env.get("export_attachments_box"),
}


# Define a custom event HANDLER for file changes in the downloads folder
class DownloadHandler(FileSystemEventHandler):
    def __init__(self, expected_file_name):
        super().__init__()
        self.expected_file_name = expected_file_name
        self.downloaded_file_name = None
        self.download_completed = False

    def on_created(self, event):
        # when download is started, on created will be called
        # temp file with .crdownload is created on chrome
        LOG.debug("File Created: %s", event.src_path)

    def on_modified(self, event):
        # when download is completed, on modified will be called
        # LOG.debug("File modified: %s", event.src_path)
        # check either if file name is same or similar by extension
        # account for file name with timestamp by checking if file name
        # ends with the same extension
        modified_file_name = os.path.basename(event.src_path)
        file_name_same = modified_file_name == self.expected_file_name
        file_name_similar = modified_file_name.endswith(
            os.path.splitext(self.expected_file_name)[1]
        )
        file_name_match = file_name_same or file_name_similar
        ends_with_crdownload = modified_file_name.endswith(".crdownload")
        unexpected_file = not file_name_match and not ends_with_crdownload
        if file_name_match and self.download_completed is False:
            LOG.debug(f"Downloaded: {event.src_path}")
            LOG.info("Downloaded: {}".format(modified_file_name))
            self.download_completed = True
            # if file name is similar and not same, set downloaded file name
            if file_name_similar and file_name_same is False:
                self.downloaded_file_name = modified_file_name
        elif unexpected_file and self.download_completed is False:
            # edge case for when file name is not same or similar
            LOG.debug(f"Downloaded: {event.src_path}")
            LOG.info("Downloaded: {}".format(modified_file_name))
            self.download_completed = True
            self.downloaded_file_name = modified_file_name


HANDLER = colorlog.StreamHandler()
FORMATTER = colorlog.ColoredFormatter(
    env.get("log_format"),
    datefmt="%d-%m-%Y %H:%M:%S",
    reset=True,
    log_colors={
        "DEBUG": "cyan",
        "INFO": "green",
        "WARNING": "yellow",
        "ERROR": "red",
        "CRITICAL": "red,bg_white",
    },
    secondary_log_colors={},
    style="%",
)
FHFORMATTER = formatter = logging.Formatter(
    env.get("file_log_format"), datefmt="%a, %d %b %Y %H:%M:%S"
)
HANDLER.setFormatter(FORMATTER)
FH = logging.FileHandler("{}.log".format(os.path.basename(__file__).replace(".py", "")))
FH.setFormatter(FHFORMATTER)
LOG = colorlog.getLogger(__name__)
LOG.addHandler(FH)
LOG.addHandler(HANDLER)
# set log level to DEBUG / INFO / WARNING / ERROR / CRITICAL
if env.get("dev_mode").lower() == "true":
    # set log level to debug if dev mode true
    LOG.setLevel(colorlog.DEBUG)
else:
    LOG.setLevel(colorlog.INFO)


def get_total_forms(total_forms):
    """Returns total number of forms rounded up to nearest whole number"""
    # 1 page = 25 forms
    return -(-int(total_forms) // 25)


def login_optimus(browser):
    """
    Login to Bentley Webapp

    Args:
        browser (WebDriver): Selenium webdriver object

    Raises:
        TimeoutException: TimeoutException when element is not found
        LoginError: LoginError when login fails
    """
    LOG.info("Logging In...")
    browser.get(BENTLEY_LOGIN_URL)
    try:
        # set wait for default webdriver wait time
        wait = WebDriverWait(
            browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        )
        # get email field
        email = wait.until(
            EC.presence_of_element_located((By.ID, env.get("email_field")))
        )
        # use action to send keys to email field
        action = ActionChains(browser)
        action.send_keys_to_element(email, USERNAME).perform()
        # set wait for element wait time
        wait = WebDriverWait(
            browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        )
        # get next btn
        next_btn = wait.until(
            EC.element_to_be_clickable((By.ID, env.get("sign_in_btn")))
        )
        action.click(next_btn).perform()
        # wait for password field to be present
        password = wait.until(
            EC.presence_of_element_located((By.ID, env.get("pw_field")))
        )
        # get password field
        action.send_keys_to_element(password, PASSWORD).perform()
        # get sign in btn
        sign_in_btn = wait.until(
            EC.element_to_be_clickable((By.ID, env.get("sign_in_btn")))
        )
        action.click(sign_in_btn).perform()
        # wait for text in div to be present for up to 30s till user enters 2fa
        wait = WebDriverWait(
            browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        )
        wait.until(EC.presence_of_element_located((By.XPATH, env.get("pingid_div"))))
        LOG.critical("PingID Authentication Required!")
        # wait until 'Change Password' Button is present, user is logged in
        wait.until(EC.presence_of_element_located((By.ID, env.get("change_pw_btn"))))
        LOG.info("PingID Authenticated!")
        # if login was successful, navigate to project page and save cookies
        browser.get(ALL_SYNCHRO_URL)
        # wait till 'All projects' text is present
        wait.until(EC.presence_of_element_located((By.XPATH, env.get("all_proj_div"))))
        LOG.info("Login Successful!")
        # save cookies to file
        save_cookies(browser)
        return True
    except TimeoutException:
        raise LoginError("Login Failed!")


def navigate_to_page(
    browser,
    url=PROJ_URL_PLACEHOLDER,
    wait_condition="//h2[contains(text(), 'My work')]",
):
    """_summary_

    Args:
        browser (WebDriver): Selenium webdriver object
        url (str, optional): Target URL. Defaults to PROJ_URL_PLACEHOLDER.
        wait_condition (str, optional): Element to wait for using webdriver.
        Defaults to "//h2[contains(text(), 'My work')]".
        Make sure this is an element that is present on the page you are
        navigating to. If not TimeOut will be raised.
    Raises:
        TimeoutException: TimeoutException when element is not found
        NavigationError: NavigationError when navigation fails
    """
    LOG.debug("GET:> {}".format(url))
    browser.get(url)
    # if browser title is 'Sign in to your account', raise error and relogin
    if browser.title == "Choose an Account":
        LOG.info("Cookies conflict, restarting script...")
        raise CookiesInvalidError("Cookies Conflict!")
    else:
        # wait for a condition to be present for up to 30s
        wait = WebDriverWait(
            browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        )
        # wait till 'wait_condition' is present
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, wait_condition)))
            LOG.debug("GET:> {} OK!".format(url))
        except TimeoutException:
            msg = "Wait Timeout! \nTarget: {} \nActual: {}\nCondition: {}".format(
                url, browser.current_url, wait_condition
            )
            raise NavigationError(msg)


def get_proj_form_types(browser):
    """Get list of form types in work project"""
    LOG.debug("Getting list of form types in work project...")
    # find and click work tab
    click_work_tab(browser)
    # get form types element list
    form_types_elem_list = get_form_types_elem_list(browser)
    # generate the form types list
    form_types_list = [form_type.text for form_type in form_types_elem_list]
    return form_types_list


def click_work_tab(browser):
    """Click work tab in nav bar"""
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    action = ActionChains(browser)
    # get the work tab in li element with and title='Work' and click it
    work_tab = wait.until(EC.element_to_be_clickable((By.XPATH, env.get("work_tab"))))
    action.click(work_tab).perform()
    # wait till work projects nav bar container is present
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, env.get("work_proj"))))


def get_form_types_elem_list(browser):
    """Get and return form types element list"""
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    form_nav_bar = env.get("form_nav_bar")
    form_types = env.get("form_types")
    # wait till work project form items in list tab are present
    form_types_nav_bar = wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, form_nav_bar))
    )
    form_types_elem_list = form_types_nav_bar.find_elements(By.CLASS_NAME, form_types)
    return form_types_elem_list


def refresh_page_form_types(browser):
    """Refresh page and wait for form types nav bar to be present"""
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    form_nav_bar = env.get("form_nav_bar")
    try:
        LOG.debug("Refreshing page...")
        browser.refresh()
        # wait for page to refresh
        time.sleep(REFRESH_WAIT_TIME)
        # click work tab
        click_work_tab(browser)
        # check for presence of form nav bar
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, form_nav_bar)))
        LOG.debug("Page refreshed!")
    except (TimeoutException, NoSuchElementException):
        LOG.warning("Page refresh failed!")


def export_forms_data(browser, form_types_list, proj_folder, proj_name):
    """
    Exports all forms dada in project to excel and pdf

    Args:
        browser (WebDriver): Selenium webdriver object
        form_types_list (list): form types list
        proj_folder (str): project folder path string
        proj_name (str): project name string
    """
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    action = ActionChains(browser)
    # strip the form type of any whitespaces
    form_types_list = [form_type.strip() for form_type in form_types_list]
    # loop through all the list items
    for form_type in form_types_list:
        # ignore 'My Work' form type
        if form_type == "My work":
            continue
        retry_count = 0
        form_found = False
        while form_found is False and retry_count <= MAX_RETRY:
            LOG.info("Form: {}".format(form_type))
            try:
                # refresh the form_types_elem_list to get the latest elements
                form_types_elem_list = get_form_types_elem_list(browser)
                cur_form_type = form_types_elem_list[form_types_list.index(form_type)]
                # click on the form type
                action.click(cur_form_type).perform()
                # wait till form type page is loaded
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//h3[contains(text(), '{}')]".format(form_type))
                    )
                )
                # set form_found to True
                form_found = True
            except (TimeoutException, NoSuchElementException):
                LOG.warning(
                    "{} Form Not Found! Retrying {}/{}".format(
                        form_type, retry_count + 1, MAX_RETRY
                    )
                )
                # refresh page if get form types nav bar is timed out
                refresh_page_form_types(browser)
                retry_count += 1
        if retry_count >= MAX_RETRY:
            LOG.warning(
                "Skipping... {}, Form Not Found After {} Retry!".format(
                    form_type, MAX_RETRY
                )
            )
            setup_export_log(proj_name, form_type)
            # update export log with form not found error
            EXPORT_LOG[proj_name]["forms"][form_type][
                "forms_export_error"
            ] = "Form Not Found!"
            continue
        else:
            # check if form type is archived and/or empty
            archive = check_is_archived(browser, form_type)
            is_empty = check_is_empty(browser)
            if is_empty is False:
                # if no empty container, get the forms
                setup_export_log(proj_name, form_type)
                # create project folder for form type
                form_type_folder = os.path.join(proj_folder, form_type)
                if not os.path.exists(form_type_folder):
                    os.makedirs(form_type_folder)
                # export forms data to excel
                export_forms_excel(
                    browser, form_type, form_type_folder, proj_name, archive
                )
                # pdf_form_type_folder = os.path.join(form_type_folder, "PDFs")
                pdf_form_type_folder = os.path.join(form_type_folder)
                # create pdf folder in form type folder
                if not os.path.exists(pdf_form_type_folder):
                    os.makedirs(pdf_form_type_folder)
                # export forms data to pdf
                export_forms_pdf(
                    browser, form_type, pdf_form_type_folder, proj_name, archive
                )
            else:
                LOG.info("No forms in {}".format(form_type))
                continue
    return True


def check_is_archived(browser, form_type):
    """Check if form type is archived"""
    archived_container = env.get("archived_container")
    try:
        is_archived = WebDriverWait(
            browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        ).until(EC.presence_of_all_elements_located((By.XPATH, archived_container)))
        # if archived container is present, set archive to true
        if is_archived:
            LOG.info("{} is Archived".format(form_type))
            return True
    except TimeoutException:
        return False


def check_is_empty(browser):
    """Check if form type is empty"""
    empty_container = env.get("empty_container")
    try:
        empty_container = WebDriverWait(
            browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        ).until(EC.presence_of_all_elements_located((By.XPATH, empty_container)))
        # if empty container is present, skip to next form type
        if empty_container:
            return True
    except TimeoutException:
        return False


def refresh_page_export(browser, wait_element):
    """Refresh page and wait for page element to be present"""
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    try:
        LOG.debug("Refreshing page...")
        browser.refresh()
        # buffer time for page to refresh and load
        time.sleep(REFRESH_WAIT_TIME)
        # wait till element is present
        wait.until(EC.presence_of_all_elements_located((By.XPATH, wait_element)))
        LOG.debug("Page refreshed!")
    except (TimeoutException, NoSuchElementException):
        LOG.warning("Page refresh failed!")


def export_forms_excel(browser, form_type, target_folder, proj_name, archive=False):
    """
    Export all forms data in form type to excel

    Args:
        browser (Webdriver): Selenium webdriver object
        form_type (object): form type object
        target_folder (string): target folder to move the exported excel to
    """
    excel_file_name = "{}.xlsx".format(form_type)
    cur_item = (None, None)
    retry_count = 0
    is_exported = False
    # set wait for default webdriver wait time
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    # loop till export is successful or retry count is more than max retry
    while is_exported is False and retry_count <= MAX_RETRY:
        try:
            cur_item = ("Table Row", "tr")
            # wait till table row is present
            wait.until(EC.presence_of_element_located((By.XPATH, table_row)))
            LOG.info("{} | {} > Export Excel".format(proj_name, form_type))
            # export all data to excel main function
            do_export_forms_data_excel_main(browser, archive)
            # check if excel file is downloaded
            download_completed, new_file_name = await_download_complete(excel_file_name)
            if download_completed:
                if new_file_name:
                    # rename excel file to new file name if any
                    excel_file_name = new_file_name
                # move exported excel file to project folder
                move_file_to_target_folder(target_folder, excel_file_name)
                # set is_exported to True
                is_exported = True
                # set excel_exported to True for current form type
                EXPORT_LOG[proj_name]["forms"][form_type]["excel_exported"] = True
            else:
                LOG.warning("Download failed!")
        except TimeoutException:
            LOG.warning(not_found_msg(cur_item))
            LOG.warning("Export Excel Failed!")
            # refresh page if any element is timed out
            refresh_page_export(browser, table_row)
            if retry_count < MAX_RETRY:
                LOG.warning(
                    "Export Excel > {} | Retry: {}/{}".format(
                        form_type, retry_count + 1, MAX_RETRY
                    )
                )
            # increase retry count by 1
            retry_count += 1
    # if retry count is more than max retry, skip to export to pdf
    if retry_count >= MAX_RETRY:
        LOG.warning(
            "Export Excel > {} | Max Retry: {} reached!".format(form_type, MAX_RETRY)
        )
        LOG.warning("Skipping to Export PDF...")
        EXPORT_LOG[proj_name]["forms"][form_type]["excel_export_error"] = not_found_msg(
            cur_item
        )


def do_export_forms_data_excel_main(browser, archive):
    """
    Export all forms data in form type to excel main function

    Args:
        browser (Webdriver): Selenium webdriver object
        archive (bool): True if form type is archived. Defaults to False.
    """
    # set wait for element wait time
    wait = WebDriverWait(
        browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    action = ActionChains(browser)
    # if form is archived
    if archive is True:
        cur_item = ("Export all data to Excel", "btn")
        export_btn = wait.until(
            EC.presence_of_element_located((By.XPATH, archive_export_excel))
        )
        # use action to click export btn
        action.move_to_element(export_btn).click().perform()
        LOG.debug(found_msg(cur_item))
    # if form is not archived
    else:
        # click the menu btn
        cur_item = ("Menu", "btn")
        menu_btn = wait.until(EC.presence_of_element_located((By.XPATH, three_dots)))
        action.click(menu_btn).perform()
        LOG.debug(found_msg(cur_item))
        # click the export to excel btn
        cur_item = ("Export all data to Excel", "btn")
        # get the export to excel btn
        export_excel_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, export_excel))
        )
        # use action to click export btn
        action.move_to_element(export_excel_btn).click().perform()
        LOG.debug(found_msg(cur_item))


def export_forms_pdf(browser, form_type, target_folder, proj_name, archive=False):
    """
    Export forms in form type to pdf

    Args:
        browser (Webdriver): Selenium webdriver object
        form_type (object): form type object
        target_folder (string): target folder to move the exported pdf to
        proj_name (string): project name
        archive (bool, optional): True if form type is archived.
        Defaults to False.
    """
    # set wait for short webdriver wait time
    wait = WebDriverWait(
        browser, SHORT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    try:
        # wait and get the total number of forms in the form type
        total_forms = wait.until(
            EC.presence_of_element_located((By.XPATH, total_forms_item))
        )
        # split the text to get the total number of forms
        total_forms = int(total_forms.text.split("of")[-1].strip())
        # get the total number of pages (1 page = 25 forms)
        total_pages = get_total_forms(total_forms)
    except TimeoutException:
        total_pages = 1
        rows = browser.find_elements(By.XPATH, table_rows)
        total_forms = len(rows)
    LOG.debug("{} | {} Total Forms: {}".format(proj_name, form_type, total_forms))
    LOG.debug("{} | {} Total Pages: {}".format(proj_name, form_type, total_pages))
    # update export log with total forms
    EXPORT_LOG[proj_name]["forms"][form_type]["total_forms"] = total_forms
    # if total pages is more than 1, export all forms in each page
    if total_pages > 1:
        multi_page_export_forms_pdf(
            browser, form_type, target_folder, proj_name, total_pages, archive
        )
    # if total pages is 1, export all forms in single page
    else:
        single_page_export_forms_pdf(
            browser, form_type, target_folder, proj_name, total_forms, archive
        )


def multi_page_export_forms_pdf(
    browser, form_type, target_folder, proj_name, total_pages, archive=False
):
    """Export all forms in each page to pdf"""
    cur_item = (None, None)
    # if there are more than 1 page, loop through each page
    LOG.debug("More than 1 page")
    for page in range(1, total_pages + 1):
        # set retry count to 0 for each run
        retry_count = 0
        # set is_exported to False for each run
        is_exported = False
        LOG.debug(
            "{} | {} | Page: {}/{}".format(proj_name, form_type, page, total_pages)
        )
        # loop till export is successful or retry count is more than max retry
        while is_exported is False and retry_count <= MAX_RETRY:
            # reset active page number to 0 for each try
            active_page_num = 0
            # Find the active page element
            try:
                # try and wait till active page is present
                cur_item = ("Active Page", "span")
                active_page_num = get_active_page_num(browser)
                # print all the page numbers
                LOG.debug(
                    "Current Page: {} | Goto Page: {}".format(active_page_num, page)
                )
                # If active page is not current page, go next page
                if active_page_num != page:
                    go_next_page(browser)
                else:
                    cur_page_total_forms = len(
                        browser.find_elements(By.XPATH, table_rows)
                    )
                    # if active page is current page, stay on current page
                    is_exported = stay_on_current_page(
                        browser,
                        page,
                        total_pages,
                        form_type,
                        target_folder,
                        proj_name,
                        archive,
                    )
                    if is_exported is False:
                        # refresh page if export failed
                        LOG.warning("Export PDF Failed!")
                        cur_item = ("Table Rows", "tr")
                        refresh_page_export(browser, table_rows)
                        if retry_count < MAX_RETRY:
                            LOG.warning(
                                "Export PDF > {} | Page: {}/{} | Retry: {}/{}".format(
                                    form_type,
                                    page,
                                    total_pages,
                                    retry_count + 1,
                                    MAX_RETRY,
                                )
                            )
                        # increase retry count by 1
                        retry_count += 1
                    else:
                        # update export log with total exported forms
                        EXPORT_LOG[proj_name]["forms"][form_type][
                            "total_exported_forms"
                        ] += cur_page_total_forms
            except (
                TimeoutException,
                StaleElementReferenceException,
                NoSuchElementException,
            ):
                LOG.warning(not_found_msg(cur_item))
        if retry_count >= MAX_RETRY:
            LOG.warning(
                "Export PDF > {} | Page: {}/{} | Max Retry: {} reached!".format(
                    form_type, page, total_pages, MAX_RETRY
                )
            )
            LOG.warning("Skipping to next page...")
            # add page to pdf export error list
            error_page_list = EXPORT_LOG[proj_name]["forms"][form_type][
                "pdfs_export_error"
            ]
            error_page_list.append({"page": page, "error": not_found_msg(cur_item)})
    # if all pages pdf export is successful, set pdfs_exported to True
    if EXPORT_LOG[proj_name]["forms"][form_type]["pdfs_export_error"] == []:
        EXPORT_LOG[proj_name]["forms"][form_type]["pdfs_exported"] = True


def get_active_page_num(browser):
    """Get the current active page number"""
    # try and wait till active page is present
    wait = WebDriverWait(
        browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    active_page = wait.until(
        EC.presence_of_element_located((By.XPATH, active_page_item))
    )
    if active_page.text:
        active_page_num = int(active_page.text)
    else:
        active_page_num = 0
    return active_page_num


def go_next_page(browser):
    """Function to go to next page"""
    wait = WebDriverWait(
        browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    action = ActionChains(browser)
    LOG.debug("Going Next Page")
    # get the next page item btn using next page button
    cur_item = ("Next Page", "btn")
    # wait for next page item btn to be present
    next_page_item_btn = wait.until(
        EC.presence_of_element_located((By.XPATH, next_page_item))
    )
    # wait some time before click next page
    time.sleep(SHORT_DELAY)
    # Click the "Next Page" button
    action.click(next_page_item_btn).perform()
    LOG.debug(found_msg(cur_item))


def stay_on_current_page(
    browser, page, total_pages, form_type, target_folder, proj_name, archive
):
    """Function for stuff to do on current page"""
    LOG.debug("Staying on Current Page")
    # Continue with exporting forms in current page
    is_exported = do_export_forms_pdf_main(
        browser, form_type, target_folder, proj_name, page, total_pages, archive
    )
    return is_exported


def single_page_export_forms_pdf(
    browser, form_type, target_folder, proj_name, total_forms, archive=False
):
    """Export all forms in single page to pdf"""
    cur_item = (None, None)
    # if there is only 1 page, export forms in current page
    LOG.debug("Only 1 page")
    retry_count = 0
    is_exported = False
    # Continue with exporting forms in current page
    while is_exported is False and retry_count <= MAX_RETRY:
        # export all forms in current page
        is_exported = do_export_forms_pdf_main(
            browser, form_type, target_folder, proj_name, 1, 1, archive
        )
        if is_exported is False:
            try:
                # refresh page if export failed
                LOG.warning("Export PDF Failed!")
                cur_item = ("Table Rows", "tr")
                refresh_page_export(browser, table_rows)
                if retry_count < MAX_RETRY:
                    LOG.warning(
                        "Export PDF > {} | Retry: {}/{}".format(
                            form_type, retry_count + 1, MAX_RETRY
                        )
                    )
                # increase retry count by 1
                retry_count += 1
            except (
                TimeoutException,
                StaleElementReferenceException,
                NoSuchElementException,
            ):
                LOG.warning(not_found_msg(cur_item))
        else:
            # if export is successful, set pdfs_exported to True
            EXPORT_LOG[proj_name]["forms"][form_type]["pdfs_exported"] = True
            # update export log with total exported forms
            EXPORT_LOG[proj_name]["forms"][form_type][
                "total_exported_forms"
            ] += total_forms
    if retry_count >= MAX_RETRY:
        LOG.warning(
            "Export PDF > {} | Max Retry: {} reached!".format(form_type, MAX_RETRY)
        )
        LOG.warning("Skipping to next form type...")
        # add page to pdfs export error list
        error_page_list = EXPORT_LOG[proj_name]["forms"][form_type]["pdfs_export_error"]
        error_page_list.append({"page": 1, "error": not_found_msg(cur_item)})


def do_export_forms_pdf_main(
    browser, form_type, target_folder, proj_name, page, total_pages, archive=False
):
    """
    Function to select all forms in current page and export to pdf

    Args:
        wait (WebDriverWait): WebDriverWait object
        browser (Webdriver): Webdriver object
        target_folder (str): str of target folder to move the pdfs zip file to
        archive (bool, optional): is form archived?
    """
    cur_item = (None, None)
    # set wait for default webdriver wait time
    wait = WebDriverWait(
        browser, DEFAULT_WEBDRIVER_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    action = ActionChains(browser)
    # wait till page is loaded
    wait.until(EC.presence_of_element_located((By.XPATH, table_row)))
    # set wait for element wait time
    wait = WebDriverWait(
        browser, ELEMENT_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
    )
    LOG.info(
        "{} | {} | Export PDF Page: {}/{}".format(
            proj_name, form_type, page, total_pages
        )
    )
    try:
        # click the select all checkbox
        cur_item = ("Select all", "checkbox")
        select_all = wait.until(
            EC.presence_of_element_located((By.XPATH, select_all_box))
        )
        action.click(select_all).perform()
        LOG.debug(found_msg(cur_item))
        time.sleep(SHORT_DELAY)  # give time for page to load
        # if form is archived
        if archive is True:
            cur_item = ("Export to PDF", "btn")
            export_pdf_btn = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "{}".format(archive_export_pdf))
                )
            )
            # use action to click export to pdf btn
            action.click(export_pdf_btn).perform()
            LOG.debug(found_msg(cur_item))
        # if form is not archived
        else:
            is_tippy_box = False
            is_export_pdf_btn = False
            # ensure tippy box is present before clicking export to pdf btn
            while is_tippy_box is False or is_export_pdf_btn is False:
                try:
                    # click the menu btn
                    cur_item = ("Menu", "btn")
                    menu_btn = wait.until(
                        EC.element_to_be_clickable((By.XPATH, three_dots))
                    )
                    action.click(menu_btn).perform()
                    LOG.debug(found_msg(cur_item))
                    time.sleep(SHORT_DELAY)  # give time for tippy box to load
                    # check if tippy box is present
                    cur_item = ("Tippy Box", "div")
                    wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, "{}".format(tippy_box))
                        )
                    )
                    LOG.debug(found_msg(cur_item))
                    is_tippy_box = True
                    cur_item = ("Export to PDF", "btn")
                    # click the export to pdf btn
                    export_pdf_btn = wait.until(
                        EC.visibility_of_element_located((By.XPATH, export_to_pdf))
                    )
                    action.click(export_pdf_btn).perform()
                    LOG.debug(found_msg(cur_item))
                    is_export_pdf_btn = True
                except TimeoutException:
                    LOG.warning(not_found_msg(cur_item))
                    is_tippy_box = False
                    is_export_pdf_btn = False
        # do sub function to export pdfs
        is_exported = do_export_forms_pdf_sub(browser, target_folder)
    except (TimeoutException, StaleElementReferenceException, NoSuchElementException):
        LOG.warning(not_found_msg(cur_item))
        is_exported = False
    return is_exported


def do_export_forms_pdf_sub(browser, target_folder):
    cur_item = (None, None)
    try:
        cur_item = ("Export modal", "div")
        # try and wait for export modal to be present
        wait = WebDriverWait(
            browser, MENU_WAIT_TIME, ignored_exceptions=IGNORED_EXCEPTIONS
        )
        wait.until(EC.presence_of_element_located((By.XPATH, export_modal)))
        # loop through all the checkbox items and get the checkboxes
        action = ActionChains(browser)
        for name, checkbox_xpath in checkboxes_xpath_dict.items():
            cur_item = (name, "checkbox")
            checkbox = browser.find_element(By.XPATH, "{}".format(checkbox_xpath))
            # use action to click checkbox
            action.click(checkbox).perform()
            LOG.debug(found_msg(cur_item))
        # get export btn to export
        cur_item = ("Export", "btn")
        export_btn = browser.find_element(By.XPATH, "{}".format(export_btn_xpath))
        # get export btn if span text is Export
        # inside toolbar btns
        # use action to click export btn
        action.click(export_btn).perform()
        LOG.debug(found_msg(cur_item))
        # downloaded zip file name should be in this format:
        # SYNCHRO_export_yyyy_mm_dd.zip
        pdfs_file_name = "SYNCHRO_export_{}.zip".format(time.strftime("%Y_%m_%d"))
        # wait for pdfs zip file to be downloaded
        download_completed, new_file_name = await_download_complete(pdfs_file_name)
        if download_completed:
            if new_file_name:
                # rename pdfs file name to new file name if any
                pdfs_file_name = new_file_name
            # get the file extension
            pdfs_file_name_ext = os.path.splitext(pdfs_file_name)[-1]
            # if file extension is zip, unzip the file
            if pdfs_file_name_ext == ".zip":
                # move the pdfs zip file to project folder
                move_file_to_target_folder(target_folder, pdfs_file_name)
                new_file_path = os.path.join(target_folder, pdfs_file_name)
                # unzip the pdfs zip file in form type folder
                extract_file(new_file_path, target_folder)
            else:
                # move the pdf file to project folder
                move_file_to_target_folder(target_folder, pdfs_file_name)
            return True
        else:
            LOG.warning("Download failed!")
            return False
    except (TimeoutException, StaleElementReferenceException, NoSuchElementException):
        LOG.warning(not_found_msg(cur_item))
        return False


def await_download_complete(file_name, timeout=600, sleep_frequency=1):
    """
    Waits for a download to complete, returns True if download completed,
    False otherwise

    Args:
        file_name (string): file name to wait for
        timeout (int, optional): await download timeout in seconds.
        Defaults to 600.
        sleep_frequency (int, optional): how long to wait in seconds.
        Defaults to 1.
    """
    # Create an observer to watch the downloads folder for a specific file
    LOG.debug("Downloading: {}".format(file_name))
    event_handler = DownloadHandler(file_name)
    observer = Observer()
    observer.schedule(event_handler, path=TEMP_OUTPUT_PATH, recursive=False)
    observer.start()

    # Wait for the specific file (e.g., "downloaded_file.zip")
    # to appear in the downloads folder
    # Adjust the timeout as needed
    start_time = time.time()
    while event_handler.download_completed is False:
        if time.time() - start_time > timeout:
            LOG.error("Download timeout exceeded.")
            return (False, None)
        time.sleep(sleep_frequency)
    new_file_name = event_handler.downloaded_file_name
    if new_file_name:
        LOG.warning(
            "Downloaded Filename Mismatch! Expected: {} | Actual: {}".format(
                file_name, new_file_name
            )
        )
    # Stop the observer
    observer.stop()
    observer.join()
    return (True, new_file_name)


def found_msg(cur_item):
    """Return found message"""
    return "[{}] {} found".format(cur_item[0], cur_item[1])


def not_found_msg(cur_item):
    """Return not found message"""
    return "[{}] {} not found".format(cur_item[0], cur_item[1])


def move_file_to_target_folder(target_folder, file_name):
    """Move file to target folder"""
    org_file_path = os.path.join(TEMP_OUTPUT_PATH, file_name)
    new_file_path = os.path.join(target_folder, file_name)
    # replace the file if it already exists
    if os.path.exists(new_file_path):
        os.remove(new_file_path)
    os.rename(org_file_path, new_file_path)
    LOG.debug("Moved: {}\nTo: {}".format(org_file_path, new_file_path))


def extract_file(file_abs_path, target_folder):
    """Extract contents of zip file"""
    # file_name = os.path.basename(file_abs_path)
    LOG.debug("Extracting: {}".format(file_abs_path))
    # unzip the file
    with ZipFile(file_abs_path, "r") as zipObj:
        # Extract all the contents of zip file in current directory
        zipObj.extractall(target_folder)
    # delete the zip file
    os.remove(file_abs_path)
    LOG.debug("Extracted: {}".format(file_abs_path))


def get_project_name(browser):
    """Get project name"""
    project_name = browser.find_element(By.CLASS_NAME, "description-text")
    # get the project name from the title
    proj_name = project_name.get_attribute("title").strip()
    LOG.info("Project Name: {}".format(proj_name))
    return proj_name


def load_cookies(browser, cookies):
    """Load cookies from file"""
    LOG.debug("Loading cookies...")
    # if cookies are valid, load them into browser. Navigate to
    # All Synchro URL first in order to load cookies
    browser.get(ALL_SYNCHRO_URL)
    browser.delete_all_cookies()
    # add cookies to browser
    for cookie in cookies:
        browser.add_cookie(cookie)
    browser.refresh()
    LOG.info("Cookies Loaded!")


def save_cookies(browser):
    """Save cookies to file"""
    LOG.info("Saving Cookies...")
    # get logged in cookies and save them to file using pickle
    cookies = browser.get_cookies()
    pickle.dump(cookies, open(COOKIES_FILE, "wb"))
    LOG.info("Cookies Saved!")


def check_cookies(cookies):
    """Check if cookies are expired"""
    for cookie in cookies:
        # check if session cookie has expired
        # skip google analytics cookie as it has no expiry
        if (
            "expiry" in cookie
            and cookie["expiry"] < int(time.time())
            and cookie["name"] != "_gat_gtag_UA_17568443_1"
        ):
            m = "Expired cookie: {} | Expiry: {} | Current Time: {} |".format(
                cookie["name"], cookie["expiry"], int(time.time())
            )
            raise CookiesInvalidError(m)
    # if no cookie is expired
    LOG.info("Cookies are Valid!")


def read_cookies():
    """Read cookies from file"""
    # check if cookies file exists
    LOG.debug("Reading cookies from {}...".format(COOKIES_FILE))
    if not os.path.exists(COOKIES_FILE):
        msg = "{} not found!".format(COOKIES_FILE)
        raise CookiesFileNotFoundError(msg)
    else:
        LOG.debug("{} found!".format(COOKIES_FILE))
        # read cookies from pickle file
        with open(COOKIES_FILE, "rb") as f:
            cookies = pickle.load(f)
        return cookies


def cookies_invalid_runtime(browser, project_urls):
    """Stuff to do when a cookie invalid or cookie file not found is raised"""
    logged_in = False
    try_count = 0
    while logged_in is False and try_count <= MAX_RETRY:
        try:
            # login with clean session
            browser.delete_all_cookies()
            logged_in = login_optimus(browser)
        except LoginError as e:
            LOG.warning(e)
            logged_in = False
            LOG.warning("Retrying Login... {}/{}".format(try_count, MAX_RETRY))
            try_count += 1
    # go back to main_runtime with new logged in session cookies
    main_runtime(browser, project_urls)


def main_runtime(browser, project_urls):
    """Stuff to do when script is running with no errors"""
    proj_export_error = ""
    # loop through project urls
    for project_url in project_urls:
        retry_count = 0
        export_done = False
        while export_done is False and retry_count <= MAX_RETRY:
            try:
                project_url = project_url.strip().replace("\n", "").replace("\r", "")
                if not project_url:
                    continue
                # navigate to project page
                navigate_to_page(browser, url=project_url)
                # get project name
                proj_name = get_project_name(browser)
                # setup export log for project
                setup_export_log(proj_name)
                # create project folder
                proj_folder = os.path.join(OUTPUT_PATH, proj_name)
                if not os.path.exists(proj_folder):
                    os.makedirs(proj_folder)
                # get list of form types in project and export forms data
                form_types_list = get_proj_form_types(browser)
                export_forms_data(browser, form_types_list, proj_folder, proj_name)
                # set export_done to True if no errors
                EXPORT_LOG[proj_name]["export_done"] = True
                LOG.info("{} - All Data Exported!".format(proj_name))
                export_done = True
            except KeyboardInterrupt:
                LOG.warning("Keyboard Interrupt!")
            except ALL_ERRORS as e:
                proj_export_error = str(e)
                LOG.error(e)
                if retry_count < MAX_RETRY:
                    LOG.warning("{} Export Failed!".format(proj_name))
                    LOG.warning("Retrying... {}/{}".format(retry_count + 1, MAX_RETRY))
                retry_count += 1
        if retry_count >= MAX_RETRY:
            LOG.warning(
                "{} Export Failed! Max Retry: {} Reached!".format(proj_name, MAX_RETRY)
            )
            LOG.warning("Skipping to next project...")
            EXPORT_LOG[proj_name]["proj_export_error"] = proj_export_error
    LOG.info("All Projects Exported!")
    # save cookies to file
    save_cookies(browser)


def cleanup_runtime(browser, wait_time=5):
    """Stuff to do when script is closing"""
    CatchErrors = (Exception, OSError) + SELENIUM_ERROR
    # close script after wait_time seconds
    for i in range(wait_time, 0, -1):
        LOG.info("Closing Script in {}s...".format(i))
        time.sleep(1)
    try:
        # write errors during runtime of export to file
        write_export_log_to_file(EXPORT_LOG)
        # try to close browser
        LOG.info("Closing Browser...")
        if browser:
            browser.quit()
        LOG.info("Good Bye!")
        sys.exit(0)
    except CatchErrors as e:
        LOG.error(e)
        LOG.info("Good Bye!")
        sys.exit(1)


def setup_export_log(proj_name, form_type=None):
    """Setup the error log for each form type"""
    # setup the form_type error log
    if proj_name not in EXPORT_LOG:
        if not form_type:
            EXPORT_LOG[proj_name] = {
                "export_done": False,
                "proj_export_error": "",
                "forms": {},
            }
        else:
            EXPORT_LOG[proj_name] = {
                "export_done": False,
                "proj_export_error": "",
                "forms": {
                    form_type: {
                        "total_forms": 0,
                        "total_exported_forms": 0,
                        "forms_export_error": "",
                        "excel_exported": False,
                        "excel_export_error": "",
                        "pdfs_exported": False,
                        "pdfs_export_error": [],
                    }
                },
            }
    else:
        if not form_type:
            EXPORT_LOG[proj_name] = {
                "export_done": False,
                "proj_export_error": "",
                "forms": {},
            }
        else:
            EXPORT_LOG[proj_name]["forms"][form_type] = {
                "total_forms": 0,
                "total_exported_forms": 0,
                "forms_export_error": "",
                "excel_exported": False,
                "excel_export_error": "",
                "pdfs_exported": False,
                "pdfs_export_error": [],
            }


def create_folders():
    """Create folders if they don't exist, clear temp folder"""
    LOG.debug("Creating/Clearing output & temp folders...")
    if not os.path.exists(OUTPUT_PATH):
        os.makedirs(OUTPUT_PATH)
    if not os.path.exists(TEMP_OUTPUT_PATH):
        os.makedirs(TEMP_OUTPUT_PATH)
    if not os.path.exists(EXPORT_LOG_DIR):
        os.makedirs(EXPORT_LOG_DIR)
    # clear temp folder
    for file in os.listdir(TEMP_OUTPUT_PATH):
        file_path = os.path.join(TEMP_OUTPUT_PATH, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            LOG.error(e)


def write_export_log_to_file(logs):
    """Write export logs to file"""
    if logs:
        LOG.debug("Saving Export Log...")
        # write dictionary to file using json
        with open(EXPORT_LOG_FILE, "w") as f:
            json.dump(logs, f, indent=4)
        # get the export log filename created when script started
        export_log_filename = os.path.basename(EXPORT_LOG_FILE)
        LOG.info("{} Saved!".format(export_log_filename))
    else:
        LOG.info("No Export Logs to Save!")


def read_project_urls():
    """Read project urls from file"""
    # check if project urls file exists
    LOG.debug("Reading project urls from {}...".format(PROJ_URLS_FILE))
    if not os.path.exists(PROJ_URLS_FILE):
        msg = "{} not found!".format(PROJ_URLS_FILE)
        raise ProjectUrlsFileNotFoundError(msg)
    else:
        LOG.debug("{} found!".format(PROJ_URLS_FILE))
        # read project urls from file
        with open(PROJ_URLS_FILE, "r") as f:
            project_urls = f.readlines()
        return project_urls


def main():
    """Main function"""
    msg = "DEV MODE: {}".format(env.get("dev_mode"))
    LOG.info(msg)
    # create folders if they don't exist
    create_folders()
    browser = None
    try:
        # read project urls from file
        project_urls = read_project_urls()
        # start webdriver only if project urls file is found
        browser = webdriver.Chrome(options=OPTIONS)
        # maximize browser window
        browser.maximize_window()
        # read cookies from file
        cookies = read_cookies()
        # checks cookies expiry and login if expired
        check_cookies(cookies)
        # load cookies into browser
        load_cookies(browser, cookies)
        # run script
        main_runtime(browser, project_urls)
    except ProjectUrlsFileNotFoundError as e:
        LOG.error(e)
        LOG.error("Please ensure {} exists first!".format(PROJ_URLS_FILE))
    except (CookiesInvalidError, CookiesFileNotFoundError) as e:
        LOG.warning(e)
        cookies_invalid_runtime(browser, project_urls)
    except KeyboardInterrupt:
        LOG.warning("Keyboard Interrupt!")
    except ALL_ERRORS as e:
        # if any other exception occurs
        LOG.error("Unhandled Error Occured!")
        LOG.error(e)
    finally:
        cleanup_runtime(browser)


if __name__ == "__main__":
    main()
