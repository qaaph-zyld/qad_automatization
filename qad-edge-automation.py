import os
import time
import logging
import argparse
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from pywinauto import Desktop, Application
import shutil
import pywinauto.keyboard as keyboard
import psutil
import tkinter as tk
from tkinter import messagebox
from pywinauto.findwindows import find_windows

class QADAutomation:
    def __init__(self, username: str, password: str, state_id: str = None, force: bool = False):
        """
        Initialize QAD automation
        
        Args:
            username (str): QAD username
            password (str): QAD password
            state_id (str): Optional state ID for custom folder navigation
            force (bool): Force execution even if QAD processes are running
        """
        self.username = username
        self.password = password
        self.state_id = state_id
        self.force = force
        self.logger = logging.getLogger(__name__)
        
        # Setup logging
        self._setup_logging()
        
        # Check for existing QAD processes first (unless force is True)
        if not self.force and not self._check_and_handle_existing_qad():
            raise Exception("Please close all QAD windows and try again")
            
        # Initialize Edge WebDriver with specific options
        self._setup_driver()
        
        # Initialize Edge WebDriver with specific options
        edge_options = webdriver.EdgeOptions()
        edge_options.use_chromium = True
        edge_options.add_argument('--start-maximized')
        
        # Add protocol handler for qadsh
        edge_options.add_argument('--protocol-handler.allowed-origin=qadsh://*')
        edge_options.add_argument('--protocol-handler.excluded-schemes')
        
        self.driver = webdriver.Edge(options=edge_options)
        self.logger.info("Edge WebDriver initialized successfully")

    def _setup_logging(self):
        """
        Set up logging configuration
        """
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / f'qad_automation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        
    def _check_and_handle_existing_qad(self) -> bool:
        """
        Check for existing QAD processes and show popup if found
        Returns True if no QAD processes found or user closed them
        """
        while True:
            qad_processes = []
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'qad' in proc.info['name'].lower():
                        qad_processes.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not qad_processes:
                return True
                
            # Create popup window
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            
            response = messagebox.showwarning(
                "QAD Already Running",
                "Please close all QAD windows before continuing.\n\n" +
                "Click OK once you have closed all QAD windows, or Cancel to exit.",
                type=messagebox.OKCANCEL
            )
            
            if response == 'cancel':
                return False
                
            # Give some time for processes to close
            time.sleep(2)
            
    def _setup_driver(self):
        """
        Initialize Edge WebDriver with required options
        """
        edge_options = webdriver.EdgeOptions()
        edge_options.use_chromium = True
        edge_options.add_argument('--start-maximized')
        
        # Add protocol handler for qadsh
        edge_options.add_argument('--protocol-handler.allowed-origin=qadsh://*')
        edge_options.add_argument('--protocol-handler.excluded-schemes')
        
        self.driver = webdriver.Edge(options=edge_options)
        self.logger.info("Edge WebDriver initialized successfully")

    def setup_driver(self) -> None:
        """
        Configure and initialize Edge WebDriver with appropriate settings
        """
        try:
            edge_options = Options()
            edge_options.add_experimental_option('prefs', {
                'download.default_directory': str(self.temp_dir),
                'download.prompt_for_download': False,
                'download.directory_upgrade': True,
                'safebrowsing.enabled': True
            })
            
            # Automatically handle SSL errors and certificate warnings
            edge_options.add_argument('--ignore-certificate-errors')
            edge_options.add_argument('--ignore-ssl-errors')
            
            # Initialize Edge WebDriver
            self.driver = webdriver.Edge(options=edge_options)
            self.wait = WebDriverWait(self.driver, 20)
            self.logger.info("Edge WebDriver initialized successfully")
            
        except WebDriverException as e:
            self.logger.error(f"Failed to initialize Edge WebDriver: {str(e)}")
            raise

    def wait_for_element(self, by: By, value: str, timeout: int = 20) -> None:
        """
        Wait for element to be clickable with explicit timeout
        
        Args:
            by: Selenium By locator strategy
            value: Element identifier
            timeout: Maximum wait time in seconds
        """
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            return element
        except TimeoutException:
            self.logger.error(f"Element {value} not found within {timeout} seconds")
            raise

    def _find_login_window(self, timeout=30):
        """
        Find the QAD login window and return the window and app objects
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Look for windows with "Login" in the title
                windows = find_windows(title_re=".*Login.*")
                if windows:
                    window_handle = windows[0]
                    app = Application().connect(handle=window_handle)
                    window = app.window(handle=window_handle)
                    self.logger.info(f"Found window with title pattern: {window.window_text()}")
                    return app, window
            except Exception as e:
                self.logger.debug(f"Error finding login window: {str(e)}")
            
            time.sleep(2)
        
        self.logger.error("Could not find QAD login window")
        return None, None

    def _verify_qad_window_exists(self, timeout=30) -> bool:
        """
        Verify that QAD window exists
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\QAD\QAD Enterprise Applications\container\QAD.Client.exe")
                windows = app.windows()
                if len(windows) > 0:
                    return True
            except:
                time.sleep(2)
        return False

    def _handle_protocol_dialog(self) -> bool:
        """
        Handle the protocol dialog that appears when opening QAD via URL
        Returns True if successful, False otherwise
        """
        try:
            self.logger.info("STEP 2: Waiting for protocol dialog to appear...")
            time.sleep(5)  # Wait for dialog to appear
            
            # Try multiple key combinations to handle the dialog
            for attempt in range(3):
                self.logger.info(f"Protocol dialog handling attempt {attempt + 1}/3")
                
                # Method 1: Try using Tab and Enter
                self.logger.info("STEP 3: Pressing Tab to focus on Open button...")
                keyboard.send_keys('{TAB}')
                time.sleep(1)
                
                self.logger.info("STEP 4: Pressing Enter to click Open...")
                keyboard.send_keys('{ENTER}')
                time.sleep(2)
                
                # Method 2: Try Alt+O for Open
                if attempt >= 1:
                    self.logger.info("Trying Alt+O shortcut...")
                    keyboard.send_keys('%o')  # Alt+O
                    time.sleep(2)
                
                # Method 3: Try different key combinations
                if attempt >= 2:
                    self.logger.info("Trying additional key combinations...")
                    # Try Space
                    keyboard.send_keys(' ')
                    time.sleep(1)
                    # Try Enter again
                    keyboard.send_keys('{ENTER}')
                    time.sleep(1)
                    # Try Tab + Enter again
                    keyboard.send_keys('{TAB}{ENTER}')
                    time.sleep(1)
                
                # Verify QAD window appears
                self.logger.info("STEP 5: Verifying QAD window appears...")
                start_time = time.time()
                while time.time() - start_time < 30:  # Wait up to 30 seconds
                    try:
                        # Check for QAD window
                        qad_windows = self._check_existing_qad_windows()
                        if qad_windows > 0:
                            self.logger.info("SUCCESS: QAD window detected")
                            return True
                    except:
                        pass
                    
                    time.sleep(2)
                
                # If we're still here, the dialog may not have been handled correctly
                if attempt < 2:
                    self.logger.warning(f"Protocol dialog handling attempt {attempt + 1} failed, trying again...")
                    time.sleep(2)
            
            self.logger.error("FAILED: QAD window not detected after protocol handling")
            return False
            
        except Exception as e:
            self.logger.error(f"Error handling protocol dialog: {str(e)}")
            return False

    def login(self) -> None:
        """
        Log into the QAD desktop application using URL protocol
        """
        try:
            self.logger.info("STEP 1: Starting QAD application via URL protocol...")
            
            # Check for existing QAD windows and store them
            self.existing_qad_windows = self._get_qad_window_list()
            existing_count = len(self.existing_qad_windows)
            self.logger.info(f"Found {existing_count} existing QAD windows")
            
            # Open QAD using URL protocol
            self.logger.info(f"Opening QAD with state ID: {self.state_id}")
            
            # Set up Edge for QAD protocol
            self.logger.info("Setting up Edge for QAD protocol...")
            
            # Navigate to QAD URL with state ID
            qad_url = f"qadsh://browse/invoke?state-id={self.state_id}"
            self.logger.info(f"Navigating to QAD URL: {qad_url}")
            self.driver.get(qad_url)
            self.logger.info("Successfully navigated to QAD URL")
            
            # Handle protocol dialog
            self.logger.info("STEP 2: Handling protocol dialog...")
            time.sleep(5)  # Wait for dialog to appear
            
            # Press Tab to focus on Open button
            self.logger.info("STEP 3: Pressing Tab to focus on Open button...")
            keyboard.send_keys('{TAB}')
            time.sleep(1)
            
            # Press Enter to click Open
            self.logger.info("STEP 4: Pressing Enter to click Open...")
            keyboard.send_keys('{ENTER}')
            time.sleep(5)  # Wait for QAD to start
            
            # Wait for QAD login window to appear
            self.logger.info("STEP 5: Waiting for QAD login window...")
            time.sleep(5)  # Wait for login window to appear
            
            # Enter credentials
            self.logger.info("STEP 6: Entering credentials...")
            
            # Enter username using keyboard
            self.logger.info("Entering username...")
            keyboard.send_keys(self.username)
            time.sleep(1)
            
            # Tab to password field
            self.logger.info("STEP 7: Username entered, moving to password...")
            keyboard.send_keys('{TAB}')
            time.sleep(1)
            
            # Enter password
            self.logger.info("Entering password...")
            keyboard.send_keys(self.password)
            time.sleep(1)
            
            # Submit form
            self.logger.info("STEP 8: Submitting login form...")
            keyboard.send_keys('{ENTER}')
            
            # Wait for main window to appear
            self.logger.info("Waiting for main window to appear after login...")
            time.sleep(20)  # Wait 20 seconds for QAD to load
            
            self.logger.info("SUCCESS: Login sequence completed")
            
        except Exception as e:
            self.logger.error(f"Login failed: {str(e)}")
            raise
    
    def _check_existing_qad_windows(self) -> int:
        """
        Check for existing QAD.Client windows and return the count
        """
        try:
            windows = find_windows(title_re=".*QAD Enterprise Applications.*")
            return len(windows)
        except Exception as e:
            self.logger.debug(f"Error checking for QAD windows: {str(e)}")
            return 0
    
    def _get_qad_window_list(self) -> list:
        """
        Get a list of existing QAD windows with their handles and titles
        """
        try:
            windows = find_windows(title_re=".*QAD Enterprise Applications.*")
            window_info = []
            for handle in windows:
                try:
                    app = Application().connect(handle=handle)
                    window = app.window(handle=handle)
                    window_info.append({
                        'handle': handle,
                        'title': window.window_text()
                    })
                    self.logger.debug(f"Found existing QAD window: {window.window_text()}")
                except Exception as e:
                    self.logger.debug(f"Error getting window info: {str(e)}")
            return window_info
        except Exception as e:
            self.logger.debug(f"Error listing QAD windows: {str(e)}")
            return []
    
    def _identify_new_qad_window(self) -> dict:
        """
        Identify the newly opened QAD window by comparing current windows with existing ones
        Returns the new window info or None if not found
        """
        try:
            # Get current list of QAD windows
            current_windows = self._get_qad_window_list()
            
            if not hasattr(self, 'existing_qad_windows'):
                self.logger.warning("No record of existing windows, using first available window")
                return current_windows[0] if current_windows else None
            
            # Find windows that weren't in the original list
            existing_handles = [w['handle'] for w in self.existing_qad_windows]
            new_windows = [w for w in current_windows if w['handle'] not in existing_handles]
            
            if new_windows:
                self.logger.info(f"Identified new QAD window: {new_windows[0]['title']}")
                return new_windows[0]
            
            # If no new window found but there are current windows, use the first one
            if current_windows:
                self.logger.warning("No new window identified, using first available window")
                return current_windows[0]
            
            return None
        except Exception as e:
            self.logger.error(f"Error identifying new QAD window: {str(e)}")
            return None
    
    def export_to_excel(self) -> None:
        """
        Export data to Excel using the QAD desktop app
        """
        try:
            self.logger.info("STEP 11: Starting Excel export process...")
            
            # Wait longer for QAD menu to fully load
            self.logger.info("Waiting for QAD menu to fully load...")
            time.sleep(20)  # Wait 20 seconds for QAD menu to load completely
            
            # Identify the newly opened QAD window
            new_window_info = self._identify_new_qad_window()
            
            if not new_window_info:
                raise Exception("No QAD windows found for export")
            
            # Connect to the identified QAD window
            try:
                app = Application().connect(handle=new_window_info['handle'])
                main_window = app.window(handle=new_window_info['handle'])
                self.logger.info(f"Connected to QAD window: {main_window.window_text()}")
            except Exception as e:
                self.logger.error(f"Failed to connect to identified QAD window: {str(e)}")
                raise Exception("Could not connect to identified QAD window")
            
            # Start the export sequence
            self.logger.info("STEP 12: Starting export sequence...")
            
            # Focus the window
            main_window.set_focus()
            time.sleep(2)
            
            # Press Alt to open menu
            self.logger.info("Pressing Alt key...")
            keyboard.send_keys('%')
            time.sleep(1)
            
            # Press Enter to select first menu item
            self.logger.info("Pressing Enter...")
            keyboard.send_keys('{ENTER}')
            time.sleep(1)
            
            # Press Down Arrow and Enter to navigate to export
            self.logger.info("Pressing Down Arrow and Enter...")
            keyboard.send_keys('{DOWN}{ENTER}')
            time.sleep(1)
            
            # Press Down Arrow and Enter again for Excel export
            self.logger.info("Pressing Down Arrow and Enter again...")
            keyboard.send_keys('{DOWN}{ENTER}')
            time.sleep(0.1)
            
            # Wait for Excel file to open
            self.logger.info("Waiting for Excel file to open...")
            time.sleep(5)  # Adjust this timeout as needed
            
            self.logger.info("SUCCESS: Export sequence completed")
            
        except Exception as e:
            self.logger.error(f"Export to Excel failed: {str(e)}")
            raise

    def cleanup(self) -> None:
        """
        Clean up resources and close browser
        """
        try:
            if hasattr(self, 'driver'):
                self.driver.quit()
                self.logger.info("Browser session closed successfully")
        except Exception as e:
            self.logger.error(f"Cleanup failed: {str(e)}")

def main():
    """
    Main entry point
    """
    parser = argparse.ArgumentParser(description='QAD Automation Script')
    parser.add_argument('--username', type=str, help='QAD username')
    parser.add_argument('--password', type=str, help='QAD password')
    parser.add_argument('--state-id', type=str, help='QAD state ID for custom folder navigation')
    parser.add_argument('--force', action='store_true', help='Force execution even if QAD processes are running')
    
    args = parser.parse_args()
    
    # Get credentials from environment if not provided
    username = args.username or os.getenv('QAD_USERNAME')
    password = args.password or os.getenv('QAD_PASSWORD')
    
    if not username or not password:
        raise ValueError("Username and password must be provided either as arguments or environment variables")
    
    try:
        qad = QADAutomation(username, password, args.state_id, args.force)
        qad.login()
        qad.export_to_excel()
    except Exception as e:
        logging.error(f"Automation failed: {str(e)}")
        raise

if __name__ == "__main__":
    main()
