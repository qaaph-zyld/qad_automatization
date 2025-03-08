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

class QADAutomation:
    def __init__(self, username: str, password: str, state_id: str = None):
        """
        Initialize QAD automation
        
        Args:
            username (str): QAD username
            password (str): QAD password
            state_id (str): Optional state ID for custom folder navigation
        """
        self.username = username
        self.password = password
        self.state_id = state_id
        self.logger = logging.getLogger(__name__)
        
        # Setup logging
        self._setup_logging()
        
        # Check for existing QAD processes first
        if not self._check_and_handle_existing_qad():
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

    def _find_login_window(self, timeout=30) -> tuple:
        """
        Find the QAD login window and return the window and app objects
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Try to connect to QAD.Client process
                app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\QAD\QAD Enterprise Applications\container\QAD.Client.exe")
                
                # Try different window title patterns
                window_patterns = [
                    ".*QAD.*",
                    "Login",
                    ".*Enterprise.*Applications.*",
                    "QAD Enterprise Applications"
                ]
                
                for pattern in window_patterns:
                    try:
                        window = app.window(title_re=pattern)
                        if window.exists():
                            self.logger.info(f"Found window with title pattern: {pattern}")
                            return app, window
                    except:
                        continue
                        
            except Exception as e:
                self.logger.debug(f"Window search attempt failed: {str(e)}")
            
            time.sleep(2)
        
        raise Exception("Login window not found after timeout")

    def _handle_protocol_dialog(self) -> bool:
        """
        Handle the Edge protocol dialog that appears when opening QAD URL
        Returns True if successful
        """
        try:
            self.logger.info("STEP 2: Waiting for protocol dialog to appear...")
            time.sleep(5)  # Wait longer for dialog to appear
            
            self.logger.info("STEP 3: Pressing Tab to focus on Open button...")
            keyboard.send_keys('{TAB}')
            time.sleep(1)
            
            self.logger.info("STEP 4: Pressing Enter to click Open...")
            keyboard.send_keys('{ENTER}')
            
            # Wait and verify QAD window appears
            self.logger.info("STEP 5: Verifying QAD window appears...")
            if self._verify_qad_window_exists():
                self.logger.info("SUCCESS: QAD window detected")
                return True
            else:
                self.logger.error("FAILED: QAD window not detected after protocol handling")
                return False
            
        except Exception as e:
            self.logger.error(f"Protocol dialog handling failed: {str(e)}")
            return False

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

    def login(self) -> None:
        """
        Log into the QAD desktop application using URL protocol
        """
        try:
            self.logger.info("STEP 1: Starting QAD application via URL protocol...")
            
            # Check for existing QAD windows
            existing_windows = self._check_existing_qad_windows()
            self.logger.info(f"Found {existing_windows} existing QAD windows")
            
            # Navigate to QAD URL with proper encoding
            if self.state_id:
                url = f"qadsh://browse/invoke?state-id={self.state_id}"
                self.logger.info(f"Opening QAD with state ID: {self.state_id}")
            else:
                url = "qadsh://browse/invoke"
            
            # Handle the URL protocol
            try:
                # First ensure Edge is the default browser for qadsh protocol
                self.logger.info("Setting up Edge for QAD protocol...")
                self.driver.get('edge://settings/defaultBrowser')
                time.sleep(2)
                
                # Now navigate to the QAD URL
                self.logger.info(f"Navigating to QAD URL: {url}")
                self.driver.get(url)
                self.logger.info("Successfully navigated to QAD URL")
                
                # Handle Edge protocol dialog and verify success
                if not self._handle_protocol_dialog():
                    raise Exception("Failed to handle protocol dialog")
                
                self.logger.info("STEP 6: Waiting for QAD login window...")
                time.sleep(15)  # Wait for QAD to fully launch
                
                # Try multiple times to connect to the login window
                max_attempts = 3
                for attempt in range(max_attempts):
                    try:
                        self.logger.info(f"STEP 7: Attempting to connect to QAD login window (attempt {attempt + 1}/{max_attempts})")
                        
                        # Find the login window
                        app, login_window = self._find_login_window()
                        
                        self.logger.info("STEP 8: Found login window, entering credentials...")
                        
                        # Enter username and password
                        username_field = login_window.child_window(auto_id="username")
                        if not username_field.exists():
                            # Try alternate IDs
                            username_field = login_window.child_window(auto_id="_txtUserName")
                        
                        username_field.type_keys(self.username)
                        
                        self.logger.info("STEP 9: Username entered, moving to password...")
                        keyboard.send_keys('{TAB}')
                        time.sleep(1)
                        
                        password_field = login_window.child_window(auto_id="password")
                        if not password_field.exists():
                            # Try alternate IDs
                            password_field = login_window.child_window(auto_id="_txtPassword")
                            
                        password_field.type_keys(self.password)
                        
                        self.logger.info("STEP 10: Submitting login form...")
                        keyboard.send_keys('{ENTER}')
                        
                        # Wait longer for login to process and main window to appear
                        self.logger.info("Waiting for main window to appear after login...")
                        time.sleep(20)
                        self.logger.info("SUCCESS: Login sequence completed")
                        return
                        
                    except Exception as e:
                        self.logger.warning(f"Login attempt {attempt + 1} failed: {str(e)}")
                        if attempt == max_attempts - 1:
                            raise
                        time.sleep(5)
                
            except Exception as e:
                self.logger.error(f"Failed to navigate to URL: {str(e)}")
                raise
                
        except Exception as e:
            self.logger.error(f"Login failed: {str(e)}")
            raise

    def _check_existing_qad_windows(self) -> int:
        """
        Check for existing QAD.Client windows and return the count
        """
        try:
            app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\QAD\QAD Enterprise Applications\container\QAD.Client.exe")
            windows = app.windows()
            return len(windows)
        except:
            return 0

    def export_to_excel(self) -> None:
        """
        Export data to Excel using the QAD desktop app
        """
        try:
            self.logger.info("STEP 11: Starting Excel export process...")
            
            # Connect to QAD application
            app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\QAD\QAD Enterprise Applications\container\QAD.Client.exe")
            
            # Wait and retry a few times to find the main window
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    # Get all QAD windows
                    windows = app.windows()
                    self.logger.info(f"Found {len(windows)} QAD windows")
                    
                    # Find the main window (should be the largest/most recently active)
                    main_window = None
                    for window in windows:
                        try:
                            title = window.window_text()
                            self.logger.info(f"Found window: {title}")
                            if title and "QAD" in title and "Login" not in title:
                                main_window = window
                                self.logger.info(f"Selected main window: {title}")
                                break
                        except Exception as e:
                            self.logger.warning(f"Error checking window: {str(e)}")
                            continue
                    
                    if main_window:
                        break
                    else:
                        self.logger.info(f"Main window not found on attempt {attempt + 1}, waiting...")
                        time.sleep(10)
                        
                except Exception as e:
                    self.logger.warning(f"Error on attempt {attempt + 1}: {str(e)}")
                    if attempt < max_attempts - 1:
                        time.sleep(10)
                    continue
            
            if not main_window:
                raise Exception("Could not find main QAD window after multiple attempts")
            
            self.logger.info("STEP 12: Looking for export button...")
            
            # Try to find and click the export button
            try:
                # First try to use keyboard shortcuts
                self.logger.info("Using keyboard shortcuts for export...")
                
                # Focus the window first
                main_window.set_focus()
                time.sleep(1)
                
                # Try Alt+A to open Actions menu
                self.logger.info("Opening Actions menu...")
                keyboard.send_keys('%a')  # Alt+A
                time.sleep(1)
                
                # Try to find Export in the menu
                self.logger.info("Looking for Export option...")
                keyboard.send_keys('e')  # E for Export
                time.sleep(1)
                
                # Press Enter to confirm
                keyboard.send_keys('{ENTER}')
                
                # Wait for export dialog/process
                time.sleep(5)
                self.logger.info("SUCCESS: Export process completed")
                
            except Exception as e:
                self.logger.error(f"Failed to trigger export: {str(e)}")
                raise
            
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
    
    args = parser.parse_args()
    
    # Get credentials from environment if not provided
    username = args.username or os.getenv('QAD_USERNAME')
    password = args.password or os.getenv('QAD_PASSWORD')
    
    if not username or not password:
        raise ValueError("Username and password must be provided either as arguments or environment variables")
    
    try:
        qad = QADAutomation(username, password, args.state_id)
        qad.login()
        qad.export_to_excel()
    except Exception as e:
        logging.error(f"Automation failed: {str(e)}")
        raise

if __name__ == "__main__":
    main()
