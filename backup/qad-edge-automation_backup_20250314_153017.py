import os
import sys
import time
import logging
import argparse
import tkinter as tk
from tkinter import messagebox
import psutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pywinauto
from pywinauto import Application
from pywinauto.findwindows import find_windows
from pywinauto.keyboard import send_keys
from pathlib import Path
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

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
        if not self.force:
            self._check_and_close_qad_applications()
            
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
        
    def _check_and_close_qad_applications(self):
        """
        Check for existing QAD applications and prompt user to close them
        Verifies that applications are actually closed before proceeding
        """
        while True:
            # Check for QAD processes
            qad_processes = self._get_qad_processes()
            
            if not qad_processes:
                self.logger.info("No QAD applications running, proceeding with automation")
                return True
                
            # Create popup window to notify user
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            
            process_info = "\n".join([f"- {proc.info['name']} (PID: {proc.info['pid']})" for proc in qad_processes[:5]])
            if len(qad_processes) > 5:
                process_info += f"\n- ... and {len(qad_processes) - 5} more"
                
            message = (
                f"Found {len(qad_processes)} QAD application(s) running:\n\n"
                f"{process_info}\n\n"
                "Please close all QAD applications before continuing.\n\n"
                "Click OK once you have closed all QAD applications, or Cancel to exit."
            )
            
            response = messagebox.showwarning(
                "QAD Applications Running",
                message,
                type=messagebox.OKCANCEL
            )
            
            if response == 'cancel':
                self.logger.info("User canceled automation due to running QAD applications")
                raise Exception("Automation canceled by user")
                
            # Wait a moment for processes to close
            self.logger.info("Waiting for QAD applications to close...")
            time.sleep(3)
            
            # Verify that applications are actually closed
            qad_processes = self._get_qad_processes()
            if not qad_processes:
                self.logger.info("All QAD applications successfully closed")
                return True
            else:
                self.logger.warning(f"QAD applications still running ({len(qad_processes)} found)")
                # Loop will continue and prompt user again
    
    def _get_qad_processes(self):
        """
        Get a list of running QAD processes
        """
        qad_processes = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name'].lower()
                # Check for various QAD-related process names
                if any(qad_term in proc_name for qad_term in ['qad', 'progress', 'sqlservr']):
                    qad_processes.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        return qad_processes

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
        
        # Add options to bypass Edge sign-in prompts
        edge_options.add_argument('--disable-blink-features=AutomationControlled')
        edge_options.add_argument('--no-first-run')
        edge_options.add_argument('--no-default-browser-check')
        edge_options.add_argument('--disable-extensions')
        edge_options.add_argument('--disable-sync')
        edge_options.add_argument('--disable-gpu')
        
        # Add user data directory to maintain session
        user_data_dir = os.path.join(os.environ['LOCALAPPDATA'], 'QAD_Automation_Profile')
        if not os.path.exists(user_data_dir):
            os.makedirs(user_data_dir)
        edge_options.add_argument(f'--user-data-dir={user_data_dir}')
        
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

    def _find_qad_windows(self):
        """Find all QAD windows using various title patterns"""
        qad_windows = []
        
        # Try to find windows with QAD in the title
        try:
            qad_pattern_windows = find_windows(title_re=".*QAD.*")
            self.logger.info(f"Found {len(qad_pattern_windows)} windows with pattern '.*QAD.*'")
            qad_windows.extend(qad_pattern_windows)
        except Exception as e:
            self.logger.warning(f"Error finding QAD windows: {str(e)}")
        
        # Try to find windows with Enterprise in the title
        try:
            enterprise_windows = find_windows(title_re=".*Enterprise.*")
            self.logger.info(f"Found {len(enterprise_windows)} windows with pattern '.*Enterprise.*'")
            qad_windows.extend(enterprise_windows)
        except Exception as e:
            self.logger.warning(f"Error finding Enterprise windows: {str(e)}")
        
        # Try to find windows with Browse in the title (common in QAD)
        try:
            browse_windows = find_windows(title_re=".*Browse.*")
            if browse_windows:
                self.logger.info(f"Found {len(browse_windows)} windows with pattern '.*Browse.*'")
                qad_windows.extend(browse_windows)
        except Exception as e:
            self.logger.warning(f"Error finding Browse windows: {str(e)}")
            
        # Filter out any windows that are clearly not QAD windows (e.g., editor windows)
        filtered_windows = []
        for window_handle in qad_windows:
            try:
                window = Application().connect(handle=window_handle).window(handle=window_handle)
                title = window.window_text()
                # Skip windows that are clearly not QAD windows
                if "Windsurf" in title or "Visual Studio" in title or "Code" in title:
                    self.logger.info(f"Skipping non-QAD window: {title}")
                    continue
                filtered_windows.append(window_handle)
            except Exception as e:
                self.logger.warning(f"Error checking window title: {str(e)}")
                
        self.logger.info(f"Found {len(filtered_windows)} QAD windows")
        return filtered_windows

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
                send_keys('{TAB}')
                time.sleep(1)
                
                self.logger.info("STEP 4: Pressing Enter to click Open...")
                send_keys('{ENTER}')
                time.sleep(2)
                
                # Method 2: Try Alt+O for Open
                if attempt >= 1:
                    self.logger.info("Trying Alt+O shortcut...")
                    send_keys('%o')  # Alt+O
                    time.sleep(2)
                
                # Method 3: Try different key combinations
                if attempt >= 2:
                    self.logger.info("Trying additional key combinations...")
                    # Try Space
                    send_keys(' ')
                    time.sleep(1)
                    # Try Enter again
                    send_keys('{ENTER}')
                    time.sleep(1)
                    # Try Tab + Enter again
                    send_keys('{TAB}{ENTER}')
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
            
            # Check for existing QAD windows
            qad_windows = self._find_qad_windows()
            self.logger.info(f"Found {len(qad_windows)} existing QAD windows")
            
            # Default state ID from URLs.md if none provided
            if not self.state_id:
                self.state_id = "413cf726-8a34-49b5-a9ee-02f9bffa42fc"  # EDI_Customer state-id
                self.logger.info(f"Using default state ID: {self.state_id}")
            
            self.logger.info(f"Opening QAD with state ID: {self.state_id}")
            
            # Setup Edge for QAD protocol
            self.logger.info("Setting up Edge for QAD protocol...")
            
            # QAD URL with state ID
            qad_url = f"qadsh://browse/invoke?state-id={self.state_id}"
            self.logger.info(f"Navigating to QAD URL: {qad_url}")
            
            # Navigate to QAD URL with retry mechanism
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    self.driver.get(qad_url)
                    self.logger.info("Successfully navigated to QAD URL")
                    break
                except Exception as e:
                    self.logger.warning(f"Attempt {attempt+1}/{max_retries} failed: {str(e)}")
                    if attempt == max_retries - 1:
                        raise Exception(f"Failed to navigate to QAD URL after {max_retries} attempts")
                    time.sleep(2)
            
            self.logger.info("STEP 2: Handling protocol dialog...")
            time.sleep(5)  # Wait for protocol dialog to appear
            
            # Press Tab to focus on Open button
            self.logger.info("STEP 3: Pressing Tab to focus on Open button...")
            send_keys("{TAB}")
            time.sleep(1)
            
            # Press Enter to click Open
            self.logger.info("STEP 4: Pressing Enter to click Open...")
            send_keys("{ENTER}")
            time.sleep(5)  # Wait for QAD to start
            
            # Wait for QAD login window
            self.logger.info("STEP 5: Waiting for QAD login window...")
            app, window = self._find_login_window()
            
            if not app or not window:
                # Try again with a different approach if login window not found
                self.logger.warning("Login window not found with standard approach, trying alternative method...")
                time.sleep(5)
                
                # Try to find any window with QAD in the title
                windows = find_windows(title_re=".*QAD.*")
                if windows:
                    window_handle = windows[0]
                    app = Application().connect(handle=window_handle)
                    window = app.window(handle=window_handle)
                    self.logger.info(f"Found QAD window with title: {window.window_text()}")
                else:
                    raise Exception("Could not find QAD login window after multiple attempts")
            
            # Enter credentials
            self.logger.info("STEP 6: Entering credentials...")
            
            # Enter username
            self.logger.info("Entering username...")
            window.set_focus()
            send_keys(self.username)
            time.sleep(1)
            
            # Move to password field
            self.logger.info("STEP 7: Username entered, moving to password...")
            send_keys("{TAB}")
            time.sleep(1)
            
            # Enter password
            self.logger.info("Entering password...")
            send_keys(self.password)
            time.sleep(1)
            
            # Submit login form
            self.logger.info("STEP 8: Submitting login form...")
            send_keys("{ENTER}")
            
            # Wait for login to complete and main window to appear
            self.logger.info("Waiting for main window to appear after login...")
            time.sleep(30)  # Increased wait time for main window
            
            # Verify login success by checking for QAD windows
            qad_windows = self._find_qad_windows()
            if not qad_windows:
                raise Exception("Login failed: No QAD windows found after login attempt")
                
            self.logger.info("SUCCESS: Login sequence completed")
            return True
            
        except Exception as e:
            self.logger.error(f"Login failed: {str(e)}")
            # Take screenshot for debugging
            try:
                screenshot_path = os.path.join('logs', f'login_error_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png')
                self.driver.save_screenshot(screenshot_path)
                self.logger.info(f"Screenshot saved to {screenshot_path}")
            except:
                pass
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
        Export data from QAD to Excel
        """
        try:
            self.logger.info("STEP 11: Starting Excel export process...")
            
            # Wait for QAD menu to fully load
            self.logger.info("Waiting for QAD menu to fully load...")
            time.sleep(30)  # Increased wait time for menu to load
            
            # Find QAD windows
            qad_windows = self._find_qad_windows()
            if not qad_windows:
                raise Exception("No QAD windows found for export")
            
            # Connect to the first QAD window
            window_handle = qad_windows[0]
            app = Application().connect(handle=window_handle)
            window = app.window(handle=window_handle)
            
            self.logger.info(f"Connected to QAD window: {window.window_text()}")
            
            # Set focus on the QAD window
            window.set_focus()
            time.sleep(2)
            
            # Press Alt to open menu
            self.logger.info("STEP 12: Opening menu with Alt key...")
            window.type_keys("%")  # Alt key
            time.sleep(2)
            
            # Press Enter to select first menu item
            self.logger.info("STEP 13: Pressing Enter to select first menu item...")
            window.type_keys("{ENTER}")
            time.sleep(2)
            
            # Press Down Arrow and Enter to navigate to export
            self.logger.info("STEP 14: Pressing Down Arrow and Enter...")
            window.type_keys("{DOWN}{ENTER}")
            time.sleep(2)
            
            # Press Down Arrow and Enter again for Excel export
            self.logger.info("STEP 15: Pressing Down Arrow and Enter again for Excel export...")
            window.type_keys("{DOWN}{ENTER}")
            time.sleep(5)  # Wait for Excel export dialog
            
            # Handle Excel export dialog if needed
            self.logger.info("STEP 16: Waiting for Excel file to open...")
            time.sleep(10)  # Wait for export to complete
            
            self.logger.info("STEP 17: Export completed successfully")
            
            # Wait for Excel file to be saved
            time.sleep(5)
            
            return True
            
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
