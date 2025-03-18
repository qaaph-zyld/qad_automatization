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
from pywinauto import Desktop, Application
from pywinauto.findwindows import find_windows
from pywinauto.keyboard import send_keys
from pathlib import Path
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui

class QADAutomation:
    def __init__(self, username: str, password: str, state_id: str = None, force: bool = False):
        """Initialize QAD automation"""
        self.username = username
        self.password = password
        self.state_id = state_id
        self.force = force
        self.driver = None
        self.excel_file_path = None
        
        # Set up logging
        self.logger = self._setup_logging()
        
        # Check for existing QAD processes first (unless force is True)
        if not self.force:
            self._check_and_close_qad_applications()
            
    def _setup_logging(self):
        """Set up logging configuration"""
        logger = logging.getLogger(__name__)
        
        # Create logs directory if it doesn't exist
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        # Set up file handler
        log_file = os.path.join(log_dir, f"qad_automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.INFO)
        
        # Set up console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Set up formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Add handlers to logger
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        logger.setLevel(logging.INFO)
        
        return logger

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

    def _init_edge(self):
        """
        Initialize Edge WebDriver with required options
        """
        try:
            # Initialize Edge WebDriver with specific options
            edge_options = webdriver.EdgeOptions()
            edge_options.use_chromium = True
            edge_options.add_argument("--start-maximized")
            
            # Add protocol handler for qadsh
            edge_options.add_argument('--protocol-handler.allowed-origin=qadsh://*')
            
            # Initialize Edge WebDriver
            self.driver = webdriver.Edge(options=edge_options)
            self.logger.info("Edge WebDriver initialized successfully")
            return True
        except Exception as e:
            self.logger.error(f"Failed to initialize Edge WebDriver: {str(e)}")
            return False

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
        """
        Find all QAD windows
        """
        try:
            self.logger.info("Finding QAD windows...")
            
            # Import pywinauto here to avoid circular imports
            from pywinauto import Desktop, Application
            
            # Get all windows
            all_windows = Desktop(backend="uia").windows()
            
            # Filter for QAD windows
            qad_windows = []
            for window in all_windows:
                try:
                    window_text = window.window_text()
                    if any(term in window_text for term in ["QAD", "Progress", "QAD Enterprise Applications"]):
                        self.logger.info(f"Found QAD window: {window_text}")
                        qad_windows.append(window)
                except:
                    pass
            
            if not qad_windows:
                self.logger.warning("No QAD windows found")
            else:
                self.logger.info(f"Found {len(qad_windows)} QAD windows")
                
            return qad_windows
        except Exception as e:
            self.logger.error(f"Error finding QAD windows: {str(e)}")
            return []
    
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

    def login(self):
        """
        Login to QAD with provided credentials
        """
        try:
            self.logger.info("STEP 8: Logging in to QAD...")
            
            # Wait for login window to appear
            time.sleep(5)
            
            # Enter username
            self.logger.info(f"Entering username: {self.username}")
            pyautogui.write(self.username)
            
            # Tab to password field
            pyautogui.press('tab')
            time.sleep(1)
            
            # Enter password
            self.logger.info("Entering password")
            pyautogui.write(self.password)
            
            # Press Enter to submit
            pyautogui.press('enter')
            self.logger.info("Login submitted, waiting for QAD to load...")
            
            # Wait for QAD to load
            time.sleep(30)
            
            self.logger.info("QAD login successful")
            return True
        except Exception as e:
            self.logger.error(f"Login failed: {str(e)}")
            return False

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
    
    def _export_to_excel(self):
        """Export QAD data to Excel"""
        try:
            self.logger.info("STEP 11: Starting Excel export process...")
            
            # Wait for QAD menu to fully load
            self.logger.info("Waiting for QAD menu to fully load...")
            time.sleep(30)  # Increased wait time for menu to load
            
            # Find all QAD windows
            self.logger.info("STEP 12: Finding QAD windows...")
            qad_windows = self._find_qad_windows()
            
            if not qad_windows:
                self.logger.error("No QAD windows found")
                return False
                
            # Connect to the first QAD window found
            self.logger.info(f"Found {len(qad_windows)} QAD windows, connecting to the first one...")
            qad_window = qad_windows[0]
            
            # Set focus on the QAD window
            self.logger.info("STEP 13: Setting focus on QAD window...")
            qad_window.set_focus()
            time.sleep(2)
            
            # Press Alt key (%) to open the main menu
            self.logger.info("STEP 14: Pressing Alt key to open main menu...")
            pyautogui.press('alt')
            time.sleep(2)
            
            # Press Enter to select first menu item
            self.logger.info("STEP 15: Pressing Enter to select first menu item...")
            pyautogui.press('enter')
            time.sleep(2)
            
            # Press Down Arrow and Enter to navigate to export
            self.logger.info("STEP 16: Pressing Down Arrow and Enter to navigate to export...")
            pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(2)
            
            # Press Down Arrow and Enter again for Excel export
            self.logger.info("STEP 17: Pressing Down Arrow and Enter again for Excel export...")
            pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(5)  # Wait for Excel export dialog
            
            # Wait for Excel file to open
            self.logger.info("STEP 18: Waiting for Excel file to open...")
            time.sleep(10)
            
            # Save the Excel file with a specific name
            self.logger.info("STEP 19: Saving Excel file as EDI_Demand.xlsx...")
            
            # Press Alt+F+A to open Save As dialog
            pyautogui.hotkey('alt', 'f', 'a')
            time.sleep(2)
            
            # Press Y3 to select Excel 97-2003 format
            pyautogui.press('y')
            pyautogui.press('3')
            time.sleep(2)
            
            # Type "EDI_Demand" as the file name
            pyautogui.write("EDI_Demand")
            time.sleep(1)
            
            # Press Enter to save
            pyautogui.press('enter')
            time.sleep(2)
            
            # Press Enter again to confirm (if dialog appears)
            pyautogui.press('enter')
            time.sleep(2)
            
            # Set the Excel file path
            self.excel_file_path = r"C:\Users\ajelacn\AppData\Local\Temp\Shell\EDI_Demand.xlsx"
            self.logger.info(f"Excel file saved to: {self.excel_file_path}")
            
            return True
        except Exception as e:
            self.logger.error(f"Export to Excel failed: {str(e)}")
            return False

    def _start_qad(self):
        """Start QAD application using URL protocol"""
        try:
            # Load QAD URL from URLs.md
            self.logger.info("Loading QAD URL from URLs.md")
            with open("URLs.md", "r") as f:
                url_content = f.read()
                # Extract URL from the file
                qad_url = url_content.strip().split(" ")[0]
                
            if not qad_url:
                self.logger.error("Failed to load QAD URL from URLs.md")
                return False
                
            self.logger.info("Opening QAD with URL from URLs.md")
            self.logger.info(f"Navigating to QAD URL: {qad_url}")
            
            # Navigate to QAD URL
            self.driver.get(qad_url)
            
            # Wait for protocol dialog
            self.logger.info("QAD URL loaded, waiting for protocol dialog...")
            time.sleep(5)
            
            # Handle protocol dialog
            self.logger.info("Handling protocol dialog...")
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('enter')
            
            # Wait for QAD login window
            self.logger.info("Protocol dialog handled, waiting for QAD login window...")
            time.sleep(10)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error starting QAD: {str(e)}")
            return False

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

    def run(self):
        """Run the QAD automation process"""
        try:
            # Check and close existing QAD applications first
            if not self.force and not self._check_and_close_qad_applications():
                self.logger.error("User canceled the operation")
                return False
            
            # Initialize Edge WebDriver only after checking for QAD applications
            self.logger.info("Initializing Edge WebDriver...")
            if not self._init_edge():
                self.logger.error("Failed to initialize Edge WebDriver")
                return False
                
            # Start QAD application
            if not self._start_qad():
                self.logger.error("Failed to start QAD application")
                return False
                
            # Login to QAD
            if not self.login():
                self.logger.error("Failed to login to QAD")
                return False
                
            # Export to Excel
            if not self._export_to_excel():
                self.logger.error("Failed to export to Excel")
                return False
                
            # Return the path to the exported Excel file
            return self.excel_file_path
                
        except Exception as e:
            self.logger.error(f"Automation failed: {str(e)}")
            return False
        finally:
            # Clean up
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

def main():
    """Main function"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='QAD Automation')
    parser.add_argument('--username', required=True, help='QAD username')
    parser.add_argument('--password', required=True, help='QAD password')
    parser.add_argument('--state-id', help='QAD state ID for custom folder navigation')
    parser.add_argument('--force', action='store_true', help='Force execution even if QAD processes are running')
    args = parser.parse_args()
    
    # Set up logging for main function
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger("main")
    
    try:
        # Create QAD automation instance
        logger.info("Creating QAD automation instance...")
        qad = QADAutomation(args.username, args.password, args.state_id, args.force)
        
        # Run QAD automation
        logger.info("Starting QAD automation process...")
        excel_file_path = qad.run()
        
        # Return the Excel file path to be used by other scripts
        if excel_file_path:
            logger.info(f"QAD automation completed successfully")
            print(f"EXCEL_FILE_PATH:{excel_file_path}")
            return 0
        else:
            logger.error("QAD automation failed")
            return 1
    except Exception as e:
        logger.error(f"Error in QAD automation: {str(e)}")
        return 1
    
if __name__ == '__main__':
    sys.exit(main())
