#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Full QAD Automation Script
--------------------------
This script runs the complete QAD automation process by:
1. Checking Edge browser status
2. Running qad-edge-automation.py to export data from QAD to Excel
3. Running analyze_demand.py to analyze the exported data

Usage:
    python run_full_automation.py --username <username> --password <password> [--state-id <state-id>] [--force]
"""

import os
import sys
import time
import psutil
import argparse
import subprocess
import logging
import winreg
import pyautogui
import pywinauto
from datetime import datetime
from pywinauto import Application
from pywinauto.findwindows import find_windows
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def check_edge_status():
    """Check if Microsoft Edge is running and if not, try to start it"""
    logger = logging.getLogger(__name__)
    
    def is_edge_running():
        """Check if Edge is running"""
        for proc in psutil.process_iter(['name']):
            try:
                if 'msedge.exe' in proc.info['name'].lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return False
    
    def get_edge_path():
        """Get Microsoft Edge installation path from registry"""
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)
            path = winreg.QueryValue(key, None)
            winreg.CloseKey(key)
            return path
        except WindowsError:
            return None
    
    # Check if Edge is already running
    if is_edge_running():
        logger.info("Microsoft Edge is already running")
        return True
    
    # Try to start Edge
    logger.info("Microsoft Edge is not running, attempting to start it...")
    edge_path = get_edge_path()
    
    if not edge_path:
        logger.error("Could not find Microsoft Edge installation path")
        return False
    
    try:
        subprocess.Popen([edge_path])
        logger.info("Started Microsoft Edge")
        
        # Wait for Edge to initialize
        max_retries = 10
        for i in range(max_retries):
            if is_edge_running():
                logger.info("Microsoft Edge is now running")
                time.sleep(5)  # Give Edge time to fully initialize
                return True
            time.sleep(1)
        
        logger.error("Microsoft Edge failed to start properly")
        return False
    except Exception as e:
        logger.error(f"Failed to start Microsoft Edge: {str(e)}")
        return False

def setup_logging():
    """Set up logging configuration"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    log_file = os.path.join(log_dir, f"full_automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stderr)
        ]
    )
    
    return logging.getLogger(__name__)

def find_qad_windows():
    """Find all QAD windows using various title patterns"""
    logger = logging.getLogger(__name__)
    qad_windows = []
    
    # Try to find windows with QAD in the title
    try:
        qad_pattern_windows = find_windows(title_re=".*QAD.*")
        logger.info(f"Found {len(qad_pattern_windows)} windows with pattern '.*QAD.*'")
        qad_windows.extend(qad_pattern_windows)
    except Exception as e:
        logger.warning(f"Error finding QAD windows: {str(e)}")
    
    # Try to find windows with Enterprise in the title
    try:
        enterprise_windows = find_windows(title_re=".*Enterprise.*")
        logger.info(f"Found {len(enterprise_windows)} windows with pattern '.*Enterprise.*'")
        qad_windows.extend(enterprise_windows)
    except Exception as e:
        logger.warning(f"Error finding Enterprise windows: {str(e)}")
    
    # Try to find windows with Browse in the title (common in QAD)
    try:
        browse_windows = find_windows(title_re=".*Browse.*")
        if browse_windows:
            logger.info(f"Found {len(browse_windows)} windows with pattern '.*Browse.*'")
            qad_windows.extend(browse_windows)
    except Exception as e:
        logger.warning(f"Error finding Browse windows: {str(e)}")
        
    # Filter out any windows that are clearly not QAD windows
    filtered_windows = []
    for window_handle in qad_windows:
        try:
            window = Application().connect(handle=window_handle).window(handle=window_handle)
            title = window.window_text()
            if "Windsurf" in title or "Visual Studio" in title or "Code" in title:
                logger.info(f"Skipping non-QAD window: {title}")
                continue
            filtered_windows.append(window_handle)
        except Exception as e:
            logger.warning(f"Error checking window title: {str(e)}")
            
    logger.info(f"Found {len(filtered_windows)} QAD windows")
    return filtered_windows

def handle_qad_export(driver, logger):
    """Handle the QAD export process"""
    try:
        # Wait for QAD menu to load
        logger.info("Waiting for QAD menu to load...")
        time.sleep(30)
        
        # Find QAD windows
        logger.info("Finding QAD windows...")
        qad_windows = find_qad_windows()
        if not qad_windows:
            logger.error("No QAD windows found")
            return False
        
        # Focus first QAD window
        logger.info("Focusing first QAD window...")
        window = Application().connect(handle=qad_windows[0]).window(handle=qad_windows[0])
        window.set_focus()
        time.sleep(2)
        
        # Press Alt
        logger.info("Pressing Alt...")
        pyautogui.press('alt')
        time.sleep(1)
        
        # Press Enter
        logger.info("Pressing Enter...")
        pyautogui.press('enter')
        time.sleep(1)
        
        # Press Down Arrow + Enter twice
        for _ in range(2):
            logger.info("Pressing Down Arrow + Enter...")
            pyautogui.press('down')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
        
        # Wait for Excel to open
        logger.info("Waiting for Excel to open...")
        time.sleep(10)
        
        # Press Alt > F > A
        logger.info("Opening Save As dialog...")
        pyautogui.press('alt')
        time.sleep(1)
        pyautogui.press('f')
        time.sleep(1)
        pyautogui.press('a')
        time.sleep(1)
        
        # Press Y > 3 for Excel format
        logger.info("Selecting Excel format...")
        pyautogui.press('y')
        time.sleep(1)
        pyautogui.press('3')
        time.sleep(1)
        
        # Type EDI_Demand and save
        logger.info("Saving file as EDI_Demand...")
        pyautogui.write('EDI_Demand')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        
        # Press Tab > Enter for potential overwrite
        logger.info("Handling potential overwrite...")
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)  # Wait for save to complete
        
        # Press Alt > F > C to close Excel
        logger.info("Closing Excel...")
        pyautogui.press('alt')
        time.sleep(1)
        pyautogui.press('f')
        time.sleep(1)
        pyautogui.press('c')
        time.sleep(1)
        
        excel_file_path = os.path.join(os.environ['TEMP'], 'Shell', 'EDI_Demand.xlsx')
        logger.info(f"Excel file saved to: {excel_file_path}")
        return excel_file_path
        
    except Exception as e:
        logger.error(f"Error during QAD export: {str(e)}")
        return False

def main():
    """Main function to run the full automation process"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Full QAD Automation Process')
    parser.add_argument('--username', required=True, help='QAD username')
    parser.add_argument('--password', required=True, help='QAD password')
    parser.add_argument('--state-id', help='QAD state ID for custom folder navigation')
    parser.add_argument('--force', action='store_true', help='Force execution even if QAD processes are running')
    args = parser.parse_args()
    
    # Set up logging
    logger = setup_logging()
    logger.info("Starting full QAD automation process")
    
    try:
        # Step 0: Check Edge browser status
        logger.info("Step 0: Checking Microsoft Edge status")
        if not check_edge_status():
            logger.error("Failed to ensure Microsoft Edge is running")
            return 1
        
        # Step 1: Run QAD Edge Automation to export data
        logger.info("Step 1: Running QAD Edge Automation to export data")
        
        # Initialize Edge WebDriver
        logger.info("Initializing Edge WebDriver...")
        edge_options = Options()
        edge_options.add_argument('--start-maximized')
        edge_options.add_argument('--protocol-handler.allowed-origin=qadsh://*')
        driver = webdriver.Edge(options=edge_options)
        
        try:
            # Load QAD URL from URLs.md
            logger.info("Loading QAD URL from URLs.md")
            with open("URLs.md", "r") as f:
                url_content = f.read()
                qad_url = url_content.strip().split(" ")[0]
            
            if not qad_url:
                logger.error("Failed to load QAD URL from URLs.md")
                return 1
            
            # Navigate to QAD URL
            logger.info(f"Navigating to QAD URL: {qad_url}")
            driver.get(qad_url)
            
            # Wait for protocol dialog
            logger.info("Waiting for protocol dialog...")
            time.sleep(5)
            
            # Handle protocol dialog
            logger.info("Handling protocol dialog...")
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('enter')
            
            # Wait for QAD login window
            logger.info("Waiting for QAD login window...")
            time.sleep(10)
            
            # Login to QAD
            logger.info("Logging in to QAD...")
            pyautogui.write(args.username)
            time.sleep(1)
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.write(args.password)
            time.sleep(1)
            pyautogui.press('enter')
            
            # Wait for QAD to load
            logger.info("Waiting for QAD to load...")
            time.sleep(30)
            
            # Handle QAD export process
            excel_file_path = handle_qad_export(driver, logger)
            if not excel_file_path:
                logger.error("Failed to export data from QAD")
                return 1
            
            # Step 2: Run Analyze Demand to process the exported data
            logger.info("Step 2: Running Analyze Demand to process the exported data")
            
            # Build the command for analyze_demand.py
            analyze_cmd = [
                sys.executable, 
                "analyze_demand.py",
                "--excel-file", excel_file_path
            ]
            
            # Run the Analyze Demand script
            logger.info(f"Running command: {' '.join(analyze_cmd)}")
            analyze_process = subprocess.run(analyze_cmd, capture_output=True, text=True)
            
            # Log the output regardless of success/failure
            if analyze_process.stdout:
                logger.info("Analyze Demand output:")
                for line in analyze_process.stdout.splitlines():
                    logger.info(f"  {line}")
                    
            if analyze_process.stderr:
                logger.info("Analyze Demand error output:")
                for line in analyze_process.stderr.splitlines():
                    logger.info(f"  {line}")
            
            # Check if the Analyze Demand was successful
            if analyze_process.returncode != 0:
                logger.error(f"Analyze Demand failed with return code {analyze_process.returncode}")
                return 1
                
            logger.info("Full QAD automation process completed successfully")
            return 0
            
        finally:
            # Clean up
            try:
                driver.quit()
            except:
                pass
            
    except Exception as e:
        logger.error(f"Error during full automation process: {str(e)}")
        return 1
        
if __name__ == "__main__":
    sys.exit(main())
