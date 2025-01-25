### Get details for all recent POs from ShipHero

# This script extends E_get_po_details.py to fetch all POs created since the last run
# It maintains a timestamp of the last successful run and fetches all POs created after that time
# The script processes multiple POs and saves their details to the same CSV format

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import time
import os
from dotenv import load_dotenv
import json
import streamlit as st
import pandas as pd

@st.cache_resource
def get_driver():
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--headless")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()),
        options=options
    )

def load_last_run_timestamp():
    try:
        # Use Streamlit's session state to persist the timestamp
        if 'last_run' not in st.session_state:
            st.session_state.last_run = None
        return st.session_state.last_run
    except Exception as e:
        print(f"Error loading timestamp: {e}")
        return None

def save_last_run_timestamp():
    try:
        current_time = datetime.now()
        st.session_state.last_run = current_time
    except Exception as e:
        print(f"Error saving timestamp: {e}")

def get_recent_pos():
    # Load environment variables
    load_dotenv()
    username = os.getenv('SHIPHERO_SANDBOX_USERNAME')
    password = os.getenv('SHIPHERO_SANDBOX_PASSWORD')

    # Initialize the driver using the cached function
    print("Setting up the Selenium WebDriver...")
    driver = get_driver()

    try:
        # Navigate to login page and perform login
        print("Navigating to login page...")
        driver.get("https://app.shiphero.com/account/login")

        # Login steps (updated selectors)
        email_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#username"))
        )
        email_input.clear()
        email_input.send_keys(username)

        continue_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Continue']"))
        )
        continue_button.click()

        password_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#password"))
        )
        password_input.clear()
        password_input.send_keys(password)

        continue_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Continue']"))
        )
        continue_button.click()

        # Wait for dashboard and navigate to PO page
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#your_orders_info"))
        )
        
        # Navigate to purchase orders page
        driver.get("https://app.shiphero.com/dashboard/purchase-orders")
        
        # Wait for the table to load
        table = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "your_orders"))
        )

        # Wait for processing to complete
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located((By.ID, "your_orders_processing"))
        )

        # Sort by Created Date in reverse chronological order
        print("Sorting table by most recent date first...")
        created_date_header = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//th[contains(text(), 'Created Date')]"))
        )
        
        time.sleep(1)
        
        # Click until we get descending order (newest first)
        attempts = 0
        while attempts < 3:
            sort_attribute = created_date_header.get_attribute("aria-sort")
            if sort_attribute == "descending":
                print("Table is sorted by newest first")
                break
            print(f"Current sort: {sort_attribute}. Clicking to change sort order...")
            created_date_header.click()
            time.sleep(1)
            attempts += 1

        # Wait for table to refresh after sorting
        time.sleep(2)
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located((By.ID, "your_orders_processing"))
        )

        # Get last run timestamp
        last_run = load_last_run_timestamp()
        if last_run:
            print(f"Last run was at: {last_run}")
        else:
            print("No previous run found - will process all POs")

        # First pass: Collect all PO links that need processing
        print("Collecting PO links to process...")
        po_links_to_process = []
        
        # Get the sorted PO rows
        po_rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#your_orders tbody tr"))
        )

        for row in po_rows:
            try:
                # Get creation date
                date_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)")  # Created Date column
                date_text = date_cell.text.strip()
                
                print(f"Found date in table: {date_text}")
                
                try:
                    created_date = datetime.strptime(date_text, '%m/%d/%Y %I:%M %p')
                except ValueError as e:
                    print(f"Warning: Could not parse date '{date_text}' from table: {str(e)}")
                    continue

                # Check if PO is newer than last run
                if not last_run or created_date > last_run:
                    po_link = row.find_element(By.CSS_SELECTOR, ".btn-info").get_attribute("href")
                    po_number = row.find_element(By.CSS_SELECTOR, "td:nth-child(3)").text  # PO Number column
                    print(f"Found new PO to process: {po_number} from {created_date}")
                    po_links_to_process.append(po_link)
                else:
                    print(f"Reached PO from {created_date} which is older than last run")
                    break

            except Exception as e:
                print(f"Error checking PO row: {str(e)}")
                continue

        print(f"Found {len(po_links_to_process)} POs to process")

        # Instead of writing to a local file, collect the data in a list
        po_data = []
        
        # Process POs and collect data
        for po_url in po_links_to_process:
            try:
                print(f"Processing PO at: {po_url}")
                driver.get(po_url)
                
                # Wait for product table
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "po-order-items"))
                )
                
                # Get PO details
                po_number = driver.find_element(By.NAME, "po_number").get_attribute("value")
                warehouse = driver.find_element(By.XPATH, "//label[text()='Warehouse']/following-sibling::div/strong").text
                total_qty = driver.find_element(By.XPATH, "//label[text()='Total Quantity:']/following-sibling::div/strong").text
                vendor = driver.find_element(By.XPATH, "//label[text()='Vendor:']/following-sibling::div/strong").text
                
                print(f"Processing PO {po_number} from vendor: {vendor}")

                # Get product rows
                product_rows = driver.find_elements(By.CSS_SELECTOR, "table.po-order-items tbody tr")
                
                # Instead of writing directly to file, collect in memory
                current_time = time.strftime("%Y-%m-%d %H:%M:%S")
                for prod_row in product_rows:
                    name_element = prod_row.find_element(By.CSS_SELECTOR, "td:nth-child(2) a")
                    product_name = name_element.text.replace(',', ' ')
                    
                    sku_text = prod_row.find_element(By.CSS_SELECTOR, "td:nth-child(2)").text
                    sku_line = [line for line in sku_text.split('\n') if 'Sku:' in line][0]
                    sku = sku_line.split('Sku: ')[1].strip()
                    
                    qty_input = prod_row.find_element(By.CSS_SELECTOR, "input.qty_input")
                    ordered_qty = qty_input.get_attribute("value")
                    
                    po_data.append({
                        'Time': current_time,
                        'PO Number': po_number,
                        'Warehouse': warehouse.replace(',', ' '),
                        'Vendor': vendor.replace(',', ' '),
                        'Total Qty': total_qty,
                        'Name': product_name,
                        'SKU': sku,
                        'Ordered Qty': ordered_qty,
                        'PO URL': po_url
                    })
                
                print(f"Processed PO: {po_number}")

            except Exception as e:
                print(f"Error processing PO {po_url}: {str(e)}")
                continue

        # Convert collected data to DataFrame
        if po_data:
            df = pd.DataFrame(po_data)
            
            # Display the data in Streamlit
            st.write("### Recent PO Details")
            st.dataframe(df)
            
            # Offer download button
            csv = df.to_csv(index=False)
            st.download_button(
                label="Download PO Details as CSV",
                data=csv,
                file_name="po_details.csv",
                mime="text/csv"
            )
        
        # Save current run timestamp
        save_last_run_timestamp()
        print("Successfully processed all recent POs")
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        return False
    finally:
        driver.quit()

if __name__ == "__main__":
    st.title("ShipHero PO Details Extractor")
    
    if st.button("Get Recent POs"):
        with st.spinner("Fetching PO details..."):
            success = get_recent_pos()
            if success:
                st.success("Successfully extracted all recent PO details")
            else:
                st.error("Failed to extract PO details") 