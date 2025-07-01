import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException # Added NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import re

def scrape_ipo_subscription_data(excel_file_path, url_column_name="URL_for_IPO_details"):
    """
    Scrapes QIB, NII, RII, and Total subscription data from IPO company pages.

    Args:
        excel_file_path (str): Path to the Excel file containing IPO links.
        url_column_name (str): The name of the column in the Excel file that
                                contains the corrected IPO URLs.

    Returns:
        pandas.DataFrame: A DataFrame containing the scraped data, or None if an error occurs.
    """
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path)
        
        # Get the list of URLs from the specified column
        ipo_urls = df[url_column_name].dropna().tolist()

    except FileNotFoundError:
        print(f"Error: Excel file not found at {excel_file_path}")
        return None
    except KeyError:
        print(f"Error: Column '{url_column_name}' not found in the Excel file.")
        print(f"Available columns: {df.columns.tolist()}")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

    # Setup Chrome options
    chrome_options = Options()
    # Keep headless mode commented out for now to observe scrolling
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--log-level=3") # Suppress verbose logging
    chrome_options.add_argument("--disable-gpu") 
    chrome_options.add_argument("--window-size=1920,1080") 
    chrome_options.add_argument("--start-maximized") 
    chrome_options.add_argument("--incognito") 
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"]) 

    # --- NEW: Additional options to try and suppress more verbose logging ---
    # These might help with some of the GPU/TensorFlow Lite messages, but might not remove all.
    chrome_options.add_argument("--disable-logging")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-setuid-sandbox")
    
    # Initialize Chrome WebDriver
    try:
        service = Service(ChromeDriverManager().install()) 
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(60) # Increased page load timeout
    except Exception as e:
        print(f"Error initializing WebDriver: {e}")
        return None

    scraped_data = []

    for url in ipo_urls:
        company_name = "N/A" 
        qib_sub = "N/A"
        nii_sub = "N/A"
        rii_sub = "N/A"
        total_sub = "N/A"

        try:
            print(f"Navigating to: {url}")
            driver.get(url)

            # --- NEW: Scroll down to ensure content loads ---
            # Scroll to the bottom of the page
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(3) # Give some time for content to load after scrolling

            # Scroll back up slightly if the target element might be at the very bottom
            # and obscured by footers or sticky elements. This might not be necessary.
            # driver.execute_script("window.scrollTo(0, document.body.scrollHeight - 200);")
            # time.sleep(1) 
            # --- END NEW SCROLLING ---

            # Wait for the table caption to be present (adjust timeout as needed)
            WebDriverWait(driver, 30).until( 
                EC.presence_of_element_located((By.XPATH, "//caption[contains(text(), 'IPO Bidding Live Updates from BSE, NSE')]"))
            )
            
            # Get the page source after dynamic content loads
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Extract company name from the page title or a prominent heading
            title_match = re.search(r'^(.*?) IPO', driver.title)
            if title_match:
                company_name = title_match.group(1).strip()
            else:
                # Fallback to URL parsing for company name
                url_parts = url.split('/')
                for part in url_parts:
                    if '-ipo' in part:
                        company_name = part.replace('-ipo', '').replace('-', ' ').title()
                        break
            
            # Find the table by its caption
            ipo_table = soup.find('caption', string=lambda text: text and 'IPO Bidding Live Updates from BSE, NSE' in text)
            
            if ipo_table:
                # Find the parent table element
                table_element = ipo_table.find_parent('table')
                if table_element:
                    # Find all table rows in the tbody
                    rows = table_element.find('tbody').find_all('tr')
                    
                    if rows:
                        # The last row contains the final subscription data
                        last_row = rows[-1]
                        
                        # Extract data using data-title attribute
                        qib_td = last_row.find('td', {'data-title': lambda x: x and 'QIB' in x})
                        nii_td = last_row.find('td', {'data-title': lambda x: x and 'NII' in x})
                        rii_td = last_row.find('td', {'data-title': lambda x: x and 'RII' in x})
                        total_td = last_row.find('td', {'data-title': lambda x: x and 'Total' in x})

                        qib_sub = qib_td.get_text(strip=True) if qib_td else "N/A"
                        nii_sub = nii_td.get_text(strip=True) if nii_td else "N/A"
                        rii_sub = rii_td.get_text(strip=True) if rii_td else "N/A"
                        total_sub = total_td.get_text(strip=True) if total_td else "N/A"
                    else:
                        print(f"No rows found in the table for {company_name} ({url})")
                else:
                    print(f"Parent table element not found for caption for {company_name} ({url})")
            else:
                print(f"IPO Bidding Live Updates from BSE, NSE table caption not found for {company_name} ({url})")

        except TimeoutException:
            print(f"Timeout while loading or finding element on {url}. Page took too long to respond. Skipping this URL.")
        except NoSuchElementException:
            print(f"Element not found on {url} after scrolling. Table structure might be different. Skipping this URL.")
        except WebDriverException as e:
            print(f"WebDriver error for {url}: {e}. This might indicate a browser crash or disconnection. Skipping this URL.")
        except Exception as e:
            print(f"Generic error scraping {url}: {e}")

        # Append the collected data
        scraped_data.append({
            "Company Name": company_name,
            "IPO Link": url,
            "QIB Subscription": qib_sub,
            "NII Subscription": nii_sub,
            "RII Subscription": rii_sub,
            "Total Subscription": total_sub
        })
        time.sleep(1) # Be polite and wait a bit between requests

    driver.quit() # Close the browser

    # Create a DataFrame from the scraped data
    output_df = pd.DataFrame(scraped_data)
    return output_df

# --- How to use the scraper ---
# IMPORTANT: Replace 'your_excel_file.xlsm' with the actual path to your Excel file.
# And replace 'Corrected_Link_Column' with the exact name of the column
# where you stored the /ipo links.
excel_input_file = 'GMP_enabled.xlsm' 
url_col = 'URL_for_IPO_details' # Example column name

print("Starting IPO data scraping...")
scraped_ipo_df = scrape_ipo_subscription_data(excel_input_file, url_col)

if scraped_ipo_df is not None:
    # Save the results to a new Excel file
    output_excel_file = 'scraped_ipo_subscription_data.xlsx'
    scraped_ipo_df.to_excel(output_excel_file, index=False)
    print(f"\nScraping complete! Data saved to '{output_excel_file}'")
    print("\nSample of scraped data:")
    print(scraped_ipo_df.head())
else:
    print("\nScraping failed or no data was retrieved.")
