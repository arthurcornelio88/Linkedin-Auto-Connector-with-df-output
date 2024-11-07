from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common import action_chains
from webdriver_manager.chrome import ChromeDriverManager
from configparser import ConfigParser
from colorama import Fore, Style, init
import time, requests
from urllib.parse import quote
import pandas as pd
import re #Arthur's modif
import nltk #Arthur's modif
import datetime #Arthur's modif
import os #Arthur's modif
from urllib.parse import unquote #Arthur's modif

# Initialize colorama
init(autoreset=True)

# Initialize config parser
config = ConfigParser()
config_file = 'setup.ini'
config.read(config_file)

#Arthur's modif
linkedin_url_list = []
linkedin_profile_name = []
#End modif

# Download necessary NLTK resources (run this once)
nltk.download('punkt')

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--log-level=2")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")

    # service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome( options=chrome_options)
    return driver

def save_cookie(driver:webdriver.Chrome):
    """Save the cookie to the setup.ini file"""
    li_at_cookie = driver.get_cookie('li_at')['value']
    config.set('LinkedIn', 'li_at', li_at_cookie)
    with open(config_file, 'w') as f:
        config.write(f)

def login_with_cookie(driver:webdriver.Chrome, li_at):
    """Attempt to login with the existing 'li_at' cookie"""
    print(Fore.YELLOW + "Attempting to log in with cookie...")
    driver.get("https://www.linkedin.com")
    driver.add_cookie(
        {
            "name": "li_at",
            "value": f"{li_at}",
            "path": "/",
            "secure": True,
        }
    )
    driver.refresh()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "global-nav-typeahead")))
    print(Fore.GREEN + "[INFO] Logged in with cookie successfully.")

def select_location(driver:webdriver.Chrome, location:str):
    """Select the location in the LinkedIn search filter"""
    try:
        print("Selecting location")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchFilter_geoUrn"))).click()
        time.sleep(1)
        location_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Add a location']")))
        location_input.send_keys(location)
        time.sleep(2)
        driver.find_element(By.XPATH,f"//*[text()='{location.title()}']").click()
        time.sleep(1)
        driver.find_element(By.XPATH,"//button[@aria-label='Apply current filter to show results']").click()
        time.sleep(3)
    except Exception as e:
        print(Fore.RED + f"[INFO] Error selecting location: {e}")

def login_with_credentials(driver:webdriver.Chrome, email:str, password:str):
    """Login using credentials and handle verification code if required"""
    print(Fore.YELLOW + "Logging in with credentials...")
    driver.get("https://www.linkedin.com/login")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))

    driver.find_element(By.ID, "username").send_keys(email)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()

    WebDriverWait(driver, 10).until(
        lambda d: d.find_element(By.ID, "global-nav-typeahead") or
        "Enter the code" in d.page_source
    )

    if "Enter the code" in driver.page_source:
        verification_code = input("[+] Enter the verification code sent to your email: ")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "input__email_verification_pin")))
        driver.find_element(By.ID, "input__email_verification_pin").send_keys(verification_code)
        driver.find_element(By.ID, "email-pin-submit-button").click()

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "global-nav-typeahead")))
    print(Fore.GREEN + "[INFO] Logged in with credentials successfully.")
    save_cookie(driver)

def extract_name(url):
    """
    Extracts the name from a LinkedIn profile URL, using a combination of regex and NLP.

    Args:
        url: The LinkedIn profile URL.

    Returns:
        The extracted name, or None if not found.
    """

    # Split the URL at '?' to remove query parameters
    base_url = url.split('?')[0]

    # Use a regular expression to find the name, allowing for hyphens and spaces
    match = re.search(r"in/([\w ]+(?:-[^\d\W]+)*)(?!\d+)", base_url)

    if match:
        name = unquote(match.group(1).replace("-", " "))
        name = name.replace("%20", " ")

        # Attempt further splitting using NLP if the name contains no spaces
        # or has more than 3 consecutive word characters (potential concatenation)
        if " " not in name or re.search(r"\w{4,}", name):
            words = nltk.word_tokenize(name)
            name = " ".join(words)

        return ' '.join(word.capitalize() for word in name.split())
    else:
        return None
#End modif

def send_connection_request(driver: webdriver.Chrome, limit:str, letter:str, include_notes: bool, message_letter:str, keyword:str, location:str, connection_degree:str):
    """Send a connection request to the specified LinkedIn profile"""

    cnt = 0
    cnt2 = 1
    last_button_found_time = time.time()

    # Define actions here to make it accessible throughout the function
    actions = action_chains.ActionChains(driver)

    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

        # Check for "No results found." message
        try:
            driver.find_element(By.XPATH, "//*[contains(text(), 'No results found.')]")
            print(Fore.RED + "[WARNING] No search results found. Stopping the script.")
            return  # Exit the function if no results
        except NoSuchElementException:
            pass  # Continue if there are search results

        while cnt < limit:
            try:
                # Determine button text and XPaths based on message_letter
                button_text = "Message" if message_letter else "Connect"
                url_xpath_template = f'(//*[text()="{button_text}"]/../../../..//span[@class="entity-result__title-line entity-result__title-line--2-lines "]//a)[{{cnt2}}]'
                name_xpath_template = f'(//*[text()="{button_text}"]/../../../../..//span[@class="entity-result__title-line entity-result__title-line--2-lines "]//a//span)[{{cnt2}}]'

                # Print the XPaths being used for debugging
                print(f"Searching for {button_text} buttons using XPath: //*[text()='{button_text}']/..")

                connect_buttons = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, f"//*[text()='{button_text}']/.."))
                )
                print(f"Number of connect buttons found: {len(connect_buttons)}")

                # Update last_button_found_time
                last_button_found_time = time.time()

                # Reset cnt2 whenever we re-find connect_buttons
                cnt2 = 1

                # Iterate through all connect buttons on the current page
                for connect_button in connect_buttons:
                    if cnt >= limit:
                        break

                    try:
                        print("Cnt : ", cnt, cnt2)

                        # Check if cnt2 is within the bounds of connect_buttons
                        if cnt2 <= len(connect_buttons):
                            if message_letter == "":
                                # Print the url_xpath for debugging
                                url_xpath = url_xpath_template.format(cnt2=cnt2)
                                print(f"Extracting URL using XPath: {url_xpath}")

                                linkedin_container = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, url_xpath)))
                                linkedin_url = linkedin_container.get_attribute('href')

                                # Extract name after getting the linkedin_url
                                name = linkedin_container.text.split(' ')[0].title()

                                # Send connection request with "Connect" button
                                cnt += 1
                                actions.move_to_element(connect_button).perform()
                                time.sleep(1)
                                connect_button.click()

                                time.sleep(1)  # Adjust based on how quickly the modal appears

                                if not include_notes:
                                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Send now"]'))).click()
                                else:
                                    add_note_button = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Add a note"]')))
                                    add_note_button.click()

                                    message_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//textarea[@name="message"]')))
                                    message_box.send_keys(letter.replace("{name}", name).replace("{fullName}", name))
                                    time.sleep(1)  # Adjust based on preference
                                    send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send invitation"]')
                                    driver.execute_script("arguments[0].click();", send_button)

                            elif message_letter != "":
                                # Print the url_xpath and name_xpath for debugging
                                url_xpath = url_xpath_template.format(cnt2=cnt2)
                                name_xpath = name_xpath_template.format(cnt2=cnt2)
                                print(f"Extracting URL using XPath: {url_xpath}")
                                print(f"Extracting name using XPath: {name_xpath}")

                                linkedin_url = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, url_xpath))).get_attribute('href')
                                full_name = driver.find_element(By.XPATH, name_xpath).text.split(' ')[0].title().replace("view", "").replace('\n', '')

                                # Send connection request with "Message" button
                                cnt += 1
                                actions.move_to_element(connect_button).perform()
                                time.sleep(1)
                                connect_button.click()
                                time.sleep(2)

                                try:
                                    driver.find_element(By.XPATH, "//h2[text()='No free personalized invitations left']")
                                    print(Fore.RED + "[ERROR] No free personalized invitations left.")
                                    return
                                except: pass

                                message_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@role='textbox']")))
                                message_box.clear()
                                message_box.send_keys(message_letter.replace("{name}", full_name).replace("{fullName}", full_name))
                                time.sleep(1)
                                send_button = driver.find_element(By.XPATH, '//button[text()="Send"]')
                                send_button.click()
                                time.sleep(1)
                                driver.find_element(By.XPATH, "//button[@class='msg-overlay-bubble-header__control artdeco-button artdeco-button--circle artdeco-button--muted artdeco-button--1 artdeco-button--tertiary ember-view']").click()
                                time.sleep(2)

                            print(Fore.GREEN + f"[INFO] Connection request sent successfully to {linkedin_url}")
                            print("Calling extract_name...")
                            name_in_url = extract_name(linkedin_url)
                            linkedin_url_list.append(linkedin_url)
                            linkedin_profile_name.append(name_in_url)
                            print(f"Appended URL: {linkedin_url}, Name: {name_in_url}")
                            print("---------------------------------------------------------------------------------------------------------------")
                            time.sleep(10)

                            # Check for "Search limit reached" message
                            try:
                                driver.find_element(By.XPATH, "//h2[text()='Search limit reached.']")
                                print(Fore.RED + "[WARNING] Search limit reached. Stopping the script.")
                                break  # Exit the loop if the limit is reached
                            except NoSuchElementException:
                                pass  # No limit reached message, continue

                        else:
                            print(f"Skipping button {cnt2} as it's out of bounds (total buttons: {len(connect_buttons)})")

                        cnt2 += 1  # Increment cnt2 after processing or skipping the button

                        # Re-find connect buttons after each request to avoid stale element references
                        if message_letter == "":
                            connect_buttons = WebDriverWait(driver, 5).until(
                                EC.presence_of_all_elements_located((By.XPATH, "//*[text()='Connect']/.."))
                            )
                        elif message_letter != "":
                            connect_buttons = WebDriverWait(driver, 5).until(
                                EC.presence_of_all_elements_located((By.XPATH, "//*[text()='Message']/.."))
                            )

                    except Exception as e:
                        print(e)

                # Move to the next page only if the limit is not reached
                if cnt < limit:
                    try:
                        # Scroll to the bottom before trying to find the "Next" button
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(2)

                        # Check if the "Next" button is present and enabled
                        next_button = driver.find_elements(By.XPATH, "//button[@aria-label='Next']")
                        if next_button and next_button[0].is_enabled():
                            # Wait for the "Next" button to be clickable
                            next_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Next']"))
                            )

                            # Scroll the "Next" button into view
                            driver.execute_script("arguments[0].scrollIntoView();", next_button)

                            # Click the "Next" button
                            next_button.click()

                            time.sleep(3.4)  # Allow time for the new page to load

                            # Reset cnt2 for the new page
                            cnt2 = 1

                        else:
                            print(Fore.YELLOW + "[INFO] No more pages or 'Next' button not enabled. Stopping the script.")
                            break  # Exit the loop if there are no more pages

                    except Exception as e:
                        print(e)

                        # Check if no new buttons were found for 20 seconds
                        if time.time() - last_button_found_time > 20:
                            print(Fore.RED + "[WARNING] No new buttons found for 20 seconds. Stopping the script.")
                            break  # Exit the loop if stuck

            except TimeoutException:
                print("No connect buttons found after waiting.")

                # Check if no new buttons were found for 20 seconds
                if time.time() - last_button_found_time > 20:
                    print(Fore.RED + "[WARNING] No new buttons found for 20 seconds. Stopping the script.")
                    break  # Exit the loop if stuck

    except Exception as e:
        print(Fore.RED + f"[INFO] No profile found or an error occurred. Details: {e}")

    finally:  # This block will always execute, even if there's an exception
        # Arthur modif
        # Create and save the DataFrame here
        df = pd.DataFrame({'Linkedin Profile name': linkedin_profile_name, 'Linkedin URL': linkedin_url_list})

        # Create the new row DataFrame
        new_row_data = {
        'Linkedin Profile name': 'Connexion degre / Keywords / Location',
        'Linkedin URL': connection_degree + " " + keyword + " " + location
        }
        new_row_df = pd.DataFrame([new_row_data])

        #   Concatenate the new row and the original DataFrame
        updated_df = pd.concat([new_row_df, df], ignore_index=True)

        # Update the original DataFrame
        df = updated_df

        # Get the current date and time
        now = datetime.datetime.now()
        date_str = now.strftime("%Y%m%d")  # Format for folder name
        time_str = now.strftime("%H-%M-%S")  # Format for filename

        # Create the folder if it doesn't exist
        folder_name = "dfs/" + date_str
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

        #hyphenating keywords
        hyphenated_keyword = "-".join(keyword.split())
        # Include the timestamp in the filename and save within the folder
        filename = os.path.join(folder_name, f'{time_str}_linkedin-name-url_keyword_{hyphenated_keyword}.xlsx')


        df.to_excel(filename, index=False)

        print(df)
        print(f"DataFrame saved to {filename}")  # Debug print

        driver.quit()  # Quit the driver here as well
        #End modif

def main():
    print(Fore.CYAN + "[-] Please enter your search criteria:")
    message = ''
    message_letter = ''
    include_note = False
    connection_degree = input(Fore.MAGENTA + "[+] Enter the connection degree (1st, 2nd, 3rd): " + Fore.RESET)
    if connection_degree.lower() not in ['1st', '2nd', '3rd']:
        print(Fore.RED + "[ERROR] Invalid connection degree. Please enter 1st, 2nd, or 3rd.")
        connection_degree = input(Fore.MAGENTA + "[+] Enter the connection degree (1st, 2nd, 3rd): " + Fore.RESET)
    keyword = input(Fore.MAGENTA + "[+] Enter the keyword for the search: " + Fore.RESET)
    location = input(Fore.MAGENTA + "[+] Enter the location: " + Fore.RESET)
    if connection_degree.lower() == '1st':
        message_letter = input(Fore.MAGENTA + "[+] Enter the message letter for the connection request: " + Fore.RESET)
    if message_letter == "":
        include_note = input(Fore.MAGENTA + "[+] Do you want to include a note in the connection request? (y/n): " + Fore.RESET)
        if include_note.lower() == 'y':
            include_note = True
            message = input(Fore.MAGENTA + "[+] Enter the personalized message to send with connection requests: " + Fore.RESET)
        else:
            include_note = False
    limit = int(input(Fore.MAGENTA + "[+] Enter the maximum number of connection requests to send: " + Fore.RESET))
    li_at = input(Fore.MAGENTA + "[+] Enter the li_at of Linkedin: " + Fore.RESET)
    print("----------------------------------------------------------------")
    driver = setup_driver()

    try:
        login_with_cookie(driver, li_at)
    except Exception as e:
        print(Fore.RED + f"[INFO] Cookie login failed: {e}\n" + Fore.YELLOW + "Attempting login with credentials.")
        email = config.get('LinkedIn', 'email')
        password = config.get('LinkedIn', 'password')
        login_with_credentials(driver, email, password)

    network_mapping = {
        "1st": "%5B%22F%22%5D",
        "2nd": "%5B%22S%22%5D",
        "3rd": "%5B%22O%22%5D"
    }
    network_code = network_mapping.get(connection_degree, "")

    search_url = f"https://www.linkedin.com/search/results/people/?keywords={keyword.replace(' ','%20').lower()}&locations={location.replace(' ','%20')}&network={network_code}&origin=FACETED_SEARCH"
    print(Fore.YELLOW + f"[INFO] Navigating to search URL: {search_url}")
    driver.get(search_url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "global-nav-typeahead")))
    if location != "":
        select_location(driver, location)
    send_connection_request(driver=driver, limit=limit, letter=message, include_notes=include_note, message_letter=message_letter, keyword=keyword, location=location, connection_degree=connection_degree)
    driver.quit()

if __name__ == "__main__":
    main()
