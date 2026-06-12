from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import random
from lxml import html
import pandas as pd
import re
import os
from datetime import datetime
import pytz
from datetime import datetime
from datetime import datetime, timedelta
import subprocess



chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_experimental_option("excludeSwitches",["enable-automation"])
service_obj = Service(f"{os.path.dirname(os.path.realpath(__file__))}/chromedriver.exe")
driver = webdriver.Chrome(options=chrome_options,service=service_obj)



# Create the folder
try:
    os.mkdir(f'{os.path.dirname(os.path.realpath(__file__))}/google scrap data')  
    print(f"Folder google scrap data created successfully.")
except FileExistsError:
    print(f"Folder google scrap data already exists.")   

with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
        file.write('')
# Get current time in UTC
utc_now = datetime.utcnow()

# Define the time zone for Morocco
morocco_timezone = pytz.timezone('Africa/Casablanca')

# Convert UTC time to Moroccan time
morocco_now = utc_now.replace(tzinfo=pytz.utc).astimezone(morocco_timezone)

# Format the date and time
folder_name = str(morocco_now.strftime("%d-%m-%Y"))

# Create the folder
try:
    os.mkdir(f'{os.path.dirname(os.path.realpath(__file__))}/google scrap data/{folder_name}')
    print(f"Folder '{folder_name}' created successfully.")
except FileExistsError:
    print(f"Folder '{folder_name}' already exists.")      

with open(f'{os.path.dirname(os.path.realpath(__file__))}/search.txt', 'r', encoding='utf-8') as file:
    name = file.readline()

searching = name


file_path = f'{os.path.dirname(os.path.realpath(__file__))}/google scrap data/{folder_name}/data_{name}.xlsx'

def data_extract(url):
    title = ''
    code_postal = ''
    address = ''
    phone_number1=''
    telephone=''
    gsm=''
    thelink =''
    category=''
    rating=''
    number_of_rating=''
    global folder_name
    print("extract started")
    while True :
        try :
            
            driver.get(url)
            break
        except Exception as e:
            
            print('Error 1')
            print(f'error in opening link {url} {e}')
            with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                file.write(f'error in opening link {url} {e}')
            quit()
    
    with open(f'{os.path.dirname(os.path.realpath(__file__))}/last_link.txt', 'w', encoding='utf-8') as file:
        file.write(str(url))
        
   

    phone1 = '//*[starts-with(@data-item-id, "phone")]'
    location = '//*[@data-item-id="address"]'
    site = '//*[@data-item-id="authority"]'

    other_xpaths = [location,site,phone1]

    for elem in other_xpaths :
        try:
            
            xpath = elem
            print('here all good !')
            try :
            
                if elem == phone1 :
                    wait = WebDriverWait(driver, 5)
                    
                    while True:
                        try:
                            # Wait for the element to be present
                            element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                            # Extract text
                            text = element.text.strip()
                            
                            if len(text) > 0:  # Check if text is not empty
                                phone_number = text.split('\n')[1] #if '\n' in text else text
                                print(phone_number)
                                phone_number1 = phone_number
                                if '+32' in phone_number :
                                    phone_number = phone_number.replace('+32', '')
                                if ' ' in phone_number :
                                    phone_number = phone_number.replace(' ', '')
                                if phone_number[0] == 0 :
                                    phone_number = phone_number[1:]
                                    
                                phone = phone_number

                                if len(phone) == 8 :
                                    telephone = phone
                                else :
                                    telephone = ''
                                    
                                if len(str(phone)) == 9 and str(phone)[0] == '4':
                                    gsm = phone
                                else :
                                    gsm = ''

                                print(phone)
                                if len(phone_number1) == 0 :
                                    return
                                break  # Break the loop if text is not empty
                        except Exception as e:
                            print('didnt get phone')
                            telephone = ''
                            gsm = ''
                            phone_number1 = ''
                            print(f"Error occurred : {e}")
                            return

            except Exception:
                print('didnt get phone')
                telephone = ''
                gsm = ''
                phone_number1 = ''
                return
            


            try :

                if elem == location :
                    wait = WebDriverWait(driver, 5)
                    while True:
                        try:
                            # Wait for the element to be present
                            element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                            # Extract text
                            text = element.text.strip()
                            if len(text) > 0:  # Check if text is not empty
                                location = text.split('\n')[1] #if '\n' in text else text
                                address = location
                                pattern = r'(\d+)[^\d]*$'
                                match = re.search(pattern, address)

                                if match:
                                    code_postal = match.group(1)
                                
                                print(address)
                                print(code_postal)
                                break  # Break the loop if text is not empty
                        except Exception as e:
                            address = ''
                            code_postal = ''
                            print(f"Error occurred while retrieving : {e}")
                            break  # Break the loop if an exception occurs


            except Exception as e:
                address = ''
                code_postal = ''
                pass

            try :

                if elem == site :
                    wait = WebDriverWait(driver,1)
                    while True:
                        try:
                            # Wait for the element to be present
                            element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                            # Extract text
                            text = element.text.strip()
                            if len(text) > 0:  # Check if text is not empty
                                thelink = text.split('\n')[1] #if '\n' in text else text
                                thelink = thelink
                                print(thelink)
                                break  # Break the loop if text is not empty
                        except Exception as e:
                            thelink = ''
                            print(f"Error occurred while retrieving : {e}")
                            break  # Break the loop if an exception occurs
                    

            except Exception as e:
                thelink = ''
                pass


        except Exception as e:
            print("Error:", e)
            continue

    wait = WebDriverWait(driver, 3)
    while True :
        try :
            title1 = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/div[1]/h1'
            element = wait.until(EC.presence_of_element_located((By.XPATH, title1)))
            title = str(element.text.encode('utf-8').decode('utf-8'))
            if len(str(title))!=0:
                print(title)
                break
        except Exception as e:
            break
    

    

    try :

        category = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/span/span/button'
        element = wait.until(EC.presence_of_element_located((By.XPATH, category)))
        category = str(element.text.encode('utf-8').decode('utf-8'))
        print(category)

    except Exception as e:
        category = ''
        pass    

    try :

        rating = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[1]/div[2]/span[1]/span[1]'
        element = wait.until(EC.presence_of_element_located((By.XPATH, rating)))
        rating = str(element.text.encode('utf-8').decode('utf-8'))
        print(rating)

    except Exception as e:
        rating = ''
        pass
    
    print(len(rating))
    
    try :

        number_of_rating = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[1]/div[2]/span[2]/span/span'
        element = wait.until(EC.presence_of_element_located((By.XPATH, number_of_rating)))
        number_of_rating = str(element.text.encode('utf-8').decode('utf-8')).replace(" ","")
        if len(number_of_rating) == 0:
            if len(rating) != 0:
                print('rating exist---------------------------------------')
                
                while True :
                    
                    number_of_rating = '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[1]/div[2]/span[2]/span/span'
                    element = wait.until(EC.presence_of_element_located((By.XPATH, number_of_rating)))
                    number_of_rating = str(element.text.encode('utf-8').decode('utf-8')).replace(" ","")
                    if len(number_of_rating)!=0:
                        print('rating total is : '+number_of_rating)
                        
                        break

    except Exception as e:
        print('error in rating-----------')
        number_of_rating = ''
        pass


    if len(phone_number1) != 0 :
        data = {

            'name':title,
            'code postal':code_postal,
            'adresse':address,
            'phone':phone_number1,
            'telephone':telephone,
            'GSM':gsm,
            'website':thelink,
            'category':category,
            'rating':rating,
            'total_of_rating':number_of_rating[1:-1],
            'google_map_link':url,

        }
        
        
        try : 
                # Check if the file exists
                if os.path.exists(file_path):
                    # Read existing data from the Excel file
                    existing_df = pd.read_excel(file_path)

                    # Create a new DataFrame with the new data
                    new_df = pd.DataFrame([data])

                    # Concatenate the existing DataFrame and the new DataFrame
                    updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                else:
                    # If the file doesn't exist, create a new DataFrame with the new data
                    updated_df = pd.DataFrame([data])

                # Save the updated DataFrame to the Excel file
                updated_df.to_excel(file_path, index=False)

                print('DONE !!')
        except Exception as e:
            print("Error in code postal :", e)
            pass
        

        # try : 
        #     with open(f'{os.path.dirname(os.path.realpath(__file__))}/searchauto.txt', 'r', encoding='utf-8') as file:
        #         all_lines = file.readlines()

        #     # Use regex to find the first number in the line
                
        #     for line in all_lines :
        #         match = re.search(r'\b\d+\b', line)

        #         if match:
        #             codpostal = str(int(match.group()))
                    
        #             if code_postal == codpostal :
        #                 print("code postal found !!!")
        #                 break

        #     if code_postal == codpostal :
            
        #         with open(f'{os.path.dirname(os.path.realpath(__file__))}/searchauto.txt', 'r', encoding='utf-8') as file:
        #             Fline1 = file.readline()
                
        #         #match = re.search(r'(.+?)\s+\b\d+\b', name)

        #         #if match:
        #             #name = match.group(1)

        #         def extract_string_before_code_postal(input_string):
        #             # Split the input string using "code postal" as the delimiter
        #             parts = input_string.split("code postal", 1)

        #             # Extract the part of the string before "code postal"
        #             string_before_code_postal = parts[0].strip()

        #             return string_before_code_postal
        #         name = extract_string_before_code_postal(Fline1)

        #         file_path = f'{os.path.dirname(os.path.realpath(__file__))}/google scrap data/{folder_name}/data {name} belgique.xlsx'

        #         # Check if the file exists
        #         if os.path.exists(file_path):
        #             # Read existing data from the Excel file
        #             existing_df = pd.read_excel(file_path)

        #             # Create a new DataFrame with the new data
        #             new_df = pd.DataFrame([data])

        #             # Concatenate the existing DataFrame and the new DataFrame
        #             updated_df = pd.concat([existing_df, new_df], ignore_index=True)
        #         else:
        #             # If the file doesn't exist, create a new DataFrame with the new data
        #             updated_df = pd.DataFrame([data])

        #         # Save the updated DataFrame to the Excel file
        #         updated_df.to_excel(file_path, index=False)

        #         print('DONE !!')
        # except Exception as e:
        #     print("Error in code postal :", e)
        #     pass




def searching():
        with open(f'{os.path.dirname(os.path.realpath(__file__))}/search.txt', 'r', encoding='utf-8') as file:
            search_all = file.readlines()
        for each_search in search_all :
            if search_all.index(each_search) == 1 :
                driver.quit()
                quit()
            all_links = []
                
            while True :
                try :
                    
                    driver.get('https://https://www.google.com/maps')
                    break
                except Exception as e:
                    print('Error 1')
                    print(f'error in link of google map {e}')
                    with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                        file.write(f'error in link of google map {e}')
                    driver.quit()
                    quit()
            while True :
                try :   
                    
                    wait = WebDriverWait(driver, 10)
                    search_input = '//*[@id="searchboxinput"]'
                    input_field = wait.until(EC.element_to_be_clickable((By.XPATH, search_input)))
                    for char in each_search :
                        input_field.send_keys(char)
                        time.sleep(random.uniform(0.01, 0.2))
                    break
                except Exception as e:
                    driver.quit()
                    quit()
            print(each_search)
            search = '//*[@id="searchbox-searchbutton"]'

            try :
                click_search = wait.until(EC.element_to_be_clickable((By.XPATH, search)))
                click_search.click()
                time.sleep(2)
                click_search = wait.until(EC.element_to_be_clickable((By.XPATH, search)))
                click_search.click()
            except Exception as e:
                with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                    file.write(f'error 2 in clicking search {e}')
                

                print('Error 2 in clicking search box')
                
                driver.quit()
                
                quit()
            time.sleep(random.uniform(3, 5))
            try:
                elem = 'https://www.gstatic.com/images/icons/material/system_gm/2x/info_gm_grey_18dp.png'
                xpath_expression = f"//img[contains(@src, '{elem}')]"
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath_expression))
                )
                src_value = element.get_attribute('src')
                #print("Found src attribute:", src_value)
                xpath = driver.execute_script(
                    "function getXPath(element) {"
                    "  if (element.id !== '')"
                    "    return '//*[@id=\"' + element.id + '\"]';"
                    "  if (element === document.body)"
                    "    return element.tagName.toLowerCase();"
                    ""
                    "  var ix = 0;"
                    "  var siblings = element.parentNode.childNodes;"
                    ""
                    "  for (var i = 0; i < siblings.length; i++) {"
                    "    var sibling = siblings[i];"
                    ""
                    "    if (sibling === element)"
                    "      return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';"
                    ""
                    "    if (sibling.nodeType === 1 && sibling.tagName === element.tagName)"
                    "      ix++;"
                    "  }"
                    "}"
                    "return getXPath(arguments[0]);", element)             
                scrollnow = xpath.replace('/div[1]/div[1]/img[1]', '/div[2]/div[1]').replace('[1]', '')
                scrollnow = wait.until(EC.presence_of_element_located((By.XPATH, scrollnow)))
                scrollnow.click()
                print('lets scroll now ...')
            except Exception as e:
                print(f'error in scroll {e}')
                pass
            # Get the current time
            start_time = datetime.now()
            lastlinkfound = ''
            # Set the duration for 25 seconds
            duration = timedelta(seconds=25)
            # Calculate the end time
            end_time = start_time + duration
            while True:
                
                current_url = driver.current_url
                if current_url.startswith('https://https://www.google.com/maps/place/') :
                    all_links.append(current_url)
                    data_extract(current_url)
                        
                    z = 0
                    break
                    
                z = 1
                print("scrolling")
                actions = ActionChains(driver)
                try_scrolling = """
                    return (() => {
                        const anchors = Array.from(document.querySelectorAll('a'));
                        const placeUrl = anchors.find(anchor => 
                            anchor.href.startsWith('https://https://www.google.com/maps/place/')
                        );
                        
                        if (placeUrl) {
                            // Uncomment to scroll to the element first if needed
                            // placeUrl.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            placeUrl.focus();
                            return 'Focused on URL';
                        }
                        return 'URL not found';
                    })();
                    """
                driver.execute_script(try_scrolling)
                actions.send_keys(Keys.END).perform()
                text = driver.find_element(By.XPATH, "//body").text
                end = 'Vous êtes arrivé à la fin de la liste.'
                end1 = "You've reached the end of the list."
                try :
                    links = driver.find_elements(By.TAG_NAME, 'a')
                    filtered_links = [link.get_attribute('href') for link in links if link.get_attribute('href') and link.get_attribute('href').startswith('https://https://www.google.com/maps/place/')]
                    for elem in filtered_links :
                        all_links.append(elem)
                except Exception as e:
                    print("scrolling")
                    actions = ActionChains(driver)
                    actions.send_keys(Keys.END).perform()
                    pass
                all_links = list(dict.fromkeys(all_links))
                if len(all_links) == 0 :
                    break
                if len(all_links) == 120 :
                    for lienx in all_links :
                        try :
                                data_extract(lienx)
                        except Exception as e:
                            print('Error 1')
                            print(f'error in opening link {lienx} {e}')
                            with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                                file.write(f'error in opening link {lienx} {e}')
                            quit()
                    
                    z = 0
                    break
                elif end in text:
                    print('the end')
                    z = 0
                    for lienx in all_links :
                        try :
                                data_extract(lienx)
                        except Exception as e:
                            print('Error 1')
                            print(f'error in opening link {lienx} {e}')
                            with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                                file.write(f'error in opening link {lienx} {e}')
                            quit()
                    break
                elif end1 in text :
                    print('the end1')
                    z = 0
                    for lienx in all_links :
                        try :
                                data_extract(lienx)
                        except Exception as e:
                            print('Error 1')
                            print(f'error in opening link {lienx} {e}')
                            with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                                file.write(f'error in opening link {lienx} {e}')
                            quit()
                    break
                
                print(f'trying to fix that : {datetime.now() - end_time}')
                
                if datetime.now() > end_time  :
                    if len(all_links) == 0 :
                            z = 0
                    if all_links[-1] == lastlinkfound:
                        print(f'Total links remained the same after 1 minute: {len(list(dict.fromkeys(all_links)))}')
                        print('the end2')
                        for lienx in all_links :
                            try :
                                    data_extract(lienx)
                            except Exception as e:
                                print('Error 1')
                                print(f'error in opening link {lienx} {e}')
                                with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                                    file.write(f'error in opening link {lienx} {e}')
                                quit()
                        z = 0
                    else:
                        print(f'One minute has passed. Total links: {len(list(dict.fromkeys(all_links)))}')
                        if len(all_links) != 0:
                            lastlinkfound = all_links[-1]
                            print('last link changed')
                        if len(all_links) == 0 :
                            z = 0
                        # Get the current time
                        start_time = datetime.now()

                        # Set the duration for 25 seconds
                        duration = timedelta(seconds=25)

                        # Calculate the end time
                        end_time = start_time + duration
                if z == 0 :
                    break
            
try :
    
    searching()
except Exception as e :
    print(f'error occured while searching  {e}')
    with open(f'{os.path.dirname(os.path.realpath(__file__))}/Errors.txt', 'w', encoding='utf-8') as file:
                    file.write(f'error occured while searching  {e}')
driver.quit()
quit()
