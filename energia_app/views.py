from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementNotInteractableException, ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework import status
from .models import UploadedFile
from .serializers import UploadedFileSerializer
import openpyxl
import os
from django.conf import settings

passed_rows = []
failed_rows = []

plan_xpath_mapping = {
    'Standard Electricity': "/html/body/div[1]/main/div/div[2]/div[5]/div[2]/div/div/div[1]/div/div/div[2]/div[1]/input",
    'Smart Data': "/html/body/div[1]/main/div/div[2]/div[6]/div[1]/div[1]/div/div/div[2]/div[1]/input",
    'Smart Drive': "/html/body/div[1]/main/div/div[2]/div[6]/div[2]/div[1]/div/div/div[2]/div[1]/input",
    'Smart 24 Hour': "/html/body/div[1]/main/div/div[2]/div[6]/div[3]/div[1]/div/div/div[2]/div[1]/input",
    'Smart Day Night': "/html/body/div[1]/main/div/div[2]/div[6]/div[4]/div[1]/div/div/div[2]/div[1]/input",
    'Energia SST': "/html/body/div[1]/main/div/div[2]/div[6]/div[5]/div[1]/div/div/div[2]/div[1]/input",
    'Standard Dual': "/html/body/div[1]/main/div/div[2]/div[5]/div/div[1]/div/div/div[2]/div[1]/input",
    'Smart Data Dual': "/html/body/div[1]/main/div/div[2]/div[6]/div[1]/div[1]/div/div/div[2]/div[1]/input",
    'Smart Drive Dual': "/html/body/div[1]/main/div/div[2]/div[6]/div[2]/div[1]/div/div/div[2]/div[1]/input",
    'Smart 24 Hour Dual': "/html/body/div[1]/main/div/div[2]/div[6]/div[3]/div[1]/div/div/div[2]/div[1]/input",
    'Smart Day Night Dual': "/html/body/div[1]/main/div/div[2]/div[6]/div[4]/div[1]/div/div/div[2]/div[1]/input",
    'Gas Offer': "/html/body/div[1]/main/div/div[2]/div[5]/div/div[1]/div/div/div[2]/div[1]/input"
}

residential_type_mapping = {
    'Apartment': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[3]/div",
    'Detached': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[4]/div",
    'Semi-detached': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[5]/div",
    'Terrace': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[6]/div",
    'Town House': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[7]/div",
    'Agricultural': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[8]/div",
    'Cottage': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[9]/div",
    'Site': "/html/body/div[1]/main/div[2]/div[2]/div[2]/div[3]/div[4]/div[10]/div"
}

bedroom_mapping = {
    '1': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[11]/div[2]/div[2]/div/div[1]",
    '2': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[11]/div[2]/div[2]/div/div[2]",
    '3': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[11]/div[2]/div[2]/div/div[3]",
    '4': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[11]/div[2]/div[2]/div/div[4]",
    '5': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[11]/div[2]/div[2]/div/div[5]",
    '6+': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[11]/div[2]/div[2]/div/div[6]"
}

extraroom_mapping = {
    '1': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[13]/div[2]/div[2]/div/div[1]",
    '2': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[13]/div[2]/div[2]/div/div[2]",
    '3': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[13]/div[2]/div[2]/div/div[3]",
    '4': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[13]/div[2]/div[2]/div/div[4]",
    '5': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[13]/div[2]/div[2]/div/div[5]",
    '6+': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[13]/div[2]/div[2]/div/div[6]"  
}

no_of_people_mapping = {
    '1': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[1]",
    '2': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[2]",
    '3': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[3]",
    '4': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[4]",
    '5': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[5]",
    '6': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[6]",
    '7': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[8]",
    '8': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[9]",
    '9+': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[15]/div[2]/div[2]/div/div[10]"
}

usage_of_gas_mapping = {
    'cooking': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[18]/div[2]/div/div/div[1]/div[1]/div",
    'heating': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[18]/div[2]/div/div/div[2]/div[1]",
    'both': "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[4]/div[18]/div[2]/div/div/div[3]/div[1]"
}
        
class FileUploadView(APIView):
    parser_classes = (MultiPartParser, FormParser)

    def post(self, request, *args, **kwargs):
        file_serializer = UploadedFileSerializer(data=request.data)
        if file_serializer.is_valid():
            file_instance = file_serializer.save()
            output_file_path = process_file(file_instance.file.path)
            output_file_url = request.build_absolute_uri(os.path.join(settings.MEDIA_URL, 'output_files', 'output.xlsx'))
            return Response({
                "message": "File uploaded and processed successfully!",
                "file_path": file_instance.file.url,
                "output_file_url": output_file_url  
            }, status=status.HTTP_201_CREATED)
        else:
            return Response(file_serializer.errors, status=status.HTTP_400_BAD_REQUEST)

def process_file(file_path):
    def initialize_driver():
        driver = webdriver.Chrome()
        driver.get("https://www.energia.ie/home")
        driver.maximize_window()
        time.sleep(2)
        return driver


    def accept_cookies(driver):
        try:
            accept_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div/div[1]/div/div[2]/div/button[2]"))
            )
            accept_button.click()
            time.sleep(1)
            print("Accepted cookies.")
        except TimeoutException:
            print("Cookie accept button not found.")


    def retry_click(driver, xpath, retries=2):
        while retries > 0:
            try:
                click_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                click_button.click()
                return True
            except (TimeoutException, StaleElementReferenceException) as e:
                print(f"Retrying click due to error: {e}")
                retries -= 1
                time.sleep(1)
        print(f"Failed to click after {retries} retries")
        return False


    def click_yes_or_no(driver, condition):
        try:
            if condition == "Yes":
                yes_input = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//label[contains(text(),"Yes")]/preceding-sibling::input[@type="checkbox"]'))
                )
                driver.execute_script("arguments[0].click();", yes_input)
                print("Clicked 'Yes' button.")
            else:
                pass
        except TimeoutException:
            print(f"Button with text '{condition}' not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def select_tariff_by_name(driver, tariff_name, retries=1):
        def check_tariff(driver):
            try:
                tariff_xpath = f"//p[contains(@class, 'product-title-text') and text()='{tariff_name}']/ancestor::div[contains(@class, 'inner-product')]"
                tariff_card = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, tariff_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", tariff_card)
                time.sleep(1)
                tariff_card.click()
                print(f"Tariff '{tariff_name}' selected.")
                return True
            except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException):
                return False

        for attempt in range(retries):
            try:
                if check_tariff(driver):
                    return
                # Scroll down and check again
                current_position = driver.execute_script("return window.pageYOffset;")
                end_of_page = driver.execute_script("return document.body.scrollHeight;")
                step = 500
                delay = 2
                while current_position < end_of_page:
                    driver.execute_script(f"window.scrollBy(0, {step});")
                    time.sleep(delay)
                    current_position += step
                    end_of_page = driver.execute_script("return document.body.scrollHeight;")
                    if check_tariff(driver):
                        return
                print(f"Attempt {attempt + 1}: Tariff '{tariff_name}' not found. Retrying...")
                if attempt == retries - 1:
                    try:
                        see_more_button_xpath = "//button[contains(@class, 'show-more-text') and text()='See more plans']"
                        see_more_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, see_more_button_xpath))
                        )
                        see_more_button.click()
                        print("Clicked 'See more plans' button.")
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, f"//p[contains(@class, 'product-title-text') and text()='{tariff_name}']"))
                        )
                        if check_tariff(driver):
                            return
                    except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException) as e:
                        print(f"Final attempt: Tariff '{tariff_name}' not found on additional plans page.")
                    except Exception as e:
                        print(f"An error occurred while checking additional plans: {e}")
            except Exception as e:
                print(f"An error occurred while checking initial page: {e}")

        print(f"Failed to select tariff '{tariff_name}' after {retries} retries.")


    def fill_text_field_by_placeholder(driver, placeholder, text):
        try:
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@placeholder="{placeholder}"]'))
            )
            field.clear()
            field.send_keys(text)
            print(f'Filled {placeholder} with {text}')
        except Exception as e:
            print(f'Failed to fill field {placeholder}: {e}')
    def fill_text_field_by_placeholder1(driver, placeholder, text):
        try:
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@placeholder="{placeholder}"]'))
            )
            field.click()
            field.clear()
            field.send_keys(text)
            print(f'Filled {placeholder} with {text}')
        except Exception as e:
            print(f'Failed to fill field {placeholder}: {e}')
    


    def fill_text_field(driver, field_name, text, retries=1, delay=1):
        attempt = 0
        while attempt < retries:
            try:
                field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f'//input[@id="{field_name}"]'))
                )
                field.clear()
                field.send_keys(text)
                print(f"Filled {field_name} with {text}")
                return True
            except Exception as e:
                print(f"Attempt {attempt + 1} to fill field {field_name} failed: {e}")
                attempt += 1
                time.sleep(delay)
        print(f"Failed to fill field {field_name} with {text} after {retries} attempts")
        return False


    def select_residential_type(driver, residential_type):
        if residential_type in residential_type_mapping:
            residential_type_xpath = residential_type_mapping[residential_type]
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, residential_type_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", element)
                print(f'Selected residential type: {residential_type}')
            except TimeoutException:
                print(f'Could not find residential type button for: {residential_type}')
            except Exception as e:
                print(f"An error occurred while selecting residential type: {e}")
        else:
            print(f'Residential type "{residential_type}" not recognized.')


    def select_bedrooms(driver, bedrooms):
        bedrooms = str(bedrooms)
        if bedrooms in bedroom_mapping:
            bedroom_xpath = bedroom_mapping[bedrooms]
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, bedroom_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", element)
                print(f'Selected number of bedrooms: {bedrooms}')
            except TimeoutException:
                print(f'Could not find bedrooms button for: {bedrooms}')
            except Exception as e:
                print(f"An error occurred while selecting bedrooms: {e}")
        else:
            print(f'Bedrooms "{bedrooms}" not recognized.')


    def select_extra_rooms(driver, extra_rooms):
        extra_rooms = str(extra_rooms)
        if extra_rooms in extraroom_mapping:
            extraroom_xpath = extraroom_mapping[extra_rooms]
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, extraroom_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", element)
                print(f'Selected number of extra rooms: {extra_rooms}')
            except TimeoutException:
                print(f'Could not find extra rooms button for: {extra_rooms}')
            except Exception as e:
                print(f"An error occurred while selecting extra rooms: {e}")
        else:
            print(f'Extra rooms "{extra_rooms}" not recognized.')


    def select_no_of_people(driver, no_of_people):
        no_of_people = str(no_of_people)
        if no_of_people in no_of_people_mapping:
            no_of_people_xpath = no_of_people_mapping[no_of_people]
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, no_of_people_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", element)
                print(f'Selected number of people: {no_of_people}')
            except TimeoutException:
                print(f'Could not find number of people button for: {no_of_people}')
            except Exception as e:
                print(f"An error occurred while selecting number of people: {e}")
        else:
            print(f'Number of people "{no_of_people}" not recognized.')


    def select_gas_usage(driver, gas_usage):
        gas_usage = str(gas_usage)
        if gas_usage in usage_of_gas_mapping:
            gas_usage_xpath = usage_of_gas_mapping[gas_usage]
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, gas_usage_xpath))
                )
                time.sleep(1)
                driver.execute_script("arguments[0].click();", element)
                print(f'Selected gas usage: {gas_usage}')
            except TimeoutException:
                print(f'Could not find gas usage button for: {gas_usage}')
            except Exception as e:
                print(f"An error occurred while selecting gas usage: {e}")
        else:
            print(f'Gas usage "{gas_usage}" not recognized.')


    def fill_text_field_and_tab(driver, field_name, text):
        try:
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, field_name))
            )
            field.clear()
            field.send_keys(text)
            field.send_keys(Keys.TAB)
            print(f"Filled field {field_name} with text {text} and pressed Tab.")
        except Exception as e:
            print(f"Failed to fill field {field_name}: {e}")


    def click_element_with_js(driver, xpath, retries=1, delay=1):
        attempt = 0
        while attempt < retries:
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                driver.execute_script("arguments[0].click();", element)
                print(f"Clicked element with XPath: {xpath}")
                return True
            except Exception as e:
                print(f"Attempt {attempt + 1} failed: {e}")
                attempt += 1
                time.sleep(delay)
        print(f"Failed to click element with XPath: {xpath} after {retries} attempts")
        return False


    def fill_text_field_with_js(driver, field_name, text):
        try:
            text = str(text)
            print(f"Locating the field with id: {field_name}")
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@id="{field_name}"]'))
            )
            print(f"Field {field_name} located. Clearing the field.")
            driver.execute_script(f"arguments[0].value = '';", field)
            print(f"Filling field {field_name} with text using JavaScript.")
            driver.execute_script(f"arguments[0].value = '{text}'; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", field)
            final_value = field.get_attribute('value')
            if final_value != text:
                raise ValueError(f"Final text '{final_value}' does not match expected text '{text}'")
            print(f"Field {field_name} filled with text '{text}' successfully using JavaScript.")
        except Exception as e:
            print(f"Failed to fill field {field_name} with text '{text}' using JavaScript: {e}")


    def fill_text_field_with_js_placeholder(driver, field_name, text):
        try:
            text = str(text)
            print(f"Locating the field with placeholder: {field_name}")
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@placeholder="{field_name}"]'))
            )
            print(f"Field {field_name} located. Clearing the field.")
            driver.execute_script(f"arguments[0].value = '';", field)
            print(f"Filling field {field_name} with text using JavaScript.")
            for char in text:
                driver.execute_script(f"arguments[0].value += '{char}'; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", field)
                time.sleep(0.1)
            final_value = field.get_attribute('value')
            if final_value != text:
                raise ValueError(f"Final text '{final_value}' does not match expected text '{text}'")
            print(f"Field {field_name} filled with text '{text}' successfully using JavaScript.")
        except TimeoutException:
            print(f"Failed to locate field with placeholder: '{field_name}' within the specified timeout.")
        except Exception as e:
            print(f"Failed to fill field {field_name} with text '{text}' using JavaScript: {e}")


    def select_title(driver, title):
        try:
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "title"))
            )
            select = Select(dropdown_element)
            select.select_by_visible_text(title)
            print(f"Selected option: {title}")
        except TimeoutException:
            print("Dropdown element not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def select_houseowner_status(driver, houseowner_status):
        try:
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "home_ownership_status"))
            )
            select = Select(dropdown_element)
            select.select_by_visible_text(houseowner_status)
            print(f"Selected option: {houseowner_status}")
        except TimeoutException:
            print("Dropdown element not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def security_question_drop(driver, security_ques):
        try:
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "security_question"))
            )
            select = Select(dropdown_element)
            select.select_by_visible_text(security_ques)
            print(f"Selected option: {security_ques}")
        except TimeoutException:
            print("Dropdown element not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def corresponding_dropdown(driver, correc_drop):
        try:
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "drpCountry"))
            )
            select = Select(dropdown_element)
            select.select_by_visible_text(correc_drop)
            print(f"Selected option: {correc_drop}")
        except TimeoutException:
            print("Dropdown element not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def corres_county_dropdown(driver, correc_drop_county):
        try:
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "drpCounty"))
            )
            select = Select(dropdown_element)
            select.select_by_visible_text(correc_drop_county)
            print(f"Selected option: {correc_drop_county}")
        except TimeoutException:
            print("Dropdown element not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def click_radio_button(driver, mobile):
        try:
            if mobile == 'Mobile phone':
                # radio_button = WebDriverWait(driver, 10).until(
                #     EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div/div[3]/div[2]/form/div[15]/div/label/span[2]"))
                # )
                # driver.execute_script("arguments[0].scrollIntoView(true);", radio_button)
                # radio_button.click()
                pass
            else:
                radio_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div/div[3]/div[2]/form/div[16]/div/label/span[2]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", radio_button)
                radio_button.click()

            print(f"Radio button for {mobile} clicked.")
        except TimeoutException:
            print(f"Radio button for value '{mobile}' not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def click_anywhere(driver):
        try:
            element_to_click = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//body'))
            )
            element_to_click.click()
            print("Clicked on the screen to defocus the input field.")
        except Exception as e:
            print(f"Failed to click on the screen: {e}")


    def click_yes_or_no_js(driver, condition, div_id):
        try:
            if condition == "Yes":
                checkbox_xpath = f"//div[@id='{div_id}']//input[@type='checkbox']"
                checkbox = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, checkbox_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", checkbox)
                time.sleep(1)
                retries = 3
                while retries > 0:
                    if checkbox.is_selected():
                        print(f"Checkbox in div '{div_id}' is set to '{condition}' using JavaScript.")
                        break
                    else:
                        driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(1)
                        retries -= 1
                else:
                    print(f"Failed to set checkbox in div '{div_id}' to '{condition}'.")
            else:
                print(f"No action needed for condition: {condition}")
        except TimeoutException:
            print(f"Checkbox in div '{div_id}' not found.")
        except Exception as e:
            print(f"An error occurred: {e}")


    def fill_text_field_with_js_ID(driver, field_id, text):
        try:
            text = str(text)
            print(f"Locating the field with id: {field_id}")
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@id="{field_id}"]'))
            )
            print(f"Field {field_id} located. Clearing the field.")
            driver.execute_script(f"arguments[0].value = '';", field)
            print(f"Filling field {field_id} with text using JavaScript.")
            for char in text:
                driver.execute_script(f"arguments[0].value += '{char}'; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", field)
                time.sleep(0.1)
            final_value = field.get_attribute('value')
            if final_value != text:
                raise ValueError(f"Final text '{final_value}' does not match expected text '{text}'")
            print(f"Field {field_id} filled with text '{text}' successfully using JavaScript.")
        except TimeoutException:
            print(f"Failed to locate field with id: '{field_id}' within the specified timeout.")
        except Exception as e:
            print(f"Failed to fill field {field_id} with text '{text}' using JavaScript: {e}")
    def click_checkbox_with_js(driver, xpath, retries=2, delay=1):
        attempt = 0
        while attempt < retries:
            try:
                checkbox = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                time.sleep(0.5)  # Small delay to ensure the element is in view
                driver.execute_script("arguments[0].click();", checkbox)
                print(f"Clicked checkbox with JavaScript: {xpath}")
                return True  # Exit the function if successful
            except Exception as e:
                print(f"Attempt {attempt + 1} to click checkbox with XPath {xpath} failed: {e}")
                attempt += 1
                time.sleep(delay)  # Wait before retrying
        print(f"Failed to click checkbox with XPath: {xpath} after {retries} attempts")
        return False  # Return False if all attempts fail

    def fill_text_field_with_js_Iban(driver, field_name, text):
        try:
            text = str(text)
            print(f"Locating the field with id: {field_name}")
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@id="{field_name}"]'))
            )
            print(f"Field {field_name} located. Clearing the field.")
            driver.execute_script("arguments[0].value = '';", field)
            print(f"Filling field {field_name} with text using JavaScript.")
            driver.execute_script("""
                var inputField = arguments[0];
                var text = arguments[1];
                inputField.value = text;
                inputField.dispatchEvent(new Event('input', { bubbles: true }));
                inputField.dispatchEvent(new Event('change', { bubbles: true }));
            """, field, text)
            final_value = field.get_attribute('value')
            if final_value != text:
                raise ValueError(f"Final text '{final_value}' does not match expected text '{text}'")
            print(f"Field {field_name} filled with text '{text}' successfully using JavaScript.")
        except Exception as e:
            print(f"Failed to fill field {field_name} with text '{text}' using JavaScript: {e}")


    df = pd.read_excel(file_path,dtype={'GPRN': str, 'Mobile Number': str,'MPRN':str,'Day':str,'Month':str, 'Year':str, 'Gas Rate':str,'electricity_day_meter':str,'electricity_night_meter':str,'electricity_meter':str})

    for index, row in df.iterrows():
        try:
            driver = initialize_driver()
            accept_cookies(driver)

            retry_click(driver, "/html/body/header/div[2]/div[2]/div[2]/div[2]/a")

            retry_click(driver, "/html/body/div[1]/main/div/div[2]/div[3]/div[3]/button")
            if row['Category'] == 'Gas':
                retry_click(driver, '/html/body/div[1]/main/div/div[4]/div/form/div/div/div[2]/div[2]/label[3]/span[2]')
            elif row['Category'] == 'Electricity and Gas':
                retry_click(driver, '/html/body/div[1]/main/div/div[4]/div/form/div/div/div[2]/div[2]/label[2]/span[2]')
            elif row['Category'] == 'Electricity':
                retry_click(driver, '/html/body/div[1]/main/div/div[4]/div/form/div/div/div[2]/div[2]/label[1]/span[2]')
            else:
                print("Energy type not found")

            if row['Contract Term'] == '12 months':
                retry_click(driver, '/html/body/div[1]/main/div/div[4]/div/form/div/div/div[2]/div[4]/label[1]/span[2]')
            elif row['Contract Term'] == 'Longer than 12 months':
                retry_click(driver, '/html/body/div[1]/main/div/div[4]/div/form/div/div/div[2]/div[4]/label[2]/span[2]')
            elif row['Contract Term'] == 'All':
                retry_click(driver, '/html/body/div[1]/main/div/div[4]/div/form/div/div/div[2]/div[4]/label[3]/span[2]')
            else:
                print("Contract type not found")

            click_yes_or_no(driver, row['Ev_home'])

            retry_click(driver, "/html/body/div[1]/main/div/div[4]/div/form/div/div/div[3]/div/input")
            print('Apply button clicked')

            if row['Meter_type'] == 'Smart':
                retry_click(driver, "/html/body/div[1]/main/div/div[2]/div[2]/ul/li[2]/a/span")
            else:
                retry_click(driver, "/html/body/div[1]/main/div/div[2]/div[2]/ul/li[1]/a/span")
                print("Standard")
            print('Meter type selected')

            select_tariff_by_name(driver, row['Tariff Title'])

            allgood = "//input[@type='button' and contains(@class, 'button--primary') and @value='All good! Start my sign up']"
            click_element_with_js(driver, allgood)
            print('Clicked start sign up')
            time.sleep(1)
            if row['Category'] == 'Gas':
                pass
            else:
                fill_text_field_with_js(driver, 'mprn_inputtariff_confirmation', row['MPRN'])
                retry_click(driver, "/html/body/div[1]/main/div/div[4]/div/form/div/div/div[3]/div/div[1]/div[2]/button")

            if row['payment_preference'] == 'Variable':
                retry_click(driver, "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[1]/div")
                takeme = "//input[@type='button' and contains(@class, 'button--largeprimary') and contains(@class, 'col-12') and @value='Take me forward!']"
                click_element_with_js(driver, takeme)
                print('Clicked Variable') 
            elif row['payment_preference'] == 'Level Pay':
                retry_click(driver, "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[3]/div[3]/div")
                print("Clicked Level Pay")
                try:
                    select_residential_type(driver, row['Residential_type'])
                    time.sleep(1)
                except Exception as e:
                    print(f"An error occurred: {e}")
                select_bedrooms(driver, row['Bedrooms'])
                select_extra_rooms(driver, row['extra_rooms'])
                select_no_of_people(driver, row['no_of_people'])
                if row['Category'] == 'Gas' or row['Category'] == 'Electricity and Gas':
                    select_gas_usage(driver, row['Gas_usage'])
                else:
                    pass
                fill_text_field_and_tab(driver, 'txtBillDirectDebitDate', row['Preferred Day of Billing'])
                time.sleep(1)
                driver.execute_script(f"window.scrollBy(0, 400);")
                allgood = "//input[@type='submit' and @form='direct_debit_day' and contains(@class, 'button--largeprimary') and @value='Take me forward!']"
                click_element_with_js(driver, allgood)
            else:
                print('Preference not found')

            time.sleep(3)
            title = row['Salutation']
            select_title(driver, title)
            fill_text_field_by_placeholder(driver, 'First name', row['First Name'])
            fill_text_field_by_placeholder(driver, 'Last name', row['Surname'])
            fill_text_field(driver, 'dayfield', row['Day'])
            fill_text_field(driver, 'monthfield', row['Month'])
            fill_text_field(driver, 'yearfield', row['Year'])

            houseowner_status = row['Occupancy Status']
            select_houseowner_status(driver, houseowner_status)
            fill_text_field_by_placeholder(driver, 'Email address', row['Email Address'])
            click_anywhere(driver)
            fill_text_field(driver, 'ConfirmedEmailAddress', row['Confirm Email Address'])
            fill_text_field_with_js_placeholder(driver, 'Mobile number', row['Mobile Number'])
            # mobile = row['Phone Type']
            # print(mobile)
            # click_radio_button(driver, mobile)
            # if mobile == 'Mobile phone':
            #     fill_text_field(driver, 'Mobile number', row['Mobile Number'])
            #     pass
            # else:
            #     fill_text_field(driver, 'Landline', row['Phone Number'])
            time.sleep(3)
            security_ques = row['security_question']
            security_question_drop(driver, security_ques)
            fill_text_field_by_placeholder(driver, 'Your answer', row['security_answer'])
            click_anywhere(driver)
            time.sleep(1)

            checkbox_xpath_click = "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div/div[3]/div[2]/form/div[22]/div/div[2]/label/span"
            click_checkbox_with_js(driver, checkbox_xpath_click)
            checkbox_xpath_click1 = "/html/body/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div/div[3]/div[2]/form/div[22]/div/div[4]/label/span"
            click_checkbox_with_js(driver, checkbox_xpath_click1)
            click_yes_or_no(driver, row['vulnerable_customer'])
            allgood = "//input[@type='submit' and @id='personal_detail_submit' and @form='personal_detail' and contains(@class, 'button--primary') and @value=\"Let's go!\"]"
            click_element_with_js(driver, allgood)
            if row['Category'] == 'Gas':
                fill_text_field_with_js_ID(driver, 'gprn_input', row['GPRN'])
                time.sleep(1)
                driver.execute_script(f"window.scrollBy(0, 300);")
                validate_click = "//button[@id='gprnvalidate_button' and @class='validate__button' and @form='gprn_validation_form' and text()='Validate']"
                click_element_with_js(driver, validate_click)
                fill_text_field_with_js_placeholder(driver, 'Eg.5555', row['Gas Rate'])
                driver.execute_script(f"window.scrollBy(0, 300);")
                if row['Gas_meter_location'] == 'Outside':
                    try:
                        radio_label = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[@class='button__radio--message' and text()='Outside']"))
                        )
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", radio_label)
                        print("Radio label 'Outside' clicked using JavaScript.")
                    except Exception as e:
                        print(f"No Radio label 'Outside' clicked using JavaScript: {e}")

            elif row['Category'] == 'Electricity':
                try:
                    # Try to find and fill the 24 Hour meter field
                    try:
                        # Locate the 24 Hour input field by its associated label "24 Hour"
                        hour_24_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '24 Hour')]/following::input[@placeholder='Eg.5555']"))
                        )
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(hour_24_input))
                        hour_24_input.clear()
                        hour_24_input.send_keys(row['electricity_meter'])
                        print(f"Typed value '{row['electricity_meter']}' into 24 Hour electricity meter field.")
                    except TimeoutException:
                        print("24 Hour meter input field not found, skipping. Checking for Day and Night meters...")

                        # Proceed to check for Day and Night fields if 24 Hour is not found
                        try:
                            # Locate the Day input field by its associated label "Day"
                            day_input = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Day')]/following::input[@placeholder='Eg.5555']"))
                            )
                            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(day_input))
                            day_input.clear()
                            day_input.send_keys(row['electricity_day_meter'])
                            print(f"Typed value '{row['electricity_day_meter']}' into Day electricity meter field.")
                        except TimeoutException:
                            print("Day meter input field not found, skipping.")

                        # Locate the Night input field by its associated label "Night"
                        try:
                            night_input = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Night')]/following::input[@placeholder='Eg.5555']"))
                            )
                            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(night_input))
                            night_input.clear()
                            night_input.send_keys(row['electricity_night_meter'])
                            print(f"Typed value '{row['electricity_night_meter']}' into Night electricity meter field.")
                        except TimeoutException:
                            print("Night meter input field not found, skipping.")
                    
                except Exception as e:
                    print(f"Error while filling electricity meter readings: {e}")


                # Always check the meter location
                if row['Electricity_meter_location'] == 'Outside':
                    try:
                        # Always attempt to click the 'Outside' radio button
                        radio_label = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[@class='button__radio--message' and text()='Outside']"))
                        )
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", radio_label)
                        print("Radio label 'Outside' clicked using JavaScript.")
                    except Exception as e:
                        print(f"No Radio label 'Outside' clicked using JavaScript: {e}")

            elif row['Category'] == 'Electricity and Gas':
                if row['Electricity_meter_location'] == 'Outside':
                    try:
                        radio_label = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[@class='button__radio--message' and text()='Outside']"))
                        )
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", radio_label)
                        print("Radio label 'Outside' clicked using JavaScript.")
                    except Exception as e:
                        print(f"No Radio label 'Outside' clicked using JavaScript: {e}")

                fill_text_field_with_js_ID(driver, 'gprn_input', row['GPRN'])
                driver.execute_script(f"window.scrollBy(0, 300);")
                clickme_xpath = "//button[@id='gprnvalidate_button' and @class='validate__button' and @form='gprn_validation_form' and text()='Validate']"
                click_element_with_js(driver, clickme_xpath)
                fill_text_field_with_js_placeholder(driver, 'Eg.5555', row['Gas Rate'])
                driver.execute_script(f"window.scrollBy(0, 300);")
                if row['Gas_meter_location'] == 'Outside':
                    try:
                        radio_labell = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div/div[2]/div/div[2]/div[3]/div[1]/form/div[2]/div/div[7]/div/label/span[2]"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView(true);", radio_labell)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", radio_labell)
                        print(f"Clicked the radio button at position .")
                    except Exception as e:
                        print(f"Failed to click the radio button at position : {e}")

            else:
                print("No Energy type found")
            if row['Category'] == 'Gas':
                if row['is_gas_name'] == 'Yes':
                    click_yes_or_no_js(driver, row['is_gas_name'], 'gasowner')
                else:
                    pass
            elif row['Category'] == 'Electricity':
                if row['is_elec_name'] == 'Yes': 
                    click_yes_or_no_js(driver, row['is_elec_name'], 'electricowner')
                else:
                    pass
            elif row['Category'] == 'Electricity and Gas':
                if row['is_elec_name'] == 'Yes': 
                    click_yes_or_no_js(driver, row['is_elec_name'], 'electricowner')
                else:
                    pass
                if row['is_gas_name'] == 'Yes':
                    click_yes_or_no_js(driver, row['is_gas_name'], 'gasowner')
                else:
                    pass
            else:
                print("Not Found Electric name, Gas name")

            if row['Is Home Address same as Billing Address?'] == 'Yes':  
                click_yes_or_no_js(driver, row['Is Home Address same as Billing Address?'], 'addressowner')
                correc_drop = row['Billing Country']
                corresponding_dropdown(driver, correc_drop)
                fill_text_field(driver, 'txtAddressline1', row['Home Address'])
                fill_text_field(driver, 'txtCity', row['Home City'])
                correc_drop_county = row['Home Region']
                corres_county_dropdown(driver, correc_drop_county)
                fill_text_field_with_js_ID(driver, 'txtEirCode', row['Home Eircode Code'])
                click_anywhere(driver)
            else:
                pass

            click_next = "//input[@type='submit' and @class='button--directdebit--largeprimary col-9 col-sm-8 col-md-6 col-lg-7' and @form='proeprty_details' and @value='Continue to Direct Debit setup']"
            click_element_with_js(driver, click_next)
            fill_text_field(driver, 'accountName', row['Account Name'])
            click_Iban = "//a[@class='bank__IBAN' and contains(text(), 'I’d like to use my IBAN instead')]"
            click_element_with_js(driver, click_Iban)
            time.sleep(2)
            fill_text_field_with_js_Iban(driver, 'ibannumber', row['IBAN'])
            do_it_xpath = "/html/body/div[1]/main/div/div[2]/div/div[2]/div[3]/div[2]/input"
            click_element_with_js(driver, do_it_xpath)
            checkbox_xpath_1 = '/html/body/div[1]/main/div/div[2]/div/div[2]/div[3]/div[1]/div[5]/div[2]/div/div/div[2]/div/div[3]/div[2]/label/span[2]'
            click_checkbox_with_js(driver, checkbox_xpath_1)
            time.sleep(1)
            debit_authorisation_button_xpath = "//button[@type='button' and @id='debit_authorisation_button' and contains(@class, 'button--largeprimary') and text()='Okay']"
            click_element_with_js(driver, debit_authorisation_button_xpath)
            
            click_next = "//input[@id='review_submit' and @type='submit' and @value='All good. Complete my switch to Energia']"
            click_element_with_js(driver, click_next)
            passed_rows.append(row.to_dict())
            time.sleep(5)
            
            
        except Exception as e:
            print(f"Failed to process row {index}: {e}")
            failed_rows.append(row.to_dict())
        finally:
            driver.quit()
            time.sleep(1)

    # Write to Excel after all rows have been processed
    df_passed = pd.DataFrame(passed_rows)
    df_failed = pd.DataFrame(failed_rows)

    # Define the output directory and file path within your Django project
    output_dir = os.path.join(settings.MEDIA_ROOT, 'output_files')  # Save files in media/output_files/
    os.makedirs(output_dir, exist_ok=True)  # Create the directory if it doesn't exist

    output_file_name = 'output.xlsx'
    output_file_path = os.path.join(output_dir, output_file_name)

    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        df_passed.to_excel(writer, sheet_name='Passed', index=False)
        df_failed.to_excel(writer, sheet_name='Failed', index=False)

    print(f"Processed rows saved to {output_file_path}")

    # Return the file path after everything is processed and written
    return output_file_path