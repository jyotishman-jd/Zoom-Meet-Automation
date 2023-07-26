import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyperclip
from selenium.webdriver.common.action_chains import ActionChains
import yaml
from selenium.webdriver.common.keys import Keys


def schedule_zoom_meeting(meeting_topic, meeting_date, meeting_time,meeting_timep,attendee_email,duration):
    # Loading the gmail credentials
    with open("credentials.yml") as f:
       credentials = yaml.load(f, Loader=yaml.FullLoader)

    # Importing from Credentials
    email = credentials["username"]
    password = credentials["password2"]


    # Set up the Chrome WebDriver
    driver = webdriver.Chrome()

    # Log into Zoom
    driver.get('https://zoom.us/signin')

    driver.maximize_window() #This maximizes the Browser Window , I did this to match the Xpaths



    # Enter email
    email_field = WebDriverWait(driver, 20).until(
     EC.presence_of_element_located((By.ID, 'email')))
    email_field.send_keys(email)

    # Enter password
    password_field = driver.find_element(By.ID, 'password')
    password_field.send_keys(password)

    # Click on Sign In button
    signin_button = WebDriverWait(driver, 20).until(
     EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "Sign In")]')))
    signin_button.click()

    # Add a delay after clicking Sign In
    time.sleep(3) 

    # Wait for the page to load after signing in
    WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="onboarding"]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div[1]/div/div[2]/div[1]/button/span')))
    # Create a meeting
    driver.get('https://zoom.us/meeting/schedule')

    # Wait for the page to load for scheduling a meeting
    WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[2]/form/div[1]/div[1]/div/div[1]/div[1]/label')))

    # Enter meeting details
    
    topic_field = driver.find_element(By.ID, 'topic')
    topic_field.send_keys(meeting_topic)

    # Date set :
    date_field = driver.find_element(By.ID, 'mt_time')
    date_value = meeting_date  
    driver.execute_script("arguments[0].removeAttribute('readonly')", date_field)
    date_field.clear()
    date_field.send_keys(date_value)
    date_field.click()
    
    #Set time :
    
    # Find the time input field
    time_field = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[2]/form/div[1]/div[2]/div/div[2]/div[2]/div[1]/div[1]/div/div/input')
    time_field.click()
    

    # Set your own time value 
    desired_time = meeting_time  
    time_field.send_keys(desired_time)
    
    # Press Enter to confirm the value
    time_field.send_keys(Keys.ENTER)
    
    #Set AM or PM
    if meeting_timep == 'PM' or meeting_timep == 'pm' :
        timep_dropdown = driver.find_element(By.ID, 'start_time2')
        timep_dropdown.click() 
        timep_option = driver.find_element(By.XPATH,'//*[@id="select-item-start_time2-1"]/div')
        timep_option.click()

    

    # Set the meeting duration
    duration_dropdown = driver.find_element(By.XPATH, '//*[@id="start_time"]')
    duration_dropdown.click()
    if duration == 1 :
         
        duration_option = driver.find_element(By.XPATH, '//*[@id="select-item-start_time-1"]/div')
        duration_option.click()
    elif duration == 2 :
        duration_option = driver.find_element(By.XPATH, '//*[@id="select-item-start_time-2"]/div')
        duration_option.click()
    elif duration == 3 :
        duration_option = driver.find_element(By.XPATH, '//*[@id="select-item-start_time-3"]/div')
        duration_option.click()
    elif duration == 4 :
        duration_option = driver.find_element(By.XPATH, '//*[@id="select-item-start_time-4"]/div')
        duration_option.click()
    elif duration == 5 :
        duration_option = driver.find_element(By.XPATH, '//*[@id="select-item-start_time-5"]/div')
        duration_option.click()
    elif duration == 6 :
        duration_option = driver.find_element(By.XPATH, '//*[@id="select-item-start_time-6"]/div')
        duration_option.click()                
            


    # Enter attendee details
    

    # Locate the attendee field
    attendee_field = driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[3]/div[3]/div/div/div[2]/div[2]/div/div/div[2]/form/div[1]/div[6]/div/div/div[2]/div[1]/div[1]/div[1]/div/div/div/input')
    attendee_field.clear()
    attendee_field.send_keys(attendee_email)

    #Waitfor attendee name to appear:
    WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="select-item-pmc-item-0"]/div/p')))


    # Select the attendee option
    
    attendee_option = driver.find_element(By.XPATH, '//*[@id="select-item-pmc-item-0"]/div/p')
    actions = ActionChains(driver)
    actions.move_to_element(attendee_option).perform()

    attendee_option.click()

    #Saving Part :
    
    save_option = driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[3]/div[3]/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div[1]/button[1]/span')
    actions = ActionChains(driver)
    actions.move_to_element(save_option).perform()

    #Click on save button :
    save_button = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[2]/div[1]/div[1]/button[1]')))
    save2_button = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div/div[2]/div[1]/div[1]/button[1]')))
    save2_button.click()




    #Waiting for next Page to load 
    WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[4]/div[3]/div[3]/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div[1]/button[1]')))

    #Wait for next details page to load :
    WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.ID, 'tab-detail')))
    WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="registration"]/button/i')))


    # Get the meeting link
    meeting_link_element = driver.find_element(By.XPATH, '//*[@id="registration"]/button/i')
    meeting_link_element.click()

    #Link Extraction and Terminal Visuals
    print("Meeting Scheduled Successfully!")
    MeetLink = pyperclip.paste()
    Link= "Your requested Zoom Meeting has been scheduled successfully , here is your meeting link:", MeetLink
    print(Link)
    
    # Close the browser
    driver.quit()
    return Link

