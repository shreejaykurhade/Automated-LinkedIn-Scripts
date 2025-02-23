import os
import random
import threading
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from tkinter import Tk, Label, Text, Button, Entry, messagebox, ttk

# Configuration
EXCEL_PATH = "LinkedIn_Connections.xlsx"
MAX_DAILY_MESSAGES = 30
DEFAULT_MIN_DELAY = 2
DEFAULT_MAX_DELAY = 5

# Initialize the WebDriver
def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
    return webdriver.Chrome(options=options)

driver = init_driver()

# Load or create the Excel file for connections
if os.path.exists(EXCEL_PATH):
    df = pd.read_excel(EXCEL_PATH)
else:
    df = pd.DataFrame(columns=["Profile URL", "Name", "Job Title", "Location", "Message Sent"])

scraping_active = True
sent_messages_today = 0

def random_delay(min_delay=DEFAULT_MIN_DELAY, max_delay=DEFAULT_MAX_DELAY):
    time.sleep(random.uniform(min_delay, max_delay))

# Login GUI
def create_login_gui():
    login_window = Tk()
    login_window.title("LinkedIn Login")
    login_window.geometry("300x200")

    Label(login_window, text="LinkedIn Username:").pack(pady=5)
    username_entry = Entry(login_window, width=40)
    username_entry.pack()

    Label(login_window, text="LinkedIn Password:").pack(pady=5)
    password_entry = Entry(login_window, show="*", width=40)
    password_entry.pack()

    Button(login_window, text="Login", command=lambda: linkedin_login(username_entry.get(), password_entry.get(), login_window)).pack(pady=10)

    login_window.mainloop()

def linkedin_login(username, password, login_window):
    global scraping_active
    if not username or not password:
        messagebox.showwarning("Input Error", "Please enter both username and password.")
        return

    driver.get("https://www.linkedin.com/login")
    random_delay()

    # Entering credentials
    try:
        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password + Keys.RETURN)
        random_delay()
    except Exception as e:
        messagebox.showerror("Login Error", f"Failed to log in: {e}")
        return

    # Start scraping connections
    threading.Thread(target=update_excel_with_connections, daemon=True).start()
    login_window.destroy()
    open_main_gui()

# Function to scrape connections and update Excel
def update_excel_with_connections():
    global scraping_active, sent_messages_today
    connections_data = []

    driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
    random_delay()

    while scraping_active:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)  # Wait for new content to load

        connections = driver.find_elements(By.CSS_SELECTOR, ".mn-connection-card__details")
        for connection in connections:
            try:
                name = connection.find_element(By.CSS_SELECTOR, ".mn-connection-card__name").text
                job_title = connection.find_element(By.CSS_SELECTOR, ".mn-connection-card__occupation").text
                profile_url = connection.find_element(By.CSS_SELECTOR, "a").get_attribute("href")

                # Add to connections data
                connections_data.append({
                    "Profile URL": profile_url,
                    "Name": name,
                    "Job Title": job_title,
                    "Location": "N/A",
                    "Message Sent": "No"
                })
            except Exception as e:
                print(f"Error retrieving connection details: {e}")

        # Click on "Show more results" button if available
        try:
            show_more_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Show more results')]")
            if show_more_button.is_displayed():
                show_more_button.click()
                random_delay()
        except Exception:
            break  # No more results to load

    # Save connections to Excel
    connections_df = pd.DataFrame(connections_data)
    updated_df = pd.concat([df, connections_df]).drop_duplicates(subset="Profile URL", keep="first").reset_index(drop=True)
    updated_df.to_excel(EXCEL_PATH, index=False)

# Main GUI for messaging
def open_main_gui():
    main_window = Tk()
    main_window.title("LinkedIn Message Automation")
    main_window.geometry("500x600")

    Label(main_window, text="Enter Message:").pack(pady=10)
    message_box = Text(main_window, height=5, width=50)
    message_box.pack()

    progress_label = Label(main_window, text="Progress: 0%")
    progress_label.pack(pady=10)
    progress = ttk.Progressbar(main_window, orient="horizontal", length=400, mode="determinate")
    progress.pack(pady=10)

    # Define button actions
    def threaded_message_sending(send_function):
        threading.Thread(target=send_function, daemon=True).start()

    def send_to_specific_profile():
        profile_url = profile_url_entry.get()
        custom_message = message_box.get("1.0", "end-1c")

        if profile_url and custom_message:
            connect_and_message({"Profile URL": profile_url, "Name": "Custom"}, custom_message)

    def send_to_custom_profiles():
        custom_message = message_box.get("1.0", "end-1c")
        if not custom_message:
            messagebox.showwarning("Input Error", "Please enter a message to send.")
            return

        profiles = df[df["Message Sent"] != "Yes"].head(MAX_DAILY_MESSAGES)
        total_profiles = len(profiles)

        for idx, (_, profile) in enumerate(profiles.iterrows(), start=1):
            connect_and_message(profile, custom_message)

            progress["value"] = (idx / total_profiles) * 100
            progress_label.config(text=f"Progress: {int((idx / total_profiles) * 100)}%")
            main_window.update_idletasks()

    Label(main_window, text="Send to specific profile URL:").pack(pady=10)
    profile_url_entry = Entry(main_window, width=50)
    profile_url_entry.pack(pady=5)
    Button(main_window, text="Send to Specific Profile", command=lambda: threaded_message_sending(send_to_specific_profile)).pack(pady=5)

    Button(main_window, text="Send to Custom Profiles", command=lambda: threaded_message_sending(send_to_custom_profiles)).pack(pady=5)

    # Button to stop scraping
    def stop_scraping():
        global scraping_active
        scraping_active = False
        messagebox.showinfo("Scraping Stopped", "The scraping process has been stopped.")

    Button(main_window, text="Stop Scraping", command=stop_scraping).pack(pady=20)
    main_window.mainloop()

def connect_and_message(profile, custom_message):
    global sent_messages_today

    if sent_messages_today >= MAX_DAILY_MESSAGES:
        messagebox.showinfo("Limit Reached", "You have reached the daily limit for messages.")
        return

    user_url = profile["Profile URL"]
    name = profile["Name"]
    personalized_message = f"Hello {name}, {custom_message}"

    driver.get(user_url)
    random_delay()

    try:
        connect_button = driver.find_element(By.XPATH, "//button[text()='Connect']")
        connect_button.click()
        random_delay()

        add_note_button = driver.find_element(By.XPATH, "//button[text()='Add a note']")
        add_note_button.click()

        message_box = driver.find_element(By.XPATH, "//textarea[@name='message']")
        message_box.send_keys(personalized_message)

        send_button = driver.find_element(By.XPATH, "//button[text()='Send']")
        send_button.click()
        random_delay()

        # Update Excel with message sent status
        df.loc[df["Profile URL"] == user_url, "Message Sent"] = "Yes"
        df.to_excel(EXCEL_PATH, index=False)
        sent_messages_today += 1

        print(f"Connected and messaged {name} successfully!")
    except Exception as e:
        print(f"Could not connect to {name}: {e}")

# Start the program
create_login_gui()
