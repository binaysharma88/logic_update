from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import random
import string
from selenium.common.exceptions import TimeoutException
from logger import log_with_time
import time
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException, WebDriverException
)
import os
from reports import log_login_status,save_successful_recipients


    
def upload_pdf(driver,wait,wait1,wait2, path, log_fn=log_with_time):
    
    try:
    
       
        driver.get("https://apps.docusign.com/send/documents")
        time.sleep(10)

        
        start_now_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-qa="manage-sidebar-actions-meerkat-meerkat_create_envelope-text"]'))
        )
        driver.execute_script("arguments[0].click();", start_now_btn)
        print("✅ Clicked 'Start Now' to begin new envelope")

        #wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Start Now"]'))).click()
        try:
            wait2.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-qa="tutorial-got-it"]'))).click()
        except:
            pass
        
        try:
            toast_container = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "DsUiToastMessages"))
            )

            # Wait for a toast child with the expected message
            toast_message = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[@id="DsUiToastMessages"]//div[contains(text(), "maximum number of envelope drafts")]'))
            )
            toast_text = driver.execute_script("""
                const toast = document.getElementById('DsUiToastMessages');
                return toast ? toast.innerText : '';
            """)

            if "maximum number of envelope drafts" in toast_text.lower():
                log_fn("🚫 Envelope draft limit exceeded detected via toast.", "RED")
                return "draft_limit_exceeded"
        except Exception as e:
            #log_fn(f"⚠️ Toast check failed: {e}") 
            print(f"⚠️ Toast check failed: {e}")   

        except TimeoutException:
            # No toast appeared
            pass
        wait.until(EC.element_to_be_clickable((By.XPATH, '//button[.//span[text()="Upload"]]'))).click()
                
        
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]'))).send_keys(path)
        try:
            wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-qa="file-thumb-text"]')))
            log_fn(f"✅ File upload confirmed for: {path}")
        except TimeoutException:
            log_fn("❌ File upload not confirmed in time.", "RED")
            return False

        time.sleep(1)
        log_fn(f"PDF uploaded: {path}")
        print(f"PDF uploaded: {path}")
        
        error_message = driver.find_elements(
            By.XPATH,
            '//div[contains(text(), "Unable to access the envelope")]'
        )

        if error_message:
            log_fn("⚠️ DocuSign reported 'Unable to access the envelope'. Attempting recovery...", "YELLOW")
            try:
                start_now = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//span[@data-qa="manage-sidebar-actions-meerkat-meerkat_create_envelope-text"]'))
                )
                driver.execute_script("arguments[0].click();", start_now)
                log_fn("🔁 Clicked 'Start Now' again to recover.")
                # Try uploading again
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]'))).send_keys(path)
            except Exception as e:
                log_fn(f"❌ Failed to recover from envelope error: ", "RED")
                print(f"❌ Failed to recover from envelope error: {e}", "RED")
                return False

        try:
            add_recipient_btn = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-qa="homepage-favorite-template-header" and text()="Add recipients"]'))
            )
            add_recipient_btn.click()
            print("👤 Clicked 'Add recipients' successfully")
        except TimeoutException:
            print("❌ 'Add recipients' button not found or clickable in time.")

        try:
            add_message_btn = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-qa="homepage-favorite-template-header" and text()="Add message"]'))
            )
            add_message_btn.click()
            print("💬 Clicked 'Add message'")
        except TimeoutException:
            print("❌ 'Add message' not found")    

        time.sleep(1)
        return True
    except TimeoutException as e:
        #log_fn("❌ Timeout while trying to upload the PDF.")
        print(f"❌ Timeout while trying to upload the PDF{e}.")
        return False
    


def add_recipients(wait1, driver, rows, log_fn=log_with_time, log_fn_partial=None):
    for _, row in rows.iterrows():
        try:
            # 🕒 Wait for any existing toast to disappear
            try:
                WebDriverWait(driver, 5).until_not(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '[data-qa="ds-toast-content-text"]'))
                )
            except TimeoutException:
                log_fn("⚠️ Toast error message did not disappear in time", "YELLOW")

            # ➕ Click "Add Recipient" button
            wait1.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-qa="recipients-add"]'))).click()
            time.sleep(0.5)

            # 🔍 Find the newly added row (usually the last one)
            recipient_rows = driver.find_elements(By.CSS_SELECTOR, '[data-qa="recipient-row"]')
            current_row = recipient_rows[-1]

            # 📝 Fill Name and Email
            name_input = current_row.find_element(By.CSS_SELECTOR, '[data-qa="recipient-name-inner-wrapper"] input')
            email_input = current_row.find_element(By.CSS_SELECTOR, '[data-qa="recipient-email-wrapper"] input')

            name_input.send_keys(row['Name'])
            email_input.send_keys(row['Email'])

            # ✅ Confirm the fields were set correctly
            wait1.until(lambda d: name_input.get_attribute("value").strip() == row['Name'].strip())
            wait1.until(lambda d: email_input.get_attribute("value").strip() == row['Email'].strip())

            log_fn(f"🟢 Added recipient: {row['Name']} ({row['Email']})")

            # 🔽 Set recipient type to "Receives a Copy"
            dropdown = current_row.find_element(By.CSS_SELECTOR, '[data-qa="recipient-type"]')
            dropdown.click()
            wait1.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-qa="carbonCopies"]'))).click()

        except Exception as e:
            log_fn(f"❌ Error adding recipient {row['Email']}: {str(e)}")



def set_email_content_with_body(driver, subject, body, log_fn=log_with_time):
    wait = WebDriverWait(driver, 10)

    try:
        subject_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-qa="prepare-subject"]')))
        subject_input.send_keys(Keys.CONTROL + "a")
        subject_input.send_keys(Keys.DELETE)
        subject_input.send_keys(subject)

        try:
            wait.until(lambda d: subject_input.get_attribute("value").strip() == subject.strip())
            log_fn("✅ Subject added and verified")
        except TimeoutException:
            log_fn("⚠️ Subject filled, but verification timed out", "YELLOW")
    except Exception as e:
        log_fn(f"❌ Failed to set subject", "RED")
        print(f"❌ Failed to set subject: {e}", "RED")
        return False

    try:
        message_input = driver.find_element(By.CSS_SELECTOR, '[data-qa="prepare-message"]')
        message_input.send_keys(Keys.CONTROL + "a")
        message_input.send_keys(Keys.DELETE)
        message_input.send_keys(body)

        try:
            wait.until(lambda d: message_input.get_attribute("value").strip() == body.strip())
            log_fn("✅ Message added and verified")
        except TimeoutException:
            log_fn("⚠️ Message filled, but verification timed out", "YELLOW")
    except Exception as e:
        log_fn(f"❌ Failed to set message: ", "RED")
        print(f"❌ Failed to set message: {e}", "RED")
        return False

    return True

def send_envelope(driver, wait, wait1,wait2, log_fn=log_with_time):
    try:

        try:
            wait2.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-qa="tutorial-got-it"]'))).click()
            print("🧹 Tutorial popup dismissed.")
        except TimeoutException:
            pass

        # Click the "Next" or "Add Fields" button if required
        wait2.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-qa="footer-add-fields-link"]'))).click()
        time.sleep(1)

        try:
            toast_container = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "DsUiToastMessages"))
            )

            # Wait for a toast child with the expected message
            toast_message = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[@id="DsUiToastMessages"]//div[contains(text(), "maximum number of envelope drafts")]'))
            )
            toast_text = driver.execute_script("""
                const toast = document.getElementById('DsUiToastMessages');
                return toast ? toast.innerText : '';
            """)

            if "maximum number of envelope drafts" in toast_text.lower():
                log_fn("🚫 Envelope draft limit exceeded detected via toast.", "RED")
                return "draft_limit_exceeded"
        except Exception as e:
            #log_fn(f"⚠️ Toast check failed: {e}") 
            print(f"⚠️ Toast check failed: {e}")

        try:
            wait2.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-qa="tutorial-got-it"]'))).click()
            print("🧹 Tutorial popup dismissed.")
        except TimeoutException:
            pass

        # Click the "Send" button
        wait2.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-qa="footer-send-button"]'))).click()
        time.sleep(2)

        # Check if "Send Without Fields" confirmation is needed
        send_without_fields_buttons = driver.find_elements(By.CSS_SELECTOR, '[data-qa="send-without-fields"]')
        if send_without_fields_buttons:
            send_without_fields_buttons[0].click()
            print("✅ (without fields confirmation) clicked.")
        else:
            print("✅ clicked on send.")

        # Check for envelope allowance error
        elements = driver.find_elements(By.XPATH, '//div[@data-qa="ds-toast-content-text" and contains(text(), "envelope allowance for the account has been exceeded")]')
        if elements:
            return "limit_exceeded"

        # ✅ Wait for "Envelope Sent" confirmation
        try:
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//h1[@data-qa="sent-text" and contains(text(),"Your Envelope Was Sent")]')
            ))
            log_fn("📨 Envelope successfully sent. Returning to home...", "green")
            time.sleep(10)
            
            # Redirect to Home page
            '''driver.get("https://account.docusign.com/")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-qa="dashboard-navigation"]'))'''
            
        except TimeoutException:
            log_fn("⚠️ Envelope confirmation not found. Skipping home redirect.", "yellow")

        return "success"

    except TimeoutException:
        print("❌ Timeout while trying to send the envelope. The send button may not be clickable.")
        return False
    except Exception as e:
        print(f"❌ Error sending envelope: {str(e)}")
        return False


def process_single_round(driver, wait, wait1, wait2, recipients, pdf_path, subject, body,
                         stop_event, log_fn, log_fn_partial, report_path, email, password,
                         round_num, rounds, completed_accounts, rounds_completed,
                         recipients_df, total_sent_recipients, round_successful_recipients,stopped,completed):
    

    
    MAX_UPLOAD_TRIES = 2
    upload_success = False

    for attempt in range(MAX_UPLOAD_TRIES):
        upload_result = upload_pdf(driver, wait, wait1, wait2, pdf_path, log_fn=log_fn)

        if upload_result == "draft_limit_exceeded":
            log_login_status(report_path, email, password, "Draft limit exceeded", completed + round_num)
            completed_accounts.append(email)
            log_fn("🚫 Exiting account due to draft limit exceeded.", "RED")
            return {"status": "exit_account"}

        elif upload_result:
            upload_success = True
            #time.sleep(2)
            break
        else:
            log_fn(f"⚠️ Upload attempt {attempt + 1} failed. Retrying...")
            #time.sleep(2)

    if not upload_success:
        log_fn("❌ Upload failed after 2 attempts. Skipping this round.", "RED")
        log_login_status(report_path,email,password, f"PDF upload failed in Round {completed + round_num + 1}", completed + round_num)
        return {"status": "skip_round"}

    if stop_event and stop_event.is_set():
        log_fn("🔴 Automation stopped by user.", level="red")
        return {"status": "stopped"}

    

    add_recipients(wait, driver, recipients, log_fn=log_fn)
                    
    if stop_event and stop_event.is_set():
        log_fn("🔴 Automation stopped by user.", level="red")
        return {"status": "stopped"}
    
    if not set_email_content_with_body(driver, subject, body, log_fn=log_fn):
        log_fn("❌ Skipping round due to failure in setting subject/body.", "RED")
        return {"status": "skip_round"}


    if stop_event and stop_event.is_set():
        log_fn("🔴 Automation stopped by user.", level="red")
        return {"status": "stopped"}
    
    send_result = send_envelope(driver, wait, wait1,wait2, log_fn=log_fn)
    if send_result =="draft_limit_exceeded":
        log_fn("draft limit exeeded swithcing account")
        completed_accounts.append(email)
        return {"status": "exit_account"}
    if send_result == "success":
        sent_emails = recipients['Email'].tolist()
        #total_sent_recipients += len(sent_emails)
        #log_fn(f"✅ Envelope sent for {email} in Round {round_num + 1}.")
        log_fn(" =====================================================",level ="green")
        log_fn_partial([
            ("✅ Envelope sent for ", "GREEN"),
            (f"{email} in Round {round_num + 1}.", "INFO")
        ])
        for _, row in recipients.iterrows():
            #log_fn(f"   🟢 {row['Name']} - {row['Email']}")
            log_fn_partial([
                (" 🟢 ", "GREEN"),  # ✅ This part green
                (f"{row['Name']} - {row['Email']}", "INFO")  # ✅ This part normal (INFO color)
                ])
        log_fn(" =====================================================\n",level ="green")    
        #round_num =+ 1    
        
        for _, recipient_row in recipients.iterrows():
            round_successful_recipients.append((
                email,
                recipient_row['Name'],
                recipient_row['Email']
            ))
        save_successful_recipients(report_path,round_successful_recipients, log_fn)
        
    elif send_result == "limit_exceeded":
        log_fn(f"🚫 Envelope limit exceeded for {email}. Logging out.")
        log_login_status(report_path,email,password, "Limit Exceeded, account removed", completed+round_num)
        completed_accounts.append(email)

        return {"status": "exit_account"}      
    else:   # send_result == "failed"   
        log_fn(f"⚠️ Envelope sending failed for {email} in Round {round_num + 1}. Skipping recipient deletion.") 
    
    return {
            "status": "success",
            "used_emails": sent_emails,
            "added_successes": round_successful_recipients
        }