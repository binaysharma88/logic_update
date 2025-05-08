from pathlib import Path
import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from pdf_generator import create_pdf_for_batch, generate_transaction_id, initialize_pdfkit
from send_envelope import process_single_round
from login import login_to_docusign
from popuphandle import handle_not_now_popup
from save_updated_data import save_updated_data
from logger import log_with_time
from reports import initialize_reports, log_login_status
from checks import is_driver_alive
from datetime import datetime

def is_file_locked(filepath):
    try:
        with open(filepath, "a"):
            return False
    except PermissionError:
        return True

def run_automation(base_folder, log_fn=log_with_time, log_fn_partial=None, stop_event=None,
                   headless=True, use_generated_pdf=False, use_random_size_generated_pdf=False):
    try:
        initialize_pdfkit(base_folder)

        # Paths
        excel_path = base_folder / "email_details.xlsx"
        pdf_folder = base_folder / "PDF"
        templates_folder = base_folder / "templates"
        generated_pdfs_folder = base_folder / "generated_pdfs"
        report_path = initialize_reports()

        # Checks
        if is_file_locked(excel_path):
            log_fn("🚫 Please close 'email_details.xlsx' before running automation.")
            return
        if is_file_locked(report_path):
            log_fn("🚫 Please close 'reports.xlsx' before running automation.")
            return
        if not excel_path.exists():
            raise FileNotFoundError(f"email_details.xlsx not found in {base_folder}")

        # Read Excel
        accounts_df = pd.read_excel(excel_path, sheet_name="Accounts")
        recipients_df = pd.read_excel(excel_path, sheet_name="Recipients")
        send_plan_df = pd.read_excel(excel_path, sheet_name="SendPlan")
        pdf_generate_df = pd.read_excel(excel_path, sheet_name="PdfGenerate")

        if accounts_df.empty or accounts_df['Email'].dropna().empty:
            log_fn("⚠️ No accounts available to process.")
            return
        
        if recipients_df.empty or recipients_df['Email'].dropna().empty:
            log_fn("⚠️ No recipient available to process.")
            return

        if log_fn_partial is None:
            def log_fn_partial(parts): log_fn("".join([x[0] for x in parts]))

        log_fn_partial([("✅ Excel loaded successfully.\n", "GREEN")])

    except Exception as e:
        log_fn(f"❌ Automation setup failed: {e}")
        return

    completed_accounts = []
    total_sent_recipients = 0
    stopped = False

    log_fn_partial([("===============🔄 Starting Automation===============\n", "GREEN")])

    for _, account in accounts_df.iterrows():
        if stop_event and stop_event.is_set():
            if not stopped:
                log_fn_partial([
                    ("🔴 Automation stopped by user.", "RED"),
                    (f" Stopped at account: {account['Email']}", "INFO")
                ])
                stopped = True
            break
        if len(recipients_df) < 5:
                log_fn("🚫 Not enough recipients remaining.")
                break

        email = account.get("Email", "").strip()
        password = account.get("Password", "").strip()
        rounds = int(account.get("Rounds", 0))

        if not email or not password:
            log_fn(f"⚠️ Skipping account with missing credentials.")
            continue

        total_needed_recipients = rounds * 5

        # Setup WebDriver
        options = webdriver.ChromeOptions()
        options.add_argument("--incognito")
        options.add_argument("--window-size=1000,1080")
        if headless:
            options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 15)
        wait1 = WebDriverWait(driver, 7)
        wait2 = WebDriverWait(driver, 2)

        log_fn(f"\n🔐 Attempting login with {email}...")
        login_result = login_to_docusign(driver, email, password, wait, wait1, wait2, log_fn=log_fn)

        if not login_result:
            log_fn_partial([("❌ Login failed for ", "RED"), (f"{email}.", "INFO")])
            log_login_status(report_path, email, password, "Login Failed, account removed", 0)
            completed_accounts.append(email)
            driver.quit()
            continue

        log_fn_partial([("✅ Login successful for ", "GREEN"), (f"{email}", "INFO")])
        handle_not_now_popup(wait2, log_fn)
        #time.sleep(5)

        completed = 0  # If you're tracking already-sent rounds, adjust this
        if completed >= rounds:
            log_fn(f"✅ No rounds left for {email}. Skipping.")
            log_login_status(report_path, email, password, "All round completed, account removed", completed)
            completed_accounts.append(email)
            driver.quit()
            continue

        pdf_generate_queue = pdf_generate_df.sample(frac=1).reset_index(drop=True).to_dict('records')
        send_plan_queue = send_plan_df.sample(frac=1).reset_index(drop=True).to_dict('records')
        round_num = 0
        
        while round_num < (rounds - completed):
            if stop_event and stop_event.is_set():
                stopped = True
                break
            if len(recipients_df) < 5:
                log_fn("🚫 Not enough recipients remaining.")
                break

            log_fn(f"➡️ Round {completed + round_num + 1} for {email}...")

            
            recipients = recipients_df.iloc[:5].reset_index(drop=True)

            if not send_plan_queue:
                send_plan_queue = send_plan_df.sample(frac=1).reset_index(drop=True).to_dict('records')

            subject, body, pdf_path = None, None, None

            if use_generated_pdf:
                if not pdf_generate_queue:
                    pdf_generate_df = pd.read_excel(excel_path, sheet_name="PdfGenerate")
                    pdf_generate_queue = pdf_generate_df.sample(frac=1).reset_index(drop=True).to_dict('records')

                plan = pdf_generate_queue.pop(0)
                pdf_generate_queue.append(plan)
                transaction_id = generate_transaction_id()
                subject = plan['Subject']
                body_template = plan['Body']
                amount = plan['Amount']
                number = plan['Number']
                product = plan['Product']
                date = datetime.now().strftime("%Y-%m-%d")

                body = body_template.format(
                    TransactionID=transaction_id,
                    Amount=amount,
                    Number=number,
                    Product=product,
                    Date=date
                )

                pdf_path = create_pdf_for_batch(
                    batch_recipients=recipients.to_dict('records'),
                    product=product,
                    body_content=body_template,
                    number=number,
                    amount=amount,
                    transaction_id=transaction_id,
                    templates_folder=templates_folder,
                    output_folder=generated_pdfs_folder,
                    date=date
                )
                log_fn(f"🧾 PDF generated: {pdf_path}")

            else:
                plan = send_plan_queue.pop(0)
                send_plan_queue.append(plan)
                subject = plan['Subject']
                body = plan['Body']
                pdf_name = plan['PDF_name']
                pdf_path = os.path.join(pdf_folder, f"{pdf_name}.pdf")

            round_successful_recipients = []
            result = process_single_round(
                driver, wait, wait1, wait2,
                recipients=recipients,
                pdf_path=pdf_path,
                subject=subject,
                body=body,
                stop_event=stop_event,
                log_fn=log_fn,
                log_fn_partial=log_fn_partial,
                report_path=report_path,
                email=email,
                password=password,
                round_num=round_num,
                rounds=rounds,
                completed_accounts=completed_accounts,
                rounds_completed=0,
                recipients_df=recipients_df,
                total_sent_recipients=total_sent_recipients,
                round_successful_recipients=round_successful_recipients,
                stopped=stopped,
                completed=completed
            )

            status = result.get("status")
            if status == "exit_account":
                break
            elif status == "skip_round":
                round_num += 1
                continue
            elif status == "stopped":
                stopped = True
                break
            elif status == "success":
                round_num += 1
                total_sent_recipients += len(result["used_emails"])
                recipients_df = recipients_df[~recipients_df['Email'].isin(result["used_emails"])].reset_index(drop=True)

        if completed + round_num >= rounds:
            log_fn(f"✅ Finished all rounds for {email}. Marking account as completed.")
            completed_accounts.append(email)
            log_login_status(report_path, email, password, "All rounds completed", completed + round_num)

        log_fn_partial([("🔄 Switching to next account...\n", "GREEN")])
        driver.quit()

    # Final cleanup
    log_fn_partial([
        ("✅ Total ", "INFO"),
        (f"{total_sent_recipients}", "GREEN"),
        (" recipients sent successfully.", "INFO")
    ])
    save_updated_data(
        accounts_df=accounts_df,
        recipients_df=recipients_df,
        completed_accounts=completed_accounts,
        excel_path=excel_path,
        log_fn=log_fn
    )
    try: driver.quit()
    except: pass
    log_fn_partial([("==================✅ Automation Completed=================\n", "GREEN")])

if __name__ == "__main__":
    base_folder = Path(__file__).resolve().parent
    run_automation(base_folder)
