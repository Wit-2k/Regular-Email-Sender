import tkinter as tk
from tkinter import messagebox
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import schedule
import time
import datetime
import threading
import re
import json
import os

CONFIG_FILE = "config.json"

def get_smtp_server(email):
    """
    Determines the SMTP server and port based on the email domain.
    """
    domain_match = re.search(r"@([\w.-]+)", email)
    if not domain_match:
        return None
    domain = domain_match.group(1).lower()
    
    smtp_servers = {
        "gmail.com": ("smtp.gmail.com", 465),
        "outlook.com": ("smtp.office365.com", 587),
        "hotmail.com": ("smtp.office365.com", 587),
        "live.com": ("smtp.office365.com", 587),
        "yahoo.com": ("smtp.mail.yahoo.com", 465),
        "aol.com": ("smtp.aol.com", 465),
        "zoho.com": ("smtp.zoho.com", 465),
        "mail.com": ("smtp.mail.com", 465),
        "gmx.com": ("mail.gmx.com", 465),
        "163.com": ("smtp.163.com", 465),
        "126.com": ("smtp.126.com", 465),
        "yeah.net": ("smtp.yeah.net", 465),
        "qq.com": ("smtp.qq.com", 465),
        "foxmail.com": ("smtp.qq.com", 465),
    }
    return smtp_servers.get(domain)

class EmailSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Scheduler (Group Sending)")

        # --- UI Fields Initialization ---
        self.setup_ui()
        
        # --- Load state and set closing protocol ---
        self.load_state()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        """Creates all the UI widgets."""
        # --- Static UI Fields ---
        tk.Label(self.root, text="Sender Name:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.sender_name_entry = tk.Entry(self.root, width=40)
        self.sender_name_entry.grid(row=0, column=1, padx=10, pady=5, columnspan=2)

        tk.Label(self.root, text="Sender Email:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.sender_email_entry = tk.Entry(self.root, width=40)
        self.sender_email_entry.grid(row=1, column=1, padx=10, pady=5, columnspan=2)

        tk.Label(self.root, text="Email Password/Auth Code:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.sender_password_entry = tk.Entry(self.root, show="*", width=40)
        self.sender_password_entry.grid(row=2, column=1, padx=10, pady=5, columnspan=2)
        
        # --- Dynamic Recipient Fields ---
        tk.Label(self.root, text="Number of Recipients:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.num_recipients_entry = tk.Entry(self.root, width=10)
        self.num_recipients_entry.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        
        self.generate_button = tk.Button(self.root, text="Generate", command=self.generate_recipient_fields)
        self.generate_button.grid(row=3, column=2, padx=5, pady=5, sticky="w")

        self.recipients_frame = tk.Frame(self.root)
        self.recipients_frame.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky="w")
        self.recipient_entries = []

        # --- Email Content and Timing ---
        tk.Label(self.root, text="Send Time (HH:MM):").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.send_time_entry = tk.Entry(self.root, width=40)
        self.send_time_entry.grid(row=5, column=1, padx=10, pady=5, columnspan=2)

        tk.Label(self.root, text="Subject:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.subject_entry = tk.Entry(self.root, width=40)
        self.subject_entry.grid(row=6, column=1, padx=10, pady=5, columnspan=2)

        tk.Label(self.root, text="Body:").grid(row=7, column=0, sticky="nw", padx=10, pady=5)
        self.body_text = tk.Text(self.root, width=40, height=10)
        self.body_text.grid(row=7, column=1, padx=10, pady=5, columnspan=2)

        self.schedule_button = tk.Button(self.root, text="Schedule Email", command=self.schedule_email)
        self.schedule_button.grid(row=8, column=1, columnspan=2, pady=10)

        self.status_label = tk.Label(self.root, text="", fg="blue")
        self.status_label.grid(row=9, column=0, columnspan=3, pady=5)

    def on_closing(self):
        """Handles the window closing event."""
        self.save_state()
        self.root.destroy()

    def save_state(self):
        """Saves the current state of the input fields to a JSON file."""
        state = {
            "sender_name": self.sender_name_entry.get(),
            "sender_email": self.sender_email_entry.get(),
            "send_time": self.send_time_entry.get(),
            "subject": self.subject_entry.get(),
            "body": self.body_text.get("1.0", "end-1c"),
            "recipients": [entry.get() for entry in self.recipient_entries]
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=4)
        except Exception as e:
            print(f"Error saving state: {e}")

    def load_state(self):
        """Loads the state from the JSON file if it exists."""
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                state = json.load(f)
            
            # --- Restore static fields ---
            self.sender_name_entry.insert(0, state.get("sender_name", ""))
            self.sender_email_entry.insert(0, state.get("sender_email", ""))
            self.send_time_entry.insert(0, state.get("send_time", ""))
            self.subject_entry.insert(0, state.get("subject", ""))
            self.body_text.insert("1.0", state.get("body", ""))
            
            # --- Restore dynamic recipient fields ---
            recipients = state.get("recipients", [])
            if recipients:
                self.num_recipients_entry.insert(0, str(len(recipients)))
                self.generate_recipient_fields()
                for i, email in enumerate(recipients):
                    if i < len(self.recipient_entries):
                        self.recipient_entries[i].insert(0, email)
        except (json.JSONDecodeError, KeyError, Exception) as e:
            print(f"Error loading state from {CONFIG_FILE}: {e}")

    def generate_recipient_fields(self):
        """Clears old recipient fields and generates new ones based on user input."""
        for widget in self.recipients_frame.winfo_children():
            widget.destroy()
        self.recipient_entries.clear()

        try:
            num = int(self.num_recipients_entry.get())
            if num <= 0:
                raise ValueError("Number must be positive")
        except (ValueError, TypeError):
            messagebox.showerror("Input Error", "Please enter a valid positive number for recipients.")
            return

        for i in range(num):
            label = tk.Label(self.recipients_frame, text=f"Recipient #{i+1} Email:")
            label.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = tk.Entry(self.recipients_frame, width=40)
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.recipient_entries.append(entry)

    def validate_inputs(self):
        """Validates all user inputs, including dynamic recipient fields."""
        # (Validation logic remains the same as the previous version)
        # Validate static fields
        static_entries = {
            "Sender Name": self.sender_name_entry, "Sender Email": self.sender_email_entry,
            "Email Password/Auth Code": self.sender_password_entry, "Send Time (HH:MM)": self.send_time_entry,
            "Subject": self.subject_entry
        }
        for name, entry in static_entries.items():
            if not entry.get().strip():
                messagebox.showerror("Input Error", f"'{name}' cannot be empty.")
                return False
        
        if not self.body_text.get("1.0", "end-1c").strip():
            messagebox.showerror("Input Error", "'Body' cannot be empty.")
            return False

        # Validate email and time formats
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        sender_email = self.sender_email_entry.get()
        if not re.match(email_regex, sender_email):
            messagebox.showerror("Input Error", "Invalid 'Sender Email' format.")
            return False

        time_regex = r'^([01]\d|2[0-3]):([0-5]\d)$'
        if not re.match(time_regex, self.send_time_entry.get()):
            messagebox.showerror("Input Error", "Invalid 'Send Time' format. Please use HH:MM (e.g., 09:30).")
            return False

        # Validate supported provider
        self.smtp_settings = get_smtp_server(sender_email)
        if not self.smtp_settings:
            messagebox.showerror("Input Error", f"Email provider for '{sender_email}' is not supported.")
            return False
            
        # Validate dynamic recipient fields
        if not self.recipient_entries:
            messagebox.showerror("Input Error", "Please generate recipient fields and provide at least one recipient.")
            return False
        
        for i, entry in enumerate(self.recipient_entries):
            email = entry.get().strip()
            if not email:
                messagebox.showerror("Input Error", f"Recipient #{i+1} email cannot be empty.")
                return False
            if not re.match(email_regex, email):
                messagebox.showerror("Input Error", f"Invalid format for Recipient #{i+1} email: {email}")
                return False

        return True

    def send_email(self):
        """Prepares and sends the email to all recipients."""
        # (Email sending logic remains the same as the previous version)
        sender_name = self.sender_name_entry.get()
        sender_email = self.sender_email_entry.get()
        sender_password = self.sender_password_entry.get()
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", "end-1c")
        
        recipient_emails = [entry.get() for entry in self.recipient_entries]
        
        smtp_server, smtp_port = self.smtp_settings

        message = MIMEText(body, "plain", "utf-8")
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = ", ".join(recipient_emails)
        message["Subject"] = Header(subject, "utf-8")

        try:
            if smtp_port == 465: # SSL
                server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            elif smtp_port == 587: # TLS
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
            
            with server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipient_emails, message.as_string())
            
            self.status_label.config(text=f"Success! Email sent to {len(recipient_emails)} recipients at {datetime.datetime.now().strftime('%H:%M:%S')}", fg="green")
            return schedule.CancelJob
        except Exception as e:
            self.status_label.config(text=f"Failed to send email: {e}", fg="red")

    def schedule_email(self):
        """Validates inputs and then schedules the email."""
        if self.validate_inputs():
            send_time = self.send_time_entry.get()
            schedule.clear()
            schedule.every().day.at(send_time).do(self.send_email)
            self.status_label.config(text=f"Email successfully scheduled for {send_time} every day.", fg="blue")

def run_scheduler():
    """Runs the scheduler loop."""
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSchedulerApp(root)
    
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()
    
    root.mainloop()