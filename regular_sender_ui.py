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

def get_smtp_server(email):
    """
    Determines the SMTP server and port based on the email domain.
    """
    # 正则表达式提取域名
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
        self.root.title("Email Scheduler")

        # UI布局
        fields = [
            "Sender Name:", "Sender Email:", "Email Password/Auth Code:",
            "Recipient Name:", "Recipient Email:", "Send Time (HH:MM):", "Subject:"
        ]
        
        self.entries = {}
        for i, field in enumerate(fields):
            tk.Label(root, text=field).grid(row=i, column=0, sticky="w", padx=10, pady=5)
            entry = tk.Entry(root, width=40)
            if "Password" in field:
                entry.config(show="*")
            entry.grid(row=i, column=1, padx=10, pady=5)
            self.entries[field] = entry

        tk.Label(root, text="Body:").grid(row=len(fields), column=0, sticky="nw", padx=10, pady=5)
        self.body_text = tk.Text(root, width=40, height=10)
        self.body_text.grid(row=len(fields), column=1, padx=10, pady=5)

        self.schedule_button = tk.Button(root, text="Schedule Email", command=self.schedule_email)
        self.schedule_button.grid(row=len(fields) + 1, column=1, pady=10)

        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.grid(row=len(fields) + 2, column=0, columnspan=2, pady=5)

    def validate_inputs(self):
        """Validates all user inputs."""
        # 1. 非空验证
        for field, entry in self.entries.items():
            if not entry.get().strip():
                messagebox.showerror("Input Error", f"'{field}' cannot be empty.")
                return False
        
        if not self.body_text.get("1.0", "end-1c").strip():
            messagebox.showerror("Input Error", "'Body' cannot be empty.")
            return False

        # 2. 邮箱格式验证
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        sender_email = self.entries["Sender Email:"].get()
        receiver_email = self.entries["Recipient Email:"].get()
        
        if not re.match(email_regex, sender_email):
            messagebox.showerror("Input Error", "Invalid 'Sender Email' format.")
            return False
        if not re.match(email_regex, receiver_email):
            messagebox.showerror("Input Error", "Invalid 'Recipient Email' format.")
            return False

        # 3. 时间格式验证
        time_regex = r'^([01]\d|2[0-3]):([0-5]\d)$'
        send_time = self.entries["Send Time (HH:MM):"].get()
        if not re.match(time_regex, send_time):
            messagebox.showerror("Input Error", "Invalid 'Send Time' format. Please use HH:MM (e.g., 09:30).")
            return False

        # 4. 邮箱服务商支持验证
        self.smtp_settings = get_smtp_server(sender_email)
        if not self.smtp_settings:
            messagebox.showerror("Input Error", f"Email provider for '{sender_email}' is not supported.")
            return False
            
        return True

    def send_email(self):
        """Prepares and sends the email."""
        # 从字典中获取值
        details = {field: entry.get() for field, entry in self.entries.items()}
        
        sender_name = details["Sender Name:"]
        sender_email = details["Sender Email:"]
        sender_password = details["Email Password/Auth Code:"]
        receiver_name = details["Recipient Name:"]
        receiver_email = details["Recipient Email:"]
        subject = details["Subject:"]
        body = self.body_text.get("1.0", "end-1c")
        
        smtp_server, smtp_port = self.smtp_settings

        message = MIMEText(body, "plain", "utf-8")
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = formataddr((receiver_name, receiver_email))
        message["Subject"] = Header(subject, "utf-8")

        try:
            if smtp_port == 465: # SSL
                with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, [receiver_email], message.as_string())
            elif smtp_port == 587: # TLS
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, [receiver_email], message.as_string())
            
            self.status_label.config(text=f"Success! Email sent to {receiver_email} at {datetime.datetime.now().strftime('%H:%M:%S')}", fg="green")
            return schedule.CancelJob # 发送成功后取消任务，防止重复发送
        except Exception as e:
            self.status_label.config(text=f"Failed to send email: {e}", fg="red")

    def schedule_email(self):
        """Validates inputs and then schedules the email."""
        if self.validate_inputs():
            send_time = self.entries["Send Time (HH:MM):"].get()
            # 清除之前的任务，以防用户多次点击按钮
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
    
    # 在后台线程中运行调度器，防止UI阻塞
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()
    
    root.mainloop()