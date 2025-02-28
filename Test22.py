import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import imaplib
import email
import json
import os
from datetime import datetime, timedelta
from collections import defaultdict
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from wordcloud import WordCloud
import smtplib
from email.mime.text import MIMEText


class EmailOrganizer:
    def __init__(self):
        self.imap_server = None
        self.email_address = None
        self.password = None
        self.rules = []
        self.load_rules()
        self.load_auto_reply_settings()

    def load_auto_reply_settings(self):
        if os.path.exists("auto_reply_settings.json"):
            with open("auto_reply_settings.json", "r") as f:
                settings = json.load(f)
                self.auto_reply_var.set(settings.get("enabled", False))
                self.auto_reply_message.delete("1.0", tk.END)
                self.auto_reply_message.insert(tk.END, settings.get("message", ""))

    def auto_reply(self, email_message):
        if not os.path.exists("auto_reply_settings.json"):
            return  # Auto-reply not configured

        with open("auto_reply_settings.json", "r") as f:
            settings = json.load(f)
            if not settings["enabled"]:
                return  # Auto-reply is disabled

        sender = email.utils.parseaddr(email_message['From'])[1]
        subject = email_message["Subject"]

        reply_subject = f"Re: {subject}"
        reply_body = settings["message"]

        msg = MIMEText(reply_body)
        msg["Subject"] = reply_subject
        msg["From"] = self.email_address
        msg["To"] = sender

        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            server.login(self.email_address, self.password)
            server.sendmail(self.email_address, sender, msg.as_string())
            server.quit()
            print(f"Auto-reply sent to {sender}")
        except Exception as e:
            print("Error sending auto-reply:", e)

    def load_rules(self):
        try:
            if os.path.exists('email_rules.json'):
                with open('email_rules.json', 'r') as f:
                    self.rules = json.load(f)
            else:
                self.rules = []
        except Exception as e:
            print(f"Error loading rules: {e}")
            self.rules = []

    def save_rules(self):
        try:
            with open('email_rules.json', 'w') as f:
                json.dump(self.rules, f, indent=2)
        except Exception as e:
            print(f"Error saving rules: {e}")

    def connect(self, email_address, password, imap_server="imap.gmail.com"):
        try:
            self.imap_server = imaplib.IMAP4_SSL(imap_server)
            self.imap_server.login(email_address, password)
            self.email_address = email_address
            self.password = password
            return True
        except Exception as e:
            print(f"Connection error: {e}")
            return False

    def analyze_emails(self, days=30):
        if not self.imap_server:
            return "Not connected to email server"

        try:
            self.imap_server.select('INBOX')
            date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
            _, messages = self.imap_server.search(None, f'(SINCE "{date}")')

            analytics = {
                'total_emails': 0,
                'sender_frequency': defaultdict(int),
                'hourly_distribution': defaultdict(int),
                'average_response_time': 0,
                'subject_keywords': defaultdict(int),
                'email_sizes': [],
                'attachment_types': defaultdict(int)
            }

            total_response_time = 0
            response_count = 0
            last_received_time = None

            for num in messages[0].split():
                _, msg_data = self.imap_server.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                analytics['total_emails'] += 1

                sender = email.utils.parseaddr(email_message['From'])[1]
                analytics['sender_frequency'][sender] += 1

                date_tuple = email.utils.parsedate_tz(email_message['Date'])
                if date_tuple:
                    local_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                    analytics['hourly_distribution'][local_date.hour] += 1

                if email_message['In-Reply-To']:
                    if last_received_time:
                        response_time = (local_date - last_received_time).total_seconds() / 60
                        total_response_time += response_time
                        response_count += 1
                last_received_time = local_date

                subject = email_message['Subject']
                if subject:
                    words = subject.lower().split()
                    for word in words:
                        if len(word) > 3:
                            analytics['subject_keywords'][word] += 1

                # Email size
                analytics['email_sizes'].append(len(email_body))

                # Attachment types
                if email_message.is_multipart():
                    for part in email_message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        file_name = part.get_filename()
                        if file_name:
                            file_ext = os.path.splitext(file_name)[1].lower()
                            analytics['attachment_types'][file_ext] += 1

            if response_count > 0:
                analytics['average_response_time'] = total_response_time / response_count

            return analytics
        except Exception as e:
            return f"Error analyzing emails: {str(e)}"

    def process_emails(self):
        if not self.imap_server:
            return "Not connected to email server"

        try:
            self.imap_server.select('INBOX')
            _, messages = self.imap_server.search(None, 'UNSEEN')

            processed = 0
            for msg_num in messages[0].split():
                _, msg_data = self.imap_server.fetch(msg_num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                self.auto_reply(email_message)  # Enable auto-reply

                for rule in self.rules:
                    if self.match_rule(email_message, rule):
                        self.imap_server.copy(msg_num, rule['folder'])
                        self.imap_server.store(msg_num, '+FLAGS', '\\Deleted')
                        processed += 1
                        break

            self.imap_server.expunge()
            return f"Processed {processed} emails"

        except Exception as e:
            return f"Error processing emails: {str(e)}"

    def match_rule(self, email_message, rule):
        if rule['condition_type'] == 'from':
            return rule['condition_value'].lower() in email_message['From'].lower()
        elif rule['condition_type'] == 'subject':
            return rule['condition_value'].lower() in email_message['Subject'].lower()
        elif rule['condition_type'] == 'body':
            return self.check_body_content(email_message, rule['condition_value'])
        return False

    def check_body_content(self, email_message, keyword):
        if email_message.is_multipart():
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":
                    return keyword.lower() in part.get_payload().lower()
        else:
            return keyword.lower() in email_message.get_payload().lower()

    def search_emails(self, query, days=30):
        if not self.imap_server:
            return "Not connected to email server"

        try:
            self.imap_server.select('INBOX')
            date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
            search_criteria = f'(SINCE "{date}") SUBJECT "{query}"'
            _, messages = self.imap_server.search(None, search_criteria)

            results = []
            for num in messages[0].split():
                _, msg_data = self.imap_server.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                subject = email_message['Subject']
                sender = email.utils.parseaddr(email_message['From'])[1]
                date = email.utils.parsedate_to_datetime(email_message['Date'])

                results.append({
                    'subject': subject,
                    'sender': sender,
                    'date': date.strftime("%Y-%m-%d %H:%M:%S")
                })

            return results
        except Exception as e:
            return f"Error searching emails: {str(e)}"


class EmailOrganizerGUI:
    def __init__(self):
        self.organizer = EmailOrganizer()
        self.window = ttkb.Window(themename="cosmo")
        self.window.title("Email Organizer & Analytics")
        self.window.geometry("1200x800")
        self.setup_gui()

    def setup_gui(self):
        self.create_menu()
        self.create_notebook()
        self.create_status_bar()
        self.create_auto_reply_tab()  # Ensure this line is included

    def create_menu(self):
        menu_bar = tk.Menu(self.window)
        self.window.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Connect", command=self.show_login_dialog)
        file_menu.add_command(label="Exit", command=self.window.quit)

        view_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Toggle Dark Mode", command=self.toggle_dark_mode)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about_dialog)
        help_menu.add_command(label="Documentation", command=self.open_documentation)

    def create_notebook(self):
        style = ttkb.Style()
        style.configure('TNotebook.Tab', padding=(20, 10))

        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(expand=True, fill="both", padx=20, pady=20)

        self.create_dashboard_tab()
        self.create_rules_tab()
        self.create_search_tab()
        self.create_settings_tab()

    def create_auto_reply_tab(self):
        auto_reply_frame = ttk.Frame(self.notebook)
        self.notebook.add(auto_reply_frame, text="Auto-Reply")

        ttk.Label(auto_reply_frame, text="Enable Auto-Reply:").pack(pady=10)
        self.auto_reply_var = tk.BooleanVar(value=False)
        auto_reply_checkbox = ttkb.Checkbutton(auto_reply_frame, variable=self.auto_reply_var, bootstyle="primary")
        auto_reply_checkbox.pack()

        ttk.Label(auto_reply_frame, text="Reply Message:").pack(pady=10)
        self.auto_reply_message = tk.Text(auto_reply_frame, height=5, width=50)
        self.auto_reply_message.insert(tk.END, "Thank you for your email. I will get back to you soon.")
        self.auto_reply_message.pack(pady=5)

        save_button = ttkb.Button(auto_reply_frame, text="Save Settings", command=self.save_auto_reply, bootstyle="success", width=20)
        save_button.pack(pady=20)

    def create_dashboard_tab(self):
        dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(dashboard_frame, text="Dashboard")

        button_frame = ttk.Frame(dashboard_frame)
        button_frame.pack(pady=20)

        analyze_button = ttkb.Button(
            button_frame,
            text="Analyze Emails",
            command=self.show_analytics,
            bootstyle="primary",
            width=20
        )
        analyze_button.pack(side=tk.LEFT, padx=10)

        process_button = ttkb.Button(
            button_frame,
            text="Process Emails",
            command=self.process_emails,
            bootstyle="success",
            width=20
        )
        process_button.pack(side=tk.LEFT, padx=10)

        self.analytics_frame = ttk.Frame(dashboard_frame)
        self.analytics_frame.pack(expand=True, fill="both", padx=20, pady=20)

    def create_rules_tab(self):
        rules_frame = ttk.Frame(self.notebook)
        self.notebook.add(rules_frame, text="Email Rules")

        add_rule_button = ttkb.Button(
            rules_frame,
            text="Add Rule",
            command=self.show_add_rule_dialog,
            bootstyle="info",
            width=20
        )
        add_rule_button.pack(pady=20)

        self.rules_tree = ttk.Treeview(rules_frame, columns=('Name', 'Conditions', 'Folder'), show='headings')
        self.rules_tree.heading('Name', text='Rule Name')
        self.rules_tree.heading('Conditions', text='Conditions')
        self.rules_tree.heading('Folder', text='Folder')
        self.rules_tree.pack(expand=True, fill="both", padx=40, pady=40)

        self.update_rules_list()

    def create_search_tab(self):
        search_frame = ttk.Frame(self.notebook)
        self.notebook.add(search_frame, text="Search Emails")

        search_entry = ttkb.Entry(search_frame, width=40, bootstyle="primary")
        search_entry.pack(pady=20)

        search_button = ttkb.Button(
            search_frame,
            text="Search",
            command=lambda: self.search_emails(search_entry.get()),
            bootstyle="primary",
            width=20
        )
        search_button.pack(pady=10)

        self.search_results = ttk.Treeview(search_frame, columns=('Subject', 'Sender', 'Date'), show='headings')
        self.search_results.heading('Subject', text='Subject')
        self.search_results.heading('Sender', text='Sender')
        self.search_results.heading('Date', text='Date')
        self.search_results.pack(expand=True, fill="both", padx=20, pady=20)

    def create_settings_tab(self):
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings")

        ttk.Label(settings_frame, text="IMAP Server:").pack(pady=10)
        self.imap_server_entry = ttkb.Entry(settings_frame, width=40, bootstyle="primary")
        self.imap_server_entry.insert(0, "imap.gmail.com")
        self.imap_server_entry.pack(pady=5)

        ttk.Label(settings_frame, text="Default Analysis Period (days):").pack(pady=10)
        self.analysis_period_entry = ttkb.Entry(settings_frame, width=40, bootstyle="primary")
        self.analysis_period_entry.insert(0, "30")
        self.analysis_period_entry.pack(pady=5)

        save_settings_button = ttkb.Button(
            settings_frame,
            text="Save Settings",
            command=self.save_settings,
            bootstyle="success",
            width=20
        )
        save_settings_button.pack(pady=20)

    def save_auto_reply(self):
        settings = {
            "enabled": self.auto_reply_var.get(),
            "message": self.auto_reply_message.get("1.0", tk.END).strip()
        }
        with open("auto_reply_settings.json", "w") as f:
            json.dump(settings, f, indent=2)
        messagebox.showinfo("Success", "Auto-Reply settings saved successfully!")

    def create_status_bar(self):
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self.window, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def show_login_dialog(self):
        dialog = ttkb.Toplevel(self.window)
        dialog.title("Email Login")
        dialog.geometry("400x400")

        ttk.Label(dialog, text="Email:").pack(pady=10)
        email_entry = ttkb.Entry(dialog, width=40, bootstyle="primary")
        email_entry.pack(pady=5)

        ttk.Label(dialog, text="Password:").pack(pady=10)
        password_entry = ttkb.Entry(dialog, width=40, show="*", bootstyle="primary")
        password_entry.pack(pady=5)

        connect_button = ttkb.Button(
            dialog,
            text="Connect",
            command=lambda: self.connect_to_email(email_entry.get(), password_entry.get(), dialog),
            bootstyle="success",
            width=20
        )
        connect_button.pack(pady=20)

    def connect_to_email(self, email_address, password, dialog):
        if self.organizer.connect(email_address, password):
            self.status_var.set("Connected to email server successfully!")
            messagebox.showinfo("Success", "Connected to email server successfully!")
            dialog.destroy()
        else:
            self.status_var.set("Error: Failed to connect to email server")
            messagebox.showerror("Error", "Failed to connect to email server")

    def show_analytics(self):
        if not self.organizer.imap_server:
            messagebox.showerror("Error", "Please connect to your email first.")
            return

        days = simpledialog.askinteger("Email Analytics", "Analyze emails from how many days ago?",
                                         minvalue=1, maxvalue=365, initialvalue=30)
        if days:
            self.status_var.set("Analyzing emails...")
            analytics = self.organizer.analyze_emails(days)
            self.status_var.set("Analysis complete.")
            self.display_analytics(analytics)

    def display_analytics(self, analytics):
        for widget in self.analytics_frame.winfo_children():
            widget.destroy()

        total_emails = analytics.get('total_emails', 0)
        average_response_time = analytics.get('average_response_time', 0)
        sender_frequency = analytics.get('sender_frequency', {})
        hourly_distribution = analytics.get('hourly_distribution', {})
        subject_keywords = analytics.get('subject_keywords', {})
        email_sizes = analytics.get('email_sizes', [])
        attachment_types = analytics.get('attachment_types', {})

        # Create a frame for text labels
        text_frame = ttk.Frame(self.analytics_frame)
        text_frame.pack(fill="x", pady=10)

        ttk.Label(text_frame, text=f"Total Emails: {total_emails}", font=("Helvetica", 14, "bold")).pack(anchor="w")
        ttk.Label(text_frame, text=f"Average Response Time: {average_response_time:.2f} minutes", font=("Helvetica", 10)).pack(anchor="w")

        # Create a figure for visualizations
        fig = plt.figure(figsize=(5, 5))

        # Hourly distribution
        ax1 = fig.add_subplot(221)
        hours = list(range(24))
        counts = [hourly_distribution.get(hour, 0) for hour in hours]
        ax1.bar(hours, counts, color='skyblue')
        ax1.set_xlabel('Hour of Day')
        ax1.set_ylabel('Number of Emails')
        ax1.set_title('Hourly Email Distribution')
        ax1.set_xticks(range(0, 24, 2))

        # Top senders
        ax2 = fig.add_subplot(222)
        top_senders = sorted(sender_frequency.items(), key=lambda x: x[1], reverse=True)[:10]
        senders, counts = zip(*top_senders)
        ax2.barh(senders, counts, color='lightgreen')
        ax2.set_xlabel('Number of Emails')
        ax2.set_title('Top 10 Senders')

        # Email size distribution
        ax3 = fig.add_subplot(223)
        ax3.hist(email_sizes, bins=20, color='lightcoral')
        ax3.set_xlabel('Email Size (bytes)')
        ax3.set_ylabel('Frequency')
        ax3.set_title('Email Size Distribution')

        # Word cloud of subject keywords
        ax4 = fig.add_subplot(224)
        wordcloud = WordCloud(width=400, height=400, background_color='white').generate_from_frequencies(subject_keywords)
        ax4.imshow(wordcloud, interpolation='bilinear')
        ax4.axis('off')
        ax4.set_title('Subject Keywords')

        plt.tight_layout()

        # Embed the matplotlib figure in the Tkinter window
        canvas = FigureCanvasTkAgg(fig, master=self.analytics_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both", pady=20)

        # Display attachment types
        if attachment_types:
            attachment_frame = ttk.Frame(self.analytics_frame)
            attachment_frame.pack(fill="x", pady=10)

            ttk.Label(attachment_frame, text="Attachment Types:", font=("Helvetica", 14, "bold")).pack(anchor="w")
            for ext, count in sorted(attachment_types.items(), key=lambda x: x[1], reverse=True)[:5]:
                ttk.Label(attachment_frame, text=f"{ext}: {count}").pack(anchor="w")

    def process_emails(self):
        if not self.organizer.imap_server:
            messagebox.showerror("Error", "Please connect to your email first.")
            return

        self.status_var.set("Processing emails...")
        result = self.organizer.process_emails()
        self.status_var.set(result)
        messagebox.showinfo("Processing Result", result)

    def show_add_rule_dialog(self):
        dialog = ttkb.Toplevel(self.window)
        dialog.title("Add Email Rule")
        dialog.geometry("500x500")

        ttk.Label(dialog, text="Rule Name:").pack(pady=10)
        rule_name_entry = ttkb.Entry(dialog, width=40, bootstyle="primary")
        rule_name_entry.pack(pady=5)

        ttk.Label(dialog, text="Condition Type:").pack(pady=10)
        condition_type = tk.StringVar()
        condition_type_combo = ttkb.Combobox(dialog, textvariable=condition_type, values=["from", "subject", "body"], bootstyle="primary")
        condition_type_combo.pack(pady=5)

        ttk.Label(dialog, text="Condition Value:").pack(pady=10)
        condition_value_entry = ttkb.Entry(dialog, width=40, bootstyle="primary")
        condition_value_entry.pack(pady=5)

        ttk.Label(dialog, text="Target Folder:").pack(pady=10)
        target_folder_entry = ttkb.Entry(dialog, width=40, bootstyle="primary")
        target_folder_entry.pack(pady=5)

        add_button = ttkb.Button(
            dialog,
            text="Add Rule",
            command=lambda: self.add_rule(rule_name_entry.get(),
                                           condition_type.get(),
                                           condition_value_entry.get(),
                                           target_folder_entry.get(),
                                           dialog),
            bootstyle="success",
            width=20
        )
        add_button.pack(pady=20)

    def add_rule(self, rule_name, condition_type, condition_value, target_folder, dialog):
        if rule_name and condition_type and condition_value and target_folder:
            rule = {
                "name": rule_name,
                "condition_type": condition_type,
                "condition_value": condition_value,
                "folder": target_folder
            }
            self.organizer.rules.append(rule)
            self.organizer.save_rules()
            self.update_rules_list()
            messagebox.showinfo("Success", "Rule added successfully!")
            dialog.destroy()
        else:
            messagebox.showerror("Error", "All fields must be filled!")

    def update_rules_list(self):
        for row in self.rules_tree.get_children():
            self.rules_tree.delete(row)
        for rule in self.organizer.rules:
            self.rules_tree.insert('', 'end', values=(rule['name'], f"{rule['condition_type']}: {rule['condition_value']}", rule['folder']))

    def search_emails(self, query):
        if not self.organizer.imap_server:
            messagebox.showerror("Error", "Please connect to your email first.")
            return

        self.status_var.set("Searching emails...")
        results = self.organizer.search_emails(query)
        self.status_var.set("Search complete.")

        for row in self.search_results.get_children():
            self.search_results.delete(row)

        if isinstance(results, list):
            for result in results:
                self.search_results.insert('', 'end', values=(result['subject'], result['sender'], result['date']))
        else:
            messagebox.showerror("Error", results)

    def toggle_dark_mode(self):
        current_theme = self.window.style.theme_use()
        if current_theme == "cosmo":
            self.window.style.theme_use("darkly")
        else:
            self.window.style.theme_use("cosmo")

    def save_settings(self):
        imap_server = self.imap_server_entry.get()
        analysis_period = self.analysis_period_entry.get()
        messagebox.showinfo("Settings Saved", f"IMAP Server: {imap_server}\nDefault Analysis Period: {analysis_period} days")

    def show_about_dialog(self):
        about_text = """
        Email Organizer & Analytics
        Version 1.0

        This application helps you processes and analyze your emails.
        It provides features like email rule management, email search,
        and detailed email analytics.

        Created by Logic Lords Team
        """
        messagebox.showinfo("About", about_text)

    def open_documentation(self):
        import webbrowser
        webbrowser.open("https://example.com/email-organizer-docs")

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = EmailOrganizerGUI()
    app.run()