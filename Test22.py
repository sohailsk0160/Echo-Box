from email.mime.text import MIMEText
import smtplib
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
import re
from PIL import Image, ImageTk


class EmailOrganizer:
    def __init__(self):
        self.imap_server = None
        self.email_address = None
        self.password = None
        self.rules = []
        self.load_rules()
        self.auto_reply_settings = self.load_auto_reply_settings()

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

    def load_auto_reply_settings(self):
        try:
            if os.path.exists('auto_reply_settings.json'):
                with open('auto_reply_settings.json', 'r') as f:
                    return json.load(f)
            else:
                return {"enabled": False, "message": ""}
        except Exception as e:
            print(f"Error loading auto-reply settings: {e}")
            return {"enabled": False, "message": ""}

    def save_auto_reply_settings(self, settings):
        try:
            with open('auto_reply_settings.json', 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            print(f"Error saving auto-reply settings: {e}")

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

            # Store the last analytics for theme switching
            self.last_analytics = analytics
            
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

                for rule in self.rules:
                    if self.match_rule(email_message, rule):
                        self.imap_server.copy(msg_num, rule['folder'])
                        self.imap_server.store(msg_num, '+FLAGS', '\\Deleted')
                        processed += 1
                        break

                # Auto-reply functionality
                if self.auto_reply_settings['enabled']:
                    self.send_auto_reply(email_message)

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

    def send_auto_reply(self, email_message):
        sender = email.utils.parseaddr(email_message['From'])[1]
        subject = "Re: " + email_message['Subject']
        body = self.auto_reply_settings['message']

        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = self.email_address
        msg['To'] = sender

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.starttls()
                smtp.login(self.email_address, self.password)
                smtp.send_message(msg)
        except Exception as e:
            print(f"Error sending auto-reply: {e}")

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
        
class AnalyticsWindow:
    def __init__(self, parent, analytics, is_dark_mode=True):
        self.window = ttkb.Toplevel(parent)
        self.window.title("Email Analytics")
        self.window.geometry("1200x800")
        self.analytics = analytics
        self.is_dark_mode = is_dark_mode
        
        # Apply theme based on parent's theme
        if is_dark_mode:
            self.window.style.theme_use("superhero")
        else:
            self.window.style.theme_use("cosmo")
            
        self.setup_ui()
        
    def setup_ui(self):
        # Create main container
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            header_frame, 
            text="Email Analytics Dashboard", 
            font=("Helvetica", 24, "bold")
        ).pack(side=tk.LEFT)
        
        # Create a frame for text labels with card-like styling
        text_frame = ttk.Frame(main_frame, style="Card.TFrame")
        text_frame.pack(fill="x", pady=10)

        # Add some visual flair with colored indicators
        stats_frame = ttk.Frame(text_frame)
        stats_frame.pack(fill="x", padx=20, pady=20)
        
        # Extract analytics data
        total_emails = self.analytics.get('total_emails', 0)
        average_response_time = self.analytics.get('average_response_time', 0)
        sender_frequency = self.analytics.get('sender_frequency', {})
        hourly_distribution = self.analytics.get('hourly_distribution', {})
        subject_keywords = self.analytics.get('subject_keywords', {})
        email_sizes = self.analytics.get('email_sizes', [])
        attachment_types = self.analytics.get('attachment_types', {})
        
        # Total emails with a visual indicator
        email_frame = ttk.Frame(stats_frame)
        email_frame.pack(side=tk.LEFT, padx=20)
        
        email_indicator = ttk.Frame(email_frame, width=15, height=40)
        email_indicator.configure(style="Success.TFrame")
        email_indicator.pack(side=tk.LEFT, padx=(0, 10))
        
        email_stats = ttk.Frame(email_frame)
        email_stats.pack(side=tk.LEFT)
        
        ttk.Label(email_stats, text="Total Emails", font=("Helvetica", 12, "bold")).pack(anchor="w")
        ttk.Label(email_stats, text=f"{total_emails}", font=("Helvetica", 20, "bold")).pack(anchor="w")
        
        # Response time with a visual indicator
        response_frame = ttk.Frame(stats_frame)
        response_frame.pack(side=tk.LEFT, padx=20)
        
        response_indicator = ttk.Frame(response_frame, width=15, height=40)
        response_indicator.configure(style="Info.TFrame")
        response_indicator.pack(side=tk.LEFT, padx=(0, 10))
        
        response_stats = ttk.Frame(response_frame)
        response_stats.pack(side=tk.LEFT)
        
        ttk.Label(response_stats, text="Avg Response Time", font=("Helvetica", 12, "bold")).pack(anchor="w")
        ttk.Label(response_stats, text=f"{average_response_time:.2f} min", font=("Helvetica", 20, "bold")).pack(anchor="w")
        
        # Add a third stat if available (e.g., total senders)
        if sender_frequency:
            sender_frame = ttk.Frame(stats_frame)
            sender_frame.pack(side=tk.LEFT, padx=20)
            
            sender_indicator = ttk.Frame(sender_frame, width=15, height=40)
            sender_indicator.configure(style="Warning.TFrame")
            sender_indicator.pack(side=tk.LEFT, padx=(0, 10))
            
            sender_stats = ttk.Frame(sender_frame)
            sender_stats.pack(side=tk.LEFT)
            
            ttk.Label(sender_stats, text="Unique Senders", font=("Helvetica", 12, "bold")).pack(anchor="w")
            ttk.Label(sender_stats, text=f"{len(sender_frequency)}", font=("Helvetica", 20, "bold")).pack(anchor="w")

        # Create tabs for different analytics views
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Overview tab with main charts
        overview_tab = ttk.Frame(notebook)
        notebook.add(overview_tab, text="Overview")
        
        # Distribution tab with more detailed charts
        distribution_tab = ttk.Frame(notebook)
        notebook.add(distribution_tab, text="Time Distribution")
        
        # Senders tab
        senders_tab = ttk.Frame(notebook)
        notebook.add(senders_tab, text="Senders Analysis")
        
        # Content tab
        content_tab = ttk.Frame(notebook)
        notebook.add(content_tab, text="Content Analysis")
        
        # Create charts for Overview tab
        self.create_overview_charts(overview_tab)
        
        # Create charts for Distribution tab
        self.create_distribution_charts(distribution_tab)
        
        # Create charts for Senders tab
        self.create_senders_charts(senders_tab)
        
        # Create charts for Content tab
        self.create_content_charts(content_tab)
        
        # Add a close button at the bottom
        close_button = ttkb.Button(
            main_frame,
            text="Close",
            command=self.window.destroy,
            bootstyle="secondary",
            width=15
        )
        close_button.pack(pady=10)
        
    def create_overview_charts(self, parent):
        # Create a figure for visualizations with a modern style
        plt.style.use('ggplot')
        fig = plt.figure(figsize=(12, 8))
        fig.patch.set_facecolor('#f0f0f0' if not self.is_dark_mode else '#2a2a2a')
        
        # Adjust subplot parameters for better spacing
        plt.subplots_adjust(hspace=0.4, wspace=0.4)
        
        # Extract analytics data
        hourly_distribution = self.analytics.get('hourly_distribution', {})
        sender_frequency = self.analytics.get('sender_frequency', {})
        email_sizes = self.analytics.get('email_sizes', [])
        subject_keywords = self.analytics.get('subject_keywords', {})

        # Hourly distribution
        ax1 = fig.add_subplot(221)
        hours = list(range(24))
        counts = [hourly_distribution.get(hour, 0) for hour in hours]
        bars = ax1.bar(hours, counts, color='#5cb85c', alpha=0.7)
        ax1.set_xlabel('Hour of Day', fontsize=10)
        ax1.set_ylabel('Number of Emails', fontsize=10)
        ax1.set_title('Hourly Email Distribution', fontweight='bold', fontsize=12)
        ax1.set_xticks(range(0, 24, 3))  # Show fewer x-ticks
        ax1.tick_params(axis='both', which='major', labelsize=9)  # Smaller tick labels
        ax1.grid(True, linestyle='--', alpha=0.7)
        
        # Add value labels on top of bars, but only for bars with significant height
        for bar in bars:
            height = bar.get_height()
            if height > max(counts) * 0.1:  # Only label bars that are at least 10% of the max height
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', fontsize=8)

        # Top senders
        ax2 = fig.add_subplot(222)
        top_senders = sorted(sender_frequency.items(), key=lambda x: x[1], reverse=True)[:5]
        if top_senders:
            # Truncate long email addresses for better display
            senders = [s[0][:15] + '...' if len(s[0]) > 15 else s[0] for s in top_senders]
            counts = [s[1] for s in top_senders]
            bars = ax2.barh(senders, counts, color='#5bc0de', alpha=0.7)
            ax2.set_xlabel('Number of Emails', fontsize=10)
            ax2.set_title('Top 5 Senders', fontweight='bold', fontsize=12)
            ax2.tick_params(axis='both', which='major', labelsize=9)  # Smaller tick labels
            ax2.grid(True, linestyle='--', alpha=0.7)
            
            # Add value labels with better positioning
            for bar in bars:
                width = bar.get_width()
                ax2.text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                        f'{int(width)}', ha='left', va='center', fontsize=8)
        else:
            ax2.text(0.5, 0.5, 'No sender data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax2.transAxes)

        # Email size distribution
        ax3 = fig.add_subplot(223)
        if email_sizes:
            n, bins, patches = ax3.hist(email_sizes, bins=15, color='#d9534f', alpha=0.7)  # Fewer bins
            ax3.set_xlabel('Email Size (KB)', fontsize=10)
            ax3.set_ylabel('Frequency', fontsize=10)
            ax3.set_title('Email Size Distribution', fontweight='bold', fontsize=12)
            ax3.tick_params(axis='both', which='major', labelsize=9)  # Smaller tick labels
            ax3.grid(True, linestyle='--', alpha=0.7)
            
            # Format x-axis to show KB instead of bytes with better formatting
            from matplotlib.ticker import FuncFormatter
            def kb_formatter(x, pos):
                return f'{x/1024:.0f}K' if x < 1024*1024 else f'{x/(1024*1024):.1f}M'
            ax3.xaxis.set_major_formatter(FuncFormatter(kb_formatter))
            
            # Limit the number of x-ticks to prevent overcrowding
            ax3.locator_params(axis='x', nbins=6)
        else:
            ax3.text(0.5, 0.5, 'No email size data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax3.transAxes)

        # Word cloud of subject keywords
        ax4 = fig.add_subplot(224)
        if subject_keywords:
            # Filter to only include more significant keywords
            significant_keywords = {k: v for k, v in subject_keywords.items() 
                                  if v > max(subject_keywords.values()) * 0.05}
            
            if significant_keywords:
                wordcloud = WordCloud(
                    width=400, 
                    height=400, 
                    background_color='white' if not self.is_dark_mode else '#2a2a2a',
                    colormap='viridis',
                    max_words=50,  # Limit number of words
                    contour_width=1,
                    contour_color='steelblue',
                    min_font_size=8,  # Ensure minimum font size
                    max_font_size=40  # Limit maximum font size
                ).generate_from_frequencies(significant_keywords)
                ax4.imshow(wordcloud, interpolation='bilinear')
                ax4.axis('off')
                ax4.set_title('Subject Keywords', fontweight='bold', fontsize=12)
            else:
                ax4.text(0.5, 0.5, 'No significant keywords found', 
                        horizontalalignment='center', verticalalignment='center',
                        transform=ax4.transAxes)
                ax4.axis('off')
        else:
            ax4.text(0.5, 0.5, 'No subject keyword data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax4.transAxes)
            ax4.axis('off')

        plt.tight_layout(pad=3.0)  # Increased padding between subplots

        # Embed the matplotlib figure in the Tkinter window
        canvas_frame = ttk.Frame(parent, style="Card.TFrame")
        canvas_frame.pack(expand=True, fill="both", pady=10, padx=10)
        
        canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both", padx=20, pady=20)
        
    def create_distribution_charts(self, parent):
        # Create a container frame
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Extract analytics data
        hourly_distribution = self.analytics.get('hourly_distribution', {})
        
        # Create a figure for time distribution
        plt.style.use('ggplot')
        fig = plt.figure(figsize=(12, 8))
        fig.patch.set_facecolor('#f0f0f0' if not self.is_dark_mode else '#2a2a2a')
        
        # Hourly distribution as a line chart
        ax1 = fig.add_subplot(211)
        hours = list(range(24))
        counts = [hourly_distribution.get(hour, 0) for hour in hours]
        
        # Add a line chart
        ax1.plot(hours, counts, marker='o', linestyle='-', color='#5cb85c', linewidth=2, markersize=8)
        ax1.set_xlabel('Hour of Day', fontsize=12)
        ax1.set_ylabel('Number of Emails', fontsize=12)
        ax1.set_title('Hourly Email Distribution (Line Chart)', fontweight='bold', fontsize=14)
        ax1.set_xticks(range(0, 24, 1))  # Show all hours
        ax1.tick_params(axis='both', which='major', labelsize=10)
        ax1.grid(True, linestyle='--', alpha=0.7)
        
        # Add value labels for each point
        for i, count in enumerate(counts):
            if count > 0:
                ax1.text(i, count + 0.3, f'{count}', ha='center', va='bottom', fontsize=9)
        
        # Group by time of day (morning, afternoon, evening, night)
        ax2 = fig.add_subplot(212)
        
        # Define time periods
        time_periods = {
            'Night (0-6)': sum(hourly_distribution.get(h, 0) for h in range(0, 6)),
            'Morning (6-12)': sum(hourly_distribution.get(h, 0) for h in range(6, 12)),
            'Afternoon (12-18)': sum(hourly_distribution.get(h, 0) for h in range(12, 18)),
            'Evening (18-24)': sum(hourly_distribution.get(h, 0) for h in range(18, 24))
        }
        
        periods = list(time_periods.keys())
        values = list(time_periods.values())
        
        # Create a pie chart
        wedges, texts, autotexts = ax2.pie(
            values, 
            labels=periods, 
            autopct='%1.1f%%',
            startangle=90,
            colors=['#5bc0de', '#5cb85c', '#f0ad4e', '#d9534f'],
            wedgeprops={'width': 0.5, 'edgecolor': 'w'},
            textprops={'fontsize': 12}
        )
        
        # Equal aspect ratio ensures that pie is drawn as a circle
        ax2.axis('equal')
        ax2.set_title('Email Distribution by Time of Day', fontweight='bold', fontsize=14)
        
        # Make the percentage labels more readable
        for autotext in autotexts:
            autotext.set_fontsize(10)
            autotext.set_weight('bold')
            autotext.set_color('white')
        
        plt.tight_layout(pad=3.0)
        
        # Embed the matplotlib figure in the Tkinter window
        canvas_frame = ttk.Frame(container, style="Card.TFrame")
        canvas_frame.pack(expand=True, fill="both", pady=10, padx=10)
        
        canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both", padx=20, pady=20)
        
    def create_senders_charts(self, parent):
        # Create a container frame
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Extract analytics data
        sender_frequency = self.analytics.get('sender_frequency', {})
        
        # Create a figure for sender analysis
        plt.style.use('ggplot')
        fig = plt.figure(figsize=(12, 8))
        fig.patch.set_facecolor('#f0f0f0' if not self.is_dark_mode else '#2a2a2a')
        
        # Top 10 senders bar chart
        ax1 = fig.add_subplot(211)
        top_senders = sorted(sender_frequency.items(), key=lambda x: x[1], reverse=True)[:10]
        
        if top_senders:
            # Extract domain from email for grouping
            def get_domain(email):
                return email.split('@')[-1] if '@' in email else email
            
            # Truncate long email addresses and add domain info
            senders = []
            for email, _ in top_senders:
                domain = get_domain(email)
                username = email.split('@')[0]
                if len(username) > 10:
                    username = username[:8] + '..'
                senders.append(f"{username}@{domain}")
            
            counts = [s[1] for s in top_senders]
            
            # Create horizontal bar chart
            bars = ax1.barh(senders, counts, color='#5bc0de', alpha=0.7)
            ax1.set_xlabel('Number of Emails', fontsize=12)
            ax1.set_title('Top 10 Senders', fontweight='bold', fontsize=14)
            ax1.tick_params(axis='both', which='major', labelsize=10)
            ax1.grid(True, linestyle='--', alpha=0.7)
            
            # Add value labels
            for bar in bars:
                width = bar.get_width()
                ax1.text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                        f'{int(width)}', ha='left', va='center', fontsize=9)
        else:
            ax1.text(0.5, 0.5, 'No sender data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax1.transAxes)
        
        # Domain distribution pie chart
        ax2 = fig.add_subplot(212)
        
        if sender_frequency:
            # Group by domain
            domain_counts = defaultdict(int)
            for email, count in sender_frequency.items():
                domain = get_domain(email)
                domain_counts[domain] += count
            
            # Get top domains
            top_domains = sorted(domain_counts.items(), key=lambda x: x[1], reverse=True)[:8]
            
            # Add "Other" category for remaining domains
            if len(domain_counts) > 8:
                other_count = sum(count for domain, count in domain_counts.items() 
                                if domain not in [d[0] for d in top_domains])
                if other_count > 0:
                    top_domains.append(('Other', other_count))
            
            domains = [d[0] for d in top_domains]
            counts = [d[1] for d in top_domains]
            
            # Create pie chart
            wedges, texts, autotexts = ax2.pie(
                counts, 
                labels=domains, 
                autopct='%1.1f%%',
                startangle=90,
                colors=plt.cm.tab10.colors,
                wedgeprops={'width': 0.5, 'edgecolor': 'w'},
                textprops={'fontsize': 12}
            )
            
            # Equal aspect ratio ensures that pie is drawn as a circle
            ax2.axis('equal')
            ax2.set_title('Email Distribution by Domain', fontweight='bold', fontsize=14)
            
            # Make the percentage labels more readable
            for autotext in autotexts:
                autotext.set_fontsize(10)
                autotext.set_weight('bold')
                autotext.set_color('white')
        else:
            ax2.text(0.5, 0.5, 'No domain data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax2.transAxes)
            ax2.axis('off')
        
        plt.tight_layout(pad=3.0)
        
        canvas_frame = ttk.Frame(container, style="Card.TFrame")
        canvas_frame.pack(expand=True, fill="both", pady=10, padx=10)
        
        canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both", padx=20, pady=20)
        
    def create_content_charts(self, parent):
        # Create a container frame
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Extract analytics data
        subject_keywords = self.analytics.get('subject_keywords', {})
        email_sizes = self.analytics.get('email_sizes', [])
        attachment_types = self.analytics.get('attachment_types', {})
        
        # Create a figure for content analysis
        plt.style.use('ggplot')
        fig = plt.figure(figsize=(12, 8))
        fig.patch.set_facecolor('#f0f0f0' if not self.is_dark_mode else '#2a2a2a')
        
        # Word cloud of subject keywords - larger and more detailed
        ax1 = fig.add_subplot(211)
        if subject_keywords:
            # Filter to only include more significant keywords
            significant_keywords = {k: v for k, v in subject_keywords.items() 
                                  if v > max(subject_keywords.values()) * 0.03}  # Lower threshold for more words
            
            if significant_keywords:
                wordcloud = WordCloud(
                    width=800, 
                    height=400, 
                    background_color='white' if not self.is_dark_mode else '#2a2a2a',
                    colormap='viridis',
                    max_words=100,  # More words
                    contour_width=1,
                    contour_color='steelblue',
                    min_font_size=8,
                    max_font_size=50
                ).generate_from_frequencies(significant_keywords)
                ax1.imshow(wordcloud, interpolation='bilinear')
                ax1.axis('off')
                ax1.set_title('Subject Keywords', fontweight='bold', fontsize=14)
            else:
                ax1.text(0.5, 0.5, 'No significant keywords found', 
                        horizontalalignment='center', verticalalignment='center',
                        transform=ax1.transAxes)
                ax1.axis('off')
        else:
            ax1.text(0.5, 0.5, 'No subject keyword data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax1.transAxes)
            ax1.axis('off')
        
        # Attachment types pie chart
        ax2 = fig.add_subplot(212)
        
        if attachment_types:
            # Sort attachment types by frequency
            sorted_types = sorted(attachment_types.items(), key=lambda x: x[1], reverse=True)
            
            # Get top types and combine the rest as "Other"
            top_types = sorted_types[:6]
            other_count = sum(count for _, count in sorted_types[6:])
            
            # Add "Other" category if needed
            if other_count > 0:
                types = [t[0] for t in top_types] + ['Other']
                counts = [t[1] for t in top_types] + [other_count]
            else:
                types = [t[0] for t in top_types]
                counts = [t[1] for t in top_types]
            
            # Create pie chart
            wedges, texts, autotexts = ax2.pie(
                counts, 
                labels=types, 
                autopct='%1.1f%%',
                startangle=90,
                colors=plt.cm.Set3.colors,
                wedgeprops={'width': 0.5, 'edgecolor': 'w'},
                textprops={'fontsize': 12}
            )
            
            # Equal aspect ratio ensures that pie is drawn as a circle
            ax2.axis('equal')
            ax2.set_title('Attachment Types Distribution', fontweight='bold', fontsize=14)
            
            # Make the percentage labels more readable
            for autotext in autotexts:
                autotext.set_fontsize(10)
                autotext.set_weight('bold')
                autotext.set_color('black')
        else:
            ax2.text(0.5, 0.5, 'No attachment data available', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax2.transAxes)
            ax2.axis('off')
        
        plt.tight_layout(pad=3.0)
        
        # Embed the matplotlib figure in the Tkinter window
        canvas_frame = ttk.Frame(container, style="Card.TFrame")
        canvas_frame.pack(expand=True, fill="both", pady=10, padx=10)
        
        canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both", padx=20, pady=20)


class EmailOrganizerGUI:
    def __init__(self):
        self.organizer = EmailOrganizer()
        self.window = ttkb.Window(themename="superhero")  # Changed theme for a modern look
        self.window.title("Echo-Box")
        self.window.geometry("1200x800")
        self.auto_reply_var = tk.BooleanVar(value=self.organizer.auto_reply_settings['enabled'])
        self.auto_reply_message = tk.Text()
        self.is_dark_mode = True  # Start with dark mode
        
        # Create icons first, before they're needed
        self.create_icons()
        
        # Then set up the GUI
        self.setup_gui()

    def create_icons(self):
        # Create simple icons using PIL if not available
        self.icons = {}
        
        # Dashboard icon
        dashboard_icon = Image.new('RGBA', (24, 24), (0, 0, 0, 0))
        for i in range(3):
            for j in range(3):
                if i != 1 or j != 1:  # Skip center
                    x, y = 2 + i*8, 2 + j*8
                    for dx in range(6):
                        for dy in range(6):
                            dashboard_icon.putpixel((x+dx, y+dy), (255, 255, 255, 255))
        self.icons['dashboard'] = ImageTk.PhotoImage(dashboard_icon)
        
        # Rules icon
        rules_icon = Image.new('RGBA', (24, 24), (0, 0, 0, 0))
        for i in range(3):
            y = 4 + i*8
            for x in range(4, 20):
                rules_icon.putpixel((x, y), (255, 255, 255, 255))
            for x in range(4, 8):
                rules_icon.putpixel((x, y+2), (255, 255, 255, 255))
        self.icons['rules'] = ImageTk.PhotoImage(rules_icon)
        
        # Search icon
        search_icon = Image.new('RGBA', (24, 24), (0, 0, 0, 0))
        # Draw circle
        for x in range(24):
            for y in range(24):
                dx, dy = x-10, y-10
                dist = (dx*dx + dy*dy)**0.5
                if 6 <= dist <= 8:
                    search_icon.putpixel((x, y), (255, 255, 255, 255))
        # Draw handle
        for i in range(5):
            x, y = 16+i, 16+i
            for j in range(3):
                search_icon.putpixel((x+j, y), (255, 255, 255, 255))
                search_icon.putpixel((x, y+j), (255, 255, 255, 255))
        self.icons['search'] = ImageTk.PhotoImage(search_icon)
        
        # Settings icon
        settings_icon = Image.new('RGBA', (24, 24), (0, 0, 0, 0))
        # Draw gear
        center_x, center_y = 12, 12
        for angle in range(0, 360, 45):
            rad = angle * 3.14159 / 180
            x = int(center_x + 8 * (angle % 90 == 0) * (1 if angle < 180 else -1))
            y = int(center_y + 8 * (angle % 90 != 0) * (1 if angle < 270 and angle > 90 else -1))
            for dx in range(-2, 3):
                for dy in range(-2, 3):
                    if 0 <= x+dx < 24 and 0 <= y+dy < 24:
                        settings_icon.putpixel((x+dx, y+dy), (255, 255, 255, 255))
        # Draw center circle
        for x in range(24):
            for y in range(24):
                dx, dy = x-center_x, y-center_y
                dist = (dx*dx + dy*dy)**0.5
                if dist <= 4:
                    settings_icon.putpixel((x, y), (255, 255, 255, 255))
        self.icons['settings'] = ImageTk.PhotoImage(settings_icon)
        
        # Auto-reply icon
        reply_icon = Image.new('RGBA', (24, 24), (0, 0, 0, 0))
        # Draw arrow
        for i in range(12):
            reply_icon.putpixel((6+i, 12), (255, 255, 255, 255))
        for i in range(5):
            reply_icon.putpixel((10+i, 8+i), (255, 255, 255, 255))
            reply_icon.putpixel((10+i, 16-i), (255, 255, 255, 255))
        self.icons['reply'] = ImageTk.PhotoImage(reply_icon)

    def setup_gui(self):
        self.create_menu()
        self.create_sidebar()
        self.create_main_content()
        self.create_status_bar()

    def create_menu(self):
        menu_bar = tk.Menu(self.window)
        self.window.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Connect", command=self.show_login_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.window.quit)

        view_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Toggle Dark Mode", command=self.toggle_dark_mode)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about_dialog)
        help_menu.add_command(label="Documentation", command=self.open_documentation)

    def create_sidebar(self):
        # Create a sidebar frame
        self.sidebar = ttk.Frame(self.window, style="Sidebar.TFrame")
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=0, pady=0)
        
        # App title
        title_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        title_frame.pack(fill=tk.X, padx=10, pady=20)
        ttk.Label(title_frame, text="Echo-Box", font=("Helvetica", 16, "bold"), 
                 style="SidebarTitle.TLabel").pack(side=tk.LEFT)
        
        # Create sidebar buttons with icons
        self.create_sidebar_button("Dashboard", self.icons['dashboard'], 0)
        self.create_sidebar_button("Email Rules", self.icons['rules'], 1)
        self.create_sidebar_button("Search", self.icons['search'], 2)
        self.create_sidebar_button("Settings", self.icons['settings'], 3)
        self.create_sidebar_button("Auto-Reply", self.icons['reply'], 4)
        
        # Add a connection status indicator
        self.connection_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        self.connection_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=20)
        
        self.connection_status = ttk.Label(
            self.connection_frame, 
            text="Not Connected", 
            style="ConnectionStatus.TLabel"
        )
        self.connection_status.pack(side=tk.LEFT, padx=5)
        
        # Connect button in sidebar
        connect_btn = ttkb.Button(
            self.connection_frame,
            text="Connect",
            command=self.show_login_dialog,
            bootstyle="outline-success",
            width=10
        )
        connect_btn.pack(side=tk.RIGHT, padx=5)

    def create_sidebar_button(self, text, icon, tab_index):
        btn_frame = ttk.Frame(self.sidebar, style="SidebarBtn.TFrame")
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        btn = ttkb.Button(
            btn_frame,
            text=f" {text}",
            image=icon,
            compound=tk.LEFT,
            command=lambda idx=tab_index: self.notebook.select(idx),
            bootstyle="outline-primary",
            width=20
        )
        btn.pack(fill=tk.X, padx=5, pady=5)
        return btn

    def create_main_content(self):
        # Create a main content frame
        self.main_content = ttk.Frame(self.window)
        self.main_content.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create notebook in main content
        self.notebook = ttk.Notebook(self.main_content, style="TNotebook")
        self.notebook.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        # Apply custom styling to tabs
        style = ttkb.Style()
        style.configure('TNotebook.Tab', padding=(15, 10), font=('Helvetica', 11))
        
        self.create_dashboard_tab()
        self.create_rules_tab()
        self.create_search_tab()
        self.create_settings_tab()
        self.create_auto_reply_tab()

    def create_dashboard_tab(self):
        dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(dashboard_frame, text="Dashboard")
        
        # Header with welcome message
        header_frame = ttk.Frame(dashboard_frame)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Email Analytics Dashboard", 
            font=("Helvetica", 20, "bold")
        ).pack(side=tk.LEFT)
        
        # Action buttons in a card-like container
        action_frame = ttk.Frame(dashboard_frame, style="Card.TFrame")
        action_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Add a subtle header to the card
        ttk.Label(
            action_frame, 
            text="Quick Actions", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        button_frame = ttk.Frame(action_frame)
        button_frame.pack(padx=20, pady=(0, 20))
        
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
        
        # Analytics content area with a card-like appearance
        analytics_container = ttk.Frame(dashboard_frame, style="Card.TFrame")
        analytics_container.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)
        
        # Add a header to the analytics card
        ttk.Label(
            analytics_container, 
            text="Email Analytics", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        self.analytics_frame = ttk.Frame(analytics_container)
        self.analytics_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=(0, 20))
        
        # Add a placeholder message
        placeholder = ttk.Label(
            self.analytics_frame,
            text="Click 'Analyze Emails' to view your email analytics",
            font=("Helvetica", 12),
            foreground="#888888"
        )
        placeholder.pack(expand=True)

    def create_rules_tab(self):
        rules_frame = ttk.Frame(self.notebook)
        self.notebook.add(rules_frame, text="Email Rules")
        
        # Header
        header_frame = ttk.Frame(rules_frame)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Email Rules Management", 
            font=("Helvetica", 20, "bold")
        ).pack(side=tk.LEFT)
        
        add_rule_button = ttkb.Button(
            header_frame,
            text="Add New Rule",
            command=self.show_add_rule_dialog,
            bootstyle="success",
            width=15
        )
        add_rule_button.pack(side=tk.RIGHT)
        
        # Rules list in a card-like container
        rules_container = ttk.Frame(rules_frame, style="Card.TFrame")
        rules_container.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)
        
        # Add a header to the rules card
        ttk.Label(
            rules_container, 
            text="Active Rules", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        # Create a styled treeview for rules
        tree_frame = ttk.Frame(rules_container)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=(0, 20))
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.rules_tree = ttk.Treeview(
            tree_frame, 
            columns=('Name', 'Conditions', 'Folder'), 
            show='headings',
            style="Treeview",
            yscrollcommand=tree_scroll.set
        )
        
        # Configure the scrollbar
        tree_scroll.config(command=self.rules_tree.yview)
        
        # Configure column widths and headings
        self.rules_tree.column('Name', width=200, anchor=tk.W)
        self.rules_tree.column('Conditions', width=400, anchor=tk.W)
        self.rules_tree.column('Folder', width=200, anchor=tk.W)
        
        self.rules_tree.heading('Name', text='Rule Name')
        self.rules_tree.heading('Conditions', text='Conditions')
        self.rules_tree.heading('Folder', text='Folder')
        
        self.rules_tree.pack(expand=True, fill=tk.BOTH)
        
        # Add buttons for rule management
        button_frame = ttk.Frame(rules_container)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        edit_button = ttkb.Button(
            button_frame,
            text="Edit Rule",
            bootstyle="warning",
            width=15,
            command=lambda: messagebox.showinfo("Info", "Edit functionality would go here")
        )
        edit_button.pack(side=tk.LEFT, padx=5)
        
        delete_button = ttkb.Button(
            button_frame,
            text="Delete Rule",
            bootstyle="danger",
            width=15,
            command=lambda: messagebox.showinfo("Info", "Delete functionality would go here")
        )
        delete_button.pack(side=tk.LEFT, padx=5)
        
        self.update_rules_list()

    def create_search_tab(self):
        search_frame = ttk.Frame(self.notebook)
        self.notebook.add(search_frame, text="Search Emails")
        
        # Header
        header_frame = ttk.Frame(search_frame)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Search Your Emails", 
            font=("Helvetica", 20, "bold")
        ).pack(side=tk.LEFT)
        
        # Search box in a card-like container
        search_container = ttk.Frame(search_frame, style="Card.TFrame")
        search_container.pack(fill=tk.X, padx=20, pady=10)
        
        # Add a header to the search card
        ttk.Label(
            search_container, 
            text="Search Criteria", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        # Create a search form
        form_frame = ttk.Frame(search_container)
        form_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        ttk.Label(form_frame, text="Search Term:").grid(row=0, column=0, padx=5, pady=10, sticky=tk.W)
        search_entry = ttkb.Entry(form_frame, width=40, bootstyle="primary")
        search_entry.grid(row=0, column=1, padx=5, pady=10, sticky=tk.W)
        
        ttk.Label(form_frame, text="Time Period:").grid(row=1, column=0, padx=5, pady=10, sticky=tk.W)
        period_combo = ttkb.Combobox(
            form_frame, 
            values=["Last 7 days", "Last 30 days", "Last 90 days", "All time"],
            bootstyle="primary",
            width=38
        )
        period_combo.current(1)  # Default to 30 days
        period_combo.grid(row=1, column=1, padx=5, pady=10, sticky=tk.W)
        
        search_button = ttkb.Button(
            form_frame,
            text="Search",
            command=lambda: self.search_emails(search_entry.get()),
            bootstyle="primary",
            width=15
        )
        search_button.grid(row=2, column=1, padx=5, pady=10, sticky=tk.E)
        
        # Results in a card-like container
        results_container = ttk.Frame(search_frame, style="Card.TFrame")
        results_container.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)
        
        # Add a header to the results card
        ttk.Label(
            results_container, 
            text="Search Results", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        # Create a styled treeview for search results
        tree_frame = ttk.Frame(results_container)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=(0, 20))
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.search_results = ttk.Treeview(
            tree_frame, 
            columns=('Subject', 'Sender', 'Date'), 
            show='headings',
            style="Treeview",
            yscrollcommand=tree_scroll.set
        )
        
        # Configure the scrollbar
        tree_scroll.config(command=self.search_results.yview)
        
        # Configure column widths and headings
        self.search_results.column('Subject', width=400, anchor=tk.W)
        self.search_results.column('Sender', width=200, anchor=tk.W)
        self.search_results.column('Date', width=200, anchor=tk.W)
        
        self.search_results.heading('Subject', text='Subject')
        self.search_results.heading('Sender', text='Sender')
        self.search_results.heading('Date', text='Date')
        
        self.search_results.pack(expand=True, fill=tk.BOTH)
        
        self.search_results.pack(expand=True, fill=tk.BOTH)

    def create_settings_tab(self):
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings")
        
        # Header
        header_frame = ttk.Frame(settings_frame)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Application Settings", 
            font=("Helvetica", 20, "bold")
        ).pack(side=tk.LEFT)
        
        # Settings in a card-like container
        settings_container = ttk.Frame(settings_frame, style="Card.TFrame")
        settings_container.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)
        
        # Add a header to the settings card
        ttk.Label(
            settings_container, 
            text="Email Server Settings", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        # Create a settings form
        form_frame = ttk.Frame(settings_container)
        form_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        # IMAP Server
        ttk.Label(form_frame, text="IMAP Server:", font=("Helvetica", 11)).grid(row=0, column=0, padx=10, pady=15, sticky=tk.W)
        self.imap_server_entry = ttkb.Entry(form_frame, width=40, bootstyle="primary")
        self.imap_server_entry.insert(0, "imap.gmail.com")
        self.imap_server_entry.grid(row=0, column=1, padx=10, pady=15, sticky=tk.W)
        
        # SMTP Server
        ttk.Label(form_frame, text="SMTP Server:", font=("Helvetica", 11)).grid(row=1, column=0, padx=10, pady=15, sticky=tk.W)
        smtp_server_entry = ttkb.Entry(form_frame, width=40, bootstyle="primary")
        smtp_server_entry.insert(0, "smtp.gmail.com")
        smtp_server_entry.grid(row=1, column=1, padx=10, pady=15, sticky=tk.W)
        
        # Port
        ttk.Label(form_frame, text="SMTP Port:", font=("Helvetica", 11)).grid(row=2, column=0, padx=10, pady=15, sticky=tk.W)
        port_entry = ttkb.Entry(form_frame, width=40, bootstyle="primary")
        port_entry.insert(0, "587")
        port_entry.grid(row=2, column=1, padx=10, pady=15, sticky=tk.W)
        
        # Analysis settings section
        ttk.Label(
            settings_container, 
            text="Analysis Settings", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        analysis_frame = ttk.Frame(settings_container)
        analysis_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        # Default analysis period
        ttk.Label(analysis_frame, text="Default Analysis Period (days):", font=("Helvetica", 11)).grid(row=0, column=0, padx=10, pady=15, sticky=tk.W)
        self.analysis_period_entry = ttkb.Entry(analysis_frame, width=40, bootstyle="primary")
        self.analysis_period_entry.insert(0, "30")
        self.analysis_period_entry.grid(row=0, column=1, padx=10, pady=15, sticky=tk.W)
        
        # Chart type
        ttk.Label(analysis_frame, text="Default Chart Type:", font=("Helvetica", 11)).grid(row=1, column=0, padx=10, pady=15, sticky=tk.W)
        chart_combo = ttkb.Combobox(
            analysis_frame, 
            values=["Bar Chart", "Pie Chart", "Line Chart"],
            bootstyle="primary",
            width=38
        )
        chart_combo.current(0)
        chart_combo.grid(row=1, column=1, padx=10, pady=15, sticky=tk.W)
        
        # Save button
        save_settings_button = ttkb.Button(
            settings_container,
            text="Save Settings",
            command=self.save_settings,
            bootstyle="success",
            width=20
        )
        save_settings_button.pack(pady=20)

    def create_auto_reply_tab(self):
        auto_reply_frame = ttk.Frame(self.notebook)
        self.notebook.add(auto_reply_frame, text="Auto-Reply")
        
        # Header
        header_frame = ttk.Frame(auto_reply_frame)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Auto-Reply Configuration", 
            font=("Helvetica", 20, "bold")
        ).pack(side=tk.LEFT)
        
        # Auto-reply settings in a card-like container
        reply_container = ttk.Frame(auto_reply_frame, style="Card.TFrame")
        reply_container.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)
        
        # Add a header to the auto-reply card
        ttk.Label(
            reply_container, 
            text="Auto-Reply Settings", 
            font=("Helvetica", 14, "bold"),
            style="CardTitle.TLabel"
        ).pack(anchor=tk.W, padx=20, pady=(20, 10))
        
        # Enable/disable auto-reply
        enable_frame = ttk.Frame(reply_container)
        enable_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(enable_frame, text="Enable Auto-Reply:", font=("Helvetica", 12, "bold")).pack(side=tk.LEFT, padx=5)
        auto_reply_switch = ttkb.Checkbutton(
            enable_frame, 
            variable=self.auto_reply_var, 
            bootstyle="success-round-toggle",
            text="Enabled" if self.auto_reply_var.get() else "Disabled"
        )
        auto_reply_switch.pack(side=tk.LEFT, padx=10)
        
        # Auto-reply message
        message_frame = ttk.Frame(reply_container)
        message_frame.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)
        
        ttk.Label(message_frame, text="Auto-Reply Message:", font=("Helvetica", 12)).pack(anchor=tk.W, pady=(10, 5))
        
        # Text editor with toolbar
        toolbar_frame = ttk.Frame(message_frame)
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # Simple formatting buttons
        bold_btn = ttkb.Button(
            toolbar_frame,
            text="Bold",
            bootstyle="secondary-outline",
            width=8,
            command=lambda: messagebox.showinfo("Info", "Bold formatting would go here")
        )
        bold_btn.pack(side=tk.LEFT, padx=2)
        
        italic_btn = ttkb.Button(
            toolbar_frame,
            text="Italic",
            bootstyle="secondary-outline",
            width=8,
            command=lambda: messagebox.showinfo("Info", "Italic formatting would go here")
        )
        italic_btn.pack(side=tk.LEFT, padx=2)
        
        # Text area with scrollbar
        text_frame = ttk.Frame(message_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        scroll_y = ttk.Scrollbar(text_frame)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.auto_reply_message = tk.Text(text_frame, height=10, yscrollcommand=scroll_y.set)
        self.auto_reply_message.insert(tk.END, self.organizer.auto_reply_settings['message'])
        self.auto_reply_message.pack(fill=tk.BOTH, expand=True)
        scroll_y.config(command=self.auto_reply_message.yview)
        
        # Template dropdown
        template_frame = ttk.Frame(message_frame)
        template_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(template_frame, text="Use Template:").pack(side=tk.LEFT, padx=5)
        template_combo = ttkb.Combobox(
            template_frame, 
            values=["Out of Office", "Thank You", "Will Reply Soon", "Custom"],
            bootstyle="primary",
            width=20
        )
        template_combo.current(3)  # Default to Custom
        template_combo.pack(side=tk.LEFT, padx=5)
        
        load_template_btn = ttkb.Button(
            template_frame,
            text="Load",
            bootstyle="info-outline",
            width=8,
            command=lambda: messagebox.showinfo("Info", "Template loading would go here")
        )
        load_template_btn.pack(side=tk.LEFT, padx=5)
        
        # Save button
        save_button = ttkb.Button(
            reply_container,
            text="Save Settings",
            command=self.save_auto_reply,
            bootstyle="success",
            width=20
        )
        save_button.pack(pady=20)

    def create_status_bar(self):
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(
            self.window, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W,
            padding=(10, 5)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def show_login_dialog(self):
        dialog = ttkb.Toplevel(self.window)
        dialog.title("Email Login")
        dialog.geometry("400x450")
        
        # Add a header
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Connect to Email Server", 
            font=("Helvetica", 16, "bold")
        ).pack()
        
        # Create a form in a card-like container
        form_container = ttk.Frame(dialog, style="Card.TFrame")
        form_container.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)
        
        # Email field
        ttk.Label(form_container, text="Email Address:", font=("Helvetica", 11)).pack(anchor=tk.W, padx=20, pady=(20, 5))
        email_entry = ttkb.Entry(form_container, width=40, bootstyle="primary")
        email_entry.pack(padx=20, pady=(0, 15), fill=tk.X)
        
        # Password field
        ttk.Label(form_container, text="Password:", font=("Helvetica", 11)).pack(anchor=tk.W, padx=20, pady=(0, 5))
        password_entry = ttkb.Entry(form_container, width=40, show="*", bootstyle="primary")
        password_entry.pack(padx=20, pady=(0, 15), fill=tk.X)
        
        # Server field
        ttk.Label(form_container, text="IMAP Server:", font=("Helvetica", 11)).pack(anchor=tk.W, padx=20, pady=(0, 5))
        server_entry = ttkb.Entry(form_container, width=40, bootstyle="primary")
        server_entry.insert(0, "imap.gmail.com")
        server_entry.pack(padx=20, pady=(0, 15), fill=tk.X)
        
        # Help text
        help_text = ttk.Label(
            form_container, 
            text="Note: For Gmail, you may need to enable 'Less secure app access'\nor use an App Password if 2FA is enabled.",
            justify=tk.LEFT,
            font=("Helvetica", 9),
            foreground="#888888"
        )
        help_text.pack(padx=20, pady=(0, 20))
        
        # Connect button
        connect_button = ttkb.Button(
            form_container,
            text="Connect",
            command=lambda: self.connect_to_email(email_entry.get(), password_entry.get(), dialog),
            bootstyle="success",
            width=20
        )
        connect_button.pack(pady=20)

    def connect_to_email(self, email_address, password, dialog):
        if self.organizer.connect(email_address, password):
            self.status_var.set("Connected to email server successfully!")
            self.connection_status.config(text=f"Connected: {email_address}")
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
            
            # Show a progress dialog
            progress = ttkb.Toplevel(self.window)
            progress.title("Analyzing Emails")
            progress.geometry("300x150")
            
            ttk.Label(progress, text="Analyzing your emails...", font=("Helvetica", 12)).pack(pady=(20, 10))
            
            progress_bar = ttkb.Progressbar(
                progress, 
                bootstyle="success-striped",
                mode="indeterminate",
                length=250
            )
            progress_bar.pack(pady=10, padx=20)
            progress_bar.start()
            
            # Update the UI to show we're working
            self.window.update()
            
            # Perform the analysis
            analytics = self.organizer.analyze_emails(days)
            
            # Close the progress dialog
            progress.destroy()
            
            self.status_var.set("Analysis complete.")
            
            # Open a new window to display analytics instead of showing in the dashboard
            AnalyticsWindow(self.window, analytics, self.is_dark_mode)


    def process_emails(self):
        if not self.organizer.imap_server:
            messagebox.showerror("Error", "Please connect to your email first.")
            return

        self.status_var.set("Processing emails...")
        
        # Show a progress dialog
        progress = ttkb.Toplevel(self.window)
        progress.title("Processing Emails")
        progress.geometry("300x150")
        
        ttk.Label(progress, text="Processing your emails...", font=("Helvetica", 12)).pack(pady=(20, 10))
        
        progress_bar = ttkb.Progressbar(
            progress, 
            bootstyle="success-striped",
            mode="indeterminate",
            length=250
        )
        progress_bar.pack(pady=10, padx=20)
        progress_bar.start()
        
        # Update the UI to show we're working
        self.window.update()
        
        # Process the emails
        result = self.organizer.process_emails()
        
        # Close the progress dialog
        progress.destroy()
        
        self.status_var.set(result)
        messagebox.showinfo("Processing Result", result)

    def show_add_rule_dialog(self):
        dialog = ttkb.Toplevel(self.window)
        dialog.title("Add Email Rule")
        dialog.geometry("500x550")
        
        # Add a header
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Label(
            header_frame, 
            text="Create New Email Rule", 
            font=("Helvetica", 16, "bold")
        ).pack()
        
        # Create a form in a card-like container
        form_container = ttk.Frame(dialog, style="Card.TFrame")
        form_container.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)
        
        # Rule name
        ttk.Label(form_container, text="Rule Name:", font=("Helvetica", 11, "bold")).pack(anchor=tk.W, padx=20, pady=(20, 5))
        rule_name_entry = ttkb.Entry(form_container, width=40, bootstyle="primary")
        rule_name_entry.pack(padx=20, pady=(0, 15), fill=tk.X)
        
        # Condition type
        ttk.Label(form_container, text="Condition Type:", font=("Helvetica", 11, "bold")).pack(anchor=tk.W, padx=20, pady=(0, 5))
        condition_type = tk.StringVar()
        condition_type_frame = ttk.Frame(form_container)
        condition_type_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        condition_types = [("From", "from"), ("Subject", "subject"), ("Body", "body")]
        for i, (text, value) in enumerate(condition_types):
            radio = ttkb.Radiobutton(
                condition_type_frame,
                text=text,
                variable=condition_type,
                value=value,
                bootstyle="primary"
            )
            radio.pack(side=tk.LEFT, padx=10)
            if i == 0:
                radio.invoke()  # Select the first option by default
        
        # Condition value
        ttk.Label(form_container, text="Condition Value:", font=("Helvetica", 11, "bold")).pack(anchor=tk.W, padx=20, pady=(0, 5))
        condition_value_entry = ttkb.Entry(form_container, width=40, bootstyle="primary")
        condition_value_entry.pack(padx=20, pady=(0, 15), fill=tk.X)
        
        # Help text for condition
        help_text = ttk.Label(
            form_container, 
            text="Example: For 'From' condition, enter an email address or domain.\nFor 'Subject' or 'Body', enter keywords to match.",
            justify=tk.LEFT,
            font=("Helvetica", 9),
            foreground="#888888"
        )
        help_text.pack(padx=20, pady=(0, 15), anchor=tk.W)
        
        # Target folder
        ttk.Label(form_container, text="Target Folder:", font=("Helvetica", 11, "bold")).pack(anchor=tk.W, padx=20, pady=(0, 5))
        target_folder_entry = ttkb.Entry(form_container, width=40, bootstyle="primary")
        target_folder_entry.pack(padx=20, pady=(0, 15), fill=tk.X)
        
        # Help text for folder
        folder_help = ttk.Label(
            form_container, 
            text="Enter the name of an existing folder in your email account.\nExample: 'Work', 'Personal', 'Newsletters', etc.",
            justify=tk.LEFT,
            font=("Helvetica", 9),
            foreground="#888888"
        )
        folder_help.pack(padx=20, pady=(0, 20), anchor=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(form_container)
        button_frame.pack(pady=20)
        
        cancel_button = ttkb.Button(
            button_frame,
            text="Cancel",
            command=dialog.destroy,
            bootstyle="secondary",
            width=15
        )
        cancel_button.pack(side=tk.LEFT, padx=10)
        
        add_button = ttkb.Button(
            button_frame,
            text="Add Rule",
            command=lambda: self.add_rule(
                rule_name_entry.get(),
                condition_type.get(),
                condition_value_entry.get(),
                target_folder_entry.get(),
                dialog
            ),
            bootstyle="success",
            width=15
        )
        add_button.pack(side=tk.LEFT, padx=10)

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

         
            messagebox.showerror("Error", "Please connect to your email first.")
            return

        if not query:
            messagebox.showinfo("Info", "Please enter a search term")
            return

        self.status_var.set("Searching emails...")
        
        # Show a progress dialog
        progress = ttkb.Toplevel(self.window)
        progress.title("Searching Emails")
        progress.geometry("300x150")
        
        ttk.Label(progress, text="Searching your emails...", font=("Helvetica", 12)).pack(pady=(20, 10))
        
        progress_bar = ttkb.Progressbar(
            progress, 
            bootstyle="info-striped",
            mode="indeterminate",
            length=250
        )
        progress_bar.pack(pady=10, padx=20)
        progress_bar.start()
        
        # Update the UI to show we're working
        self.window.update()
        
        # Perform the search
        results = self.organizer.search_emails(query)
        
        # Close the progress dialog
        progress.destroy()
        
        self.status_var.set("Search complete.")

        for row in self.search_results.get_children():
            self.search_results.delete(row)

        if isinstance(results, list):
            if results:
                for result in results:
                    self.search_results.insert('', 'end', values=(result['subject'], result['sender'], result['date']))
            else:
                messagebox.showinfo("Search Results", "No emails found matching your search criteria.")
        else:
            messagebox.showerror("Error", results)

    def toggle_dark_mode(self):
        current_theme = self.window.style.theme_use()
        if current_theme == "superhero":
            self.window.style.theme_use("cosmo")
            self.is_dark_mode = False
        else:
            self.window.style.theme_use("superhero")
            self.is_dark_mode = True
            
        # Update any open analytics
        if hasattr(self, 'analytics_frame') and len(self.analytics_frame.winfo_children()) > 0:
            # Re-display analytics with new theme
            if hasattr(self.organizer, 'last_analytics'):
                self.display_analytics(self.organizer.last_analytics)

    def save_settings(self):
        imap_server = self.imap_server_entry.get()
        analysis_period = self.analysis_period_entry.get()
        
        # Add visual feedback
        self.status_var.set("Settings saved successfully!")
        
        # Show a success message with a checkmark
        success = ttkb.Toplevel(self.window)
        success.title("Settings Saved")
        success.geometry("300x150")
        
        ttk.Label(success, text="✓", font=("Helvetica", 30, "bold"), foreground="green").pack(pady=(20, 0))
        ttk.Label(success, text="Settings saved successfully!", font=("Helvetica", 12)).pack(pady=10)
        
        # Auto-close after 2 seconds
        success.after(2000, success.destroy)

    def save_auto_reply(self):
        settings = {
            "enabled": self.auto_reply_var.get(),
            "message": self.auto_reply_message.get("1.0", tk.END).strip()
        }
        self.organizer.save_auto_reply_settings(settings)
        
        # Add visual feedback
        self.status_var.set("Auto-Reply settings saved successfully!")
        
        # Show a success message
        success = ttkb.Toplevel(self.window)
        success.title("Settings Saved")
        success.geometry("300x150")
        
        ttk.Label(success, text="✓", font=("Helvetica", 30, "bold"), foreground="green").pack(pady=(20, 0))
        ttk.Label(success, text="Auto-Reply settings saved!", font=("Helvetica", 12)).pack(pady=10)
        
        # Auto-close after 2 seconds
        success.after(2000, success.destroy)

    def show_about_dialog(self):
        about_dialog = ttkb.Toplevel(self.window)
        about_dialog.title("About Echo-Box")
        about_dialog.geometry("500x400")
        
        # Create a card-like container
        about_container = ttk.Frame(about_dialog, style="Card.TFrame")
        about_container.pack(fill=tk.BOTH, padx=20, pady=20, expand=True)
        
        # App logo/icon placeholder
        logo_frame = ttk.Frame(about_container, width=100, height=100)
        logo_frame.pack(pady=20)
        
        # App title
        ttk.Label(
            about_container, 
            text="Echo-Box", 
            font=("Helvetica", 18, "bold")
        ).pack(pady=(0, 10))
        
        # Version
        ttk.Label(
            about_container, 
            text="Version 1.0", 
            font=("Helvetica", 12)
        ).pack(pady=(0, 20))
        
        # Description
        description = """
        This application helps you process and analyze your emails.
        It provides features like email rule management, email search,
        and detailed email analytics.
        
        Features:
        • Email rule management
        • Auto-reply functionality
        • Advanced email analytics
        • Email search capabilities
        • Dark/Light mode
        """
        
        desc_label = ttk.Label(
            about_container, 
            text=description, 
            font=("Helvetica", 11),
            justify=tk.CENTER,
            wraplength=400
        )
        desc_label.pack(pady=10)
        
        # Credits
        ttk.Label(
            about_container, 
            text="Created by Logic Lords Team", 
            font=("Helvetica", 10, "italic")
        ).pack(pady=(20, 10))
        
        # Close button
        close_button = ttkb.Button(
            about_container,
            text="Close",
            command=about_dialog.destroy,
            bootstyle="secondary",
            width=15
        )
        close_button.pack(pady=10)

    def open_documentation(self):
        import webbrowser
        webbrowser.open("https://example.com/email-organizer-docs")
        
        # Show a notification
        self.status_var.set("Opening documentation in your web browser...")

    def run(self):
        # Apply custom styles
        style = ttkb.Style()
        
        # Card-like frames
        style.configure("Card.TFrame", background="#ffffff" if not self.is_dark_mode else "#2a2a2a", 
                       relief="solid", borderwidth=1)
        style.configure("InnerCard.TFrame", background="#f8f9fa" if not self.is_dark_mode else "#343a40", 
                       relief="solid", borderwidth=1)
        
        # Sidebar styling
        style.configure("Sidebar.TFrame", background="#343a40")
        style.configure("SidebarBtn.TFrame", background="#343a40")
        style.configure("SidebarTitle.TLabel", background="#343a40", foreground="white")
        style.configure("ConnectionStatus.TLabel", background="#343a40", foreground="#5cb85c")
        
        # Card titles
        style.configure("CardTitle.TLabel", font=("Helvetica", 14, "bold"))
        
        # Colored frames for indicators
        style.configure("Success.TFrame", background="#5cb85c")
        style.configure("Info.TFrame", background="#5bc0de")
        style.configure("Warning.TFrame", background="#f0ad4e")
        style.configure("Danger.TFrame", background="#d9534f")
        style.configure("Primary.TFrame", background="#0275d8")
        
        self.window.mainloop()


if __name__ == "__main__":
    app = EmailOrganizerGUI()
    app.run()
