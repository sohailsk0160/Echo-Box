# Mini-Project-Sem4
Auto-Mail-Organizer (Echo Box !!!)<br>
Group Members :<br>
Mohammad Sohail <br>
Sahil Bhandare  <br>
Rohan <br>
Chinmay

<br><br>
📨 Echo Box – Auto Mail Organiser

Overview:
Echo Box is an intelligent email management tool designed to automatically organize incoming emails into predefined folders based on their content. The aim is to reduce inbox clutter and enhance user productivity by categorizing mails like OTPs, promotional offers, updates, or personal messages into dedicated folders.


---

🔧 How It Works:

1. Email Integration:
The system connects to a user's email account using protocols like IMAP or POP3 to fetch incoming emails.


2. Content Analysis:
Each email is parsed and analyzed using keyword detection, subject line parsing, and content filtering to determine its category.


3. Classification Algorithm:
The system uses conditional logic or basic machine learning (optional) to match the email content with categories like:

OTP/Verification → Moved to "OTP"

Promotions/Offers → Moved to "Promotions"

Updates/Newsletters → Moved to "Updates"

Personal/Other → Moved to "General" or a custom folder



4. Automatic Sorting:
Once categorized, the email is moved to its respective folder, keeping the primary inbox clean and organized.


5. Optional Notifications:
Users can receive alerts for high-priority folders like OTP or Personal mails.




---

💡 Technologies Used (customizable):

Language: Python 

Libraries/Tools: IMAPClient, email-parser, Pandas, Regex,TKinter

Database: local file system 

Platform: Desktop script or lightweight GUI



---

🌟 Key Benefits:

Reduces inbox clutter

Saves user time by minimizing manual mail sorting

Helps users stay focused on priority messages
