import unittest
from unittest.mock import MagicMock
from email.message import EmailMessage
from email import policy
from email.parser import BytesParser

from Test22 import EmailOrganizer


class TestAnalyzeEmailsDateHandling(unittest.TestCase):
    def test_missing_date_header_does_not_crash(self):
        organizer = EmailOrganizer()
        organizer.imap_server = MagicMock()
        organizer.imap_server.search.return_value = (
            "OK",
            [b"1"]
        )

        message = EmailMessage()
        message["From"] = "sender@example.com"
        message["Subject"] = "Test message"
        raw_message = message.as_bytes()

        organizer.imap_server.fetch.return_value = (
            "OK",
            [(b"1 (RFC822)", raw_message)]
        )

        result = organizer.analyze_emails()

        self.assertIsInstance(result, dict)
        self.assertEqual(result["total_emails"], 1)


if __name__ == "__main__":
    unittest.main()
