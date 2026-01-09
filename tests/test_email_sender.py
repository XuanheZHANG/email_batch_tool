"""
Unit tests for the email sender module
"""

import unittest
from email_batch_tool.utils.email_sender import OutlookEmailSender, BatchEmailProcessor


class TestEmailSender(unittest.TestCase):
    
    def test_imports(self):
        """Test that classes can be imported correctly"""
        # This test just verifies that the classes can be imported
        # without syntax errors or missing dependencies
        self.assertTrue(True)  # Placeholder - actual implementation would require mocking
        
    def test_sanitize_html(self):
        """Test HTML sanitization"""
        sender = OutlookEmailSender("tenant", "client", "secret", "mailbox")
        
        # Test basic HTML
        html = "<p>Hello World</p>"
        sanitized = sender.sanitize_html(html)
        self.assertIn("<p>", sanitized)
        
        # Test HTML with script tag
        html_with_script = "<p>Hello</p><script>alert('test');</script>"
        sanitized = sender.sanitize_html(html_with_script)
        self.assertNotIn("<script>", sanitized)


if __name__ == "__main__":
    unittest.main()