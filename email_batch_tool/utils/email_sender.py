"""
Email sender module for batch sending emails one by one using Microsoft Outlook.
This module implements a safe, compliant 1-to-1 email delivery system.
"""

import time
import random
import logging
import json
import base64
import re
import os
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Tuple
import msal
import requests
from bs4 import BeautifulSoup

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("email_batch.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class OutlookEmailSender:
    """
    A class to send emails one by one using Microsoft Graph API.
    """

    def __init__(self, tenant_id: str, client_id: str, client_secret: str, 
                 shared_mailbox: str):
        """
        Initialize the Outlook email sender.
        
        Args:
            tenant_id: Azure AD tenant ID
            client_id: Application (client) ID
            client_secret: Client secret
            shared_mailbox: Shared mailbox email address
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.shared_mailbox = shared_mailbox
        self.access_token = None
        self.token_expires_at = None
        self.base_url = "https://graph.microsoft.com/v1.0"
        
    def authenticate(self) -> bool:
        """
        Authenticate with Microsoft Graph API using client credentials flow.
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}"
            )
            
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                # Set token expiration time (typically 1 hour for Azure AD tokens)
                # We'll set it to 55 minutes to be safe
                self.token_expires_at = datetime.now() + timedelta(minutes=55)
                logger.info("Authentication successful")
                logger.info(f"Token expires at: {self.token_expires_at.isoformat()}")
                return True
            else:
                logger.error(f"Authentication failed: {result.get('error_description')}")
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
            
    def is_token_expired(self) -> bool:
        """
        Check if the current access token is expired or about to expire.
        
        Returns:
            bool: True if token is expired or will expire in the next 5 minutes, False otherwise
        """
        if not self.access_token or not self.token_expires_at:
            return True
            
        # Consider token expired if it will expire in the next 5 minutes
        return datetime.now() >= (self.token_expires_at - timedelta(minutes=5))
            
    def send_email(self, to_address: str, subject: str, html_body: str, attachments: Optional[List[Dict]] = None, cc_addresses: Optional[List[str]] = None) -> bool:
        """
        Send a single email to a recipient with optional inline images and CC recipients.
        
        Args:
            to_address: Recipient email address
            subject: Email subject
            html_body: HTML email body
            attachments: List of attachment dictionaries with 'contentBytes', 'name', and 'contentType'
            cc_addresses: List of CC recipient email addresses
            
        Returns:
            bool: True if email sent successfully, False otherwise
        """
        # Check if we need to authenticate or refresh token
        if not self.access_token or self.is_token_expired():
            logger.info("Access token is missing or expired. Authenticating...")
            if not self.authenticate():
                logger.error("Failed to authenticate and obtain access token.")
                return False
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Prepare email data
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": html_body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_address
                        }
                    }
                ]
            },
            "saveToSentItems": True
        }
        
        # Add CC recipients if provided
        if cc_addresses:
            cc_recipients = [
                {
                    "emailAddress": {
                        "address": cc_address
                    }
                }
                for cc_address in cc_addresses
            ]
            email_data["message"]["ccRecipients"] = cc_recipients
        
        # Add attachments if provided
        if attachments:
            email_data["message"]["attachments"] = attachments
        
        # Send the email
        try:
            url = f"{self.base_url}/users/{self.shared_mailbox}/sendMail"
            response = requests.post(url, headers=headers, json=email_data)
            
            # Check if token is expired (401 error with specific message)
            if response.status_code == 401:
                error_response = response.json()
                if "error" in error_response and error_response["error"]["code"] == "InvalidAuthenticationToken":
                    logger.warning("Access token expired during request. Re-authenticating...")
                    # Re-authenticate to get a new token
                    if self.authenticate():
                        # Retry the request with new token
                        headers["Authorization"] = f"Bearer {self.access_token}"
                        response = requests.post(url, headers=headers, json=email_data)
                    else:
                        logger.error("Failed to refresh access token.")
                        return False
            
            if response.status_code == 202:
                logger.info(f"Email sent successfully to {to_address}" + (f" with CC to {', '.join(cc_addresses)}" if cc_addresses else ""))
                return True
            else:
                logger.error(f"Failed to send email to {to_address}. Status: {response.status_code}, Response: {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Error sending email to {to_address}: {str(e)}")
            return False
            
    def extract_inline_images(self, html_content: str) -> Tuple[str, List[Dict]]:
        """
        Extract inline images from HTML content and convert them to attachments.
        
        Args:
            html_content: HTML content with img tags
            
        Returns:
            Tuple of (cleaned_html, attachments_list)
        """
        soup = BeautifulSoup(html_content, 'html.parser')
        attachments = []
        
        # Find all img tags with src attributes
        for img_tag in soup.find_all('img', src=True):
            src = img_tag['src']
            
            # Handle data URLs (base64 encoded images)
            if src.startswith('data:'):
                # Extract MIME type and base64 data
                match = re.match(r'data:(image/[^;]+);base64,(.*)', src)
                if match:
                    mime_type, base64_data = match.groups()
                    
                    # Generate a unique CID
                    cid = f"image_{len(attachments)+1}@example.com"
                    
                    # Create attachment
                    attachment = {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": f"image_{len(attachments)+1}",
                        "contentType": mime_type,
                        "contentBytes": base64_data,
                        "isInline": True,
                        "contentId": cid
                    }
                    
                    attachments.append(attachment)
                    
                    # Replace src with CID reference
                    img_tag['src'] = f"cid:{cid}"
            
            # Handle local file paths (for development/testing)
            elif src.startswith('file://') or src.startswith('/'):
                # In production, you would need to implement actual file loading
                # For now, we'll just log and skip
                logger.warning(f"Local file path detected but not supported: {src}")
            
            # Handle relative image paths (e.g., images/filename.png)
            elif not src.startswith('http') and not src.startswith('cid:'):
                try:
                    # Get the directory of the template file
                    # Assuming the HTML content is from template/email.html
                    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                    images_dir = os.path.join(project_root, 'template', 'images')
                    image_path = os.path.join(images_dir, src)
                    
                    # Alternative path resolution for relative paths like "images/filename.png"
                    if not os.path.exists(image_path) and '/' in src:
                        # Try to resolve relative to template directory
                        template_dir = os.path.join(project_root, 'template')
                        image_path = os.path.join(template_dir, src)
                    
                    # Check if file exists
                    if os.path.exists(image_path):
                        # Determine MIME type from file extension
                        _, ext = os.path.splitext(src)
                        mime_type = 'image/png'  # default
                        if ext.lower() in ['.jpg', '.jpeg']:
                            mime_type = 'image/jpeg'
                        elif ext.lower() == '.gif':
                            mime_type = 'image/gif'
                        
                        # Read and encode the image
                        with open(image_path, 'rb') as image_file:
                            base64_data = base64.b64encode(image_file.read()).decode('utf-8')
                        
                        # Generate a unique CID
                        cid = f"image_{len(attachments)+1}@example.com"
                        
                        # Create attachment
                        attachment = {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": src.split('/')[-1],  # Just the filename
                            "contentType": mime_type,
                            "contentBytes": base64_data,
                            "isInline": True,
                            "contentId": cid
                        }
                        
                        attachments.append(attachment)
                        
                        # Replace src with CID reference
                        img_tag['src'] = f"cid:{cid}"
                        logger.info(f"Processed local image: {src} -> CID: {cid}")
                    else:
                        logger.warning(f"Image file not found: {image_path}")
                        logger.warning(f"Images directory: {images_dir}")
                        logger.warning(f"Current working directory: {os.getcwd()}")
                except Exception as e:
                    logger.error(f"Error processing image {src}: {str(e)}")
                    import traceback
                    logger.error(traceback.format_exc())
        
        return str(soup), attachments
            
    def sanitize_html(self, html_content: str) -> str:
        """
        Sanitize HTML content to remove potentially problematic elements.
        
        Args:
            html_content: Raw HTML content
            
        Returns:
            str: Sanitized HTML content
        """
        try:
            soup = BeautifulSoup(html_content, 'html5lib')
            
            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.decompose()
                
            # Return cleaned HTML
            return str(soup)
        except Exception as e:
            logger.warning(f"HTML sanitization failed: {str(e)}")
            return html_content


class BatchEmailProcessor:
    """
    Process batch email sending with delays and retry logic.
    """
    
    def __init__(self, email_sender: OutlookEmailSender):
        """
        Initialize the batch processor.
        
        Args:
            email_sender: Configured OutlookEmailSender instance
        """
        self.email_sender = email_sender
        
    def send_batch(self, recipients: List[str], subject: str, html_template: str,
                   min_delay: int = 30, max_delay: int = 120, 
                   max_retries: int = 3, cc_addresses: Optional[List[str]] = None) -> dict:
        """
        Send emails to all recipients one by one with delays.
        
        Args:
            recipients: List of recipient email addresses
            subject: Email subject
            html_template: HTML email template
            min_delay: Minimum delay between emails in seconds
            max_delay: Maximum delay between emails in seconds
            max_retries: Maximum number of retries for failed emails
            cc_addresses: List of CC recipient email addresses (optional)
            
        Returns:
            dict: Summary of sending results
        """
        # Sanitize HTML template
        clean_html = self.email_sender.sanitize_html(html_template)
        
        # Extract inline images and prepare attachments
        processed_html, attachments = self.email_sender.extract_inline_images(clean_html)
        
        results = {
            "total": len(recipients),
            "sent": 0,
            "failed": 0,
            "skipped": 0,
            "details": []
        }
        
        logger.info(f"Starting batch email send to {len(recipients)} recipients")
        logger.info(f"Delay range: {min_delay}-{max_delay} seconds")
        if cc_addresses:
            logger.info(f"CC recipients: {', '.join(cc_addresses)}")
        
        for i, recipient in enumerate(recipients):
            logger.info(f"Processing recipient {i+1}/{len(recipients)}: {recipient}")
            
            # Try to send email with retries
            success = False
            for attempt in range(max_retries + 1):
                if attempt > 0:
                    logger.info(f"Retry attempt {attempt}/{max_retries} for {recipient}")
                    
                if self.email_sender.send_email(recipient, subject, processed_html, attachments, cc_addresses):
                    success = True
                    break
                elif attempt < max_retries:
                    # Check if this was a token expiration issue
                    # If so, we should re-authenticate and retry immediately
                    # Otherwise, use exponential backoff
                    delay = min(60 * (2 ** attempt), 300)  # Max 5 minutes
                    logger.warning(f"Send failed, waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    
            # Record result
            timestamp = datetime.now().isoformat()
            if success:
                results["sent"] += 1
                results["details"].append({
                    "timestamp": timestamp,
                    "recipient": recipient,
                    "status": "success"
                })
            else:
                results["failed"] += 1
                results["details"].append({
                    "timestamp": timestamp,
                    "recipient": recipient,
                    "status": "failed"
                })
                logger.error(f"Failed to send email to {recipient} after {max_retries} retries")
                
            # Add delay before next email (except for the last one)
            if i < len(recipients) - 1:
                delay = random.randint(min_delay, max_delay)
                logger.info(f"Waiting {delay} seconds before next email...")
                time.sleep(delay)
                
        logger.info(f"Batch sending completed. Sent: {results['sent']}, Failed: {results['failed']}")
        return results