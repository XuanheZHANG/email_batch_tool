"""主程序入口"""

import argparse
import json
import sys
from typing import List
from email_batch_tool.utils.email_sender import OutlookEmailSender, BatchEmailProcessor


def load_recipients(file_path: str) -> List[str]:
    """
    Load recipient email addresses from a file (one per line) or JSON array.
    
    Args:
        file_path: Path to file containing email addresses
        
    Returns:
        List of email addresses
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            
        # Try to parse as JSON array first
        try:
            recipients = json.loads(content)
            if isinstance(recipients, list):
                return [str(email).strip() for email in recipients if str(email).strip()]
        except json.JSONDecodeError:
            # If not JSON, treat as plain text with one email per line
            return [line.strip() for line in content.split('\n') if line.strip()]
            
    except FileNotFoundError:
        print(f"Error: Recipients file '{file_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error loading recipients: {str(e)}")
        sys.exit(1)


def load_html_template(file_path: str) -> str:
    """
    Load HTML email template from file.
    
    Args:
        file_path: Path to HTML template file
        
    Returns:
        HTML content as string
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(f"Error: HTML template file '{file_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error loading HTML template: {str(e)}")
        sys.exit(1)


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="Batch email sender that sends emails one by one using Microsoft Outlook",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s -r recipients.txt -t template.html -s "Subject Line" \\
    --tenant-id YOUR_TENANT_ID \\
    --client-id YOUR_CLIENT_ID \\
    --client-secret YOUR_CLIENT_SECRET \\
    --shared-mailbox shared@company.com

  %(prog)s -r recipients.json -t template.html -s "Subject Line" \\
    --config config.json
        """
    )
    
    # Required arguments
    parser.add_argument("-r", "--recipients", required=True,
                        help="Path to file containing recipient email addresses (one per line or JSON array)")
    parser.add_argument("-t", "--template", required=True,
                        help="Path to HTML email template file")
    parser.add_argument("-s", "--subject", required=True,
                        help="Email subject line")
    
    # Authentication arguments (either direct or via config file)
    auth_group = parser.add_mutually_exclusive_group(required=True)
    auth_group.add_argument("--config",
                           help="Path to JSON config file with authentication details")
    auth_group.add_argument("--tenant-id",
                           help="Azure AD tenant ID")
    
    # Additional authentication arguments when not using config
    parser.add_argument("--client-id", 
                       help="Application (client) ID")
    parser.add_argument("--client-secret",
                       help="Client secret")
    parser.add_argument("--shared-mailbox",
                       help="Shared mailbox email address")
    
    # Optional arguments
    parser.add_argument("--min-delay", type=int, default=30,
                       help="Minimum delay between emails in seconds (default: 30)")
    parser.add_argument("--max-delay", type=int, default=120,
                       help="Maximum delay between emails in seconds (default: 120)")
    parser.add_argument("--max-retries", type=int, default=3,
                       help="Maximum number of retries for failed emails (default: 3)")
    parser.add_argument("--cc", nargs="+",
                       help="Email addresses to CC (can specify multiple)")
    parser.add_argument("--output", "-o",
                       help="Path to save results JSON file")
    parser.add_argument("--dry-run", action="store_true",
                       help="Perform a dry run without sending emails")
    parser.add_argument("--version", action="version", version="1.0.0")
    
    args = parser.parse_args()
    
    # Load configuration
    if args.config:
        try:
            with open(args.config, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            print(f"Error loading config file: {str(e)}")
            sys.exit(1)
    else:
        config = {
            "tenant_id": args.tenant_id,
            "client_id": args.client_id,
            "client_secret": args.client_secret,
            "shared_mailbox": args.shared_mailbox
        }
    
    # Validate required config fields
    required_fields = ["tenant_id", "client_id", "client_secret", "shared_mailbox"]
    for field in required_fields:
        if not config.get(field):
            print(f"Error: Missing required configuration field '{field}'")
            sys.exit(1)
    
    # Load recipients and template
    recipients = load_recipients(args.recipients)
    html_template = load_html_template(args.template)
    
    print(f"Loaded {len(recipients)} recipients")
    print(f"Using subject: {args.subject}")
    print(f"Using shared mailbox: {config['shared_mailbox']}")
    
    if args.dry_run:
        print("*** DRY RUN MODE - No emails will be sent ***")
        print("Would send emails with the following parameters:")
        print(f"  Min delay: {args.min_delay} seconds")
        print(f"  Max delay: {args.max_delay} seconds")
        print(f"  Max retries: {args.max_retries}")
        return
    
    # Initialize email sender
    email_sender = OutlookEmailSender(
        tenant_id=config["tenant_id"],
        client_id=config["client_id"],
        client_secret=config["client_secret"],
        shared_mailbox=config["shared_mailbox"]
    )
    
    # Authenticate
    print("Authenticating with Microsoft Graph API...")
    if not email_sender.authenticate():
        print("Authentication failed. Exiting.")
        sys.exit(1)
    
    # Initialize batch processor
    batch_processor = BatchEmailProcessor(email_sender)
    
    # Send emails
    print("Starting batch email sending...")
    results = batch_processor.send_batch(
        recipients=recipients,
        subject=args.subject,
        html_template=html_template,
        min_delay=args.min_delay,
        max_delay=args.max_delay,
        max_retries=args.max_retries,
        cc_addresses=args.cc
    )
    
    # Output results
    print("\n" + "="*50)
    print("BATCH SENDING RESULTS")
    print("="*50)
    print(f"Total recipients: {results['total']}")
    print(f"Successfully sent: {results['sent']}")
    print(f"Failed: {results['failed']}")
    print(f"Skipped: {results['skipped']}")
    
    # Save results if output file specified
    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f"\nDetailed results saved to: {args.output}")
        except Exception as e:
            print(f"\nWarning: Failed to save results to {args.output}: {str(e)}")


if __name__ == "__main__":
    main()
