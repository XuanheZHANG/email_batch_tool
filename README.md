# Email Batch Tool

A safe, compliant batch email sender that sends emails one by one using Microsoft Outlook from a shared mailbox.

## Features

- 📧 **One-to-One Email Sending**: Sends emails individually to each recipient for privacy
- 🔐 **Secure Authentication**: Uses Microsoft Graph API with OAuth 2.0 client credentials flow
- 🖼️ **Inline Image Support**: Automatically converts local images to inline attachments
- ⏱️ **Rate Limiting Compliance**: Built-in delays between emails to respect Microsoft's sending limits
- 🔄 **Automatic Token Refresh**: Handles authentication token expiration automatically
- 📋 **Batch Processing**: Processes recipient lists from text files or JSON arrays
- 📝 **Template Support**: Uses HTML email templates with variable substitution
- 📤 **CC Support**: Optionally send copies to additional recipients
- 📊 **Detailed Logging**: Comprehensive logging of all email sending activities
- 🧪 **Dry Run Mode**: Test your configuration without actually sending emails

## Prerequisites

1. Microsoft 365 account with access to a shared mailbox
2. Azure AD application registration with appropriate permissions
3. Python 3.6 or higher

## Setup

### 1. Azure AD Application Registration

1. Go to Azure Portal > Azure Active Directory > App registrations
2. Create a new application registration
3. Add the following API permissions:
   - `Mail.Send` (Delegated or Application type)
   - `Mail.ReadWrite` (if you want to save to sent items)
4. Grant admin consent for these permissions
5. Create a client secret and note down:
   - Tenant ID
   - Application (client) ID
   - Client secret

### 2. Install Dependencies

```bash
# Create and activate virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install the package
pip install -e .
```

## Usage

### Basic Usage

```bash
email_batch_tool \
  --recipients recipients.txt \
  --template template/email.html \
  --subject "Coinvest with Us: Positron Series B x Primitiva Global" \
  --config config.json
```

### Using Configuration File with CC

Create a `config.json` file:

```json
{
  "tenant_id": "YOUR_AZURE_AD_TENANT_ID",
  "client_id": "YOUR_APPLICATION_CLIENT_ID",
  "client_secret": "YOUR_CLIENT_SECRET",
  "shared_mailbox": "shared-mailbox@your-company.com"
}
```

Then run:

```bash
email_batch_tool \
  --recipients recipients.txt \
  --template template/email.html \
  --subject "Your Subject Line" \
  --config config.json \
  --cc manager@company.com secretary@company.com
```

### Recipients File Formats

You can provide recipients in two formats:

1. Plain text (one email per line):
```
john.doe@example.com
jane.smith@example.com
bob.johnson@example.com
```

2. JSON array:
```json
[
  "john.doe@example.com",
  "jane.smith@example.com",
  "bob.johnson@example.com"
]
```

### Additional Options

```bash
# Dry run (no emails sent)
email_batch_tool --dry-run ...

# Custom delay range
email_batch_tool --min-delay 60 --max-delay 180 ...

# Custom retry count
email_batch_tool --max-retries 5 ...

# CC recipients
email_batch_tool --cc cc1@example.com cc2@example.com ...

# Save results to file
email_batch_tool --output results.json ...
```

## Spam Risk Mitigation

This tool minimizes spam risk by:

1. Sending emails one-by-one (never in bulk or BCC)
2. Introducing random delays between sends
3. Using the same subject and HTML body for all recipients
4. Not modifying tracking headers or adding marketing footers
5. Not including bulk-sending headers
6. Ensuring the sender appears as a normal Outlook email
7. Respecting Outlook rate limits

## Logging

All email send attempts are logged with:
- Timestamp
- Recipient email
- Success or failure status

Logs are written to both console and `email_batch.log` file.

## Development

### Install Development Dependencies

```bash
pip install -e .[dev]
```

### Run Tests

```bash
pytest
```

## Outlook-Specific Considerations

1. Rate Limits: The tool respects Outlook's rate limits by introducing delays
2. Error Handling: Transient errors are retried with exponential backoff
3. Compliance: The tool complies with Outlook policies for sending emails
4. Authentication: Uses Microsoft Graph API with proper authentication flow

## Troubleshooting

### Authentication Token Expiration

If you encounter an error like:
```
Status: 401, Response: {"error":{"code":"InvalidAuthenticationToken","message":"Lifetime validation failed, the token is expired."}}
```

This means your authentication token has expired. The tool now includes enhanced token management with the following features:

1. **Automatic Token Refresh**: The tool automatically detects expired tokens and obtains new ones before sending emails.
2. **Proactive Token Management**: Tokens are refreshed before they expire (55 minutes lifetime with refresh at 50 minutes).
3. **Retry Mechanism**: If a token expires during sending, the tool will automatically re-authenticate and retry the operation.

If you continue to experience authentication issues:

1. Ensure your `config.json` contains valid credentials:
   ```json
   {
     "tenant_id": "YOUR_AZURE_AD_TENANT_ID",
     "client_id": "YOUR_APPLICATION_CLIENT_ID",
     "client_secret": "YOUR_CLIENT_SECRET",
     "shared_mailbox": "shared-mailbox@your-company.com"
   }
   ```

2. Verify your Azure application has the required permissions:
   - Mail.Send (Application permission)
   - User.Read (Delegated permission)

3. Check that your client secret hasn't expired in the Azure portal.

### Other Common Issues

- **Image files not found**: Ensure all image references in your HTML template are correctly placed in the `template/images/` directory.
- **Rate limiting**: The tool includes built-in delays between emails to comply with Microsoft's sending limits.
- **HTML formatting**: Complex HTML may be sanitized for security reasons. Test your template before sending.
