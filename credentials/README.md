# Credentials Folder

This folder should contain your Google API credentials. These files are gitignored and should never be committed.

## Required Files

### 1. Service Account Key (for Vision, Sheets, Drive APIs)

Download from Google Cloud Console:
1. Go to IAM & Admin > Service Accounts
2. Create or select a service account
3. Keys > Add Key > Create new key > JSON
4. Save as `credentials/your-service-account.json`

Update `config/config.yaml`:
```yaml
google_cloud:
  credentials_file: "credentials/your-service-account.json"
```

### 2. OAuth2 Client Credentials (for Gmail API)

Download from Google Cloud Console:
1. Go to APIs & Services > Credentials
2. Create OAuth client ID (Desktop application)
3. Download JSON
4. Save as `credentials/gmail_oauth_credentials.json`

Update `config/config.yaml`:
```yaml
gmail:
  oauth_credentials_file: "credentials/gmail_oauth_credentials.json"
```

### 3. Gmail Token (auto-generated)

A `gmail_token.pickle` file will be created automatically when you first authenticate with Gmail. This contains your OAuth refresh token.
