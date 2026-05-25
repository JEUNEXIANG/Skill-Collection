# Outlook IMAP Setup for Himalaya

## Prerequisites

- Outlook/Office 365 account.
- **IMAP must be explicitly enabled** at Outlook.com > Settings > Mail > Sync email > "Let devices and apps use IMAP" (set to ON).
- An **app password** if you have two-factor authentication (2FA) enabled — generate one at https://account.microsoft.com/security > Advanced security options > Create a new app password.

> **Note:** Some accounts no longer support basic IMAP authentication. If `himalaya folder list` returns `"AUTHENTICATE failed."`, you must use OAuth2 instead (see OAuth2 section below).

## IMAP/SMTP Settings

| Service | Server | Port | Encryption |
|---------|--------|------|------------|
| IMAP | `outlook.office365.com` | 993 | SSL/TLS |
| SMTP | `smtp.office365.com` | 587 | STARTTLS |

## Creating an App Password (if 2FA enabled)

1. Go to [Microsoft Security](https://account.microsoft.com/security).
2. Under "Advanced security options", select "Create a new app password".
3. Copy the generated 16‑character password.
4. Use this password in your himalaya configuration (instead of your regular password).

## Configuring Himalaya

### Option 1: Interactive Wizard

Run:
```bash
himalaya account configure
```
Follow the prompts:
- Account name: `outlook`
- Email: your Outlook email address
- Display name: your name
- Backend: `imap`
- IMAP host: `outlook.office365.com`
- IMAP port: `993`
- Encryption: `tls`
- Login: your email again
- Authentication: `password`
- Password: your password (or app password)
- SMTP host: `smtp.office365.com`
- SMTP port: `587`
- SMTP encryption: `start-tls`
- SMTP login: your email
- SMTP authentication: `password`
- SMTP password: same as above

The wizard will store the password in the system keyring.

### Option 2: Manual Config File

Edit `~/.config/himalaya/config.toml`:

```toml
[accounts.outlook]
email = "your-email@outlook.com"
display-name = "Your Name"
default = true

backend.type = "imap"
backend.host = "outlook.office365.com"
backend.port = 993
backend.encryption.type = "tls"
backend.login = "your-email@outlook.com"
backend.auth.type = "password"
backend.auth.cmd = "security find-generic-password -a your-email@outlook.com -s himalaya-imap -w"

message.send.backend.type = "smtp"
message.send.backend.host = "smtp.office365.com"
message.send.backend.port = 587
message.send.backend.encryption.type = "start-tls"
message.send.backend.login = "your-email@outlook.com"
message.send.backend.auth.type = "password"
message.send.backend.auth.cmd = "security find-generic-password -a your-email@outlook.com -s himalaya-imap -w"

[accounts.outlook.folder.alias]
inbox = "INBOX"
sent = "Sent Items"
drafts = "Drafts"
trash = "Deleted Items"
```

### Option 3: OAuth2 (for accounts that reject basic auth)

If `himalaya folder list` returns `"AUTHENTICATE failed."`, Microsoft requires OAuth2.

1. **Register an app** in Azure AD:
   - Go to https://portal.azure.com > Azure Active Directory > App registrations > New registration
   - Name: `himalaya`
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: `http://localhost:7892` (type: Web)
   - Register and copy the **Application (client) ID**

2. **Create a client secret**: Certificates & secrets > New client secret > copy the value.

3. **Add API permissions**: API permissions > Add permission > Microsoft Graph > Delegated:
   - `IMAP.AccessAsUser.All`, `Mail.ReadWrite`, `Mail.Send`, `offline_access`
   - Click "Grant admin consent" (or user will consent at first login)

4. **Configure config.toml**:
```toml
[accounts.outlook]
email = "your-email@outlook.com"
display-name = "Your Name"
default = true

backend.type = "imap"
backend.host = "outlook.office365.com"
backend.port = 993
backend.encryption.type = "tls"
backend.login = "your-email@outlook.com"
backend.auth.type = "oauth2"
backend.auth.client-id = "YOUR_CLIENT_ID"
backend.auth.client-secret.cmd = "security find-generic-password -a oauth -s himalaya-client-secret -w"
backend.auth.tenant = "common"
backend.auth.auth-url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
backend.auth.token-url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

message.send.backend.type = "smtp"
message.send.backend.host = "smtp.office365.com"
message.send.backend.port = 587
message.send.backend.encryption.type = "start-tls"
message.send.backend.login = "your-email@outlook.com"
message.send.backend.auth.type = "oauth2"
backend.auth.client-id = "YOUR_CLIENT_ID"
backend.auth.client-secret.cmd = "security find-generic-password -a oauth -s himalaya-client-secret -w"
backend.auth.tenant = "common"
backend.auth.auth-url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
backend.auth.token-url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
```

5. **Store client secret**:
```bash
security add-generic-password -a "oauth" -s "himalaya-client-secret" -w "YOUR_CLIENT_SECRET"
```

6. On first run, himalaya will open a browser for you to sign in and consent.

### Storing Password in macOS Keychain

If you use the `auth.cmd` method, store the password first:

```bash
security add-generic-password -a "your-email@outlook.com" -s "himalaya-imap" -w "YOUR_PASSWORD"
```

## Testing

List folders to verify connectivity:

```bash
himalaya folder list
```

Expected output includes `INBOX`, `Sent Items`, `Drafts`, `Deleted Items`, etc.

## Troubleshooting

- **"Invalid credentials"**: Double‑check your password. If 2FA is enabled, use an app password.
- **"Connection refused"**: Ensure IMAP is enabled in your Outlook account settings.
- **"Cannot get imap password from global keyring"**: The password wasn't stored in the keyring. Run the wizard again or manually add the password as shown above.
- **"Folder not found"**: Adjust folder aliases to match Outlook's folder names (case‑sensitive).