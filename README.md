# Outlook Add-in: Email ID Viewer

A classic Outlook add-in that displays the email ID of the selected email in an on-premises Exchange Server.

## Local Development Setup

### Prerequisites
- Node.js (version 14 or higher)
- **Classic Outlook (Office 2013 SP1 or later)** - Web Add-ins support required
- Administrator access to modify registry (for sideloading method)
- Exchange Server environment (on-premises or Exchange Online)

### Installation

1. **Install dependencies:**
   ```powershell
   npm install
   ```

2. **Start the development server:**
   
   For HTTP (easier setup but limited functionality):
   ```powershell
   $env:USE_HTTP="true"; npm start
   ```
   
   For HTTPS (recommended but requires certificates):
   ```powershell
   npm start
   ```

### HTTPS Setup (Recommended)

For full Office add-in functionality, you need HTTPS. Here are two options:

#### Option 1: Using mkcert (Recommended)
1. Install mkcert: https://github.com/FiloSottile/mkcert
2. Create certificates:
   ```powershell
   mkcert -install
   mkcert localhost
   ```
3. Rename the generated files to `localhost.pem` and `localhost-key.pem`

#### Option 2: Self-signed certificates with OpenSSL
```powershell
openssl req -x509 -newkey rsa:2048 -keyout localhost-key.pem -out localhost.pem -days 365 -nodes
```

### Testing the Add-in

1. **Start the server:**
   ```powershell
   npm start
   ```

2. **Install the add-in in Classic Outlook:**
   
   **Method 1: Automated Registry Installation (Recommended)**
   1. Run PowerShell as Administrator
   2. Navigate to the project folder
   3. Execute: `.\install-addin.ps1`
   4. Restart Outlook completely
   
   **Method 2: Manual Registry Installation**
   1. Open Registry Editor (regedit) as Administrator
   2. Navigate to: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`
      - For Office 2019/2021: Use `16.0`
      - For Office 2016: Use `16.0`
      - For Office 2013: Use `15.0`
   3. Create the `Developer` key if it doesn't exist
   4. Right-click on `Developer` → New → String Value
   5. Name it with any unique name (e.g., `EmailIDViewer`)
   6. Set the value to the full path of your `manifest.xml` file (e.g., `<path_to_repo>\classic_outlook_addin\manifest.xml`)
   7. Restart Outlook
   
   **Method 3: Shared Folder Deployment**
   1. Create a network share or local folder (e.g., `C:\OutlookAddins`)
   2. Copy your `manifest.xml` to this folder
   3. In Outlook, go to File → Options → Trust Center → Trust Center Settings
   4. Click "Catalog Manifests" on the left
   5. Add the folder path in "Catalog URL" field
   6. Check "Show in menu"
   7. Click OK and restart Outlook

3. **Use the add-in in Classic Outlook:**
   - Open an email in Outlook (reading pane or separate window)
   - The add-in should appear automatically in the reading pane or email window
   - If not visible, check the ribbon for an "Add-ins" tab or look for "Email ID Viewer"
   - The add-in will display in a task pane showing the email's unique ID
   - For classic Outlook, the add-in loads based on the rules defined in the manifest (when reading messages)

### Development Notes

- The server runs on `https://localhost:3000` (or `http://localhost:3000` with USE_HTTP=true)
- All files are served from the project root and `addin/` folder
- The manifest points to the local server for development
- Make sure to update URLs in `manifest.xml` for production deployment

### Troubleshooting

**General Issues:**
1. **Certificate errors:** Accept the self-signed certificate in your browser first by visiting `https://localhost:3000`
2. **Add-in not loading:** Check the browser console in Outlook for JavaScript errors
3. **HTTPS issues:** Try using HTTP mode with `$env:USE_HTTP="true"; npm start`

**Classic Outlook Specific:**
4. **Add-in not appearing:** 
   - Verify the registry entry is correct and Outlook was restarted
   - Check Windows Event Viewer for Office/Outlook errors
   - Ensure the manifest.xml file path is accessible
5. **Trust issues:** 
   - Go to File → Options → Trust Center → Trust Center Settings → Add-ins
   - Ensure "Require Application Add-ins to be signed by Trusted Publisher" is unchecked for development
6. **Office version compatibility:** 
   - Verify you're using the correct registry path for your Office version
   - Check that your Office version supports Web Add-ins (2013 SP1 or later)
7. **Manifest validation:** 
   - Use the Office Add-in Validator: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest

### File Structure
```
classic_outlook_addin/
├── manifest.xml          # Add-in manifest file
├── server.js            # Development server
├── package.json         # Node.js dependencies
├── start-dev.bat        # Windows batch file to start server
├── start-dev.ps1        # PowerShell script to start server
├── install-addin.ps1    # Automated registry installation script
├── uninstall-addin.ps1  # Automated registry uninstall script
├── addin/
│   ├── taskpane.html    # Main UI
│   ├── taskpane.js      # JavaScript logic
│   └── taskpane.css     # Styling
└── README.md           # This file
```
