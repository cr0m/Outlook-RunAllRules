# RunAllRules

This VBA script automates the execution of all Outlook rules defined in your Outlook application. It loops through each rule in the rules collection and executes them, showing progress and handling any errors or user cancellations appropriately.

## Overview

The `RunAllRules` script performs the following tasks:
1. Initializes the Outlook application object.
2. Retrieves the default Inbox folder.
3. Gets the rules collection from the default store.
4. Loops through each rule in the collection and executes it.
5. Handles errors and user cancellations gracefully.
6. Add a button to the Outlook Ribbon to run all rules.
7. To speed up the process, hit Cancel after a few seconds to move to the next rule. The script processes top-down, so there's no need to wait for full completion.

    
## Script

```vba
Sub RunAllRules()
    ' Declare variables for Outlook objects and counters
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olInbox As Outlook.Folder
    Dim olRules As Outlook.Rules
    Dim olRule As Outlook.Rule
    Dim i As Integer

    ' Initialize Outlook application object
    Set olApp = Outlook.Application
    ' Get MAPI namespace
    Set olNamespace = olApp.GetNamespace("MAPI")
    ' Get the default Inbox folder
    Set olInbox = olNamespace.GetDefaultFolder(olFolderInbox)
    ' Get the rules collection from the default store
    Set olRules = olApp.Session.DefaultStore.GetRules()

    ' Loop through each rule in the rules collection and execute it
    For i = 1 To olRules.Count
        ' Set the current rule
        Set olRule = olRules.Item(i)
        ' Ignore errors for the following code
        On Error Resume Next
        ' Clear any existing errors
        Err.Clear
        ' Execute the current rule and show progress
        olRule.Execute ShowProgress:=True
        ' Check if the user canceled the operation
        If Err.Number = -2147221233 Then
            ' Display a message box if the rule execution was canceled by the user
            MsgBox "Execution of rule '" & olRule.Name & "' was canceled by the user.", vbExclamation
            ' Clear the error
            Err.Clear
        End If
        ' Turn error handling back to default
        On Error GoTo 0
    Next i

    ' Display a message box indicating all rules have been executed or skipped
    MsgBox "All rules have been executed or skipped.", vbInformation
End Sub
```

#### Optional: Create and Add a Signature for the Script

1. Generate a self-signed certificate:

    ```powershell
    $cert = New-SelfSignedCertificate -Type Custom -Subject "CN=RunAllRules" -KeySpec Signature -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeyLength 2048 -HashAlgorithm sha256 -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter (Get-Date).AddYears(5)
    ```

2. Convert the password to a secure string:

    ```powershell
    $pwd = ConvertTo-SecureString -String "NEWPASSWORD" -Force -AsPlainText
    ```

3. Export the certificate to a PFX file:

    ```powershell
    Export-PfxCertificate -Cert $cert -FilePath "C:\Users\USER\Documents\RunAllRules.pfx" -Password $pwd
    ```

Replace `"NEWPASSWORD"` with your desired password and update the file path as needed.

4. Import the certificate into the Personal store:

    ```powershell
    Import-PfxCertificate -FilePath "C:\Users\USER\Documents\RunAllRules.pfx" -CertStoreLocation "Cert:\CurrentUser\My" -Password $pwd
    ```

5. Add the certificate to the Trusted Root Certification Authorities (TRCA) store:

    ```powershell
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -eq "CN=RunAllRules" }
    Import-Certificate -FilePath $cert.PSPath -CertStoreLocation "Cert:\CurrentUser\Root"
    ```

6. Open Outlook and press `Alt + F11` to open the VBA editor.

7. In the VBA editor, go to `Tools` > `Digital Signature...`.

8. In the `Digital Signature` dialog box, click `Choose...`.

9. Select the certificate you created (`RunAllRules`) from the list and click `OK`.

10. Click `OK` again to close the `Digital Signature` dialog box.

11. Save your VBA project by going to `File` > `Save`.

