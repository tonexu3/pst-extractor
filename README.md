# pst-extractor
This script allows you to extract each file from PST as a separated files, normalizing the names of the files and using is subject as file names.

First of all, import your PST to Outlook and keep it open.

For using this script you have to substitute three items in the script:

#### Path to your PST file
    $pstFilePath: Absolute path to PST file. 
#### Folder where to save the EML files
    $outputFolder: Output folder. Must exist prior the script execution.
#### Outlook folder
```
Line 52:  Export-EmailsToEML -folderPath "<Outlook_Foder>"
```
This references to the folder where the emails are stored inside the PST. This usually is "Inbox" or "Sent Items". If you want to export from other folders (e.g. Sent, Drafts, etc.), call the function again with different folder names.
