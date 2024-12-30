# Make sure Outlook is running and accessible via COM
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
 
# Path to your PST file
$pstFilePath = "<pat_to_PST>"
# Folder where to save the EML files
$outputFolder = "<path_to_output_directory>"
 
# Load the PST file
$Namespace.AddStore($pstFilePath)
 
# Access the root folder and all mail folders
$RootFolder = $Namespace.Folders.Item(1)
 
# Function to export emails from a folder to EML
function Export-EmailsToEML {
    param (
        [Parameter(Mandatory=$true)]
        [string]$folderPath
    )
    $Folder = $RootFolder.Folders.Item($folderPath)
    if ($null -eq $Folder) {
        Write-Host "Folder not found: $folderPath"
        return
    }
 
    $Items = $Folder.Items
    $Items.Sort("[ReceivedTime]", $false)  # Sort by ReceivedTime, descending order
 
    # Loop through all emails in the folder
    $i=1
    foreach ($Item in $Items) {
        if ($Item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
            # Generate file name using the subject and received time
            $Subject = $Item.Subject -replace '[\\\/:*?"<>|]', "_" # Clean subject for file name
            $ReceivedTime = $Item.ReceivedTime.ToString("yyyyMMdd_HHmmss")
            $EMLFileName = "${i}_${ReceivedTime}_${Subject}.eml"
            $i = $i+1
 
            # Full path to save the EML file
            $EMLFilePath = Join-Path $outputFolder $EMLFileName
 
            # Save the email as EML
            $Item.SaveAs($EMLFilePath, 3)  # 3 = olRFC822
            Write-Host "Exported: $EMLFilePath"
        }
    }
}
 
# Example: Extract emails from the "Inbox" folder (modify the folder name as needed)
Export-EmailsToEML -folderPath "<Outlook_folder_name>"
 
# If you want to export from other folders (e.g. Sent, Drafts, etc.), call the function again with different folder names.
# Export-EmailsToEML -folderPath "Sent Items"
