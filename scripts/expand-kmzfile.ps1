<#
.Synopsis
   expand-kmzfile - Extract the KML file in a KMZ Archive. Put the extracted file in a sub-folder with the same name as the KMZ file.
.DESCRIPTION
   Extract the files in a KMZ Archive. Uses the name of the KMZ file to create a folder and extracts the contents into that folder. If the kml file in the archive
   has the default name of doc.kml the file is renamed to match the name of the KMZ file.
.EXAMPLE
   expand-kmz -kmzfile gpslog_12_29_13.kmz
   Creates a folder named gpslog_12_29_13 and places the contents of the KMZ archive in that folder. Renames the KML file to the KMZ name.
#>
param([parameter(Mandatory=$true)][string]$kmzfile)

$global:fileToExtract = $kmzfile
#$global:fileToExtract = "testfilename.kmz"

function exitWithError([string]$message, $value) {
    $message += " -- $value"
    write-error $message
    write-host -foregroundcolor cyan $message
    exit 2
}

function expand-kmzfile
{

    $currentLocation = get-location
    $currentLocation = $currentLocation.toString()
    $currentLocationWithFile = $currentLocation + "\" + $global:fileToExtract
    Write-Host "Expanding: " $currentLocationWithFile

    $fileIsFound = test-path $currentLocationWithFile
    if ($fileIsFound -eq $false) {
        exitWithError "The file you specified is not present in the current folder. Run this script in the folder containing the file." ""
    }

    #pull of the file extension - we want to test for a folder
    $lastDotPosition = $currentLocationWithFile.LastIndexOf(".")
    if ($lastDotPosition -eq -1) {
        exitWithError "Provided file name does not appear to be valid. No extension found." $kmzFile
    }
    $qualifiedFolderPath = $currentLocationWithFile.Substring(0,$lastDotPosition)

    $canCreateFolder = test-path $qualifiedFolderPath 
    if ($canCreateFolder -eq $true) {
        exitWithError "Can't create the folder needed to store the extracted files. It already exists." $qualifiedFolderPath
    }

    new-item -Itemtype directory -Path $qualifiedFolderPath
    copy $kmzfile $qualifiedFolderPath
    $fileInFolder = $qualifiedFolderPath + "\" + $kmzFile

    # this runaround is cause the win32 shell won't treat a KMZ file as a zip even though that's what it is
    $baseFileName = $kmzFile.Split(".")
    $newFileName = $baseFileName[0] + ".zip"
    $newFullPathWithNewName = $qualifiedFolderPath + "\" + $newFileName
    ren $fileInFolder $newFullPathWithNewName

    $WindowsShell = new-object -com shell.application
    $archive = $WindowsShell.NameSpace($newFullPathWithNewName)

    foreach($item in $archive.items())
    {
        $WindowsShell.Namespace($qualifiedFolderPath).copyhere($item)
    }

    # Final step - if you create a KMZ archive with google earth it places the xml in doc.kml
    # We rename that to use the name of KMZ file
    $extractedDocFilePath = $qualifiedFolderPath + "\\" + "doc.kml"
    $docFileFound = test-path $extractedDocFilePath
    if ($docFileFound -eq $true) {
        $finallyTheFileNameWeWant = $qualifiedFolderPath + "\\" + $baseFileName[0] + ".kml" 
        ren $extractedDocFilePath $finallyTheFileNameWeWant  
    }
    write-host -BackgroundColor yellow -ForegroundColor DarkBlue "The extraction completed successfully"
}

expand-kmzfile