<#
.Synopsis
   expand-kml2csv - extract the locations --the latitude/longitude pairs-- from a KML file and write them to a .csv file
.DESCRIPTION
   
.EXAMPLE
   expand-kml2csv -kmlfile gpslog_12_29_13.kml -outfile decembertracks.csv
#>

param(
    [parameter(Mandatory=$true)][string]$kmlfile,
    [parameter(Mandatory=$true)][string]$outfile   
)

$global:fileToExtract = $kmlfile
$global:fileToSave = $outfile


function exitWithError([string]$message, $value) {
    $message += " -- $value"
    #write-error $message
    write-host -foregroundcolor cyan $message
    exit 2
}

function getCurrentLocation() {
    $currentLocation = get-location
    $currentLocation = $currentLocation.toString()
    return $currentLocation
}


function expand-kml2csv
{

    $currentLocation = getCurrentLocation
    $inputFileFullPath = $currentLocation + "\" + $global:fileToExtract
    $outputFileFullPath = $currentLocation + "\" + $global:fileToSave

    Write-Host "Converting: " $inputFileFullPath
    Write-Host "Saving to: " $outputFileFullPath

    $inputfileIsFound = test-path $inputFileFullPath
    if ($inputfileIsFound -eq $false) {
        exitWithError "The KML file you specified is not present in the current folder. Run this script in the folder containing the file." ""
    }
    $outputfileIsFound = test-path $outputFileFullPath
    if ($outputfileIsFound -eq $true) {
        exitWithError "The output file you specified is already present in the current folder. We don't want to overwrite it so resolve this issue to continue." ""
    }

    # get the contents of the file
    [System.Xml.XmlDocument] $kmlDocument = New-Object System.Xml.XmlDocument;
    $xml = [xml] (get-content $inputFileFullPath)
    $rowCounter = 1
    foreach ($placemark in $xml.kml.Document.Folder.placemark) {
        $theRowAsCSV = [string]$rowCounter + "," + $placemark.name + "," + $placemark.description + "," + $placemark.point.coordinates 
        Add-Content -path $outputFileFullPath -value $theRowAsCSV       
        $rowCounter += 1
    }

    #write to the output csv file

    Write-Host "Conversion Completed Successfully"

}


expand-kml2csv
