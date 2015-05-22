<#
.Synopsis
   extract-imagedata - given a folder containing JPG images with exif data -- extract the latitude/longitude pairs and write them to a .csv file
   based on example code from the Powershell Image module (which is required) 

.DESCRIPTION
   
.EXAMPLE
   extract-imagedata.ps1 -outfile "photopoints.csv" -folder "c:\images\"
   given a folder containing images -- extract the latitude/longitude pairs from each image and write them to a .csv file
#>

## --- ##
# Requires the PowerShell Image module. Available here:
# https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Image-module-caa4405a


param(
    [parameter][string]$outfile,
    [parameter][string]$folder
)

import-module image

#$global:folderpath = "C:\forest-data-images\"
$global:folderpath = $folder
#$global:fileToSave = "photolocations.csv"
$global:fileToSave = $outfile
set-location $global:folderpath


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


# return an array containing a latitude and longitude in the normalized decimal degree form
<#
    the exif object contains var that indicate the position of the data part in the string
    $ExifIDGPSLatref and $ExifIDGPSLattitude (yes spelled incorrectly)
    Not using these for now, since at this time the script only has to deal with data from
        the canon 260 -- and I know what that format is...

    NOTE - currently on handling "north" and "west" coordinates
#>
function normalizeCoordinateStringCanon260($coordinateString) {
    # parsing this string is fragile and ugly - but I don't know a better way...
    # example: 43°2'43.452"N  73°46'51.636"W, 90.1M above Sea Level
    # second param to substring is a length, not a position

    
    
    $latitudeIdentifier = 0

    $coordinateStringFirstSplit = $coordinateString.split(",")
    $latStringEndPos = $coordinateStringFirstSplit[0].indexof("`"N")
    $latString = $coordinateStringFirstSplit[0].substring(0, $latStringEndPos)

    # magic 3 - No. of spaces between the coordinate values in the string
    $lngStringStartPos = $latStringEndPos+3;
    $lngStringEndPos = $coordinateStringFirstSplit[0].indexof("`"W")
    $lngStringLength = $lngStringEndPos - $lngStringStartPos
    $lngString = $coordinateStringFirstSplit[0].substring($lngStringStartPos, $lngStringLength)
    $elevationString = $coordinateStringFirstSplit[1].trim()
    $elevationEndPos = $elevationString.indexof("M")
    $elevation = [double] $elevationString.substring(0, $elevationEndPos) 

    return ($latString, $lngString, $elevation)
    
}

# given a coordinate in dd mm ssssss - convert to decimal degree
function convertDMS2DD($coordinate) {
    
    $degreePartEndPos = $coordinate.indexof("°")
    $degreeValue = [int]$coordinate.substring(0,$degreePartEndPos) 

    $newCoordinateString = $coordinate.substring($degreePartEndPos+1)
    $minutePartEndPos = $newCoordinateString.indexof("`'")
    $minuteValue = [int]$newCoordinateString.substring(0,$minutePartEndPos) 

    $secondPartStartPos = $minutePartEndPos+1
    $secondValue = [double] $newCoordinateString.substring($secondPartStartPos)

    $decimalDegrees = $degreeValue + ($minuteValue/60) + ($secondValue/3600)
    return $decimalDegrees

}


function extract-imagedata
{

    $currentLocation = getCurrentLocation
    $outputFileFullPath = $currentLocation + "\" + $global:fileToSave

    $outputfileIsFound = test-path $outputFileFullPath
    if ($outputfileIsFound -eq $true) {
        exitWithError "The output file you specified is already present in the current folder. We don't want to overwrite it so resolve this issue to continue." ""
    }


    Write-Host "Converting: " $currentLocation
    Write-Host "Saving to: " $outputFileFullPath

    $fileList = get-childitem -filter "*.jpg" | select {$_.name}
    $rowCounter = 1
    foreach ($file in $fileList) {
        $imageFileName = $file.'$_.name'
        $fullFilePath = $currentLocation + "\" + $imageFileName
        $fileInfo = get-image $fullFilePath | get-exif
        $locationString = $fileInfo.GPS
        if ($locationString -eq "") {
            write-host "Image $imageFileName does not have GPS exif data"
            continue
        }
        
        $ddmmssCoordinates = normalizeCoordinateStringCanon260 $locationString

        # expecting the lat as the first coordinate in the pair
        $normalizedLat = convertDMS2DD($ddmmssCoordinates[0])
        $normalizedLng = convertDMS2DD($ddmmssCoordinates[1])

        #giant hack - only dealing with west coordinates - these should be "negative" in the decimal degree representation
        $normalizedLng = $normalizedLng*-1
        
        $normalizedElevation = $ddmmssCoordinates[2]

        $theRowAsCSV = [string]$rowCounter + "," + $imageFileName + "," + "19-JAN-2014" + "," + $normalizedLng + "," + $normalizedLat + "," + $normalizedElevation  
        Add-Content -path $outputFileFullPath -value $theRowAsCSV
        $rowCounter += 1
    }
    Write-Host "Image processing completed"
    Write-Host "Output file: $outfileFullPath"

}


extract-imagedata
