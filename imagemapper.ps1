<#
.SYNOPSIS
    Forensic tool to extract and log image metadata
.DESCRIPTION
    Extracts metadata from folder of imagee files. Creates a CSV log of image data
    and generates a KML file for Google Earth containing placemarks for any images
    containing GPS data.
.PARAMETER Source
    The path to the folder containing the images for examination
.PARAMETER Target
    The path where the CSV and KML output files will be stored.
.PARAMETER TargetFileName
    Case Number or any other name unique to the case
.NOTES
    File Name   :   imagemapper.ps1
    Author      :   Garrett Ed Pewitt
    Version     :   1.0
    Requires    :   PowerShell v5
.LINK
    http://www.forensicexpedition.com/tools/imagemapper.ps1
.EXAMPLE
    imagemapper.ps1 -source "C:\CaseImages" -target "C:\Cases\001\Images"
    Example showing how to run script against folder of images
.EXAMPLE
    imagemapper.ps1 -source "C:\Case\pics.zip" -target "C:\Cases\001\Images"
    Example showing how to run script against a zip file containing images
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$True, Position=1)]
    [string]$Source,
    [Parameter(Mandatory=$False, Position=2)]
    [string]$Target,
    [Parameter(Mandatory=$False, Position=3)]
    [string]$TargetFileName,
    [Parameter(Mandatory=$False, Position=4)]
    [string]$Hash
)

$CurDate = Get-Date -f "yyyyMMdd"

Function Unzip
{
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.Filesystem')
    Expand-Archive $Source -DestinationPath "$TargetPath\Images"
}

Function AddCSV
{
    # Function to manually create CSV is needed.
    $CSVContent = "`"$ImgBaseName`",`"$ImgExt`",`"$ImgFolderPath`",`"" `
    + "$ImgDateCreated`",`"$ImgDateAccessed`",`"$ImgDateModified`",`"$ImgSize`",`"" `
    + "$ImgAttributes`",`"$ImgReadOnly`",`"$ImgMD5`",`"$ImgSHA`",`"$ImgSHA256`",`"$ImgPerceivedType`",`"" `
    + "$ImgOwner`",`"$ImgDateTaken`",`"$ImgProtected`",`"$ImgCameraModel`",`"$ImageDimensions`",`"" `
    + "$ImgCameraMake`",`"$ImgLocation`",`"$ImgBitDepth`",`"$ImgHorRes`",`"$ImgWidth`",`"$ImgVertRes`",`"" `
    + "$ImgHeight`",`"$ImgExifVer`",`"$ImgEvent`",`"$ImgExpBias`",`"" `
    + "$ImgExpProgram`",`"$ImgExpTime`",`"$ImgFstop`",`"$ImgFlashMode`",`"$ImgFocalLen`",`"$ImgISOSpeed`",`"" `
    + "$ImgLensMaker`",`"$ImgLensModel`",`"$ImgLightSrc`",`"$ImgMaxAper`",`"$ImgMeterMode`",`"$ImgOrientation`",`"" `
    + "$ImgPeople`",`"$ImgProgMode`",`"$ImgSaturation`",`"$ImgSubDist`",`"$ImgWhiteBal`",`"$ImgPriority`",`"" `
    + "$ImgGPS`",`"$LatOrt`",`"$LonOrt`",`"$Alt`",`"$SeaLev`""

    Add-Content $CSVFile $CSVContent
}

Function ImageMetaDataExtract
{
    # Function to extract data from image files

}

Function AddKMLPlacemark ($ImgName, $ImgDesc, $ImgLon, $ImgLat, $ImgAlt)
{
    # Function to add placemark to KML file for images with GPS data
    Add-Content $KMLFile "<Placemark>`r`n<name>$ImgName</name>`r`n<description>$ImgDesc</description>`r`n<Point>`r`n<coordinates>$ImgLon,$ImgLat,$ImgAlt</coordinates>`r`n</Point>`r`n</Placemark>"

}

Function HashFile
{
    $MD5Hash =  Get-FileHash -Path $ImgFilePath -Algorithm MD5
    $SHAHash = Get-FileHash -Path $ImgFilePath -Algorithm SHA1
    $SHA256Hash = Get-FileHash -Path $ImgFilePath -Algorithm SHA256
    Return $MD5Hash.Hash, $SHAHash.Hash, $SHA256Hash.Hash
}

# Test source path
If ($Source -ne '')
{
    Write-Host "Source Found"
    # Check to see if Source path is valid
    If ((Test-Path($Source)) -eq $true)
    {
        # If no target path is specifed use current user desktop as target.
        If ($Target -eq '')
        {
            $Target = [Environment]::GetFolderPath("Desktop")

            If ($TargetFileName -eq '')
            {
                $TargetFileName = $CurDate
            }
            $TargetPath = $Target + "\" + $TargetFileName + "-imagemapper"
            New-Item $TargetPath -Type Directory
        }
        else
        {
            Write-Host "Target not found"
        }
        
        # Create output files
        $CSVFile = "$TargetPath\$TargetFileName-Results.csv"
        $KMLFile = "$TargetPath\$TargetFileName-GeoMap.kml"

        If ($Source -like '*zip*')
        {
            Unzip $Source
            $Images = Get-ChildItem "$TargetPath\Images" -Recurse | Where { ! $_.PSIsContainer}

        }
        else {
            $TargetPath = $Source
            $Images = Get-ChildItem "$TargetPath"  -Recurse | Where { ! $_.PSIsContainer}
        }      

        # Create CSV file and add CSV File Header
        New-Item $CSVFile -Type File
        $CSVHeader = '"File Name","Extension","Directory","Creation Time","Access Time","Modified Time","File Size",' `
        + '"Attributes","Read Only","MD5 Hash","SHA1 Hash","SHA256 Hash","Perceived Type","Owner","Date Taken","Protected","Camera Model",' `
        + '"Dimensions","Camera Make","Location","Bit Depth","Horz Res","Width","Vert Res","Height",' `
        + '"EXIF Ver","Event","Exp Bias","Exp Program","Exp Time","F Stop","Flash Mode","Focal Len","ISO Speed","Lens Maker",' `
        + '"Lens Model","Light Src","Max Aperature","Meter Mode","Orientation","People","Program Mode","Saturation","Subject Dist",' `
        + '"White Bal","Priority","GPS Data","Latitude","Longitude","Altitude","Sea Level"'
        Add-Content $CSVFile $CSVHeader

        # Create KML File and add KML File Header
        New-Item $KMLFile -Type File
        Add-Content $KMLFile "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`r`n<kml xmlns=`"http://www.opengis.net/kml/2.2`">`r`n<Folder>`r`n<name>$TargetFileName-Images</name>"

        [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");

        ForEach ($Image in $Images)
        {
            # Get File Metadata
            $COM = New-Object -COMObject Shell.Application
            $Folder = Split-Path $Image.FullName
            $File = Split-Path $Image.FullName -leaf
            $COMFolder = $COM.Namespace($Folder)
            $COMFile = $COMFolder.ParseName($File)
            $MData = New-Object -TypeName PSCustomObject -Property @{
                Name = $COMfolder.GetDetailsOf($COMFile,0)
                Size = $COMfolder.GetDetailsOf($COMFile,1)
                Type = $COMfolder.GetDetailsOf($COMFile,2)
                DateModified = $COMfolder.GetDetailsOf($COMFile,3)
                DateCreated = $COMfolder.GetDetailsOf($COMFile,4)
                DateAccessed = $COMfolder.GetDetailsOf($COMFile,5)
                Attributes = $COMfolder.GetDetailsOf($COMFile,6)
                PerceivedType = $COMfolder.GetDetailsOf($COMFile,9)
                Owner = $COMfolder.GetDetailsOf($COMFile,10)
                DateTaken = $COMfolder.GetDetailsOf($COMFile,12)
                Protected = $COMfolder.GetDetailsOf($COMFile,29)
                CameraModel = $COMfolder.GetDetailsOf($COMFile,30)
                Dimensions = $COMfolder.GetDetailsOf($COMFile,31)
                CameraMake = $COMfolder.GetDetailsOf($COMFile,32)
                Location = $COMfolder.GetDetailsOf($COMFile,41)
                BitDepth = $COMfolder.GetDetailsOf($COMFile,169)
                HorRes = $COMfolder.GetDetailsOf($COMFile,170)
                Width = $COMfolder.GetDetailsOf($COMFile,171)
                VertRes = $COMfolder.GetDetailsOf($COMFile,172)
                Height = $COMfolder.GetDetailsOf($COMFile,173)
                FolderPath = $COMfolder.GetDetailsOf($COMFile,186)
                EXIFVer = $COMfolder.GetDetailsOf($COMFile,250)
                Event = $COMfolder.GetDetailsOf($COMFile,251)
                ExpBias = $COMfolder.GetDetailsOf($COMFile,252)
                ExpProgram = $COMfolder.GetDetailsOf($COMFile,253)
                ExpTime = $COMfolder.GetDetailsOf($COMFile,254)
                Fstop = $COMfolder.GetDetailsOf($COMFile,255)
                FlashMode = $COMfolder.GetDetailsOf($COMFile,256)
                FocalLen = $COMfolder.GetDetailsOf($COMFile,257)
                MMFocalLen = $COMfolder.GetDetailsOf($COMFile,258)
                ISOSpeed = $COMfolder.GetDetailsOf($COMFile,259)
                LensMaker = $COMfolder.GetDetailsOf($COMFile,260)
                LensModel = $COMfolder.GetDetailsOf($COMFile,261)
                LightSrc = $COMfolder.GetDetailsOf($COMFile,262)
                MaxAper = $COMfolder.GetDetailsOf($COMFile,263)
                MeterMode = $COMfolder.GetDetailsOf($COMFile,264)
                Orientation = $COMfolder.GetDetailsOf($COMFile,265)
                People = $COMfolder.GetDetailsOf($COMFile,266)
                ProgMode = $COMfolder.GetDetailsOf($COMFile,267)
                Saturation = $COMfolder.GetDetailsOf($COMFile,268)
                SubDist = $COMfolder.GetDetailsOf($COMFile,269)
                WhiteBal = $COMfolder.GetDetailsOf($COMFile,270)
                Priority = $COMfolder.GetDetailsOf($COMFile,271)
                }

                $ImgName = $MData.Name
                $ImgSize = $MData.Size
                $ImgType = $MData.Type
                $ImgDateModified = $MData.DateModified
                $ImgDateCreated = $MData.DateCreated
                $ImgDateAccessed = $MData.DateAccessed
                $ImgAttributes = $MData.Attributes
                $ImgPerceivedType = $MData.PerceivedType
                $ImgOwner = $MData.Owner
                $ImgDateTaken = $MData.DateTaken
                $ImgProtected = $MData.Protected
                $ImgCameraModel = $MData.CameraModel
                $ImgDimensions = $MData.Dimensions
                $ImgCameraMake = $MData.CameraMake
                $ImgLocation = $MData.Location
                $ImgBitDepth = $MData.BitDepth
                $ImgHorRes = $MData.HorRes
                $ImgWidth = $MData.Width
                $ImgVertRes = $MData.VertRes
                $ImgHeight = $MData.Height
                $ImgFolderPath = $MData.FolderPath
                $ImgEXIFVer = $MData.EXIFVer
                $ImgEvent = $MData.Event
                $ImgExpBias = $MData.ExpBias
                $ImgExpProgram = $MData.ExpProgram
                $ImgExpTime = $MData.ExpTime
                $ImgFstop = $MData.Fstop
                $ImgFlashMode = $MData.FlashMode
                $ImgFocalLen = $MData.FocalLen
                $ImgMMFocalLen = $MData.MMFocalLen
                $ImgISOSpeed = $MData.ISOSpeed
                $ImgLensMaker = $MData.LensMaker
                $ImgLensModel = $MData.LensModel
                $ImgLightSrc = $MData.LightSrc
                $ImgMaxAper = $MData.MaxAper
                $ImgMeterMode = $MData.MeterMode
                $ImgOrientation = $MData.Orientation
                $ImgPeople = $MData.People
                $ImgProgMode = $MData.ProgMode
                $ImgSaturation = $MData.Saturation
                $ImgSubDist = $MData.SubDist
                $ImgWhiteBal = $MData.WhiteBal
                $ImgPriority = $MData.Priority
                $ImgExt = $Image.Extension
                $ImgBaseName = $Image.basename
                $ImgLastAccess = $Image.LastAccessTime
                $ImgReadOnly = $Image.IsReadOnly
                $ImgFilePath = $Image.FullName

            $Hashes = Hashfile
            $ImgMD5 = $Hashes[0]
            $ImgSHA = $Hashes[1]
            $ImgSHA256 = $Hashes[2]

            Try 
            {
                $img = New-Object -TypeName system.drawing.bitmap -ArgumentList $ImgFilePath;
            }
            Catch 
            {
                $FileStatus = "Error"
            }

            # Set default values
            $GPSInfo = $true
            $ImgGPS = "TRUE"
            $Encode = New-Object System.Text.ASCIIEncoding

            Try
            {
                $LatNS = $Encode.GetString($img.GetPropertyItem(1).Value)
            }
            Catch
            {
                $GPSInfo = $False
                $ImgGPS = "FALSE"
            }

            If ($GPSInfo -eq $true)
            {
                $LatDeg = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(2).Value, 0)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(2).Value, 4)))
                $LatMin = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(2).Value, 8)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(2).Value, 12)))
                $LatSec = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(2).Value, 16)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(2).Value, 20)))

                $LonEW = $Encode.GetString($img.GetPropertyItem(3).Value)
                $LonDeg = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(4).Value, 0)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(4).Value, 4)))
                $LonMin = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(4).Value, 8)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(4).Value, 12)))
                $LonSec = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(4).Value, 16)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(4).Value, 20)))

                Try
                {
                    $SeaLev = 1 - ([System.BitConverter]::ToInt32($img.GetPropertyItem(6).Value, 0))
                    $Alt = (([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(6).Value, 0)) / ([Decimal][System.BitConverter]::ToInt32($img.GetPropertyItem(6).Value, 4)))

                }
                Catch
                {
                    $SeaLev = 0
                    $Alt = 0
                    Write-Host "Altitude not found"
                }

                # Convert to decimal Degrees
                If ($LatNS -eq 'S')
                {
                    $LatOrt = "-"   
                }
                If ($LonEW -eq 'W')
                {
                    $LonOrt = "-"
                }
                $LatDec = ($LatDeg + ($LatMin/60) + ($LatSec/3600))
                $LonDec = ($LonDeg + ($LonMin/60) + ($LonSec/3600))

                $LatOrt = $LatOrt + $LatDec
                $LonOrt = $LonOrt + $LonDec

                # Add information to KML File
                $TestDesc = "$ImgBaseName`r`n$ImgFolderPath`r`n`r`nCreate Time: $ImgDateCreated `r`nAccess Time: $ImgDateAccessed`r`nModified Time: $ImgDateModified`r`n`r`nMD5 Hash: $ImgMD5`r`nSHA1 Hash: $ImgSHA`r`nSHA256 Hash: $ImgSHA256`r`n`r`nLatitude: $Latort`r`nLongitude: $LonOrt`r`nAltitude: $Alt`r`nAbove Sea Level: $SeaLev"
                AddKMLPlacemark $ImgName $TestDesc $LonOrt $LatOrt $Alt

                # Add information to CSV File
                AddCSV
            }
            else
            {
                # Add information to CSV File
                AddCSV
            }

            $LatDeg = $null
            $LatMin = $null
            $LatSec = $null
            $LonDeg = $null
            $LonMin = $null
            $LonSec = $null
            $SeaLev = $null
            $Alt = $null
            $LatNS = $null
            $LonEW = $null
            $LatOrt = $null
            $LonOrt = $null
            $LatDec = $null
            $LonDec = $null
        }

        # KML Footer
        Add-Content $KMLFile "</Folder>`r`n</kml>"
    }
    else
    {
        Write-Host "Source Test Path Failed"
    }
}