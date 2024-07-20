<#
.SYNOPSIS
Downloads image tiles from a IIIF manifest.

.DESCRIPTION
This script downloads images from a IIIF manifest specified by the ManifestUrl parameter. It allows you to specify a download folder, page range, and scale factor for the images.

.PARAMETER ManifestUrl
The URL of the IIIF manifest.

.PARAMETER DownloadFolder
The folder where the downloaded images will be saved. If not specified, the current folder will be used.

.PARAMETER PageRange
The range of pages to download. If not specified, all pages will be downloaded. Non-existant pages will be ignored.

.PARAMETER ForceDownload
If specified, the script will overwrite existing files in the download folder.

.PARAMETER ScaleFactor
The scale factor to apply to the downloaded images. It will be ignored if it's not in the manifest and will default to 1.

.EXAMPLE
powershell.exe -Sta -File .\iiifetch.ps1 -ManifestUrl "https://example.com/manifest.json" -DownloadFolder "C:\Downloads" -PageRange "1-5" -ForceDownload -ScaleFactor 2
Downloads the images from the specified manifest URL, saves them in the "C:\Downloads" folder, downloads only pages 1 to 5, overwrites existing files, and applies a scale factor of 2.

#>

[CmdletBinding()]
param (
  [Parameter(Mandatory=$true)]
  [string]$ManifestUrl,
  [string]$DownloadFolder,
  [string]$PageRange,
  [bool]$ForceDownload,
  [int]$ScaleFactor
)

function Get-JsonData {
  param (
    [string]$Url
  )

  # Try to retrieve the JSON data
  try {
    # Use Invoke-RestMethod to retrieve the JSON data
    $jsonData = Invoke-RestMethod -Uri $Url -Method Get

    # Return the parsed JSON data
    return $jsonData
  }
  catch {
    # Notify the user that there was an error
    Write-Host "Error: $($_.Exception.Message)"
    # Exit the script
    exit
  }
}

# Prompt the user for the manifest URL, download folder, and page range (optional)
function Get-UserInput {
  param (
    [string]$ManifestUrl,
    [string]$DownloadFolder,
    [string]$PageRange
  )

  # Prompt the user for the manifest URL
  $ManifestUrl = Read-Host "Enter manifest url"

  # Prompt the user for the download folder
  $DownloadFolder = Read-Host "Enter download folder (default: current folder)"

  # Prompt the user for the page range
  $PageRange = Read-Host "Enter page range (default: all pages)"

  # Set default values if the user did not enter any
  if ($DownloadFolder -eq "") {
    $DownloadFolder = "."
  }

  # Return the user input
  return $ManifestUrl, $DownloadFolder, $PageRange
}

# page range parser, supports printer-friendly page ranges (e.g. 1-3, 5, 7-9, (1,3,5))
function Get-PageRange {
  param (
    [string]$PageRange
  )

  # Split the page range into an array of pages
  $pages = $PageRange -split ","

  # Create an array to store the page numbers
  $pageNumbers = @()

  # Loop through each page
  foreach ($page in $pages) {
    # Check if the page is a range
    if ($page -like "*-*") {
      # Split the range into a start and end page
      $currentRange = $page -split "-"

      # Loop through each page in the range
      for ($i = [int]$currentRange[0]; $i -le [int]$currentRange[1]; $i++) {
        # Add the page number to the array
        $pageNumbers += $i
      }
    }
    else {
      # Add the page number to the array
      $pageNumbers += [int]$page
    }
  }

  # Return the page numbers
  return $pageNumbers
}

# get the required pages from the manifest json data
function Get-RequiredPages {
  param (
    [object]$JsonData,
    [object]$PageNumbers
  )

  # Create an array to store the required pages
  $requiredPages = @()

  # if no page range is specified, download all pages
  if ($PageNumbers -eq 0 -or $PageNumbers.Count -eq 0) {
    # Loop through each page in the manifest
    foreach ($page in $JsonData.sequences[0].canvases) {
      $pageManifest = $page.images[0].resource.service.'@id'
      # Add the page to the array
      $requiredPages += [PSCustomObject]@{
        Label    = $page.label
        Manifest = $pageManifest
      }
    }
  }
  else {
    # Loop through each page number
    foreach ($pageNumber in $pageNumbers) {
      # Loop through each page in the manifest
      foreach ($page in $JsonData.sequences[0].canvases) {
        # Check if the page number matches the current page
        if ([string]$pageNumber -eq [string]$page.label) {
          $pageManifest = $page.images[0].resource.service.'@id'
          # Add the page to the array
          $requiredPages += [PSCustomObject]@{
            Label    = $page.label
            Manifest = $pageManifest
          }
        }
      }
    }
  }

  # Return the required pages
  return $requiredPages
}

# get the page manifests
function Get-PageManifests {
  param (
    [object]$Pages
  )

  # Create an array to store the page manifests
  $pageManifests = @()
  $index = 0
  # Loop through each required page
  $pageCount = $Pages.Count
  if ($null -eq $pageCount -or $pageCount -eq 0) {
    $pageCount = 1
  }

  # create manifests folder if it doesn't exist
  $manifestsFolder = Join-Path -Path $DownloadFolder -ChildPath "manifests"
  if (-not (Test-Path -Path $manifestsFolder)) {
    New-Item -Path $manifestsFolder -ItemType Directory
  }
  foreach ($requiredPage in $Pages) {
    $index++
    Write-Progress -Activity "Fetching Page Manifests" -Status "Fetching page $index of $($pageCount)" -PercentComplete ($index / $pageCount * 100)
    try {
      # check if page manifest exists in folder
      $manifestFilename = [string]::Format("{0}\{1}.json", $manifestsFolder, $requiredPage.Label)
      if ($ForceDownload -eq $true -or -not (Test-Path -Path $manifestFilename)) {
        # download page manifest
        Invoke-WebRequest -Uri $requiredPage.Manifest -Method Get -OutFile $manifestFilename
      }
      # get page manifest data from file
      $manifestData = Get-Content -Path $manifestFilename | ConvertFrom-Json
      $pageManifests += [PSCustomObject]@{
        Label    = $requiredPage.Label
        Manifest = $manifestData
      }
    }
    catch {
      # Notify the user that there was an error
      Write-Host "Error: $($_.Exception.Message)"
      # Exit the script
      exit
    }
  }
  Write-Progress -Activity "Fetching Page Manifests" -Status "Fetched all pages" -Completed
  # Return the page manifests
  return $pageManifests
}

# get all tiles per page
function Get-PageTiles {
  param (
    [object]$Manifest,
    [string]$TilesFolder,
    [int]$PageNumber,
    [string]$PageLabel
  )

  $urlBase = $Manifest.'@id'
  # image format, choose jpg if available, otherwise first available format
  $format = 'jpg'
  if ($null -ne $Manifest.profile[1].formats) {
    $formats = $Manifest.profile[1].formats
    if ($formats -contains 'jpg') {
      $format = $formats[0]
    }
  }

  # image dimensions
  $width = $Manifest.width
  $height = $Manifest.height

  # tile dimensions
  $tileWidth = $Manifest.tiles[0].width
  # we assume tileHeight to be the same as tileWidth if not specified.
  # TODO: check if this is a valid assumption.
  $tileHeight = $Manifest.tiles[0].height
  if ($null -eq $tileHeight) {
    $tileHeight = $tileWidth
  }

  $defaultScaleFactor = $ScaleFactor
  # if no scale factor is given default to 1
  if ($null -eq $defaultScaleFactor) {
    $defaultScaleFactor = 1
  }

  if ($null -ne $Manifest.tiles[0].scaleFactors) {
    $scaleFactors = $Manifest.tiles[0].scaleFactors
    $scaleFactors = $scaleFactors | Sort-Object
    if (-not ($scaleFactors -contains $defaultScaleFactor)) {
      $defaultScaleFactor = $scaleFactors[0]
    }
  }

  # region dimensions
  $regionWidth = $Manifest.tiles[0].width * $defaultScaleFactor
  $regionHeight = $Manifest.tiles[0].height * $defaultScaleFactor

  # scaled image dimensions
  $scaledWidth = $width * $defaultScaleFactor
  $scaledHeight = $height * $defaultScaleFactor

  $tilesData = @()
  $tileCount = 1
  $rowCounter = 0
  $colCounter = 0
  $y = 0

  # row loop
  while ($y -lt $height) {
    $x = 0
    $colCounter = 0

    #col loop
    while ($x -lt $width) {
      $region = 'full'
      if ($scaledWidth -gt $tileWidth -or $scaledHeight -gt $tileHeight) {
        $region = [string]::Format("{0},{1},{2},{3}", $x, $y, $regionWidth, $regionHeight)
      }
      $scaledRemainingWidth = [math]::ceiling(($width - $x) / $defaultScaleFactor)
      $currentTileWidth = $tileWidth
      if ($scaledRemainingWidth -lt $tileWidth) {
        $currentTileWidth = $scaledRemainingWidth
      }
      $curretTileHeight = $tileHeight
      $scaledRemainingHeight = [math]::ceiling(($height - $y) / $defaultScaleFactor)
      if ($scaledRemainingHeight -lt $tileHeight) {
        $currentTileHeight = $scaledRemainingHeight
      }

      $tileUrl = [string]::Format("{0}/{1}/{2},/0/default.{3}", $urlBase, $region, $currentTileWidth, $format)
      # download tile image
      # Invoke-WebRequest -Uri $tileUrl -Method Get -OutFile ([string]::Format("{0}\{1}.{2}", $TilesFolder, $tileCount, $format))
      if ($defaultScaleFactor -eq 1) {
        $filename = [string]::Format("{0}\{1}_{2}.{3}", $TilesFolder, $region, $currentTileWidth, $format)
      } else {
        $filename = [string]::Format("{0}\{1}_{2}x{3}.{4}", $TilesFolder, $region, $currentTileWidth, $defaultScaleFactor, $format)
      }
      $tilesData += [PSCustomObject]@{
        Url = $tileUrl
        Label = $tileCount
        X = $ColCounter
        Y = $RowCounter
        Filename = $filename
      }
      $tileCount++
      $colCounter++
      $x += $regionWidth
    }
    $y += $regionHeight
    $rowCounter++
  }

  Invoke-PageTilesDownload -TilesData $tilesData -PageLabel $PageLabel

  $newImageWidth = [math]::ceiling($width/$defaultScaleFactor)
  $newImageHeight = [math]::ceiling($height/$defaultScaleFactor)
  # convert to int
  $newImageWidth = [int]$newImageWidth
  $newImageHeight = [int]$newImageHeight
  Merge-PageTiles -Width $newImageWidth -Height $newImageHeight -PageLabel $PageLabel -TilesData $tilesData -SCaleFactor $defaultScaleFactor
}

# download page tiles
function Invoke-PageTilesDownload {
  param (
    [object]$TilesData,
    [string]$PageLabel
  )
  $index = 0
  foreach ($tile in $TilesData) {
    $index++
    Write-Progress -Activity "Fetching Tiles" -Status "Fetching tile $index/$($TilesData.Count) for page $PageLabel" -PercentComplete ($index / $TilesData.Count * 100)

    # if the file already exists, skip it unless ForceDownload is set to true
    if ($ForceDownload -eq $true -or -not (Test-Path -Path $tile.Filename)) {
      Invoke-WebRequest -Uri $tile.Url -Method Get -OutFile $tile.Filename
    }
  }
  Write-Progress -Activity "Fetching Tiles" -Status "Fetched all tiles for page $PageLabel" -Completed
}

# combine page tiles
function Merge-PageTiles {
  param (
    [int]$Width,
    [int]$Height,
    [string]$PageLabel,
    [object]$TilesData,
    [int]$ScaleFactor
  )

  # if combined file already exists, skip
  $savePath = [string]::Format("{0}\combined-{1}.jpg", $DownloadFolder, $PageLabel)
  if ($ScaleFactor -ne 1) {
    $savePath = [string]::Format("{0}\combined-{1}x{2}.jpg", $DownloadFolder, $PageLabel, $ScaleFactor)
  }
  if (Test-Path -Path $savePath) {
    return
  }

  # loop through tiles and calculate real width and height
  $realWidth = 0
  $realHeight = 0
  $seenRows = @()
  $seenCols = @()
  foreach ($tile in $TilesData) {
    $tileImage = New-Object System.Drawing.Bitmap($tile.Filename)
    $tileImageWidth = $tileImage.Width
    $tileImageHeight = $tileImage.Height
    if ($seenRows -notcontains $tile.X) {
      $realWidth += $tileImageWidth
      $seenRows += $tile.X
    }
    if ($seenCols -notcontains $tile.Y) {
      $realHeight += $tileImageHeight
      $seenCols += $tile.Y
    }

  }

  # create a bitmap to store the combined image
  $combinedImage = New-Object System.Drawing.Bitmap($realWidth, $realHeight)

  $index = 0
  $currentX = 0
  $currentY = 0
  $currentRow = 0
  $prevTileImageHeight = 0

  foreach ($tile in $TilesData) {
    $index++
    Write-Progress -Activity "Combining Tiles" -Status "Combining tile $index/$($TilesData.Count) for page $PageLabel" -PercentComplete ($index / $TilesData.Count * 100)

    $tileImage = New-Object System.Drawing.Bitmap($tile.Filename)
    $tileImageWidth = $tileImage.Width
    $tileImageHeight = $tileImage.Height

    # if we are on a new row, reset the x position and increment the y position
    if ($currentRow -ne $tile.Y) {
      $currentRow = $tile.Y
      $currentX = 0
      $currentY += $prevTileImageHeight
    }

    $prevTileImageHeight = $tileImage.Height
    $tileImageRect = New-Object System.Drawing.Rectangle($currentX, $currentY, $tileImageWidth, $tileImageHeight)
    $tileImageGraphics = [System.Drawing.Graphics]::FromImage($combinedImage)
    $tileImageGraphics.DrawImage($tileImage, $tileImageRect)

    $currentX += $tileImageWidth
  }

  # save the combined image
  $combinedImage.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
  Write-Progress -Activity "Combining Tiles" -Status "Combined tiles for page $PageLabel" -Completed
}

# get all pages tiles
function Get-AllPagesTiles {
  param (
    [object]$PageManifests
  )
  $index = 0
  # Loop through each page manifest
  foreach ($pageManifest in $pageManifests) {
    $index++
    # Create a folder for the page
    $pageFolder = Join-Path -Path $DownloadFolder -ChildPath $pageManifest.Label
    if (-not (Test-Path -Path $pageFolder)) {
      New-Item -Path $pageFolder -ItemType Directory
    }
    # Get the page tiles
    Get-PageTiles -Manifest $pageManifest.Manifest -TilesFolder $pageFolder -PageNumber $index -PageLabel $pageManifest.Label
  }
}

function Convert-StringToValidFileName {
  param (
      [string]$InputString
  )

  # Remove invalid characters
  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join '_'
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  $sanitizedString = ($InputString -replace $re)
  $sanitizedString = $inputString -replace '[\\\/\:\*\?\"\<\>\|\s\.\^,\[\]\(\)]', '_'
  $sanitizedString = $inputString -replace '[\\\/\:\*\?\"\<\>\|\s\.\^,\[\]\(\)]', '_'
  $sanitizedString = $sanitizedString -replace '_+', '_'
  $sanitizedString = $sanitizedString -replace '^_', ''
  $sanitizedString = $sanitizedString -replace '_$', ''

  return $sanitizedString
}

  Add-Type -Assembly System.Drawing

  Write-Host "This script only supports IIIF 2 and does not check for manifest validity"

  # Get the user input
  # $ManifestUrl, $DownloadFolder, $PageRange = Get-UserInput
  Write-Progress -Activity "Fetching Manifest" -Status "Fetching..."
  $JsonData = Get-JsonData -Url $ManifestUrl
  Write-Progress -Activity "Fetching Manifest" -Status "Fetched" -Completed

  # create the download folder if it does not exist
  $BookTitle = Convert-StringToValidFileName -InputString $JsonData.label
  Write-Host "Book: ${bookTitle}"

  if ($null -eq $DownloadFolder -or $DownloadFolder -eq "") {
    $DownloadFolder = "."
  }
  $DownloadFolder = Join-Path -Path $DownloadFolder -ChildPath $BookTitle
  if (-not (Test-Path -Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory
  }

  # get pages and all that
  $PageNumbers = Get-PageRange -PageRange $PageRange
  $RequiredPages = Get-RequiredPages -JsonData $JsonData -PageNumbers $PageNumbers
  $PageManifests = Get-PageManifests -Pages $RequiredPages
  Get-AllPagesTiles -PageManifests $PageManifests
  Write-Host "All done."
  exit

