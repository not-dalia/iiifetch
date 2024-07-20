# IIIFetch

IIIFetch is a PowerShell script that allows users to download and assemble high-quality images from IIIF (International Image Interoperability Framework) manifests.

## Features

- Download image tiles from IIIF manifests
- Assemble downloaded tiles into complete images
- Support for specifying page ranges
- Adjustable scale factor for image quality

## Requirements

- PowerShell 5.1 or higher
- .NET Framework 4.5 or higher (for System.Drawing assembly)

## Usage

```powershell
powershell.exe -Sta -File .\iiifetch.ps1 -ManifestUrl "https://example.com/manifest.json" -DownloadFolder "C:\Downloads" -PageRange "1-5" -ForceDownload -ScaleFactor 2
```

### Parameters

- `ManifestUrl` (Required): The URL of the IIIF manifest.
- `DownloadFolder`: The folder where the downloaded images will be saved. If not specified, the current folder will be used.
- `PageRange`: The range of pages to download. If not specified, all pages will be downloaded.
- `ForceDownload`: If specified, the script will overwrite existing files in the download folder, otherwise it will resume download from last file.
- `ScaleFactor`: The scale factor to apply to the downloaded images. It will be ignored if it's not in the manifest and will default to 1.

## Limitations

- This script only supports IIIF version 2 manifests.
- The script does not check for manifest validity.
