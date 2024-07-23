# Remove the first 27 characters from the processor name
$processorNameShort = if ($processor.Name.Length -gt 20) { $processor.Name.Substring(27, [math]::Min($processor.Name.Length - 27, 20)) + "" } else { $processor.Name }

# Remove " Notebook PC" from the model name
$modelNameShort = if ($cs.Model -like "* Notebook PC") { $cs.Model.Replace(" Notebook PC", "") } else { $cs.Model }

# Define the URL for the ZXing.Net NuGet package
$nugetUrl = "https://www.nuget.org/api/v2/package/ZXing.Net/0.16.8"

# Define the path to save the downloaded package and extracted DLL
$nugetPackagePath = Join-Path -Path $PSScriptRoot -ChildPath "ZXing.Net.zip"
$extractionSubfolder = Join-Path -Path $PSScriptRoot -ChildPath "ZXingExtraction"
$dllPath = Join-Path -Path $extractionSubfolder -ChildPath "lib/netstandard2.0/zxing.dll"

# Path to Handle.exe
$handlePath = Join-Path -Path $PSScriptRoot -ChildPath "Handle.exe"

# Function to download the NuGet package
function Download-NuGetPackage {
    if (-Not (Test-Path $nugetPackagePath)) {
        Write-Output "Downloading NuGet package..."
        Invoke-WebRequest -Uri $nugetUrl -OutFile $nugetPackagePath
    } else {
        Write-Output "NuGet package already exists."
    }
}

# Function to extract the DLL from the NuGet package
function Extract-DLL {
    # Ensure no processes are using the DLL
    Close-DLLHandles

    # Remove existing extraction directory if it exists
    if (Test-Path $extractionSubfolder) {
        Write-Output "Removing existing extraction subfolder..."
        try {
            Remove-Item -Recurse -Force $extractionSubfolder -ErrorAction Stop
        } catch {
            Write-Output "Initial attempt to remove extraction subfolder failed: $_"
        }
    }

    # Wait and retry logic to ensure the folder is removed
    $attempts = 0
    while ((Test-Path $extractionSubfolder) -and ($attempts -lt 10)) {
        Start-Sleep -Seconds 2
        try {
            Remove-Item -Recurse -Force $extractionSubfolder -ErrorAction Stop
        } catch {
            Write-Output "Attempt $($attempts + 1) to remove extraction subfolder failed: $_"
        }
        $attempts++
    }

    if (Test-Path $extractionSubfolder) {
        Write-Error "Failed to remove existing extraction subfolder."
        return $false
    }

    # Create the extraction subfolder
    New-Item -ItemType Directory -Path $extractionSubfolder

    # Extract the DLL from the NuGet package
    Write-Output "Extracting NuGet package..."
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($nugetPackagePath, $extractionSubfolder)

    # Confirm DLL extraction
    if (Test-Path $dllPath) {
        Write-Output "DLL extracted successfully: $dllPath"
        return $true
    } else {
        Write-Error "Failed to extract DLL: $dllPath"
        return $false
    }
}

# Function to close handles to zxing.dll using Handle.exe
function Close-DLLHandles {
    if (-Not (Test-Path $handlePath)) {
        Write-Output "Handle.exe not found. Skipping handle closing."
        return
    }
    
    Write-Output "Checking for open handles to $dllPath..."
    $handleOutput = & $handlePath $dllPath
    $regexPattern = [regex]"\s+pid:\s+(\d+)\s+type:\s+(\w+)\s+(\w+)\s+$([regex]::Escape($dllPath))"
    $matches = $regexPattern.Matches($handleOutput)

    if ($matches.Count -gt 0) {
        foreach ($match in $matches) {
            $pid = [int]$match.Groups[1].Value
            Write-Output "Terminating process ID $pid using $dllPath"
            Stop-Process -Id $pid -Force
        }
        Start-Sleep -Seconds 2 # Give it some time to release handles
    } else {
        Write-Output "No open handles found."
    }
}

# Function to load the ZXing.Net assembly
function Load-ZXingAssembly {
    try {
        Write-Output "Loading ZXing.Net assembly..."
        Add-Type -Path $dllPath
        Write-Output "Loaded types from ZXing assembly:"
        [Reflection.Assembly]::LoadFrom($dllPath).GetTypes() | ForEach-Object { Write-Output $_.FullName }
        return $true
    } catch {
        Write-Error "Error loading ZXing.Net assembly: $_"
        return $false
    }
}

# Download and extract the NuGet package
Download-NuGetPackage
if (-Not (Extract-DLL)) {
    Write-Output "Redownloading and extracting NuGet package..."
    Remove-Item -Force $nugetPackagePath
    Download-NuGetPackage
    if (-Not (Extract-DLL)) {
        Write-Error "Failed to extract DLL after redownloading."
        exit
    }
}

# Load the ZXing.Net assembly
if (-Not (Load-ZXingAssembly)) {
    Close-DLLHandles
    if (-Not (Load-ZXingAssembly)) {
        Write-Error "Failed to load ZXing.Net assembly after closing handles."
        exit
    }
}

# Function to create a bitmap from ZXing's BitMatrix
function Create-BitmapFromBitMatrix {
    param (
        [ZXing.Common.BitMatrix]$matrix
    )

    $width = $matrix.Width
    $height = $matrix.Height
    $bitmap = New-Object System.Drawing.Bitmap($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)

    for ($y = 0; $y -lt $height; $y++) {
        for ($x = 0; $x -lt $width; $x++) {
            $color = if ($matrix[$x, $y]) { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::White }
            $bitmap.SetPixel($x, $y, $color)
        }
    }

    return $bitmap
}

# Function to resize a bitmap
function Resize-Bitmap {
    param (
        [System.Drawing.Bitmap]$bitmap,
        [int]$width,
        [int]$height
    )

    $resizedBitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($resizedBitmap)
    $graphics.DrawImage($bitmap, 0, 0, $width, $height)
    $graphics.Dispose()
    return $resizedBitmap
}

# QR Code generation using ZXing.QrCode.QRCodeWriter
function Generate-QRCode {
    param (
        [string]$content,
        [string]$filePath
    )

    $qrCodeWriter = New-Object ZXing.QrCode.QRCodeWriter
    $qrCodeMatrix = $qrCodeWriter.encode($content, [ZXing.BarcodeFormat]::QR_CODE, 256, 256)

    $qrCodeBitmap = Create-BitmapFromBitMatrix -matrix $qrCodeMatrix
    $qrCodeBitmap.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)
    return $qrCodeBitmap
}

# QR Code generation
$cs = Get-WmiObject -Class Win32_ComputerSystem
# $qrCodeContent = "$($cs.Name)"
$qrCodeContent = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
$qrCodeFilePath = Join-Path -Path "C:\" -ChildPath "logs\qrcode.png"
$qrCodeBitmap = Generate-QRCode -content $qrCodeContent -filePath $qrCodeFilePath
$qrCodeBitmap = Resize-Bitmap -bitmap $qrCodeBitmap -width 170 -height 170 # Resize QR code to larger size
Write-Output "QR Code saved to $qrCodeFilePath"

# Barcode generation using ZXing.BarcodeWriterPixelData
try {
    $writer = New-Object ZXing.BarcodeWriterPixelData
    $writer.Format = [ZXing.BarcodeFormat]::CODE_128

    # $barcodeContent = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
    $barcodeContent = "$($cs.Name)"
    $barcodePixelData = $writer.Write($barcodeContent)
    $barcodeBitmap = New-Object System.Drawing.Bitmap($barcodePixelData.Width, $barcodePixelData.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $bitmapData = $barcodeBitmap.LockBits([System.Drawing.Rectangle]::FromLTRB(0, 0, $barcodePixelData.Width, $barcodePixelData.Height), [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    [System.Runtime.InteropServices.Marshal]::Copy($barcodePixelData.Pixels, 0, $bitmapData.Scan0, $barcodePixelData.Pixels.Length)
    $barcodeBitmap.UnlockBits($bitmapData)

    $barcodeFilePath = Join-Path -Path "C:\" -ChildPath "logs\barcode.png"
    $barcodeBitmap.Save($barcodeFilePath, [System.Drawing.Imaging.ImageFormat]::Png)
    $barcodeBitmap = Resize-Bitmap -bitmap $barcodeBitmap -width 450 -height 80 # Resize barcode to smaller size
    Write-Output "Barcode saved to $barcodeFilePath"
} catch {
    Write-Error "Error generating barcode: $_"
}

try {
    # Gather system information using WMI
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $cs = Get-WmiObject -Class Win32_ComputerSystem
    $processor = Get-WmiObject -Class Win32_Processor | Select-Object -First 1
    $disk = Get-WmiObject -Class Win32_DiskDrive -Filter "Index=0"
    $SerialNumber = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber

    # Define a function to resolve model names based on SKU
    function Get-ModelName {
        param (
            [string]$sku
        )

        switch ($sku) {
            "Surface_Pro_1796" { return "Surface Pro 5" }
            "Surface_Pro_1807" { return "Surface Pro 5 LTE Advanced" }
            "Surface_Pro_6_1796_Commercial" { return "Surface Pro 6 Commercial" }
            "Surface_Pro_6_1796_Consumer" { return "Surface Pro 6 Consumer" }
            "Surface_Pro_7_1866" { return "Surface Pro 7" }
            "Surface_Pro_7+_1960" { return "Surface Pro 7+" }
            "Surface_Pro_7+_with_LTE_Advanced_1961" { return "Surface Pro 7+ LTE" }
            "Surface_Pro_8_for_Business_1983" { return "Surface Pro 8 for Business" }
            "Surface_Pro_8_1983" { return "Surface Pro 8 Consumer" }
            "Surface_Pro_8_for_Business_with_LTE_Advanced_1982" { return "Surface Pro Pro 8 LTE for Business" }
            "Surface_Pro_9_2038" { return "Surface Pro 9" }
            "Surface_Pro_9_for_Business_2038" { return "Surface Pro 9 for Business" }
            "Surface_Pro_9_With_5G_1996" { return "Surface Pro 9 with 5G" }
            "LENOVO_MT_11LV_BU_Think_FM_ThinkCentre M60e" { return "LENOVO ThinkCentre M60e" }
            default { return "Not listed - $sku" }
        }
    }

    # Extract and shorten the processor name
    function Get-ShortProcessorName {
        param (
            [string]$processorName
        )

        if ($processorName -match "i\d-\d{4,5}[A-Za-z0-9]*") {
            return $matches[0]
        }
        return $processorName
    }

    # Get the model name using the function
    $modelname = Get-ModelName -sku $cs.SystemSKUNumber
    if ($modelname -eq "Not listed - $($cs.SystemSKUNumber)") {
        $modelNameShort = if ($cs.Model -like "* Notebook PC") { $cs.Model.Replace(" Notebook PC", "") } else { $cs.Model }
    } else {
        $modelNameShort = $modelname
    }

    # Calculate the disk size in GB
    $diskSizeGB = [math]::Round($disk.Size / 1GB, 2)

    # Extract the required information and shorten processor name and model
    $processorNameShort = Get-ShortProcessorName -processorName $processor.Name

    $info = @(
        "$($os.Caption)"
        "$($modelNameShort)"
        "$($cs.Name)"
        "$($cs.SystemSKUNumber)"
        "SN: $($SerialNumber)"
        "$($processorNameShort)" +" | "+ "$($cs.TotalPhysicalMemory / 1GB -as [int]) GB" +" | "+ "${diskSizeGB} GB"
    )

    # Convert the information array to a single string
    $infoText = $info -join "`n"

    # Optional: Display the information to verify
    $info

} catch {
    Write-Host "An error occurred: $_"
}


# Function to combine images with barcode over QR code and add text next to QR code
function Combine-ImagesWithText {
    param (
        [System.Drawing.Bitmap]$qrCodeImage,
        [System.Drawing.Bitmap]$barcodeImage,
        [string]$text,
        [string]$outputFilePath
    )

    # Calculate the combined width and height
    $dpi = 300
    $labelWidthMM = 50
    $labelHeightMM = 30
    $combinedWidth = [math]::Round($labelWidthMM / 25.4 * $dpi)
    $combinedHeight = [math]::Round($labelHeightMM / 25.4 * $dpi)

    # Create a new bitmap for the combined image
    $combinedBitmap = New-Object System.Drawing.Bitmap($combinedWidth, $combinedHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($combinedBitmap)
    $graphics.Clear([System.Drawing.Color]::White) # Set background to white

    # Calculate positions for the images and text
    $qrCodeX = 40
    $qrCodeY = $combinedHeight - $qrCodeImage.Height - 40
    $barcodeX = 60
    $barcodeY = 20
    $textX = $qrCodeX + $qrCodeImage.Width - 20
    $textY = $qrCodeY + 10

    # Draw the QR code and barcode images
    $graphics.DrawImage($qrCodeImage, $qrCodeX, $qrCodeY)
    $graphics.DrawImage($barcodeImage, $barcodeX, $barcodeY)

    # Adding text next to the QR code
    $font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Regular)
    $brush = [System.Drawing.Brushes]::Black
    $graphics.DrawString($text, $font, $brush, $textX, $textY)
    
    # Rotate the combined bitmap
    # $combinedBitmap.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone)

    # Save the combined bitmap
    $combinedBitmap.SetResolution($dpi, $dpi)
    $combinedBitmap.Save($outputFilePath, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $combinedBitmap.Dispose()
}

# Ensure the directory exists before creating the combined file
$combinedFilePath = Join-Path -Path "C:\" -ChildPath "logs\combined.png"
$directoryPath = [System.IO.Path]::GetDirectoryName($combinedFilePath)
if (-Not (Test-Path -Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath
}

# Combine images with barcode over QR code and add text next to the QR code
Combine-ImagesWithText -qrCodeImage $qrCodeBitmap -barcodeImage $barcodeBitmap -text $infoText -outputFilePath $combinedFilePath
Write-Output "Combined image saved to $combinedFilePath"

# Clean up bitmaps and delete temporary files
$qrCodeBitmap.Dispose()
$barcodeBitmap.Dispose()
Remove-Item -Path $qrCodeFilePath -Force
Remove-Item -Path $barcodeFilePath -Force


# Define the printer name and file path
$printerName = "M221 Printer"
$combinedFilePath = "C:\logs\combined.png"
$customPaperName = "Custom50x30mm"
$paperWidthMM = 5000  # 50mm in hundredths of a millimeter
$paperHeightMM = 3000  # 30mm in hundredths of a millimeter

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Function to create a custom paper size using the registry
function New-CustomPaperSize {
    param (
        [string]$printerName,
        [string]$paperName,
        [int]$paperWidth,  # Width in hundredths of a millimeter
        [int]$paperHeight  # Height in hundredths of a millimeter
    )

    $keyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Forms\$paperName"
    
    # Convert dimensions to hundredths of an inch
    $widthInHundredthsOfInch = [math]::Round($paperWidth / 25.4 * 100)
    $heightInHundredthsOfInch = [math]::Round($paperHeight / 25.4 * 100)

    # Convert the dimensions to byte arrays (little-endian format)
    $widthBytes = [BitConverter]::GetBytes([int]$widthInHundredthsOfInch)
    $heightBytes = [BitConverter]::GetBytes([int]$heightInHundredthsOfInch)

    # Combine width and height bytes into a flat array
    $sizeBytes = New-Object System.Collections.ArrayList
    $sizeBytes.AddRange($widthBytes)
    $sizeBytes.AddRange($heightBytes)
    $sizeBytes = $sizeBytes.ToArray([type]::GetType("System.Byte"))

    New-Item -Path $keyPath -Force | Out-Null
    Set-ItemProperty -Path $keyPath -Name "Name" -Value $paperName
    Set-ItemProperty -Path $keyPath -Name "Size" -Value $sizeBytes
    Set-ItemProperty -Path $keyPath -Name "PrintableArea" -Value $sizeBytes
    Set-ItemProperty -Path $keyPath -Name "Width" -Value $widthInHundredthsOfInch
    Set-ItemProperty -Path $keyPath -Name "Height" -Value $heightInHundredthsOfInch
}

# Function to set the printer's default paper size
function Set-DefaultPaperSize {
    param (
        [string]$printerName,
        [string]$paperName
    )

    $printerSettingsPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts\$printerName"
    $paperSizeRegValue = Get-ItemProperty -Path $printerSettingsPath -Name "Paper Size" -ErrorAction SilentlyContinue

    if ($paperSizeRegValue) {
        Set-ItemProperty -Path $printerSettingsPath -Name "Paper Size" -Value $paperName
    } else {
        New-ItemProperty -Path $printerSettingsPath -Name "Paper Size" -Value $paperName -PropertyType String
    }
}

# Function to print the image using PowerShell
function Print-FullPageImage {
    param (
        [string]$filePath,
        [string]$printerName,
        [string]$paperName,
        [int]$paperWidthMM,  # Width in millimeters
        [int]$paperHeightMM  # Height in millimeters
    )

    $printDoc = New-Object System.Drawing.Printing.PrintDocument
    $printDoc.PrinterSettings.PrinterName = $printerName
    $printDoc.DefaultPageSettings.PaperSize = New-Object System.Drawing.Printing.PaperSize($paperName, [math]::Round($paperWidthMM / 25.4 * 100), [math]::Round($paperHeightMM / 25.4 * 100))

    $printPageEventHandler = [System.Drawing.Printing.PrintPageEventHandler]{
        param (
            [object]$sender,
            [System.Drawing.Printing.PrintPageEventArgs]$e
        )

        $image = [System.Drawing.Image]::FromFile($filePath)

        # Calculate the rectangle to fit the image
        $rect = [System.Drawing.RectangleF]::Empty
        $rect.Width = $e.PageBounds.Width
        $rect.Height = $e.PageBounds.Height

        # Ensure the image fits the rectangle maintaining aspect ratio
        $scale = [math]::Min($rect.Width / $image.Width, $rect.Height / $image.Height)
        $widthScaled = $image.Width * $scale
        $heightScaled = $image.Height * $scale

        # Center the image
        $rect.X = ($rect.Width - $widthScaled) / 2
        $rect.Y = ($rect.Height - $heightScaled) / 2
        $rect.Width = $widthScaled
        $rect.Height = $heightScaled

        $e.Graphics.DrawImage($image, $rect)
        $e.HasMorePages = $false
    }

    $printDoc.add_PrintPage($printPageEventHandler)
    $printDoc.Print()
}

# Define the printer name and file path
$printerName = "M221 Printer"
$combinedFilePath = "C:\logs\combined.png"
$customPaperName = "Custom50x30mm"
$paperWidthMM = 50  # 50mm
$paperHeightMM = 30  # 30mm

# Create the custom paper size
New-CustomPaperSize -printerName $printerName -paperName $customPaperName -paperWidth $paperWidthMM * 100 -paperHeight $paperHeightMM * 100

# Set the default paper size for the printer
Set-DefaultPaperSize -printerName $printerName -paperName $customPaperName

# Print the full-page image
Print-FullPageImage -filePath $combinedFilePath -printerName $printerName -paperName $customPaperName -paperWidthMM $paperWidthMM -paperHeightMM $paperHeightMM
