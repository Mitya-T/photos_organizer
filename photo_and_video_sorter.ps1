# Photo and Video Organizer Script
# Organizes photos and videos into Year\Month folders using metadata

param(
    [string]$SourceFolder = (Read-Host "Enter the folder path containing your photos and videos"),
    [switch]$DryRun = $false
)

# Common image and video extensions
$imageExtensions = @('*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.tiff', '*.heic', '*.raw', '*.cr2', '*.nef')
$videoExtensions = @('*.mp4', '*.mov', '*.avi', '*.mkv', '*.wmv', '*.flv', '*.m4v', '*.mpg', '*.mpeg', '*.3gp', '*.webm')
$allExtensions = $imageExtensions + $videoExtensions

# Validate source folder
if (-not (Test-Path $SourceFolder)) {
    Write-Host "Error: Folder does not exist!" -ForegroundColor Red
    exit
}

if ($DryRun) {
    Write-Host "`n*** DRY RUN MODE - No files will be moved ***`n" -ForegroundColor Yellow
}

Write-Host "`nStarting photo and video organization..." -ForegroundColor Green
Write-Host "Source folder: $SourceFolder`n" -ForegroundColor Cyan

$filesProcessed = 0
$filesMoved = 0
$filesWithMetadata = 0
$filesWithoutMetadata = 0
$filesSkipped = 0

# Function to get EXIF date from image
function Get-ExifDate {
    param([string]$FilePath)
    
    try {
        Add-Type -AssemblyName System.Drawing
        $img = [System.Drawing.Image]::FromFile($FilePath)
        
        # Property ID 36867 is DateTimeOriginal
        $propId = 36867
        
        if ($img.PropertyIdList -contains $propId) {
            $dateTakenBytes = $img.GetPropertyItem($propId).Value
            $dateTakenString = [System.Text.Encoding]::ASCII.GetString($dateTakenBytes)
            $dateTakenString = $dateTakenString.Trim([char]0)
            
            if ($dateTakenString -match '(\d{4}):(\d{2}):(\d{2}) (\d{2}):(\d{2}):(\d{2})') {
                $dateStr = "$($matches[1])-$($matches[2])-$($matches[3]) $($matches[4]):$($matches[5]):$($matches[6])"
                $img.Dispose()
                return [DateTime]::Parse($dateStr)
            }
        }
        
        $img.Dispose()
        return $null
    } catch {
        return $null
    }
}

# Function to get video metadata
function Get-VideoDate {
    param([string]$FilePath)
    
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace([System.IO.Path]::GetDirectoryName($FilePath))
        $file = $folder.ParseName([System.IO.Path]::GetFileName($FilePath))
        
        foreach ($propIndex in @(208, 12)) {
            $dateValue = $folder.GetDetailsOf($file, $propIndex)
            if ($dateValue -and $dateValue -ne "") {
                $parsedDate = $null
                if ([DateTime]::TryParse($dateValue, [ref]$parsedDate)) {
                    if ($parsedDate -le (Get-Date) -and $parsedDate -ge (Get-Date "1990-01-01")) {
                        return $parsedDate
                    }
                }
            }
        }
        
        return $null
    } catch {
        return $null
    }
}

# Function to get media date
function Get-MediaDate {
    param([System.IO.FileInfo]$File)
    
    $extension = $File.Extension.ToLower()
    
    # Try EXIF data for images
    if ($extension -in @('.jpg', '.jpeg', '.tiff', '.tif', '.heic', '.raw', '.cr2', '.nef')) {
        $exifDate = Get-ExifDate -FilePath $File.FullName
        if ($exifDate) {
            return @{Date = $exifDate; Source = "EXIF"}
        }
    }
    
    # Try video metadata
    if ($extension -in @('.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.m4v', '.mpg', '.mpeg', '.3gp', '.webm')) {
        $videoDate = Get-VideoDate -FilePath $File.FullName
        if ($videoDate) {
            return @{Date = $videoDate; Source = "VideoMetadata"}
        }
    }
    
    # Use oldest date
    $oldestDate = $File.LastWriteTime
    $dateSource = "LastWriteTime"
    
    if ($File.CreationTime -lt $oldestDate) {
        $oldestDate = $File.CreationTime
        $dateSource = "CreationTime"
    }
    
    $today = Get-Date
    if ($oldestDate.Date -eq $today.Date) {
        Write-Host "  WARNING: $($File.Name) has todays date - metadata may be missing!" -ForegroundColor Yellow
    }
    
    return @{Date = $oldestDate; Source = $dateSource}
}

# Get all files
$allFiles = @()
foreach ($ext in $allExtensions) {
    $files = Get-ChildItem -Path $SourceFolder -Filter $ext -File -Recurse:$false
    foreach ($f in $files) {
        # Only add if not already in list (prevent duplicates)
        if ($allFiles.FullName -notcontains $f.FullName) {
            $allFiles += $f
        }
    }
}

Write-Host "Found $($allFiles.Count) unique files to process`n" -ForegroundColor Cyan

if ($allFiles.Count -eq 0) {
    Write-Host "No media files found in: $SourceFolder" -ForegroundColor Yellow
    exit
}

foreach ($file in $allFiles) {
    $filesProcessed++
    
    Write-Host "`nProcessing: $($file.Name)" -ForegroundColor Cyan
    Write-Host "  Full path: $($file.FullName)" -ForegroundColor Gray
    
    $dateInfo = Get-MediaDate -File $file
    $date = $dateInfo.Date
    $source = $dateInfo.Source
    
    Write-Host "  Date found: $($date.ToString('yyyy-MM-dd HH:mm:ss')) from $source" -ForegroundColor Gray
    
    if ($source -in @("EXIF", "VideoMetadata")) {
        $filesWithMetadata++
    } else {
        $filesWithoutMetadata++
    }
    
    $year = $date.Year
    $monthNum = $date.Month.ToString("00")
    $monthName = $date.ToString("MMM").ToUpper()
    $monthFolder = "${monthNum}_${monthName}"
    
    $destPath = Join-Path $SourceFolder $year
    $destPath = Join-Path $destPath $monthFolder
    
    Write-Host "  Destination: $destPath" -ForegroundColor Gray
    
    if (-not $DryRun -and -not (Test-Path $destPath)) {
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null
        Write-Host "  Created folder: $destPath" -ForegroundColor Yellow
    }
    
    $destFile = Join-Path $destPath $file.Name
    Write-Host "  Destination file: $destFile" -ForegroundColor Gray
    Write-Host "  Source == Dest? $($file.FullName -eq $destFile)" -ForegroundColor Gray
    
    if ($file.FullName -ne $destFile) {
        $action = if ($DryRun) { "Would move" } else { "MOVING" }
        Write-Host "  $action file..." -ForegroundColor Green
        
        if (-not $DryRun) {
            try {
                # Use literal path to handle special characters and spaces
                $sourcePath = $file.FullName
                
                # Check if file exists before moving
                if (-not (Test-Path -LiteralPath $sourcePath)) {
                    Write-Host "  ERROR: Source file no longer exists!" -ForegroundColor Red
                    Write-Host "  Attempted path: $sourcePath" -ForegroundColor Red
                    continue
                }
                
                # Ensure destination directory exists
                if (-not (Test-Path -LiteralPath $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
                
                Move-Item -LiteralPath $sourcePath -Destination $destFile -Force -ErrorAction Stop
                
                # Verify the move worked
                if (Test-Path -LiteralPath $destFile) {
                    $filesMoved++
                    Write-Host "  Successfully moved and verified!" -ForegroundColor Green
                } else {
                    Write-Host "  ERROR: Move reported success but file not found at destination!" -ForegroundColor Red
                }
                
                # Check if still at source
                if (Test-Path -LiteralPath $sourcePath) {
                    Write-Host "  ERROR: File still exists at source after move!" -ForegroundColor Red
                }
            } catch {
                Write-Host "  ERROR: $_" -ForegroundColor Red
                Write-Host "  Error details: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "  File already in correct location - skipping" -ForegroundColor Yellow
        $filesSkipped++
    }
}

Write-Host "`n================================================" -ForegroundColor Green
Write-Host "Organization complete!" -ForegroundColor Green
Write-Host "Files processed: $filesProcessed" -ForegroundColor Cyan
Write-Host "Files moved: $filesMoved" -ForegroundColor Cyan
Write-Host "Files skipped (already in place): $filesSkipped" -ForegroundColor Cyan
Write-Host "Files with metadata: $filesWithMetadata" -ForegroundColor Yellow
Write-Host "Files using file dates: $filesWithoutMetadata" -ForegroundColor Yellow
if ($DryRun) {
    Write-Host "`nThis was a DRY RUN - no files were actually moved" -ForegroundColor Yellow
}
Write-Host "================================================`n" -ForegroundColor Green