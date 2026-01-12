# Video Frame Grid Extractor
# This script extracts 10 frames from each video file and creates a 2x5 grid image

# Configuration
$videoFolder = Get-Location  # Uses the current directory where you run the script
$outputFolder = "$videoFolder\Grids"      # Output folder for grid images
$videoExtensions = @("*.mp4", "*.avi", "*.mkv", "*.mov", "*.wmv", "*.flv", "*.webm", "*.m4v", "*.mpg", "*.mpeg", "*.3gp", "*.ts", "*.mts", "*.m2ts")

# Create output folder if it doesn't exist
if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
    Write-Host "Created output folder: $outputFolder" -ForegroundColor Green
}

# Get all video files
Write-Host "Looking for video files in: $videoFolder" -ForegroundColor Cyan
$videoFiles = Get-ChildItem -Path "$videoFolder\*" -Include $videoExtensions -File

if ($videoFiles.Count -eq 0) {
    Write-Host "No video files found in $videoFolder" -ForegroundColor Red
    exit
}

Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Cyan

foreach ($video in $videoFiles) {
    Write-Host "`nProcessing: $($video.Name)" -ForegroundColor Yellow
    
    # Get video duration using ffprobe
    $durationCmd = "ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 `"$($video.FullName)`""
    $duration = Invoke-Expression $durationCmd
    $duration = [double]$duration
    
    if ($duration -le 5) {
        Write-Host "  Video too short (less than 5 seconds), skipping..." -ForegroundColor Red
        continue
    }
    
    # Calculate time points (starting from 5 seconds, ending 15 seconds before end)
    $startTime = 5
    $endTime = $duration - 15
    
    if ($endTime -le $startTime) {
        Write-Host "  Video too short (need at least 20 seconds), skipping..." -ForegroundColor Red
        continue
    }
    
    $interval = ($endTime - $startTime) / 8  # 8 intervals for 9 frames
    
    # Create temporary folder for individual frames
    $tempFolder = "$outputFolder\temp_$([guid]::NewGuid())"
    New-Item -ItemType Directory -Path $tempFolder | Out-Null
    
    # Extract 9 frames
    Write-Host "  Extracting frames..." -ForegroundColor Gray
    for ($i = 0; $i -lt 9; $i++) {
        $timestamp = $startTime + ($i * $interval)
        $frameFile = "$tempFolder\frame_$($i.ToString('00')).jpg"
        
        $ffmpegCmd = "ffmpeg -ss $timestamp -i `"$($video.FullName)`" -frames:v 1 -vf `"scale=267:-1`" -q:v 5 `"$frameFile`" -y -loglevel quiet"
        Invoke-Expression $ffmpegCmd
    }
    
    # Create 3x3 grid using ffmpeg
    $outputFile = "$outputFolder\$($video.BaseName)_grid.jpg"
    Write-Host "  Creating grid image..." -ForegroundColor Gray
    
    $gridCmd = "ffmpeg -i `"$tempFolder\frame_%02d.jpg`" -filter_complex `"tile=3x3`" -q:v 5 `"$outputFile`" -y -loglevel quiet"
    Invoke-Expression $gridCmd
    
    # Clean up temporary files
    Remove-Item -Path $tempFolder -Recurse -Force
    
    Write-Host "  Created: $($video.BaseName)_grid.jpg" -ForegroundColor Green
}

Write-Host "`nAll done! Grid images saved to: $outputFolder" -ForegroundColor Cyan