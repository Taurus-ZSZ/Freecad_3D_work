<#
.SYNOPSIS
    FreeCAD Weekly Build Auto Update Script
.DESCRIPTION
    Automatically updates FreeCAD weekly builds
.NOTES
    Author: AI Assistant
    Version: 1.6 - Proper nested structure repair
#>

# Error handling settings
$ErrorActionPreference = "Stop"

# Constants
$DownloadPath = [System.Environment]::GetFolderPath("UserProfile") + "\Downloads"
$TargetPath = "D:\Program Files"
$ShortcutName = "FreeCAD Dev.lnk"
$Pattern = "FreeCAD_weekly-*-Windows-x86_64-py311.7z"

# Output functions
function Write-Info { Write-Host "[INFO] $($args[0])" -ForegroundColor Cyan }
function Write-Success { Write-Host "[SUCCESS] $($args[0])" -ForegroundColor Green }
function Write-Warning { Write-Host "[WARNING] $($args[0])" -ForegroundColor Yellow }
function Write-ErrorMsg { Write-Host "[ERROR] $($args[0])" -ForegroundColor Red }

# 1. Enhanced 7-Zip detection
function Test-7Zip {
    Write-Info "Searching for 7-Zip installation..."
    
    # Method 1: Check PATH environment variable
    Write-Info "Checking PATH environment variable..."
    $7zInPath = Get-Command "7z" -ErrorAction SilentlyContinue
    if ($7zInPath) {
        Write-Success "Found 7-Zip in PATH: $($7zInPath.Source)"
        return $7zInPath.Source
    }
    
    # Method 2: Check common installation directories
    Write-Info "Checking common installation directories..."
    $commonPaths = @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\7z.exe",
        "$env:ProgramFiles\7-Zip\7zG.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7zG.exe"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            Write-Success "Found 7-Zip: $path"
            return $path
        }
    }
    
    # Method 3: Search in all drives for 7z.exe
    Write-Info "Searching all drives for 7z.exe..."
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Ready }
    
    foreach ($drive in $drives) {
        $rootPath = $drive.Root
        Write-Info "Searching in $rootPath..."
        
        try {
            $searchPaths = @(
                "$rootPath\Program Files\*",
                "$rootPath\Program Files (x86)\*",
                "$rootPath\*"
            )
            
            foreach ($searchPath in $searchPaths) {
                $found = Get-ChildItem -Path $searchPath -Filter "7z.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($found) {
                    Write-Success "Found 7-Zip: $($found.FullName)"
                    return $found.FullName
                }
            }
        }
        catch {
            continue
        }
    }
    
    # Method 4: Check registry
    Write-Info "Checking Windows Registry for 7-Zip..."
    $registryPaths = @(
        "HKLM:\SOFTWARE\7-Zip",
        "HKLM:\SOFTWARE\WOW6432Node\7-Zip",
        "HKCU:\SOFTWARE\7-Zip"
    )
    
    foreach ($regPath in $registryPaths) {
        if (Test-Path $regPath) {
            try {
                $regValue = Get-ItemProperty -Path $regPath -Name "Path" -ErrorAction SilentlyContinue
                if ($regValue -and $regValue.Path) {
                    $7zPath = Join-Path $regValue.Path "7z.exe"
                    if (Test-Path $7zPath) {
                        Write-Success "Found 7-Zip via registry: $7zPath"
                        return $7zPath
                    }
                }
            }
            catch {
            }
        }
    }
    
    # If we get here, 7-Zip was not found
    Write-ErrorMsg "7-Zip was not found automatically."
    Write-Host "`nPlease install 7-Zip from https://www.7-zip.org/ or specify the path manually." -ForegroundColor Yellow
    
    # Ask user if they want to manually specify the path
    $manualPath = Read-Host "`nIf you know the path to 7z.exe, enter it now (or press Enter to exit)"
    if ($manualPath -and (Test-Path $manualPath)) {
        Write-Success "Using manually specified 7-Zip: $manualPath"
        return $manualPath
    }
    
    throw "7-Zip is required but was not found. Please install 7-Zip and try again."
}

# 2. Find the latest archive
function Find-LatestArchive {
    Write-Info "Searching for archives in download directory..."
    
    if (-not (Test-Path $DownloadPath)) {
        throw "Download directory does not exist: $DownloadPath"
    }
    
    $archives = Get-ChildItem -Path $DownloadPath -Filter $Pattern | Sort-Object LastWriteTime -Descending
    
    if ($archives.Count -eq 0) {
        throw "No matching archives found in $DownloadPath. Pattern: $Pattern"
    }
    
    $latestArchive = $archives[0]
    Write-Success "Found latest archive: $($latestArchive.Name)"
    
    return $latestArchive
}

# 3. Extract date from filename
function Get-DateFromFileName {
    param([string]$FileName)
    
    $datePattern = "(\d{4}\.\d{2}\.\d{2})"
    $match = [regex]::Match($FileName, $datePattern)
    
    if ($match.Success) {
        return [DateTime]::ParseExact($match.Groups[1].Value, "yyyy.MM.dd", $null)
    }
    
    throw "Cannot extract date from filename: $FileName"
}

# 4. Extract archive using 7-Zip
function Expand-ArchiveWith7Zip {
    param(
        [string]$ArchivePath,
        [string]$Destination
    )
    
    $7zPath = Test-7Zip
    
    Write-Info "Extracting to: $Destination"
    
    # Ensure destination directory exists
    if (-not (Test-Path $Destination)) {
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    }
    
    # Build folder name (remove .7z extension)
    $folderName = [System.IO.Path]::GetFileNameWithoutExtension($ArchivePath)
    $extractPath = Join-Path $Destination $folderName
    
    Write-Info "Extraction path: $extractPath"
    
    # Remove existing folder if it exists
    if (Test-Path $extractPath) {
        Write-Warning "Target folder exists, removing: $extractPath"
        Remove-Item $extractPath -Recurse -Force
    }
    
    # Extract with 7-Zip
    $arguments = @("x", "`"$ArchivePath`"", "-o`"$extractPath`"", "-y")
    
    Write-Info "Executing: $7zPath $arguments"
    $process = Start-Process -FilePath $7zPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
    
    if ($process.ExitCode -ne 0) {
        throw "7-Zip extraction failed with exit code: $($process.ExitCode)"
    }
    
    Write-Success "Extraction completed: $extractPath"
    return $extractPath
}

# 5. Remove old versions
function Remove-OldVersions {
    param([string]$KeepFolder)
    
    Write-Info "Cleaning up old versions..."
    
    if (-not (Test-Path $TargetPath)) {
        Write-Warning "Target path does not exist, skipping cleanup"
        return
    }
    
    $freecadFolders = Get-ChildItem -Path $TargetPath -Directory | Where-Object { 
        $_.Name -like "FreeCAD_weekly-*-Windows-x86_64-py311" 
    }
    
    if ($freecadFolders.Count -eq 0) {
        Write-Warning "No FreeCAD folders found, skipping cleanup"
        return
    }
    
    # Sort by date
    $sortedFolders = $freecadFolders | Sort-Object { Get-DateFromFileName $_.Name } -Descending
    
    $removedCount = 0
    foreach ($folder in $sortedFolders) {
        if ($folder.FullName -ne $KeepFolder) {
            Write-Info "Removing old version: $($folder.Name)"
            try {
                Remove-Item $folder.FullName -Recurse -Force
                $removedCount++
            }
            catch {
                Write-Warning "Failed to remove folder $($folder.Name): $_"
            }
        }
    }
    
    if ($removedCount -gt 0) {
        Write-Success "Cleaned up $removedCount old versions"
    } else {
        Write-Info "No old versions to clean up"
    }
}

# 6. Create desktop shortcut
function New-DesktopShortcut {
    param(
        [string]$TargetExe,
        [string]$ShortcutName
    )
    
    Write-Info "Creating desktop shortcut..."
    
    $desktopPath = [System.Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath $ShortcutName
    
    # Remove existing shortcut
    if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
        Write-Info "Removed old shortcut"
    }
    
    # Create shortcut
    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $TargetExe
        $Shortcut.WorkingDirectory = Split-Path $TargetExe
        $Shortcut.IconLocation = $TargetExe
        $Shortcut.Save()
        
        Write-Success "Shortcut created: $shortcutPath"
    }
    catch {
        Write-Warning "Failed to create shortcut: $_"
        Write-Warning "This might be due to security restrictions. The script will continue."
    }
}

# 7. Get the latest FreeCAD executable path
function Get-LatestFreeCADExe {
    Write-Info "Finding latest FreeCAD version..."
    
    if (-not (Test-Path $TargetPath)) {
        throw "Target path does not exist: $TargetPath"
    }
    
    $freecadFolders = Get-ChildItem -Path $TargetPath -Directory | Where-Object { 
        $_.Name -like "FreeCAD_weekly-*-Windows-x86_64-py311" 
    }
    
    if ($freecadFolders.Count -eq 0) {
        throw "No FreeCAD folders found in $TargetPath"
    }
    
    # Get latest version by date
    $latestFolder = $freecadFolders | Sort-Object { Get-DateFromFileName $_.Name } -Descending | Select-Object -First 1
    
    Write-Success "Found latest version: $($latestFolder.Name)"
    
    # Find the executable
    $exePath = Join-Path $latestFolder.FullName "bin\freecad.exe"
    
    if (-not (Test-Path $exePath)) {
        throw "FreeCAD executable not found: $exePath"
    }
    
    Write-Success "Found FreeCAD executable: $exePath"
    return $exePath
}

# 8. Improved nested structure repair
function Repair-NestedStructure {
    param([string]$FolderPath)
    
    Write-Info "Checking for nested directory structure in: $FolderPath"
    
    $folderName = Split-Path $FolderPath -Leaf
    $nestedFolder = Join-Path $FolderPath $folderName
    
    if (-not (Test-Path $nestedFolder -PathType Container)) {
        Write-Info "No nested structure found"
        return $false
    }
    
    Write-Info "Found nested structure: $nestedFolder"
    Write-Info "Repairing nested directory structure..."
    
    # Create a temporary directory in the parent of FolderPath
    $parentPath = Split-Path $FolderPath -Parent
    $tempFolder = Join-Path $parentPath "FreeCAD_temp_$(Get-Date -Format 'yyyyMMddHHmmss')"
    
    try {
        # Create temp folder
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
        
        # Move all contents from nested folder to temp folder
        Write-Info "Moving contents from nested folder to temp location..."
        $nestedContents = Get-ChildItem -Path $nestedFolder
        foreach ($item in $nestedContents) {
            $destination = Join-Path $tempFolder $item.Name
            Move-Item -Path $item.FullName -Destination $destination -Force
        }
        
        # Remove the now-empty nested folder
        Write-Info "Removing empty nested folder..."
        Remove-Item $nestedFolder -Recurse -Force
        
        # Move all contents from temp folder back to main folder
        Write-Info "Moving contents to main folder..."
        $tempContents = Get-ChildItem -Path $tempFolder
        foreach ($item in $tempContents) {
            $destination = Join-Path $FolderPath $item.Name
            Move-Item -Path $item.FullName -Destination $destination -Force
        }
        
        # Remove temp folder
        Remove-Item $tempFolder -Recurse -Force
        
        Write-Success "Successfully repaired nested directory structure"
        Write-Success "Now using clean path: $FolderPath"
        return $true
    }
    catch {
        Write-ErrorMsg "Failed to repair nested structure: $($_.Exception.Message)"
        
        # Clean up temp folder if it exists
        if (Test-Path $tempFolder) {
            Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        throw "Cannot repair nested directory structure. Please check permissions and try again."
    }
}

# Main function
function Main {
    try {
        Write-Host "=== FreeCAD Weekly Build Auto Update Script ===" -ForegroundColor Magenta
        Write-Host "Start time: $(Get-Date)" -ForegroundColor Gray
        
        # Step 1: Check for 7-Zip
        $7zPath = Test-7Zip
        
        # Step 2: Find latest archive
        $latestArchive = Find-LatestArchive
        
        # Step 3: Extract to target location
        $extractedPath = Expand-ArchiveWith7Zip -ArchivePath $latestArchive.FullName -Destination $TargetPath
        
        # Step 3.5: Repair nested directory structure if needed
        $repaired = Repair-NestedStructure -FolderPath $extractedPath
        
        # Step 4: Clean up old versions
        Remove-OldVersions -KeepFolder $extractedPath
        
        # Step 5: Get latest executable path
        $latestExe = Get-LatestFreeCADExe
        
        # Step 6: Create desktop shortcut
        New-DesktopShortcut -TargetExe $latestExe -ShortcutName $ShortcutName
        
        Write-Host "`n=== Update Completed ===" -ForegroundColor Magenta
        Write-Success "FreeCAD successfully updated to latest version"
        Write-Success "Location: $latestExe"
        Write-Success "Shortcut: $ShortcutName"
        if ($repaired) {
            Write-Success "Nested directory structure was automatically repaired"
        }
        
    }
    catch {
        Write-ErrorMsg "Script execution failed: $($_.Exception.Message)"
        Write-Host "Full error details:" -ForegroundColor Red
        Write-Host $_.Exception.StackTrace -ForegroundColor Gray
        exit 1
    }
}

# Script entry point
if ($MyInvocation.InvocationName -ne '.') {
    Main
}