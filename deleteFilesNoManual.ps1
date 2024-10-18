

#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process


Import-Module ImportExcel

# Path to the Excel file generated previously
$excelFile = "c:\..."
$sheetName = "Sheet1"  # Name of the worksheet, adjust if needed

# Function to delete files or skip them
function DeleteFiles {
    param (
        [string]$filePath,
        [DateTime]$lastAccessTime
    )

    # Check if the file exists
    if (Test-Path -Path $filePath) {
        Write-Host "`nFile: $filePath"
        Write-Host "Last Accessed: $lastAccessTime"
        $response = Read-Host "Do you want to delete this file? (Y/N/ShowDate/ShowPath/Skip/Specific)"

        if ($response -eq 'Y') {
            $finalConfirmation = Read-Host "Are you sure you want to delete this file? (Y/N)"
            if ($finalConfirmation -eq 'Y') {
                Remove-Item -Path $filePath -Force
                Write-Host "Deleted: $filePath"
            } else {
                Write-Host "Skipped: $filePath"
            }
        } elseif ($response -eq 'ShowDate') {
            Write-Host "Last Accessed Date: $lastAccessTime"
        } elseif ($response -eq 'ShowPath') {
            Write-Host "File Path: $filePath"
        } elseif ($response -eq 'Specific') {
            ManualSpecificMode
        } else {
            Write-Host "Skipped: $filePath"
        }
    } else {
        Write-Host "`nFile not found: $filePath"
    }
}

# Function to find the row index of a file by its path or access time
function FindFileIndex {
    param (
        [array]$data,
        [string]$searchPath,
        [DateTime]$searchAccessTime
    )

    # If searching by path
    if ($searchPath) {
        for ($i = 0; $i -lt $data.Count; $i++) {
            if ($data[$i].Path -eq $searchPath) {
                return $i
            }
        }
    }

    # If searching by access time
    if ($searchAccessTime) {
        for ($i = 0; $i -lt $data.Count; $i++) {
            if ($data[$i].LastAccessed.Date -eq $searchAccessTime.Date -and $data[$i].LastAccessed.ToString("HH:mm") -eq $searchAccessTime.ToString("HH:mm")) {
                return $i
            }
        }
    }

    # Return -1 if not found
    return -1
}

# Function to delete files by wildcard or specific date
function DeleteFilesByPattern {
    param (
        [array]$data,
        [DateTime]$searchAccessTime,
        [string]$wildcard
    )

    # Iterate over all files and check if they match the date and wildcard time
    foreach ($row in $data) {
        $filePath = $row.Path
        $lastAccessTime = $row.LastAccessed

        # Match files for a specific date or date and time with wildcard
        if ($lastAccessTime.Date -eq $searchAccessTime.Date) {
            if ($wildcard -eq '*') {
                DeleteFiles -filePath $filePath -lastAccessTime $lastAccessTime
            } elseif ($wildcard -match '^\d{2}:\d{2} \*$') {
                $patternTime = [DateTime]::ParseExact($wildcard.TrimEnd(' *'), "HH:mm", $null)
                if ($lastAccessTime.ToString("HH:mm") -eq $patternTime.ToString("HH:mm")) {
                    DeleteFiles -filePath $filePath -lastAccessTime $lastAccessTime
                }
            }
        }
    }
}

# Function to manually specify a file or access time
function ManualSpecificMode {
    $manualInput = Read-Host "Enter a specific file path or last accessed time (yyyy-MM-dd HH:mm)"
    $wildcard = $null
    if ($manualInput -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2} \*$') {
        $dateString = $manualInput.Split(' ')[0]
        $wildcard = $manualInput.Split(' ')[1] + ' ' + $manualInput.Split(' ')[2]
        $searchAccessTime = [DateTime]::ParseExact($dateString, "yyyy-MM-dd", $null)
        DeleteFilesByPattern -data $data -searchAccessTime $searchAccessTime -wildcard $wildcard
    } elseif ($manualInput -match '^\d{4}-\d{2}-\d{2} \*$') {
        $searchAccessTime = [DateTime]::ParseExact($manualInput.TrimEnd(' *'), "yyyy-MM-dd", $null)
        DeleteFilesByPattern -data $data -searchAccessTime $searchAccessTime -wildcard '*'
    } else {
        $searchAccessTime = [DateTime]::ParseExact($manualInput, "yyyy-MM-dd HH:mm", $null)
        $index = FindFileIndex -data $data -searchAccessTime $searchAccessTime
        if ($index -ne -1) {
            for ($i = $index; $i -lt $data.Count; $i++) {
                $filePath = $data[$i].Path
                $lastAccessTime = $data[$i].LastAccessed

                try {
                    DeleteFiles -filePath $filePath -lastAccessTime $lastAccessTime
                } catch {
                    Write-Host "`nAn error occurred: $_"
                    break
                }
            }
        } else {
            Write-Host "File not found for the specified time."
        }
    }
}

# Check if the Excel file exists
if (Test-Path -Path $excelFile) {
    # Read the Excel file
    $data = Import-Excel -Path $excelFile -WorksheetName $sheetName

    # Ask the user if they want to skip to a specific file
    $skipToFile = Read-Host "Do you want to skip to a specific file by path or access time? (Y/N)"

    $startIndex = 0

    if ($skipToFile -eq 'Y') {
        $searchType = Read-Host "Search by FilePath or AccessTime? (Path/Time)"
        if ($searchType -eq 'Path') {
            $searchPath = Read-Host "Enter the file path to search for"
            $startIndex = FindFileIndex -data $data -searchPath $searchPath
        } elseif ($searchType -eq 'Time') {
            $searchAccessTimeInput = Read-Host "Enter the last accessed date and time to search for (yyyy-MM-dd HH:mm)"
            $searchAccessTime = [DateTime]::ParseExact($searchAccessTimeInput, "yyyy-MM-dd HH:mm", $null)
            $startIndex = FindFileIndex -data $data -searchAccessTime $searchAccessTime
        }

        if ($startIndex -eq -1) {
            Write-Host "File not found. Starting from the beginning of the list."
            $startIndex = 0
        } else {
            Write-Host "Starting from the specified file at index $startIndex."
        }
    }

    for ($i = $startIndex; $i -lt $data.Count; $i++) {
        $filePath = $data[$i].Path
        $lastAccessTime = $data[$i].LastAccessed

        try {
            DeleteFiles -filePath $filePath -lastAccessTime $lastAccessTime
        } catch {
            Write-Host "`nAn error occurred: $_"
            break
        }
    }
} else {
    Write-Host "The Excel file $excelFile does not exist."
}
