#this will take files from a specific period which haven't been accessed for, 
# it then saves it to an excel file, run Deletefiles.ps1 to delete the files


#update 20XX to appropriate year 

Import-Module ImportExcel

$folder = "C:\drive path"
$OutputFile = "C:\file.xlsx"

if (Test-Path -Path $folder) {
    Write-Output "Listing files in $folder that were last accessed between 2000 and 20XX"

    $results = @()

    # Get all files last accessed between 2000 and 20XX
    $filesBetween2000And20XX = Get-ChildItem -Path $folder -Recurse -File | Where-Object { 
        $_.LastAccessTime -ge (Get-Date -Year 2000) -and $_.LastAccessTime -le (Get-Date -Year #20XX) 
    }

    # List all files that match the criteria
    $results += $filesBetween2000And20XX | ForEach-Object {
        [PSCustomObject]@{
            "Path"          = $_.FullName
            "LastAccessed"  = $_.LastAccessTime
        }
    }

    # Sort the results by LastAccessed in descending order
    $sortedResults = $results | Sort-Object -Property LastAccessed -Descending

    # Export the sorted results to Excel
    $sortedResults | Export-Excel -Path $OutputFile -AutoSize

    Write-Output "Output written to $OutputFile"
} else {
    Write-Output "The folder $folder does not exist."
}