$targetPath = "C:\Users\PAth\AppData\Roaming"  # Automatically sets to user's roaming folder
#Put any folder path into the path to calculate.
# Get top-level folders and calculate sizes
Get-ChildItem -Path $targetPath -Directory | ForEach-Object {
    $folder = $_.FullName
    $size = (Get-ChildItem -Path $folder -Recurse -ErrorAction SilentlyContinue | 
             Measure-Object -Property Length -Sum).Sum
    [PSCustomObject]@{
        Folder = $_.Name
        SizeGB = '{0:N2}' -f ($size / 1GB)
    }
} | Sort-Object -Property SizeGB -Descending | Format-Table -AutoSize
