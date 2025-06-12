# Part 1
# Define paths
$Folder = "C:\NetworkTitan\WICleanup"
$ScriptPath = "$Folder\WICleanup.vbs"
$LogPath = "$Folder\WICleanup-Log.txt"

# Create folder
if (-Not (Test-Path $Folder)) {
    New-Item -ItemType Directory -Path $Folder -Force | Out-Null
}

# Write WICleanup.vbs content directly to file
$VbsContent = @'
' WICleanup.vbs - Removes orphaned files from Windows Installer folder
Option Explicit

Dim installer, products, product, patches, patch, fso, folder, file, usedFiles, orphanedFiles
Set installer = CreateObject("WindowsInstaller.Installer")
Set fso = CreateObject("Scripting.FileSystemObject")
Set usedFiles = CreateObject("Scripting.Dictionary")
Set orphanedFiles = CreateObject("Scripting.Dictionary")

Function AddUsedFile(path)
    If Not usedFiles.Exists(LCase(path)) Then
        usedFiles.Add LCase(path), True
    End If
End Function

For Each product In installer.Products
    For Each patch In installer.Patches(product)
        AddUsedFile installer.PatchInfo(patch, "LocalPackage")
    Next
    AddUsedFile installer.ProductInfo(product, "LocalPackage")
Next

Set folder = fso.GetFolder("C:\Windows\Installer")
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "msi" Or LCase(fso.GetExtensionName(file.Name)) = "msp" Then
        If Not usedFiles.Exists(LCase(file.Path)) Then
            orphanedFiles.Add file.Path, True
        End If
    End If
Next

Dim arg, quietMode, safeMode
quietMode = False
safeMode = False

For Each arg In WScript.Arguments
    If arg = "/Q" Then quietMode = True
    If arg = "/S" Then safeMode = True
Next

If orphanedFiles.Count = 0 Then
    WScript.Echo "No orphaned files found."
Else
    For Each file In orphanedFiles.Keys
        If quietMode Then
            WScript.Echo "Orphaned: " & file
        ElseIf safeMode Then
            On Error Resume Next
            fso.DeleteFile file, True
            WScript.Echo "Deleted: " & file
        End If
    Next
End If
'@

# Save the .vbs to file
Set-Content -Path $ScriptPath -Value $VbsContent -Force
Write-Output "✅ WICleanup.vbs script written to $ScriptPath"

# Run the script in SAFE MODE to clean orphaned installers
try {
    & cscript.exe //nologo "$ScriptPath" /S > "$LogPath"
    Write-Output "✅ Cleanup complete. Log saved to $LogPath"
} catch {
    Write-Output "❌ Cleanup failed: $($_.Exception.Message)"
}

# Output tail of the log for live feedback
if (Test-Path $LogPath) {
    Write-Output "`n=== Cleanup Results ==="
    Get-Content $LogPath -Tail 10
}

# Part 2
$ScriptPath = "C:\NetworkTitan\WICleanup\WICleanup.vbs"

if (Test-Path $ScriptPath) {
    Write-Output "=== Running in PREVIEW mode ==="
    & cscript.exe //nologo "$ScriptPath" /Q
} else {
    Write-Output "❌ Script not found at $ScriptPath"
}


# Part 3. 
# === Set Paths ===
$Folder = "C:\NetworkTitan\WICleanup"
$ScriptPath = "$Folder\WICleanup.vbs"
$LogPath = "$Folder\WICleanup-Log.txt"

# Create folder if needed
if (-Not (Test-Path $Folder)) {
    New-Item -ItemType Directory -Path $Folder -Force | Out-Null
}

# Updated script content with error handling
$VbsContent = @'
Option Explicit

Dim installer, products, product, patches, patch, fso, folder, file, usedFiles, orphanedFiles
Set installer = CreateObject("WindowsInstaller.Installer")
Set fso = CreateObject("Scripting.FileSystemObject")
Set usedFiles = CreateObject("Scripting.Dictionary")
Set orphanedFiles = CreateObject("Scripting.Dictionary")

Function AddUsedFile(path)
    If Not usedFiles.Exists(LCase(path)) Then
        usedFiles.Add LCase(path), True
    End If
End Function

On Error Resume Next
For Each product In installer.Products
    Err.Clear
    For Each patch In installer.Patches(product)
        AddUsedFile installer.PatchInfo(patch, "LocalPackage")
    Next
    If Err.Number = 0 Then
        AddUsedFile installer.ProductInfo(product, "LocalPackage")
    End If
Next
On Error GoTo 0

Set folder = fso.GetFolder("C:\Windows\Installer")
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "msi" Or LCase(fso.GetExtensionName(file.Name)) = "msp" Then
        If Not usedFiles.Exists(LCase(file.Path)) Then
            orphanedFiles.Add file.Path, True
        End If
    End If
Next

Dim arg, quietMode, safeMode, logFile
quietMode = False
safeMode = False
Set logFile = fso.OpenTextFile("C:\NetworkTitan\WICleanup\WICleanup-Log.txt", 2, True)

For Each arg In WScript.Arguments
    If arg = "/Q" Then quietMode = True
    If arg = "/S" Then safeMode = True
Next

If orphanedFiles.Count = 0 Then
    WScript.Echo "No orphaned files found."
    logFile.WriteLine "No orphaned files found."
Else
    For Each file In orphanedFiles.Keys
        If quietMode Then
            WScript.Echo "Orphaned: " & file
            logFile.WriteLine "Orphaned: " & file
        ElseIf safeMode Then
            On Error Resume Next
            fso.DeleteFile file, True
            WScript.Echo "Deleted: " & file
            logFile.WriteLine "Deleted: " & file
        End If
    Next
End If

logFile.Close
'@

# Save updated script
Set-Content -Path $ScriptPath -Value $VbsContent -Force
Write-Output "✅ Updated WICleanup.vbs written to $ScriptPath"

# Run preview (no delete)
try {
    Write-Output "=== Running in Preview Mode ==="
    & cscript.exe //nologo "$ScriptPath" /Q
} catch {
    Write-Output "❌ Error running script: $($_.Exception.Message)"
}

# Last Step Delete 
cscript.exe //nologo "C:\NetworkTitan\WICleanup\WICleanup.vbs" /S
