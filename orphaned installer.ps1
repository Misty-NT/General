# === Define Paths ===
$Folder = "C:\NetworkTitan\WICleanup"
$ScriptPath = "$Folder\WICleanup.vbs"
$LogPath = "$Folder\WICleanup-Log.txt"
$PreviewPath = "$Folder\Preview.txt"

# === Ensure folder exists ===
if (-Not (Test-Path $Folder)) {
    New-Item -ItemType Directory -Path $Folder -Force | Out-Null
}

# === Write VBScript for WICleanup with error-safe logic ===
$VbsContent = @'
Option Explicit

Dim installer, product, patch, fso, folder, file, usedFiles, orphanedFiles
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

# === Save .vbs file ===
Set-Content -Path $ScriptPath -Value $VbsContent -Force
Write-Output "✅ WICleanup.vbs script saved to $ScriptPath"

# === Run Preview Mode ===
try {
    Write-Output "🔍 Running in Preview Mode (no deletions)..."
    & cscript.exe //nologo "$ScriptPath" /Q > $PreviewPath
    Write-Output "✅ Preview complete. Orphaned file list saved to $PreviewPath"
} catch {
    Write-Output "❌ Error running preview: $($_.Exception.Message)"
}

# === Display results in console ===
if (Test-Path $PreviewPath) {
    Write-Output "`n--- Preview Output ---"
    Get-Content $PreviewPath
    Write-Output "`nReview the files above. To delete them, run the following command manually:"
    Write-Output "`n    cscript.exe //nologo `"$ScriptPath`" /S"
    Write-Output "`nThat will safely delete the orphaned .msi/.msp files and log the output to:"
    Write-Output "    $LogPath"
} else {
    Write-Output "⚠️ No preview file generated — script may not have found any orphaned files."
}

# Step 2 
# cscript.exe //nologo "C:\NetworkTitan\WICleanup\WICleanup.vbs" /S
