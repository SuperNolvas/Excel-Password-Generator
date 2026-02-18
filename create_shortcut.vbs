 ' Creates a Windows desktop shortcut that opens PowerShell in
 ' %USERPROFILE%\Downloads\Excel-PASSWORD-GENERATOR
 ' Run this VBScript by double-clicking or via: cscript create_shortcut.vbs
 Option Explicit
 Dim WshShell, userProfile, folder, desktop, shortcutPath, lnk
 Set WshShell = WScript.CreateObject("WScript.Shell")
 userProfile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
Dim fso, downloadsFolder, documentsFolder, targetFolder, downloadsParent
Set fso = CreateObject("Scripting.FileSystemObject")
downloadsFolder = userProfile & "\Downloads\Excel-PASSWORD-GENERATOR"
documentsFolder = userProfile & "\Documents\Excel-Password-Generator"
downloadsParent = userProfile & "\Downloads"

' Choose the best target folder that exists
If fso.FolderExists(downloadsFolder) Then
	targetFolder = downloadsFolder
ElseIf fso.FolderExists(documentsFolder) Then
	targetFolder = documentsFolder
Else
	targetFolder = downloadsParent
End If

desktop = WshShell.SpecialFolders("Desktop")
shortcutPath = desktop & "\Excel Password Generator.lnk"
Set lnk = WshShell.CreateShortcut(shortcutPath)
lnk.TargetPath = WshShell.ExpandEnvironmentStrings("%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe")
' Build arguments depending on whether the exact project folder exists
If targetFolder = downloadsFolder Or targetFolder = documentsFolder Then
	lnk.Arguments = "-NoExit -Command ""Set-Location -LiteralPath '" & targetFolder & "'"""
	lnk.WorkingDirectory = targetFolder
Else
	lnk.Arguments = "-NoExit -Command ""Write-Host 'Target folder not found; opening Downloads'; Set-Location -LiteralPath '" & targetFolder & "'"""
	lnk.WorkingDirectory = targetFolder
End If

lnk.IconLocation = WshShell.ExpandEnvironmentStrings("%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe,0")
lnk.Description = "Open PowerShell in Downloads\Excel-PASSWORD-GENERATOR (falls back to Downloads)"
lnk.Save