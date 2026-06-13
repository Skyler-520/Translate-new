Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

' Get parameters from command line arguments
Dim exePath, workingDir, iconFilename, companyName

' Default values
exePath = ""
workingDir = ""
iconFilename = "LOGO"
companyName = ""

' Process command line arguments: exePath, workingDir, iconFilename, companyName
If WScript.Arguments.Count > 0 Then
    exePath = WScript.Arguments(0)
End If

If WScript.Arguments.Count > 1 Then
    workingDir = WScript.Arguments(1)
End If

If WScript.Arguments.Count > 2 Then
    iconFilename = WScript.Arguments(2)
End If

If WScript.Arguments.Count > 3 Then
    companyName = WScript.Arguments(3)
End If

' Check if required paths are provided
If exePath = "" Or workingDir = "" Then
    WScript.Echo "Error: Missing required parameters (exePath and workingDir)"
    WScript.Quit 1
End If

' Build shortcut name with company name and "CADCAM" suffix
Dim shortcutName
If companyName <> "" Then
    shortcutName = companyName & "CADCAM"
Else
    shortcutName = "SyntecLaserMarking"
End If

Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & shortcutName & ".lnk")

' Build icon path
iconPath = workingDir & "\Logo\" & iconFilename & ".ico"

oShellLink.TargetPath = exePath
oShellLink.WorkingDirectory = workingDir
oShellLink.IconLocation = iconPath
oShellLink.Save

WScript.Echo "Shortcut created successfully! Name: " & shortcutName & ", Icon: " & iconFilename & ".ico"