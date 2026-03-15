'====================================================================
' Script Name: FTPUpload.vbs
' Author: [Your Name]
' Created: [Date]
' Description:
'   This VBScript uploads a local file to a specified FTP server using
'   Windows Shell objects. The script demonstrates how to create 
'   an FTP connection string and automate file upload operations 
'   without external dependencies.
'
' Usage:
'   Adjust the FTP credentials, host, and file path as needed.
'   Run the script by double-clicking or via command line using:
'       cscript FTPUpload.vbs
'
' Note:
'   The wait time (waitTime) should be adjusted according to the 
'   size of the uploaded file. Large files may require more time.
'====================================================================

'Define required objects
Set oShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Define the file to upload
path = "C:\lang.txt"

'Call upload procedure
FTPUpload(path)

'--------------------------------------------------------------------
' Subroutine: FTPUpload
' Description:
'   Handles the upload of a specified file to an FTP server.
'   Uses Windows Shell to transfer the file via an FTP URL.
'--------------------------------------------------------------------
Sub FTPUpload(path)

    On Error Resume Next

    Const copyType = 16  'CopyHere operation option

    'FTP wait time (in ms) – ensure file transfer completes
    'Depends on the file size; adjust as necessary.
    waitTime = 80000

    'FTP credentials and connection parameters
    FTPUser = "ftpuser"          'FTP username
    FTPPass = "ftppassword"      'FTP password
    FTPHost = "ftp.please-change.com"  'FTP hostname or IP
    FTPDir  = "/"                'FTP directory (end with "/")

    'Build the FTP connection string
    strFTP = "ftp://" & FTPUser & ":" & FTPPass & "@" & FTPHost & FTPDir
    Set objFTP = oShell.NameSpace(strFTP)

    'Check if file exists before upload
    If objFSO.FileExists(path) Then

        Set objFile = objFSO.GetFile(path)
        strParent = objFile.ParentFolder
        Set objFolder = oShell.NameSpace(strParent)

        Set objItem = objFolder.ParseName(objFile.Name)

        Wscript.Echo "Uploading file " & objItem.Name & " to " & strFTP
        objFTP.CopyHere objItem, copyType

    End If

    'Error handling section (optional)
    If Err.Number <> 0 Then
        Wscript.Echo "Error: " & Err.Description
    End If

    'Pause execution until upload completes
    Wscript.Sleep waitTime

    MsgBox "FTP upload completed."

End Sub
