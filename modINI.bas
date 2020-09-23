Attribute VB_Name = "ModIni"

'  Terms of Agreement:

'  By using this code, you agree to the following terms:

' 1) You may use this code in your own programs and may
'    compile it into an .exe/.dll/.ocx and distribute it
'    in binary format freely and with no charge.

' 2) You MAY NOT redistribute this code (for example to a
'    web site) without written permission from the
'    original author. Failure to do so is a violation of
'    copyright laws.

' 3) You may link to this code from another website, but
'    ONLY if it is not wrapped in a frame.

'  This code is copywrited by Niklas Spångberg 2000
'  Email: nickokick@spray.se

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function WriteIni(KeySection As String, KeyKey As String, KeyValue As String) As Boolean
    Dim lngResult As Long, strFileName As String, AppPath As String
    
    ' Get the full path to the ini-file.
    AppPath = App.Path
    If right(AppPath, 1) = "\" Then
        strFileName = AppPath & "Config.ini"
    Else
        strFileName = AppPath & "\Config.ini"
    End If

    ' Write to ini-file
    lngResult = WritePrivateProfileString(KeySection, KeyKey, KeyValue, strFileName)
    
    ' Check if the write was successful
    If lngResult = 0 Then
        ' If an error occured, return False.
        WriteIni = False
    Else
        ' Return True (Successful write).
        WriteIni = True
    End If
End Function

Public Function ReadIni(KeySection As String, KeyKey As String) As String
    Dim lngResult As Long
    Dim strFileName
    Dim strResult As String * 50
    
    ' Get the full path to the ini-file.
    AppPath = App.Path
    If right(AppPath, 1) = "\" Then
        strFileName = AppPath & "Config.ini"
    Else
        strFileName = AppPath & "\Config.ini"
    End If
    
    ' Read the ini-file
    lngResult = GetPrivateProfileString(KeySection, KeyKey, strFileName, strResult, Len(strResult), strFileName)
    
    ' Check if the read was successful
    If lngResult = 0 Then
        ' If an error occured, return "error".
        ReadIni = "error"
    Else
        ' Return the value.
        ReadIni = left(strResult, lngResult)
    End If
End Function
