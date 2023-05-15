Attribute VB_Name = "Module1"
'#############################################################################
' 指定フォルダをパスワード付きでZip圧縮
'
'　zip_pw
'#############################################################################
Sub ZIP_PW(ByRef ZIP_PATH As String, ByRef ZIP_PW As String)
    Dim shtMain As Worksheet
    Dim wsh As Object
    Dim fso As Object
    Dim MaxRow As Long
    Dim i As Long
    Dim exePath As String
    Dim command As String
    Dim taisyoPath As String
    Dim zipFilePath As String
    Dim password As String
    Dim obj As Object
    Dim ret As Long
    Set shtMain = ThisWorkbook.Sheets("010")
    exePath = "C:\7-Zip\7z.exe"
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    taisyoPath = ZIP_PATH
    If fso.FileExists(taisyoPath) Then
        zipFilePath = Left(taisyoPath, Len(taisyoPath) - Len(fso.GetExtensionName(taisyoPath))) & "zip"
    Else
        zipFilePath = taisyoPath & ".zip"
    End If
    password = ZIP_PW
    command = exePath & " a"
    If password <> "" Then
        command = command & " -p" & password
    End If
    command = command & Space(1) & Chr(34) & zipFilePath & Chr(34) & Space(1) & Chr(34) & taisyoPath & Chr(34) ' 変更
    ret = wsh.Run(command, 1, True)
    Set wsh = Nothing
    Set fso = Nothing
End Sub
