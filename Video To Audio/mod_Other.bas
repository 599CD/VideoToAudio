Attribute VB_Name = "mod_Other"
Sub OpenFolder(strFolder As String)

    Dim Foldername As String
    'Foldername = "\\server\Instructions\"
    Foldername = strFolder
    
    Shell "C:\WINDOWS\explorer.exe """ & Foldername & "", vbNormalFocus

End Sub

Function ProgramsExist() As Boolean
    ProgramsExist = False
    ProgramsExist = FileExists(App.Path & "\" & "_ffmpeg.exe")
End Function

Function AddQuotes(strValue) As String
    AddQuotes = Chr(34) & strValue & Chr(34)
End Function

Public Function FileExists(PathFilename As String) As Boolean

    Dim D As String

    On Error GoTo NoFileExists
    D = ""
    D = Dir(PathFilename)
    If D <> "" Then
        FileExists = True
        Exit Function
    End If
    
NoFileExists:
    FileExists = False
    
End Function

Function GetFileName(ByVal FilePath As String) As String()
    
    Dim File(2) As String
    
    Dim strFileNameWithExt As String
    Dim strFileName As String
    Dim strFileExt As String
    
    Dim intSlash_Pos As Integer
    Dim intFilePath_Len As Integer
    
    intFilePath_Len = Len(FilePath)
    'MsgBox intFilePath_Len
    intSlash_Pos = InStrRev(FilePath, "\")
    'MsgBox intSlash_Pos
    
    'FileName with Ext
    File(0) = Right(FilePath, intFilePath_Len - intSlash_Pos)
    
    Dim FileNameAndExt() As String
    Dim max As Integer
    Dim min As Integer
    
    FileNameAndExt = Split(File(0), ".")
    max = UBound(FileNameAndExt)
    min = LBound(FileNameAndExt)
    
    'FileName
    File(1) = FileNameAndExt(min)
    'Ext
    File(2) = FileNameAndExt(max)
    
    GetFileName = File
    
End Function
