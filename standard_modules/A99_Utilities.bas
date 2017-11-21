Attribute VB_Name = "A99_Utilities"
Option Explicit
'
' Functions required to support the error handling code
'
Public Function UserIsDeveloper(ByRef wbSource As Workbook) As Boolean

    ' Tries to match the Application.userName against values in the
    ' developer list. Returns True if found
    
    Dim rngDeveloperNames As Range
    Dim cell As Range
    Dim bUserIsDeveloper As Boolean
    Dim sCellValue As String
    Dim sCurrentUserName As String
    
    bUserIsDeveloper = False
    sCurrentUserName = UCase$(currentusername())
    
    Set rngDeveloperNames = GetNamedRange(wbSource.Name, "nrDeveloperList")
    For Each cell In rngDeveloperNames.Cells
        sCellValue = UCase$(Trim$(cell.Value))
        If sCellValue = vbNullString Then
            Exit For
        End If
        If sCellValue = sCurrentUserName Then
            bUserIsDeveloper = True
            Exit For
        End If
    Next cell

    UserIsDeveloper = bUserIsDeveloper
    
End Function

Public Function GetUserName() As String

Dim result As String

On Error Resume Next

    result = Environ("username") 'Application.UserName
    If Err.Number <> 0 Then
        Err.Clear
        result = result & "<error!>"
    End If
    If result = "" Then
        result = "<unknown>"
    End If

    GetUserName = result

End Function

Public Function GetComputerName() As String

Dim result As String

On Error Resume Next

    result = Environ("computername")
    If Err.Number <> 0 Then
        Err.Clear
        result = result & "<error!>"
    End If
    If result = "" Then
        result = "<unknown>"
    End If

    GetComputerName = result

End Function

Public Function GetNamedRange(ByRef nameOfWorkbook As String, ByRef nameOfRange As String) As Range

Dim resultRange As Range

On Error Resume Next

    Set resultRange = Workbooks(nameOfWorkbook).Names(nameOfRange).RefersToRange
    If Err.Number <> 0 Then
        Set resultRange = Nothing
    End If
    
    Set GetNamedRange = resultRange
    
End Function

Public Function GetNamedRangeValue(ByRef nameOfWorkbook As String, ByRef nameOfRange As String) As String

Dim rngRange As Range
Dim sValue As String

On Error Resume Next

    Set rngRange = GetNamedRange(nameOfWorkbook, nameOfRange)
    
    If Not (rngRange Is Nothing) Then
        sValue = CStr(rngRange.Value)
        Set rngRange = Nothing
    Else
        sValue = vbNullString
    End If
    
    GetNamedRangeValue = sValue
    
End Function

Public Function WorksheetExists(ByRef wb As Workbook, ByRef sheetNameToCheck As String) As Boolean

' Returns True if a worksheet called "sheetNameToCheck"
' exists in workbookToCheck
' Simon Letten 02-Jan-2013

    Dim ws As Worksheet
    Dim result As Boolean

On Error Resume Next

    Set ws = wb.Worksheets(sheetNameToCheck)
    If Err.Number = 0 Then
        result = True
    Else
        Err.Clear
        result = False
    End If
    
    WorksheetExists = result
    
End Function

Public Function BuildPath(folderPath As String, subFolderOrFileName As String) As String

' Using FileSystemObject.BuildPath fails when
' the first path is H:
' Simon Letten 19-Nov-2013

    Dim result As String

    result = folderPath
    
    If (Right(result, 1) <> "\") And (Left(subFolderOrFileName, 1) <> "\") Then
        ' Add a backslash
        result = result & "\"
    End If
    If (Right(result, 1) = "\") And (Left(subFolderOrFileName, 1) = "\") Then
        ' Remove a backslash
        result = Left(result, Len(result) - 1)
    End If

    result = result & subFolderOrFileName
    
    BuildPath = result
    
End Function

Public Function GetFileName(ByRef sFilePathAndName As String) As String

    ' Returns the text after the last '\' in file path and name
    On Error Resume Next
    
    Dim iPosition As Integer
    Dim sResult As String
    
    iPosition = InStrRev(sFilePathAndName, "\")
    
    If iPosition > 0 Then
        sResult = Mid$(sFilePathAndName, iPosition + 1)
    Else
        sResult = vbNullString
    End If
    
    GetFileName = sResult
    
End Function

Public Function GetFileExtension(ByRef sFileName As String) As String

    ' Returns the text after the last full stop in file name
    On Error Resume Next
    
    Dim iPosition As Integer
    Dim sResult As String
    
    iPosition = InStrRev(sFileName, ".")
    
    If iPosition > 0 Then
        sResult = Mid$(sFileName, iPosition + 1)
    Else
        sResult = vbNullString
    End If
    
    GetFileExtension = sResult
    
End Function
