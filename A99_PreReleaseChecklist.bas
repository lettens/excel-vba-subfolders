Attribute VB_Name = "A99_PreReleaseChecklist"
Option Explicit

' These functions will be used by the Pre-Release Checklist code
' so that the code does not need to know any 'internal workings' of
' the workbook, i.e. the named ranges, etc
' Simon Letten 23-Oct-2017


Public Function HandlingErrors() As Boolean
    HandlingErrors = myerrors
End Function

Public Function VersionNumber() As Single

    ' Returns the version number from the named range
    
    On Error Resume Next
    
    Dim sVersionNumber As String
    Dim fVersionNumber As Single
    
    ' Load into a string variable to avoid weird small decimal place errors
    sVersionNumber = GetNamedRangeValue(ThisWorkbook.Name, "nrVC_Version")
    ' Other types of workbook might be using this range:
    ' sVersionNumber = GetNamedRangeValue(ThisWorkbook.Name,"rngVersion")
    
    If IsNumeric(sVersionNumber) Then
        fVersionNumber = CSng(sVersionNumber)
    Else
        fVersionNumber = 0
    End If

    VersionNumber = fVersionNumber
    
End Function

Public Function VersionFileName() As String

    ' Returns the workbook name & version number from the named range
    On Error Resume Next
    VersionFileName = GetNamedRangeValue(ThisWorkbook.Name, "nrVC_Filename")
    
End Function

Public Function VBAProjectIsProtected() As Boolean

    ' Returns True if the VB Project is protected
    On Error Resume Next
    VBAProjectIsProtected = (ThisWorkbook.VBProject.Protection = 1)

End Function
