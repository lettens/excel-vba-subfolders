Attribute VB_Name = "A99_FunctionsConstants"

' ---------------------------------------------------------------------------------------
' Module    : MB_FunctionsConstants
' DateTime  : 22-June-16
' Author    : Michael Beckinsale of Excel Experts
' Purpose   : Module to house constants and functions that are used frequently throughout
'             the VBA project.
' ---------------------------------------------------------------------------------------

Option Explicit
Option Private Module

Public Const myerrors As Boolean = True 'False

Public Function VersionNumber() As Single

    ' Returns the version number from the named range
    
    On Error Resume Next
    
    Dim sVersionNumber As String
    Dim fVersionNumber As Single
    
    ' Load into a string variable to avoid weird small decimal place errors
    sVersionNumber = CStr(ThisWorkbook.Names("nrVC_Version").RefersToRange.Value)
    
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
    VersionFileName = CStr(ThisWorkbook.Names("nrVC_Filename").RefersToRange.Value)
    
End Function

'   Function used to return the user login id

Function currentuserloginid() As String
    currentuserloginid = Environ("USERNAME")
End Function

'   Function to return the current username

Function currentusername() As String
    currentusername = Application.UserName
End Function

