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


'   Function used to return the user login id

Function currentuserloginid() As String
    currentuserloginid = Environ("USERNAME")
End Function

'   Function to return the current username

Function currentusername() As String
    currentusername = Application.UserName
End Function

