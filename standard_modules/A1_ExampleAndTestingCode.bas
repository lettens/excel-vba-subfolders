Attribute VB_Name = "A1_ExampleAndTestingCode"
Option Explicit


Public Sub Test1()

    Call CentralErrorHandler("Test1", "The first test", 0, "The error description", "Line:0", True)

    Call CentralErrorHandler("Test2", "The 2nd test", 0, "The error description" & vbNewLine & "With a new line inside", "Line:0", False)

End Sub

Public Sub TestExceedingWindow()

Dim i As Integer

    For i = 1 To 100
    
        Call CentralErrorHandler("Test" & i, "The test number #" & i, 0, "The error description", "Line:0", True)
    Next i
'    Call CentralErrorHandler("Test2", "The 2nd test", 0, "The error description" & vbNewLine & "With a new line inside", "Line:0", False)

End Sub

Public Sub Test3()

    Const PROC_NAME As String = "Test3"
    
    If myerrors Then On Error GoTo ErrorHandler

    Dim i As Integer

    i = 1
    i = i + 1
    i = i / 0
    
ExitProc:
    Exit Sub
    
ErrorHandler:
    Call CentralErrorHandler(PROC_NAME, "An error occurred when running the Exmaple code.", Err.Number, Err.Description)
    Resume ExitProc
Resume  ' For when stepping through the code

End Sub





