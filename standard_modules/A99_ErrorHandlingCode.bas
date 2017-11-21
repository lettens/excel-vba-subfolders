Attribute VB_Name = "A99_ErrorHandlingCode"
Option Explicit

Public Function HandlingErrors() As Boolean
    HandlingErrors = myerrors
End Function

Public Function MsgBoxTitle() As String

' returns the name of this book but
' without the extension

Dim fullStopPosn As Integer
Dim result As String

    result = ThisWorkbook.Name
    fullStopPosn = InStrRev(result, ".", , vbTextCompare)
    If fullStopPosn > 0 Then
        result = Mid$(result, 1, fullStopPosn - 1)
    End If
    
    MsgBoxTitle = result
    
End Function

Public Function LogsFolder() As String

' Returns the full path to the Logs subfolder
' of ThisWorkbook.Path.
' Returns ThisWorkbook.Path if Logs
' subfolder doesn't exist.

Dim result As String

Const LOGS_FOLDER As String = "Log_Files"
    
On Error Resume Next
    
    result = BuildPath(ThisWorkbook.Path, LOGS_FOLDER)
    
    If Dir$(result, vbDirectory) = "" Then
        result = ""
    End If
    
    ' If function failed to return a value, use folder
    ' of this workbook
    If result = "" Then
        result = ThisWorkbook.Path
    End If
    
    LogsFolder = result

End Function

Private Function LogFileName() As String

' Returns the name of the file to which errors get written
    LogFileName = Replace(MsgBoxTitle() & "_errors" & ".log", " ", "_")
End Function

Private Function ErrorWindowCountOk() As Boolean

' Returns True if the number of errors in the time window
' is below the threshold
' If above the threshold, should avoid recording any more errors

Static errorCounter As Integer
Static startOfErrorWindow As Date

Dim theMessage As String
Dim theDescription As String
Dim result As Boolean

On Error GoTo ErrorHandler

Const ERROR_WINDOW_SECS As Integer = 60
Const MAX_NUM_ERRORS As Integer = 5

Const PROC_NAME As String = "ErrorWindowCountOk"

    ' If the error window started less than ERROR_WINDOW_SECS seconds ago, then
    ' look at how many errors have been counted. If too many, do not record this error
    If DateDiff("s", startOfErrorWindow, Now) <= ERROR_WINDOW_SECS Then
        If errorCounter >= MAX_NUM_ERRORS Then
            If errorCounter = (MAX_NUM_ERRORS) Then
                errorCounter = errorCounter + 1
                theMessage = "Error reporting threshold exceeded."
                theDescription = "Halting the reporting of errors. The error reporting window is " & ERROR_WINDOW_SECS & " seconds."
                Call WriteErrorToFile(PROC_NAME, theMessage, 0, theDescription)
                Call WriteErrorToWorksheet(PROC_NAME, theMessage, 0, theDescription)
            End If
            ' Temporarily stop recording the errors
            result = False
            errorCounter = errorCounter + 1
            GoTo ExitProc
        End If
        errorCounter = errorCounter + 1
        result = True
    Else
        ' The prev error was over ERROR_WINDOW_SECS secs ago
        ' Set the counters/window
        startOfErrorWindow = Now()
        errorCounter = 1
        result = True
    End If

ExitProc:
On Error Resume Next
    ErrorWindowCountOk = result
    Exit Function
    
ErrorHandler:
' Assume the system is overloading
    result = False
    Resume ExitProc

End Function

Public Sub CentralErrorHandler(ByVal procName As String, ByVal messageForUser As String, _
    ByVal errorNumber As Long, ByVal errorDesc As String, _
    Optional ByVal otherDetails As String, Optional DisplayMsgBox As Boolean = True)

' Gets called by various procs
' Controls whether to display message box and also
' calls proc to write error to log file and worksheet
' Simon Letten 02-Jan-2014

On Error Resume Next

    If ErrorWindowCountOk() Then
        If DisplayMsgBox Then
            MsgBox messageForUser & vbNewLine & vbNewLine & errorDesc, vbExclamation, MsgBoxTitle()
        End If
        
        Call WriteErrorToFile(procName, messageForUser, errorNumber, errorDesc, otherDetails)
        Call WriteErrorToWorksheet(procName, messageForUser, errorNumber, errorDesc, otherDetails)
    End If
    
End Sub

Private Sub WriteErrorToFile(ByVal procName As String, ByVal messageForUser As String, _
    ByVal errorNumber As Long, ByVal errorDesc As String, _
    Optional ByVal otherDetails As String)

' NOTE: Requires the Microsoft Scripting Runtime reference set
' This function is passed some text and the details of an error that has occurred.
' These details are appended to the error log file.

Dim fileSysObject As Scripting.FileSystemObject
Dim outputStream As Scripting.TextStream
Dim theUserName As String
Dim theMachineName As String
Dim fileHeadings As String
Dim fileOutput As String
Dim theFilePathAndName As String
        
On Error GoTo ErrorHandler

Const SEPARATOR As String = ","

    theUserName = GetUserName()
    theMachineName = GetComputerName()
    
    Set fileSysObject = New FileSystemObject
    
    theFilePathAndName = BuildPath(LogsFolder(), LogFileName())
    
    If Not fileSysObject.FileExists(theFilePathAndName) Then
        ' Create the column headings
        fileHeadings = "Comma-separated Log file for " & ThisWorkbook.Name & " workbook" & vbNewLine _
            & "Date/Time" & SEPARATOR & "User Name" & SEPARATOR & "Machine Name" & SEPARATOR _
            & " Proc/Function" & SEPARATOR & "Message" & SEPARATOR & "Error Number" & SEPARATOR _
            & "Error Description" & SEPARATOR & "Other Details"
    End If
    
    ' Create the file if it doesn't already exist otherwise open for Appending
    Set outputStream = fileSysObject.OpenTextFile(theFilePathAndName, ForAppending, True, TristateFalse)
    If Len(fileHeadings) > 0 Then
        ' Write headings to the file
        outputStream.WriteLine (fileHeadings)
    End If
    
    fileOutput = Format$(Now(), "dd-mmm-yyyy hh:nn:ss") & SEPARATOR & theUserName & SEPARATOR & theMachineName _
        & SEPARATOR & procName & SEPARATOR & messageForUser & SEPARATOR & errorNumber & SEPARATOR & errorDesc _
        & SEPARATOR & otherDetails

    ' replace any newline chars
    ' Simon Letten 25-Feb-2010
    fileOutput = Replace(fileOutput, vbNewLine, " ")
    ' Write a line with the field names
    outputStream.WriteLine (fileOutput)

ExitProc:
On Error Resume Next
    outputStream.Close
    Set outputStream = Nothing
    Set fileSysObject = Nothing
    Exit Sub
    
ErrorHandler:
    ' Cannot do anything here but must simply exit gracefully
    Resume ExitProc

End Sub

Private Sub WriteErrorToWorksheet(ByVal procName As String, ByVal messageForUser As String, _
    ByVal errorNumber As Long, ByVal errorDesc As String, _
    Optional ByVal otherDetails As String)

' NOTE: Requires the Microsoft Scripting Runtime reference set
' This function is passed some text and the details of an error that has occurred.
' These details are appended to the error log file.
' They are also written to an "Errors" worksheet.

' Simon Letten 02-Jan-2013

Dim errorWorksheet As Worksheet
Dim outputRange As Range
Dim arrayOfValues As Variant
Dim theUserName As String
Dim theMachineName As String

On Error GoTo ErrorHandler

Const SEPARATOR As String = ","
Const ERROR_SHEET_NAME As String = "Errors Log"

    theUserName = GetUserName()
    theMachineName = GetComputerName()
    
    ReDim arrayOfValues(1 To 1, 1 To 8)
    
    ' If the errors sheet already exists, then make sure it doesn't have too many entries
    ' If doesn't exist, then create it and add headings
    If WorksheetExists(ThisWorkbook, ERROR_SHEET_NAME) Then
        Set errorWorksheet = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)
        Set outputRange = errorWorksheet.Range("A3")
        ' does the sheet contain a lot of entries?
        ' If so, delete some
        If outputRange.CurrentRegion.Rows.Count > 150 Then
            ' Delete all contents after row 100
            outputRange.CurrentRegion.Offset(RowOffset:=100).EntireRow.ClearContents
        End If
        ' Insert a new row for the error
        outputRange.EntireRow.Insert
        ' Change outputRange so it points at the new row
        Set outputRange = outputRange.Offset(RowOffset:=-1)
    Else
        Set errorWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        errorWorksheet.Visible = xlSheetHidden
        errorWorksheet.Name = ERROR_SHEET_NAME
        Set outputRange = errorWorksheet.Range("A1")
        outputRange.Resize(RowSize:=100).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
        outputRange.Value = "Error logs for " & ThisWorkbook.Name & " workbook"
        Set outputRange = outputRange.Offset(RowOffset:=1)
        ' Use an array so just write to the sheet in one action
        arrayOfValues(1, 1) = "Date/Time"
        arrayOfValues(1, 2) = "User Name"
        arrayOfValues(1, 3) = "Machine Name"
        arrayOfValues(1, 4) = "Proc/Function"
        arrayOfValues(1, 5) = "Message"
        arrayOfValues(1, 6) = "Error Number"
        arrayOfValues(1, 7) = "Error Description"
        arrayOfValues(1, 8) = "Other Details"
        outputRange.Resize(ColumnSize:=1 + UBound(arrayOfValues, 2) - LBound(arrayOfValues, 2)) = arrayOfValues
        Set outputRange = outputRange.Offset(RowOffset:=1)
    End If

    ' Write the details to errors worksheet
    ' Use an array so just write to the sheet in one action
    arrayOfValues(1, 1) = Now
    arrayOfValues(1, 2) = theUserName
    arrayOfValues(1, 3) = theMachineName
    arrayOfValues(1, 4) = procName
    arrayOfValues(1, 5) = Replace(messageForUser, vbNewLine, " ")
    arrayOfValues(1, 6) = errorNumber
    arrayOfValues(1, 7) = Replace(errorDesc, vbNewLine, " ")
    arrayOfValues(1, 8) = Replace(otherDetails, vbNewLine, " ")
    outputRange.Resize(ColumnSize:=1 + UBound(arrayOfValues, 2) - LBound(arrayOfValues, 2)) = arrayOfValues

ExitProc:
On Error Resume Next
    Erase arrayOfValues
    Exit Sub
    
ErrorHandler:
    ' Cannot do anything here but must simply exit gracefully
    Resume ExitProc

End Sub

