VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Dim varNewFileName As Variant
    Dim sNewFileName As String
    Dim sOldFilename As String
    Dim sIncrementedFileName As String
    Dim sCurrentUserName As String
    Dim sFileExtension As String
    Dim sNewExtension As String
    Dim sFileFilter As String
    Dim sFilePath As String
    Dim dVersionOld As Double
    Dim dVersionNew As Double
    Dim dIncrement As Double
    Dim vDeveloperAnswer As VbMsgBoxResult
    Dim bUserIsDeveloper As Boolean

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo MyExit

'   -----------------------------------------------------------------------------------
'   Developer control - remind them about version control and offer 'Save' or 'SaveAs'
'   -----------------------------------------------------------------------------------

    sFilePath = ThisWorkbook.Path
    sFileExtension = LCase$(GetFileExtension(ThisWorkbook.Name))
    sFileFilter = "Excel workbooks (*." & sFileExtension & "), *." & sFileExtension
    
'   Determine if the current user is a developer and ask question
    bUserIsDeveloper = UserIsDeveloper(ThisWorkbook)
    
    If bUserIsDeveloper Then
        vDeveloperAnswer = MsgBox("Do you want to 'SAVEAS' and increment the VERSION CONTROL number?", vbYesNo + vbCritical, "Version Control")
    '   Action if'SaveAs'is chosen
        If vDeveloperAnswer = vbYes Then
    '       .. cancel direct save
            Cancel = True
    '       .. increment version, build filename, and inform user
            With versioncontrol
                ' Read values in string variables firts to avoid weird Excel number issues
                dVersionOld = CDbl(CStr(.Range("nrVC_Version").Value))
                dIncrement = CDbl(CStr(.Range("nrVC_Increment").Value))
                dVersionNew = dVersionOld + dIncrement
                .Range("nrVC_Version").Value = dVersionNew
                .Calculate
                sIncrementedFileName = .Range("nrVC_Filename").Value
            End With
            
            MsgBox "The suggested filename in the following 'SaveAs' dialogue has been incremented by " & CStr(dIncrement), _
                vbOKOnly + vbInformation, "Information"
'           .. get current filename to ensure it is not used again
            sOldFilename = ThisWorkbook.Name
'           .. give the developer opportunity to inspect new filename, and change if required
'              ## The suggested approach is for the developer to change the parameters within the 'Version Control' sheet ##
            varNewFileName = Application.GetSaveAsFilename(InitialFileName:=BuildPath(sFilePath, sIncrementedFileName), _
                FileFilter:=sFileFilter)
'           .. if 'SaveAs' is cancelled revert to original version number and filename
            If varNewFileName = False Then
                With versioncontrol
                    .Range("nrVC_Version").Value = dVersionOld
                    .Calculate
                End With
                GoTo MyExit
            End If
    '       .. extract new filename
            sNewFileName = GetFileName(CStr(varNewFileName))
            sNewExtension = LCase$(GetFileExtension(sNewFileName))
            If sNewExtension <> sFileExtension Then
                MsgBox "The file extension in the new file name (." & sNewExtension & ") does not match the one currently " _
                    & "in use(." & sFileExtension & "). Unable to save!", vbCritical, MsgBoxTitle()
                GoTo MyExit
            End If
    '       .. check old and new filenames differ | check that the version number has increaed
            If UCase$(sNewFileName) = UCase$(sOldFilename) Or dVersionNew < dVersionOld Then
                MsgBox "Invalid Filename, please try again, the file has NOT been saved.", vbCritical, "Stop"
            Else
    '       .. 'SaveAs' if checks are ok.
                ThisWorkbook.SaveAs BuildPath(sFilePath, sNewFileName), FileFormat:=ThisWorkbook.FileFormat
                GoTo MyExit
            End If
        End If
    '   Normal save if no version control required
        If vDeveloperAnswer = vbNo Then
            ThisWorkbook.Save
            GoTo MyExit
        End If
    End If

'   --------------------------------------------------------------
'   User control - it is fine for end users to 'Save' or 'SaveAs'
'   --------------------------------------------------------------

MyExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

