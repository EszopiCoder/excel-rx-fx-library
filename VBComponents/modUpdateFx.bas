Attribute VB_Name = "modUpdateFx"
Option Explicit

Public Sub UpdateFx()

    ' Declare variables
    Dim ws As Worksheet
    Dim cell As Range
    Dim cellCount As Long
    Dim instCount As Long
    Dim strTempInst As Long
    Dim i As Long
    Dim findArray As Variant
    Dim replaceArray As Variant
    Dim strUpdateLog As String
    Dim saveLog As Long
    
    On Error GoTo UpdateError

    ' Arrays to change/update formulas
    ' Change these on each release if necessary
    findArray = Array("sin", "tan")
    replaceArray = Array("cos", "cot")
    strUpdateLog = "Cell Range" & vbTab & "Change Log"
    
    ' Turn on fast mode
    Call FastMode(True)
    
    ' Loop through all worksheets and cells
    For Each ws In ActiveWorkbook.Worksheets
        For Each cell In ws.UsedRange.Cells
            ' Only change cells with an equal sign and function
            If Len(cell.Formula) > 1 And Left(cell.Formula, 1) = "=" Then
                For i = LBound(findArray) To UBound(findArray)
                    strTempInst = NumInstances(cell.Formula, CStr(findArray(i)), False)
                    If strTempInst > 0 Then
                        instCount = instCount + strTempInst
                        cellCount = cellCount + 1
                        strUpdateLog = strUpdateLog & vbNewLine & ws.Name & "!" & _
                            Split(Cells(1, Val(cell.Column)).Address, "$")(1) & cell.Row & vbTab & _
                            "Replaced '" & findArray(i) & "' with '" & replaceArray(i) & "' " & _
                            strTempInst & " times"
                    End If
                    cell.Replace findArray(i), replaceArray(i), xlPart
                Next i
            End If
        Next cell
    Next ws
    
    ' Return to normal state
    Call FastMode(False)
    
    ' Exit if there are no changes
    If cellCount = 0 Then Exit Sub
    
    ' Option to save change log to text file
    saveLog = MsgBox(cellCount & " cell(s) updated" & vbNewLine & _
        "Would you like to save a log of the changes?", vbYesNo)
    If saveLog = vbYes Then
        Call SaveAsTextDoc("--Summary of Changes--" & vbNewLine & _
            "Total cell(s) updated: " & cellCount & vbNewLine & _
            "Total instance(s) updated: " & instCount & vbNewLine & _
            strUpdateLog)
    End If
    Exit Sub
    
UpdateError:
    Call FastMode(False)
    MsgBox "Run-time error '" & Err.Number & "':" & vbNewLine & Err.Description, vbInformation
End Sub

Public Sub FastMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub

Public Function NumInstances(strText As String, _
    strFind As String, Optional ByVal boolMatchCase As Boolean = True) As Long
    
    ' Validate variables
    If Len(strText) = 0 Or Len(strFind) = 0 Then
        Err.Raise vbObjectError + 513, "", _
            "Invalid argument(s) for NumInstances(): strText or strFind are null."
    End If
    
    ' If not matching case, change to uppercase
    If boolMatchCase = False Then
        strText = UCase(strText)
        strFind = UCase(strFind)
    End If
    
    ' Return number of instances
    NumInstances = (Len(strText) - Len(Replace(strText, strFind, ""))) / Len(strFind)

End Function

Public Function SaveAsTextDoc(TextPrint As String) As Boolean

    Dim intFile As Integer
    Dim TextFilePath As String
    
    ' Save in the same path as the active workbook
    TextFilePath = ActiveWorkbook.Path & "\" & _
        ActiveWorkbook.Name & " Update " & Replace(Date, "/", "-") & ".txt"
    
    'Validate that user added correct arguments
    If Len(TextPrint) = 0 Then
        MsgBox "No text to save", vbInformation
        Exit Function
    End If
    
    'Write text to text file
    intFile = FreeFile
    Open TextFilePath For Output As #intFile
        Print #intFile, TextPrint
    Close #intFile
    
    'Let user know text document was saved
    MsgBox Dir(TextFilePath) & vbNewLine & _
        "Saved successfully", vbInformation
    SaveAsTextDoc = True

End Function
