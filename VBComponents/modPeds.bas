Attribute VB_Name = "modPeds"
Option Explicit

Sub TestPeds()

    Debug.Print Rx_PEDS_AdjAge(10, 29) & " months"

End Sub

Public Function Rx_PEDS_AdjAge(ByVal Age As Integer, _
    ByVal GA As Integer) As Variant
Attribute Rx_PEDS_AdjAge.VB_Description = "Calculate adjusted age or corrected age (AdjAge).\r\nFormula: AdjAge = Chronological Age - ((40 - GA) / 4)\r\nOutput: Adjusted Age (Corrected Age) [months]"
Attribute Rx_PEDS_AdjAge.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Based off of Adjusted Age (Corrected Age) formula
    '   AdjAge = Chronological Age - ((40 - GA) / 4)
    ' Input(s)
    '   Age = Months
    '   GA = Gestational age: Age from date of mother's first day of last menstrual period to date of birth in weeks.
    ' Output:
    '   Adjusted Age in months

    Rx_PEDS_AdjAge = Age - ((40 - GA) / 4)

End Function

Public Function Rx_PEDS_GFR_BS(ByVal Height As String, _
    ByVal sCr As Single, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_PEDS_GFR_BS.VB_Description = "Calculate GFR with Bedside-Schwartz formula (GFR).\r\nFormula: GFR = 0.413 × Height / sCr\r\nOutput: Adjusted Age (Corrected Age) [months]"
Attribute Rx_PEDS_GFR_BS.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Based off of Bedside-Schwartz formula
    '   eGFR = eGFR = 0.413 * Height / sCr
    ' Input(s):
    '   Height = inches or cm
    '   sCr = mg/dL
    ' Output: eGFR = mL/min/1.73m^2

    ' Validate variables
    If Len(Height) = 0 Then
        Rx_PEDS_GFR_BS = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_PEDS_GFR_BS = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_PEDS_GFR_BS = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim Ht As String

    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to metric units
    If Metric = False Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) * 2.54
    End If
    
    ' Return eGFR
    Rx_PEDS_GFR_BS = 0.413 * Height / sCr
    
End Function

Private Function PrimeToInches(strHeight As String, _
    Optional ByVal Metric As Boolean = False) As String

    ' Input:
    '   Height = inches or cm
    ' Output:
    '   IBW inches or cm
    
    ' Declare variables
    Dim Ht As String
    Dim PrimeLoc As Byte
    
    ' Trim string
    Ht = Trim(strHeight)
    ' Remove quotation marks
    Ht = Replace(Ht, Chr(34), "")
    
    ' Convert Height Format x'y" to inches
    PrimeLoc = InStr(1, Ht, "'")
    If PrimeLoc > 0 And Metric = False Then
        'Debug.Print "Height=" & Ht
        'Debug.Print "Feet=" & Left(Ht, PrimeLoc - 1)
        'Debug.Print "Inches=" & Mid(Ht, PrimeLoc + 1, Len(Ht) - PrimeLoc - 1)
        Ht = Val(Left(Ht, PrimeLoc - 1)) * 12 + _
            Val(Mid(Ht, PrimeLoc + 1, Len(Ht) - PrimeLoc))
    End If

    ' Return inches or cm
    PrimeToInches = Ht
    
End Function

Public Sub Rx_PedsMacroArg()

    Application.MacroOptions "Rx_PEDS_AdjAge", "Calculate adjusted age or corrected age (AdjAge)." & vbNewLine & _
        "Formula: AdjAge = Chronological Age - ((40 - GA) / 4)" & vbNewLine & _
        "Output: Adjusted Age (Corrected Age) [months]", , , , , "Rx", , , , _
        Array("Number [months]", _
        "Number [weeks]")
        
    Application.MacroOptions "Rx_PEDS_GFR_BS", "Calculate GFR with Bedside-Schwartz formula (GFR)." & vbNewLine & _
        "Formula: GFR = 0.413 " & Chr(215) & " Height / sCr" & vbNewLine & _
        "Output: Adjusted Age (Corrected Age) [months]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Serum creatinine [mg/dL]", _
        "OPTIONAL Specify the units of height and weight [TRUE=Metric (Default) or FALSE=US]")
        
End Sub

