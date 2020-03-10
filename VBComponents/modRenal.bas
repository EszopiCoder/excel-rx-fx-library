Attribute VB_Name = "modRenal"
Option Explicit

Private Sub TestGFR()
    
    ' Declare variables
    Dim Age As Integer
    Dim Weight As Single
    Dim Female As Boolean
    Dim Black As Boolean
    Dim sCr As Single
    Dim Metric As Boolean
    Dim CKDEPI As Integer, MDRD As Integer
    
    ' Set variables
    Age = 45
    Weight = 250
    Female = False
    Black = False
    sCr = 1
    Metric = False
    CKDEPI = Rx_GFR_CKDEPI(Age, sCr, Female, Black)
    MDRD = Rx_GFR_MDRD(Age, sCr, Female, Black)
    
    ' Display calculations
    Debug.Print "CrCl: " & Rx_CrCl_CG(Age, Weight, sCr, Female, Metric) & " mL/min"
    Debug.Print "CKDEPI: " & CKDEPI & " mL/min/1.73m" & Chr(178) & " Category " & Rx_GFR_Class(CKDEPI)
    Debug.Print "MDRD: " & MDRD & " mL/min/1.73m" & Chr(178) & " Category " & Rx_GFR_Class(MDRD)
    
End Sub

Public Function Rx_CrCl_CG(ByVal Age As Integer, _
    ByVal Weight As Single, ByVal sCr As Single, _
    ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_CrCl_CG.VB_Description = "Calculate creatinine clearance (CrCl).\r\nFormula: CrCl = [(140 - Age) × Weight] / (72 × sCr) × 0.85 [if female]\r\nOutput: CrCl [mL/min]"
Attribute Rx_CrCl_CG.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of Cockcroft-Gault formula
    '   CrCl = [(140 - Age) * Weight] / (72 * sCr) * 0.85 [if female]
    ' Input(s):
    '   Age = Years
    '   Weight = lbs or kg
    '   sCr = mg/dL
    ' Output: CrCl = mL/min
    
    ' Validate variables
    If Age < 0 Or Weight <= 0 Or sCr < 0 Then
        Rx_CrCl_CG = CVErr(xlErrNum)
        Exit Function
    ElseIf sCr = 0 Then
        Rx_CrCl_CG = CVErr(xlErrDiv0)
        Exit Function
    End If
    
    ' Declare variables
    Dim Wt As Single
    Dim CrCl As Integer
    
    ' Convert to metric units
    If Metric = False Then
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Wt = Weight
    End If
    
    ' Calculate CrCl per sex
    If Female = True Then
        CrCl = (140 - Age) * Wt / (72 * sCr) * 0.85
    Else
        CrCl = (140 - Age) * Wt / (72 * sCr)
    End If
    
    ' Return CrCl
    Rx_CrCl_CG = CrCl
    
End Function

Public Function Rx_CrCl_SC(ByVal Age As Integer, _
    ByVal Height As String, ByVal Weight As Single, _
    ByVal sCr As Single, ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_CrCl_SC.VB_Description = "Calculate creatinine clearance (CrCl).\r\nCrCl (male) = (137 - Age) × [(0.285 × (Weight in kg)) + (12.1 × (Height in m)²] / (51 × sCr)\r\nCrCl (female) = (146 - Age) × [(0.287 × (Weight in kg)) + (9.74 × (Height in m)²] / (60 × sCr)\r\nOutput: CrCl [mL/min]"
Attribute Rx_CrCl_SC.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Based off of Cockcroft-Gault formula
    '   CrCl (male) = (137 - Age) * [(0.285 * (Weight in kg)) + (12.1 * (Height in m)^2] / (51 * sCr)
    '   CrCl (female) = (146 - Age) * [(0.287 * (Weight in kg)) + (9.74 * (Height in m)^2] / (60 * sCr)
    ' Input(s):
    '   Age = Years
    '   Height = inches or cm
    '   Weight = lbs or kg
    '   sCr = mg/dL
    ' Output: CrCl = mL/min

    ' Validate variables
    If Age < 0 Or Weight <= 0 Or sCr < 0 Then
        Rx_CrCl_SC = CVErr(xlErrNum)
        Exit Function
    ElseIf Len(Height) = 0 Or Weight <= 0 Then
        Rx_CrCl_SC = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) = 0 And Len(Height) > 0 Then
        Rx_CrCl_SC = CVErr(xlErrDiv0)
        Exit Function
    ElseIf Val(Height) < 0 And Len(Height) > 0 Then
        Rx_CrCl_SC = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_CrCl_SC = CVErr(xlErrNum)
        Exit Function
    ElseIf sCr = 0 Then
        Rx_CrCl_SC = CVErr(xlErrDiv0)
        Exit Function
    End If

    ' Declare variables
    Dim Ht As Single
    Dim Wt As Single
    Dim CrCl As Integer
    
    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to metric units
    If Metric = False Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) * 2.54
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Ht = Height
        Wt = Weight
    End If

     ' Calculate CrCl per sex
    If Female = True Then
        CrCl = (146 - Age) * ((0.287 * Wt) + (0.000974 * Ht ^ 2)) / (60 * sCr)
    Else
        CrCl = (137 - Age) * ((0.285 * Wt) + (0.00121 * Ht ^ 2)) / (51 * sCr)
    End If

    ' Return CrCl
    Rx_CrCl_SC = CrCl

End Function

Public Function Rx_GFR_Class(ByVal eGFR As Integer) As String
Attribute Rx_GFR_Class.VB_Description = "Classify GFR."
Attribute Rx_GFR_Class.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Input: GFR = mL/min/1.73m^2
    ' Output: GFR class
    
    Select Case eGFR
        Case Is >= 90
            Rx_GFR_Class = "G1: Normal or high"
        Case 60 To 89
            Rx_GFR_Class = "G2: Mildly decreased"
        Case 45 To 59
            Rx_GFR_Class = "G3a: Mildly to moderately decreased"
        Case 30 To 44
            Rx_GFR_Class = "G3b: Moderately to severely decreased"
        Case 15 To 29
            Rx_GFR_Class = "G4: Severely decreased"
        Case Is < 15
            Rx_GFR_Class = "G5: Kidney failure"
    End Select

End Function

Public Function Rx_GFR_CKDEPI(ByVal Age As Integer, _
    ByVal sCr As Single, ByVal Female As Boolean, _
    Optional ByVal Black As Boolean = False) As Variant
Attribute Rx_GFR_CKDEPI.VB_Description = "Calculate GFR with CKDEPI formula.\r\nFormula: eGFR = 141 × min(sCr/k, 1)^a × max(sCr/k, 1)^-1.209 × 0.993^Age × 1.018 [if female] × 1.159 [if Black]\r\nOutput: eGFR [mL/min/1.73m²]"
Attribute Rx_GFR_CKDEPI.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Based off of CKDEPI formula
    '   eGFR = 141 * min(sCr/k, 1)^a * max(sCr/k, 1)^-1.209 * 0.993^Age * 1.018 [if female] * 1.159 [if Black]
    ' Input(s):
    '   sCr = mg/dL
    '   Age = years
    ' Output: eGFR = mL/min/1.73m^2

    ' Validate variables
    If Age < 18 Or sCr <= 0 Then
        Rx_GFR_CKDEPI = CVErr(xlErrNum)
        Exit Function
    End If

    ' Declare variables
    Dim k As Single
    Dim a As Single
    Dim min As Single
    Dim max As Single
    Dim eGFR As Single
    
    ' Set sex constants
    If Female = True Then
        k = 0.7
        a = -0.329
    Else
        k = 0.9
        a = -0.411
    End If

    ' Set min/max constants
    If sCr < k Then
        min = sCr / k
        max = 1
    Else
        min = 1
        max = sCr / k
    End If
    
    ' Calculate eGFR
    eGFR = 141 * min ^ a * max ^ -1.209 * 0.993 ^ Age
    
    ' Female correction factor
    If Female = True Then
        eGFR = eGFR * 1.018
    End If
    
    ' Black correction factor
    If Black = True Then
        eGFR = eGFR * 1.159
    End If
    
    ' Return function
    Rx_GFR_CKDEPI = eGFR

End Function

Public Function Rx_GFR_MDRD(ByVal Age As Integer, _
    ByVal sCr As Single, ByVal Female As Boolean, _
    Optional ByVal Black As Boolean = False) As Variant
Attribute Rx_GFR_MDRD.VB_Description = "Calculate GFR with MDRD formula.\r\nFormula: eGFR = 175 × sCr^-1.154 × Age^-0.203 × 0.742 [if female] × 1.212 [if Black]\r\nOutput: eGFR [mL/min/1.73m²]"
Attribute Rx_GFR_MDRD.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of MDRD formula
    '   eGFR = 175 * sCr^-1.154 * Age^-0.203 * 0.742 [if female] * 1.212 [if black]
    ' Input(s):
    '   sCr = mg/dL
    '   Age = years
    ' Output: eGFR = mL/min/1.73m^2
    
    ' Validate variables
    If Age < 18 Or sCr <= 0 Then
        Rx_GFR_MDRD = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim eGFR As Single
    
    ' Calculate eGFR
    eGFR = 175 * sCr ^ -1.154 * Age ^ -0.203
    
    ' Female correction factor
    If Female = True Then
        eGFR = eGFR * 0.742
    End If
    
    ' Black correction factor
    If Black = True Then
        eGFR = eGFR * 1.212
    End If
    
    ' Return function
    Rx_GFR_MDRD = eGFR
    
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

Public Sub Rx_RenalMacroArg()

    Application.MacroOptions "Rx_CrCl_CG", "Calculate creatinine clearance (CrCl)." & vbNewLine & _
        "Formula: CrCl = [(140 - Age) " & Chr(215) & " Weight] / (72 " & Chr(215) & " sCr) " & Chr(215) & " 0.85 [if female]" & vbNewLine & _
        "Output: CrCl [mL/min]", , , , , "Rx", , , , _
        Array("Number [years]", _
        "Number [lbs or kg]", _
        "Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Specify the units of height and weight [TRUE=Metric (Default) or FALSE=US]")
    
    Application.MacroOptions "Rx_CrCl_SC", "Calculate creatinine clearance (CrCl)." & vbNewLine & _
        "CrCl (male) = (137 - Age) " & Chr(215) & " [(0.285 " & Chr(215) & " (Weight in kg)) + (12.1 " & Chr(215) & " (Height in m)" & Chr(178) & "] / (51 " & Chr(215) & " sCr)" & vbNewLine & _
        "CrCl (female) = (146 - Age) " & Chr(215) & " [(0.287 " & Chr(215) & " (Weight in kg)) + (9.74 " & Chr(215) & " (Height in m)" & Chr(178) & "] / (60 " & Chr(215) & " sCr)" & vbNewLine & _
        "Output: CrCl [mL/min]", , , , , "Rx", , , , _
        Array("Number [years]", _
        "Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Specify the units of height and weight [TRUE=Metric (Default) or FALSE=US]")
    
    Application.MacroOptions "Rx_GFR_Class", "Classify GFR.", , , , , "Rx", , , , _
        Array("eGFR [mL/min/1.73m" & Chr(178) & "]")
    
    Application.MacroOptions "Rx_GFR_CKDEPI", "Calculate GFR with CKDEPI formula." & vbNewLine & _
        "Formula: eGFR = 141 " & Chr(215) & " min(sCr/k, 1)^a " & Chr(215) & " max(sCr/k, 1)^-1.209 " & Chr(215) & _
            " 0.993^Age " & Chr(215) & " 1.018 [if female] " & Chr(215) & " 1.159 [if Black]" & vbNewLine & _
        "Output: eGFR [mL/min/1.73m" & Chr(178) & "]", , , , , "Rx", , , , _
        Array("Number [years]", _
        "Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Black race [TRUE=Black or FALSE=Other (Default)]")
        
    Application.MacroOptions "Rx_GFR_MDRD", "Calculate GFR with MDRD formula." & vbNewLine & _
        "Formula: eGFR = 175 " & Chr(215) & " sCr^-1.154 " & Chr(215) & " Age^-0.203 " & Chr(215) & _
            " 0.742 [if female] " & Chr(215) & " 1.212 [if Black]" & vbNewLine & _
        "Output: eGFR [mL/min/1.73m" & Chr(178) & "]", , , , , "Rx", , , , _
        Array("Number [years]", _
        "Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Black race [TRUE=Black or FALSE=Other (Default)]")
    
End Sub

