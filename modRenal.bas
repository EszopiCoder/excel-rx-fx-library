Attribute VB_Name = "modRenal"
Option Explicit
Private Const ErrorUnder18 = 1800

Private Sub TestGFR()
    
    ' Declare variables
    Dim Age As Byte
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
    CKDEPI = RxCalc_GFR_CKDEPI(sCr, Female, Age, Black)
    MDRD = RxCalc_GFR_MDRD(sCr, Female, Age, Black)
    
    ' Display calculations
    Debug.Print "CrCl: " & RxCalc_CrCl(Age, Weight, sCr, Female, Metric) & " mL/min"
    Debug.Print "CKDEPI: " & CKDEPI & " mL/min/1.73m" & Chr(178) & " Category " & RxCalc_GFR_Class(CKDEPI)
    Debug.Print "MDRD: " & MDRD & " mL/min/1.73m" & Chr(178) & " Category " & RxCalc_GFR_Class(MDRD)
    
End Sub

Public Function RxCalc_CrCl(ByVal Age As Byte, _
    ByVal Weight As Single, ByVal sCr As Single, _
    ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Single
Attribute RxCalc_CrCl.VB_Description = "Calculate creatinine clearance (CrCl).\r\nFormula: CrCl = [(140 - Age) × Weight] / (72 × sCr) × 0.85 [if female]\r\nOutput: CrCl [mL/min]"
Attribute RxCalc_CrCl.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of Cockcroft-Gault equation
    '   CrCl = [(140 - Age) * Weight] / (72 * sCr) * 0.85 [if female]
    ' Input(s):
    '   Age = Years
    '   Weight = lbs or kg
    '   sCr = mg/dL
    ' Output: CrCl = mL/min
    
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
    RxCalc_CrCl = CrCl
    
End Function

Public Function RxCalc_GFR_Class(eGFR As Integer) As String
    
    ' Input: GFR = mL/min/1.73m^2
    ' Output: GFR class
    
    Select Case eGFR
        Case Is >= 90
            RxCalc_GFR_Class = "G1: Normal or high"
        Case 60 To 89
            RxCalc_GFR_Class = "G2: Mildly decreased"
        Case 45 To 59
            RxCalc_GFR_Class = "G3a: Mildly to moderately decreased"
        Case 30 To 44
            RxCalc_GFR_Class = "G3b: Moderately to severely decreased"
        Case 15 To 29
            RxCalc_GFR_Class = "G4: Severely decreased"
        Case Is < 15
            RxCalc_GFR_Class = "G5: Kidney failure"
    End Select

End Function

Public Function RxCalc_GFR_CKDEPI(ByVal sCr As Single, _
    ByVal Female As Boolean, ByVal Age As Integer, _
    Optional ByVal Black As Boolean = False) As Single
Attribute RxCalc_GFR_CKDEPI.VB_Description = "Calculate GFR with CKDEPI formula.\r\nFormula: eGFR = 141 × min(sCr/k, 1)^a × max(sCr/k, 1)^-1.209 × 0.993^Age × 1.018 [if female] × 1.159 [if Black]\r\nOutput: eGFR [mL/min/1.73m²]"
Attribute RxCalc_GFR_CKDEPI.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Based off of CKDEPI equation
    '   eGFR = 141 * min(sCr/k, 1)^a * max(sCr/k, 1)^-1.209 * 0.993^Age * 1.018 [if female] * 1.159 [if Black]"
    ' Input(s):
    '   sCr = mg/dL
    '   Age = years
    ' Output: eGFR = mL/min/1.73m^2

    ' Validate variables
    If Age < 18 Then
        RxCalc_GFR_CKDEPI = ErrorUnder18
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
    RxCalc_GFR_CKDEPI = eGFR

End Function

Public Function RxCalc_GFR_MDRD(ByVal sCr As Single, _
    ByVal Female As Boolean, ByVal Age As Integer, _
    Optional ByVal Black As Boolean = False) As Single
Attribute RxCalc_GFR_MDRD.VB_Description = "Calculate GFR with MDRD formula.\r\nFormula: eGFR = 175 × sCr^-1.154 × Age^-0.203 × 0.742 [if female] × 1.212 [if Black]\r\nOutput: eGFR [mL/min/1.73m²]"
Attribute RxCalc_GFR_MDRD.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of MDRD equation
    '   eGFR = 175 * sCr^-1.154 * Age^-0.203 * (0.742 if female, 1.212 if black)
    ' Input(s):
    '   sCr = mg/dL
    '   Age = years
    ' Output: eGFR = mL/min/1.73m^2
    
    ' Validate variables
    If Age < 18 Then
        RxCalc_GFR_MDRD = ErrorUnder18
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
    RxCalc_GFR_MDRD = eGFR
    
End Function

Public Sub RxCalc_RenalMacroArg()

    Application.MacroOptions "RxCalc_CrCl", "Calculate creatinine clearance (CrCl)." & vbNewLine & _
        "Formula: CrCl = [(140 - Age) " & Chr(215) & " Weight] / (72 " & Chr(215) & " sCr) " & Chr(215) & " 0.85 [if female]" & vbNewLine & _
        "Output: CrCl [mL/min]", , , , , "RxCalc", , , , _
        Array("Number [years]", _
        "Number [lbs or kg]", _
        "Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Specify the units of height and weight [TRUE=Metric (Default) or FALSE=US]")
    
    Application.MacroOptions "RxCalc_GFR_CKDEPI", "Calculate GFR with CKDEPI formula." & vbNewLine & _
        "Formula: eGFR = 141 " & Chr(215) & " min(sCr/k, 1)^a " & Chr(215) & " max(sCr/k, 1)^-1.209 " & Chr(215) & _
            " 0.993^Age " & Chr(215) & " 1.018 [if female] " & Chr(215) & " 1.159 [if Black]" & vbNewLine & _
        "Output: eGFR [mL/min/1.73m" & Chr(178) & "]", , , , , "RxCalc", , , , _
        Array("Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "Number [years]", _
        "OPTIONAL Black race [TRUE=Black or FALSE=Other (Default)]")
        
    Application.MacroOptions "RxCalc_GFR_MDRD", "Calculate GFR with MDRD formula." & vbNewLine & _
        "Formula: eGFR = 175 " & Chr(215) & " sCr^-1.154 " & Chr(215) & " Age^-0.203 " & Chr(215) & _
            " 0.742 [if female] " & Chr(215) & " 1.212 [if Black]" & vbNewLine & _
        "Output: eGFR [mL/min/1.73m" & Chr(178) & "]", , , , , "RxCalc", , , , _
        Array("Serum creatinine [mg/dL]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "Number [years]", _
        "OPTIONAL Black race [TRUE=Black or FALSE=Other (Default)]")
    
End Sub

