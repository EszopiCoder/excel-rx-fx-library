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
Attribute Rx_CrCl_CG.VB_ProcData.VB_Invoke_Func = " \n21"
    
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

Public Function Rx_GFR_Class(ByVal eGFR As Integer) As String
Attribute Rx_GFR_Class.VB_Description = "Classify GFR."
Attribute Rx_GFR_Class.VB_ProcData.VB_Invoke_Func = " \n21"
    
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
Attribute Rx_GFR_CKDEPI.VB_ProcData.VB_Invoke_Func = " \n21"

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
Attribute Rx_GFR_MDRD.VB_ProcData.VB_Invoke_Func = " \n21"
    
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

Public Sub Rx_RenalMacroArg()

    Application.MacroOptions "Rx_CrCl_CG", "Calculate creatinine clearance (CrCl)." & vbNewLine & _
        "Formula: CrCl = [(140 - Age) " & Chr(215) & " Weight] / (72 " & Chr(215) & " sCr) " & Chr(215) & " 0.85 [if female]" & vbNewLine & _
        "Output: CrCl [mL/min]", , , , , "Rx", , , , _
        Array("Number [years]", _
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

