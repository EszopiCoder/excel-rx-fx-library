Attribute VB_Name = "modDM"
Option Explicit

Sub TestDM()

    ' Declare variables
    Dim TDD As Integer
    Dim ActBG As Integer
    Dim TarBG As Integer
    Dim RapidIns As Boolean
    Dim Carbs As Integer
    
    ' Set variables
    TDD = 36
    ActBG = 220
    TarBG = 120
    RapidIns = True
    Carbs = 60
    
    ' Display calculations
    Debug.Print Rx_CorrectionFactor(TDD, ActBG, TarBG, RapidIns) & " units"
    Debug.Print Rx_CarbCounting(TDD, Carbs) & " units"
    
End Sub


Public Function Rx_DM_CF(ByVal TDD As Integer, _
    ByVal Actual As Integer, Optional ByVal Target As Integer = 140, _
    Optional ByVal Rapid As Boolean = True) As Variant
Attribute Rx_DM_CF.VB_Description = "Calculate correction factor insulin dose.\r\nFormula: CF = (Actual BG - Target BG) / Insulin Sensitivity\r\nOutput: Insulin dose [units]"
Attribute Rx_DM_CF.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Input(s):
    '   Total daily dose (basal and bolus) = units
    '   Actual blood glucose = mg/dL
    '   Target blood glucose = mg/dL
    ' Output:
    '   Correction factor dose = units

    ' Validate variables
    If Actual <= Target And Actual > 0 Then
        Rx_DM_CF = 0
        Exit Function
    ElseIf TDD = 0 Then
        Rx_DM_CF = CVErr(xlErrDiv0)
        Exit Function
    ElseIf TDD < 0 Or Actual <= 0 Then
        Rx_DM_CF = CVErr(xlErrNum)
        Exit Function
    End If

    ' Declare variables
    Dim InsSens As Single

    ' Calculate insulin sensitivity (InsSens) based on insulin type
    If Rapid = True Then
        ' Rule of 1800
        InsSens = 1800 / TDD
    Else
        ' Rule of 1500
        InsSens = 1500 / TDD
    End If
    
    ' Return correction factor dose
    Rx_DM_CF = (Actual - Target) / InsSens

End Function

Public Function Rx_DM_CC(ByVal TDD As Integer, _
    ByVal Carbs As Integer) As Variant
Attribute Rx_DM_CC.VB_Description = "Calculate carb counting insulin dose.\r\nFormula: CC = Meal carbs / (500 / TDD)\r\nOutput: Insulin dose [units]"
Attribute Rx_DM_CC.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Input(s)
    '   Total Daily Dose (basal and bolus) = units
    '   Carbs = Grams
    ' Output:
    '   Carb Counting Dose = units
    
    ' Validate variables
    If TDD = 0 Then
        Rx_DM_CC = CVErr(xlErrDiv0)
        Exit Function
    ElseIf TDD < 0 Or Carbs < 0 Then
        Rx_DM_CC = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim CarbIns As Single
    
    ' Calculate carb:insulin ratio
    CarbIns = 500 / TDD
    
    ' Return carb counting dose
    Rx_DM_CC = Carbs / CarbIns
    
End Function

Public Sub Rx_DMMacroArg()
    
    Application.MacroOptions "Rx_DM_CF", "Calculate correction factor insulin dose (CF)." & vbNewLine & _
        "Formula: CF = (Actual BG - Target BG) / Insulin Sensitivity" & vbNewLine & _
        "Output: Insulin dose [units]", , , , , "Rx", , , , _
        Array("Total Daily Dose (basal+bolus) [units]", _
        "Actual blood glucose [units]", _
        "OPTIONAL Target blood glucose [units (Default: 140)]", _
        "OPTIONAL Specify insulin type [TRUE=Rapid (Default) or FALSE=Regular]")

    Application.MacroOptions "Rx_DM_CC", "Calculate carb counting insulin dose (CC)." & vbNewLine & _
        "Formula: CC = Meal carbs / (500 / TDD)" & vbNewLine & _
        "Output: Insulin dose [units]", , , , , "Rx", , , , _
        Array("Total Daily Dose (basal+bolus) [units]", _
        "Carbs [grams]")
        
End Sub


