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
    Debug.Print RxCalc_CorrectionFactor(TDD, ActBG, TarBG, RapidIns) & " units"
    Debug.Print RxCalc_CarbCounting(TDD, Carbs) & " units"
    
End Sub


Public Function RxCalc_CorrectionFactor(ByVal TDD As Integer, _
    ByVal ActualBG As Integer, Optional ByVal TargetBG As Integer = 140, _
    Optional ByVal RapidIns As Boolean = True) As Variant
Attribute RxCalc_CorrectionFactor.VB_Description = "Calculate correction factor insulin dose.\r\nFormula: CF = (Actual BG - Target BG) / Insulin Sensitivity\r\nOutput: Insulin dose [units]"
Attribute RxCalc_CorrectionFactor.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Input(s):
    '   Total daily dose (basal and bolus) = units
    '   Actual blood glucose = mg/dL
    '   Target blood glucose = mg/dL
    ' Output:
    '   Correction factor dose

    ' Validate variables
    If ActualBG <= TargetBG And ActualBG > 0 Then
        RxCalc_CorrectionFactor = 0
        Exit Function
    ElseIf TDD = 0 Then
        RxCalc_CorrectionFactor = CVErr(xlErrDiv0)
        Exit Function
    ElseIf TDD < 0 Or ActualBG <= 0 Then
        RxCalc_CorrectionFactor = CVErr(xlErrNum)
        Exit Function
    End If

    ' Declare variables
    Dim InsSens As Single

    ' Calculate insulin sensitivity (InsSens) based on insulin type
    If RapidIns = True Then
        ' Rule of 1800
        InsSens = 1800 / TDD
    Else
        ' Rule of 1500
        InsSens = 1500 / TDD
    End If
    
    ' Return correction factor dose
    RxCalc_CorrectionFactor = (ActualBG - TargetBG) / InsSens

End Function

Public Function RxCalc_CarbCounting(ByVal TDD As Integer, _
    ByVal Carbs As Integer) As Variant
Attribute RxCalc_CarbCounting.VB_Description = "Calculate carb counting insulin dose.\r\nFormula: CC = Meal carbs / (500 / TDD)\r\nOutput: Insulin dose [units]"
Attribute RxCalc_CarbCounting.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Input(s)
    '   Total Daily Dose (basal and bolus) = units
    '   Carbs = Grams
    ' Output:
    '   Carb Counting Dose = units
    
    ' Validate variables
    If TDD = 0 Then
        RxCalc_CarbCounting = CVErr(xlErrDiv0)
        Exit Function
    ElseIf TDD < 0 Or Carbs < 0 Then
        RxCalc_CarbCounting = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim CarbIns As Single
    
    ' Calculate carb:insulin ratio
    CarbIns = 500 / TDD
    
    ' Return carb counting dose
    RxCalc_CarbCounting = Carbs / CarbIns
    
End Function

Public Sub RxCalc_DMMacroArg()
    
    Application.MacroOptions "RxCalc_CorrectionFactor", "Calculate correction factor insulin dose." & vbNewLine & _
        "Formula: CF = (Actual BG - Target BG) / Insulin Sensitivity" & vbNewLine & _
        "Output: Insulin dose [units]", , , , , "RxCalc", , , , _
        Array("Total Daily Dose (basal+bolus) [units]", _
        "Actual blood glucose [units]", _
        "OPTIONAL Target blood glucose [units (Default: 140)]", _
        "OPTIONAL Specify insulin type [TRUE=Rapid (Default) or FALSE=Regular]")

    Application.MacroOptions "RxCalc_CarbCounting", "Calculate carb counting insulin dose." & vbNewLine & _
        "Formula: CC = Meal carbs / (500 / TDD)" & vbNewLine & _
        "Output: Insulin dose [units]", , , , , "RxCalc", , , , _
        Array("Total Daily Dose (basal+bolus) [units]", _
        "Carbs [grams]")
        
End Sub


