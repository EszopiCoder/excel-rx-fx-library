Attribute VB_Name = "modHtWt"
Option Explicit
Private Const ErrorUnder60 = 9999
Private Const ErrorOver60 = 8888

Private Sub TestHtWt()
    
    ' Declare variables
    Dim Height As String, Weight As Single
    Dim Metric As Boolean
    Dim Female As Boolean
    Dim BMI As Single
    
    ' Set variables
    Height = 69
    Weight = 150
    Metric = False
    Female = False
    BMI = Rx_BMI(Height, Weight, Metric)
    
    ' Display calculations
    Debug.Print "BMI: " & BMI & " kg/m^2 (" & Rx_BMI_Class(BMI) & ")"
    Debug.Print "BSA (Du Bois): " & Rx_BSA_DuBois(Height, Weight, Metric) & " m^2"
    Debug.Print "BSA (Mosteller): " & Rx_BSA_Mosteller(Height, Weight, Metric) & " m^2"
    Debug.Print "IBW: " & Rx_IBW_Devine(Height, Weight, Metric) & " kg"
    Debug.Print "AdjBW: " & Rx_AdjBW_Devine(Height, Weight, Female, Metric) & " kg"
    
End Sub

Public Function Rx_BMI(ByVal Height As String, _
    ByVal Weight As Single, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_BMI.VB_Description = "Calculate body mass index (BMI).\r\nFormula: BMI = (Weight in kg) / (Height in cm)²\r\nOutput: BMI [kg/m²]"
Attribute Rx_BMI.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of BMI formula:
    '   BMI = (Weight in kg) / (Height in cm)^2 x 10000
    ' Input(s):
    '   Height = inches or cm
    '   Weight = lbs or kg
    ' Output:
    '   BMI = kg/m^2
    
    ' Validate variables
    If Len(Height) = 0 Or Weight <= 0 Then
        Rx_BMI = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) = 0 And Len(Height) > 0 Then
        Rx_BMI = CVErr(xlErrDiv0)
        Exit Function
    ElseIf Val(Height) < 0 And Len(Height) > 0 Then
        Rx_BMI = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_BMI = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim Ht As String
    Dim BMI As Single
    
    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Calculate BMI
    BMI = Weight / Val(Ht) ^ 2
    
    ' Convert to metric units
    If Metric = False Then
        BMI = BMI * 703
    Else
        BMI = BMI * 10000
    End If
    
    ' Return BMI
    Rx_BMI = BMI
    
End Function

Public Function Rx_BMI_Class(ByVal BMI As Single) As String
Attribute Rx_BMI_Class.VB_Description = "Classify body mass index (BMI)."
Attribute Rx_BMI_Class.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Input(s):
    '   BMI = kg/m^2
    ' Output:
    '   BMI class
    
    Select Case BMI
        Case Is < 18.5
            Rx_BMI_Class = "Underweight"
        Case 18.5 To 24.9
            Rx_BMI_Class = "Normal"
        Case 25 To 29.9
            Rx_BMI_Class = "Overweight"
        Case 30 To 34.9
            Rx_BMI_Class = "Obese class I"
        Case 35 To 39.9
            Rx_BMI_Class = "Obese class II"
        Case Is >= 40
            Rx_BMI_Class = "Obese class III"
    End Select
    
End Function

Public Function Rx_BSA_DuBois(ByVal Height As String, _
    ByVal Weight As Single, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_BSA_DuBois.VB_Description = "Calculate body surface area (BSA) with Du Bois formula.\r\nFormula: BSA = 0.007184 × (Weight in kg)^0.425 × (Height in cm)^0.725\r\nOutput: BSA [m²]"
Attribute Rx_BSA_DuBois.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of Du Bois Formula:
    '   BSA = 0.007184 x (Weight in kg)^0.425 x (Height in cm)^0.725
    ' Input(s):
    '   Height = inches or cm
    '   Weight = lbs or kg
    ' Output:
    '   BSA = m^2
    
    ' Validate variables
    If Len(Height) = 0 Or Weight <= 0 Then
        Rx_BSA_DuBois = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_BSA_DuBois = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_BSA_DuBois = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim Ht As String
    Dim Wt As Single
    Dim BSA As Single
    
    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to metric units
    If Metric = False Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) * 2.54
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Wt = Weight
    End If
    
    ' Calculate BSA
    BSA = 0.007184 * Wt ^ 0.425 * Ht ^ 0.725
    
    ' Return BSA
    Rx_BSA_DuBois = BSA
    
End Function

Public Function Rx_BSA_Mosteller(ByVal Height As String, _
    ByVal Weight As Single, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_BSA_Mosteller.VB_Description = "Calculate body surface area (BSA) with Mosteller formula.\r\nFormula: BSA = Sqr[(Height in cm) × (Weight in kg)] / 60\r\nOutput: BSA [m²]"
Attribute Rx_BSA_Mosteller.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Based off of Mosteller formula:
    '   BSA = Sqr[(Height in cm) x (Weight in kg)] / 60
    ' Input(s):
    '   Height = inches or cm
    '   Weight = lbs or kg
    ' Output:
    '   BSA = m^2

    ' Validate variables
    If Len(Height) = 0 Or Weight <= 0 Then
        Rx_BSA_Mosteller = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_BSA_Mosteller = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_BSA_Mosteller = CVErr(xlErrNum)
        Exit Function
    End If

    ' Declare variables
    Dim Ht As String
    Dim Wt As String
    Dim BSA As Single
    
    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Validate Ht (If Metric=TRUE and Ht includes ')
    If InStr(1, Ht, "'") > 0 Then
        Rx_BSA_Mosteller = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Convert to metric units
    If Metric = False Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) * 2.54
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Wt = Weight
    End If
    
    ' Calculate BSA
    BSA = Sqr(Wt * Ht) / 60
    
    ' Return BSA
    Rx_BSA_Mosteller = BSA

End Function

Public Function Rx_IBW_Devine(ByVal Height As String, _
    ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_IBW_Devine.VB_Description = "Calculate ideal body weight (IBW) with Devine formula.\r\nFormula: IBW (Male) = 50kg + 2.3kg for each inch above/below 60 inches\r\nFormula: IBW (Female) = 45.5kg + 2.3kg for each inch above/below 60 inches\r\nOutput: IBW [kg]"
Attribute Rx_IBW_Devine.VB_ProcData.VB_Invoke_Func = " \n20"
    
    ' Based off of Devine formula:
    '   IBW (Male) = 50kg + 2.3kg for each inch above/below 60 inches
    '   IBW (Female) = 45.5kg + 2.3kg for each inch above/below 60 inches
    ' Input(s):
    '   Height = inches or cm
    ' Output:
    '   IBW = kg
    
    ' Validate variables
    If Len(Height) = 0 Then
        Rx_IBW_Devine = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_IBW_Devine = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_IBW_Devine = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim Ht As String
    Dim IBW As Single

    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to imperial units
    If Metric = True Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) / 2.54
    End If
    
    ' Calculate IBW per sex
    If Female = True Then
        IBW = 45.5 + 2.3 * (Val(Ht) - 60)
    Else
        IBW = 50 + 2.3 * (Val(Ht) - 60)
    End If

    ' Return IBW
    Rx_IBW_Devine = IBW
    
End Function

Public Function Rx_AdjBW_Devine(ByVal Height As String, _
    ByVal Weight As Single, ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_AdjBW_Devine.VB_Description = "Calculate adjusted body weight (AdjBW) with Devine formula.\r\nFormula: AdjBW = IBW + 0.4(Actual - IBW)\r\nOutput: AdjBW [kg]"
Attribute Rx_AdjBW_Devine.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Use only if Height >= 60 inches
    ' Formula:
    '   AdjBW = IBW + 0.4(Actual - IBW)
    ' Input(s):
    '   Height = inches or cm
    '   Weight = lbs or kg
    ' Output:
    '   AdjBW = kg
    
    ' Validate variables
    If Len(Height) = 0 Or Weight <= 0 Then
        Rx_AdjBW_Devine = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_AdjBW_Devine = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_AdjBW_Devine = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim IBW As Single
    Dim Wt As Single
    Dim AdjBW As Single
    
    ' Calculate IBW
    IBW = Rx_IBW_Devine(Height, Female, Metric)
    
    ' Convert to metric units
    If Metric = False Then
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Wt = Weight
    End If
    
    ' Calculate AdjBW
    AdjBW = IBW + 0.4 * (Wt - IBW)
    
    ' Return AdjBW
    Rx_AdjBW_Devine = AdjBW
    
End Function

Public Function Rx_LBW(ByVal Height As String, _
    ByVal Weight As Single, ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_LBW.VB_Description = "Calculate lean body weight (LBW).\r\nFormula: LBW (Male) = [9270 × (Weight in kg)] / [6680 + (216 × BMI)]\r\nFormula: LBW (Female) = [9270 × (Weight in kg)] / [8780 + (244 × BMI)]\r\nOutput: LBW [kg]"
Attribute Rx_LBW.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Formula:
    '   LBW (Male) = [9270 x (Weight in kg)] / [6680 + (216 x BMI)]
    '   LBW (Female) = [9270 x (Weight in kg)] / [8780 + (244 x BMI)]
    ' Input(s):
    '   Height = inches or cm
    '   Weight = lbs or kg
    ' Output:
    '   LBW = kg

    ' Validate variables
    If Len(Height) = 0 Or Weight <= 0 Then
        Rx_LBW = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_LBW = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_LBW = CVErr(xlErrNum)
        Exit Function
    End If

    ' Declare variables
    Dim BMI As Single
    Dim Wt As Single
    Dim LBW As Single

    BMI = Rx_BMI(Height, Weight, Metric)

    ' Convert to metric units
    If Metric = False Then
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Wt = Weight
    End If

    ' Calculate IBW per sex
    If Female = True Then
        LBW = (9270 * Wt) / (8780 + (244 * BMI))
    Else
        LBW = (9270 * Wt) / (6680 + (216 * BMI))
    End If

    ' Return IBW
    Rx_LBW = LBW

End Function

Public Function Rx_IBW_Baseline(ByVal Height As String, _
    ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_IBW_Baseline.VB_Description = "Calculate ideal body weight (IBW) under 60 inches.\r\nFormula: IBW (Male) = 50kg - 0.833kg for each inch below 60 inches\r\nFormula: IBW (Female) = 45.5kg - 0.758kg for each inch below 60 inches\r\nOutput: IBW [kg].\r\nError 8888: Height over 60 inches"
Attribute Rx_IBW_Baseline.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Use only if Height < 60 inches
    ' Based off of Baseline method:
    '   IBW (Male) = 50kg - 0.833kg for each inch below 60 inches
    '   IBW (Female) = 45.5kg - 0.758kg for each inch below 60 inches
    ' Input:
    '   Height = inches or cm
    ' Output:
    '   IBW = kg
    
    ' Validate variables
    If Len(Height) = 0 Then
        Rx_IBW_Baseline = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_IBW_Baseline = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_IBW_Baseline = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim Ht As String
    Dim IBW As Single

    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to imperial units
    If Metric = True Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) / 2.54
    End If
    
    ' Validate variables (Must be under 60 inches)
    If Val(Ht) > 60 Then
        Rx_IBW_Baseline = ErrorOver60
        Exit Function
    End If
    
    ' Calculate IBW per sex
    If Female = True Then
        IBW = 45.5 - (45.5 / 60) * (60 - Val(Ht))
    Else
        IBW = 50 - (5 / 6) * (60 - Val(Ht))
    End If

    ' Return IBW
    Rx_IBW_Baseline = IBW
    
End Function

Public Function Rx_IBW_BMI(ByVal Height As String, _
    ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_IBW_BMI.VB_Description = "Calculate ideal body weight (IBW) under 60 inches.\r\nFormula: IBW (Male) = 50 / 152.4² × 10000 × [(Height in cm) / 100]²\r\nFormula: IBW (Female) = 45.5 / 152.4² × 10000 × [(Height in cm) / 100]²\r\nOutput: IBW [kg].\r\nError 8888: Height over 60 inches"
Attribute Rx_IBW_BMI.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Use only if Height < 60 inches
    ' Based off of BMI method:
    '   IBW (Male) = 50 / 152.4^2 x 10000 x [(Height in cm) / 100]^2
    '   IBW (Female) = 45.5 / 152.4^2 x 10000 x [(Height in cm) / 100]^2
    ' Input:
    '   Height = inches or cm
    ' Output:
    '   IBW = kg
    
    ' Validate variables
    If Len(Height) = 0 Then
        Rx_IBW_BMI = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_IBW_BMI = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_IBW_BMI = CVErr(xlErrNum)
        Exit Function
    End If

    ' Declare variables
    Dim Ht As String
    Dim IBW As Single
    
    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to metric units
    If Metric = False Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) * 2.54
    End If

    ' Validate variables (Must be under 60 inches or 152.4 cm)
    If Val(Ht) > 152.4 Then
        Rx_IBW_BMI = ErrorOver60
        Exit Function
    End If
    
    ' Calculate IBW per sex
    If Female = True Then
        IBW = 45.5 * (Ht / 152.4) ^ 2
    Else
        IBW = 50 * (Ht / 152.4) ^ 2
    End If

    ' Return IBW
    Rx_IBW_BMI = IBW

End Function

Public Function Rx_IBW_Hume(ByVal Height As String, _
    ByVal Weight As Single, ByVal Female As Boolean, _
    Optional ByVal Metric As Boolean = True) As Variant
Attribute Rx_IBW_Hume.VB_Description = "Calculate ideal body weight (IBW) under 60 inches.\r\nFormula: IBW (Male) = [0.3281 × (Weight in kg)] + [0.33939 × (Height in cm)] - 29.5336\r\nFormula: IBW (Female) = [0.29569 × (Weight in kg)] + [0.41813 × (Height in cm)] - 43.2933\r\nOutput: IBW [kg]"
Attribute Rx_IBW_Hume.VB_ProcData.VB_Invoke_Func = " \n20"

    ' Use only if Height < 60 inches
    ' Based off of Hume formula:
    '   IBW (Male) = [0.3281 x (Weight in kg)] + [0.33939 x (Height in cm)] - 29.5336
    '   IBW (Female) = [0.29569 x (Weight in kg)] + [0.41813 x (Height in cm)] - 43.2933
    ' Input:
    '   Height = inches or cm
    '   Weight = lbs or kg
    ' Output:
    '   IBW = kg
    
    ' Validate variables
    If Len(Height) = 0 Or Weight <= 0 Then
        Rx_IBW_Hume = CVErr(xlErrNum)
        Exit Function
    ElseIf Val(Height) <= 0 And Len(Height) > 0 Then
        Rx_IBW_Hume = CVErr(xlErrNum)
        Exit Function
    ElseIf InStr(1, Trim(Height), "'") > 0 And Metric = True Then
        Rx_IBW_Hume = CVErr(xlErrNum)
        Exit Function
    End If
    
    ' Declare variables
    Dim Ht As String
    Dim Wt As Single
    Dim IBW As Single

    ' Convert x'y" to inches
    Ht = PrimeToInches(Height, Metric)
    
    ' Convert to metric units
    If Metric = False Then
        ' 1 inch = 2.54 cm
        Ht = Val(Ht) * 2.54
        ' ~2.2 lbs = 1 kg
        Wt = Weight / 2.20462262185
    Else
        Wt = Weight
    End If
    
    ' Validate variables (Must be under 60 inches)
    If Val(Ht) > 152.4 Then
        Rx_IBW_Hume = ErrorOver60
        Exit Function
    End If
    
    ' Calculate IBW per sex
    If Female = True Then
        IBW = (0.29569 * Wt) + (0.41813 * Ht) - 43.2933
    Else
        IBW = (0.3281 * Wt) + (0.33939 * Ht) - 29.5336
    End If

    ' Return IBW
    Rx_IBW_Hume = IBW

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

Public Sub Rx_HtWtMacroArg()
    
    Application.MacroOptions "Rx_BMI_Class", "Classify body mass index (BMI).", , , , , "Rx", , , , _
        Array("Body mass index [kg/m" & Chr(178) & "]")
    
    Application.MacroOptions "Rx_BMI", "Calculate body mass index (BMI)." & vbNewLine & _
        "Formula: BMI = (Weight in kg) / (Height in cm)" & Chr(178) & vbNewLine & _
        "Output: BMI [kg/m" & Chr(178) & "]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_BSA_DuBois", "Calculate body surface area (BSA) with Du Bois formula." & vbNewLine & _
        "Formula: BSA = 0.007184 " & Chr(215) & " (Weight in kg)^0.425 " & Chr(215) & " (Height in cm)^0.725" & vbNewLine & _
        "Output: BSA [m" & Chr(178) & "]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_BSA_Mosteller", "Calculate body surface area (BSA) with Mosteller formula." & vbNewLine & _
        "Formula: BSA = Sqr[(Height in cm) " & Chr(215) & " (Weight in kg)] / 60" & vbNewLine & _
        "Output: BSA [m" & Chr(178) & "]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_IBW_Devine", "Calculate ideal body weight (IBW) with Devine formula." & vbNewLine & _
        "Formula: IBW (Male) = 50kg + 2.3kg for each inch above/below 60 inches" & vbNewLine & _
        "Formula: IBW (Female) = 45.5kg + 2.3kg for each inch above/below 60 inches" & vbNewLine & _
        "Output: IBW [kg]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_AdjBW_Devine", "Calculate adjusted body weight (AdjBW) with Devine formula." & vbNewLine & _
        "Formula: AdjBW = IBW + 0.4(Actual - IBW)" & vbNewLine & _
        "Output: AdjBW [kg]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_LBW", "Calculate lean body weight (LBW)." & vbNewLine & _
        "Formula: LBW (Male) = [9270 " & Chr(215) & " (Weight in kg)] / [6680 + (216 " & Chr(215) & " BMI)]" & vbNewLine & _
        "Formula: LBW (Female) = [9270 " & Chr(215) & " (Weight in kg)] / [8780 + (244 " & Chr(215) & " BMI)]" & vbNewLine & _
        "Output: LBW [kg]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_IBW_Baseline", "Calculate ideal body weight (IBW) under 60 inches." & vbNewLine & _
        "Formula: IBW (Male) = 50kg - 0.833kg for each inch below 60 inches" & vbNewLine & _
        "Formula: IBW (Female) = 45.5kg - 0.758kg for each inch below 60 inches" & vbNewLine & _
        "Output: IBW [kg]." & vbNewLine & "Error " & ErrorOver60 & ": Height over 60 inches", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")
        
    Application.MacroOptions "Rx_IBW_BMI", "Calculate ideal body weight (IBW) under 60 inches." & vbNewLine & _
        "Formula: IBW (Male) = 50 / 152.4" & Chr(178) & " " & Chr(215) & " 10000 " & Chr(215) & " [(Height in cm) / 100]" & Chr(178) & vbNewLine & _
        "Formula: IBW (Female) = 45.5 / 152.4" & Chr(178) & " " & Chr(215) & " 10000 " & Chr(215) & " [(Height in cm) / 100]" & Chr(178) & vbNewLine & _
        "Output: IBW [kg]." & vbNewLine & "Error " & ErrorOver60 & ": Height over 60 inches", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")

    Application.MacroOptions "Rx_IBW_Hume", "Calculate ideal body weight (IBW) under 60 inches." & vbNewLine & _
        "Formula: IBW (Male) = [0.3281 " & Chr(215) & " (Weight in kg)] + [0.33939 " & Chr(215) & " (Height in cm)] - 29.5336" & vbNewLine & _
        "Formula: IBW (Female) = [0.29569 " & Chr(215) & " (Weight in kg)] + [0.41813 " & Chr(215) & " (Height in cm)] - 43.2933" & vbNewLine & _
        "Output: IBW [kg]", , , , , "Rx", , , , _
        Array("Sample formats: 5'10" & Chr(34) & " or 70 [inches or cm]", _
        "Number [lbs or kg]", _
        "Boolean [TRUE=Female or FALSE=Male]", _
        "OPTIONAL Measurement units of height and weight [TRUE=Metric (Default) or FALSE=US]")
    
End Sub

