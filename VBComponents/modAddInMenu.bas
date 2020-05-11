Attribute VB_Name = "modAddInMenu"
Option Explicit
Dim RxFxList As Variant

'*********************************XML CODE*********************************
'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
'   <ribbon>
'      <tabs>
'         <tab idMso="TabFormulas">
'            <group id="RxFxLib" label="Rx Function Library">
'               <gallery id="RxFx"
'                   label="Rx Functions" columns="1"
'                   imageMso = "GroupFunctionLibrary"
'                   getItemCount = "RxFx_getItemCount"
'                   getItemLabel = "RxFx_getItemLabel"
'                   getItemScreentip = "RxFx_getItemScreentip"
'                   getItemSupertip = "RxFx_getItemSupertip"
'                   onAction = "RxFx_Click"
'                   showItemLabel = "true"
'                   size="large">
'                 <button id="insertFx"
'                    imageMso = "GroupFunctionLibrary"
'                    label = "Insert Function"
'                    screentip="Insert Function (Shift+F3)"
'                    supertip = "Work with the formula in the current cell. You can easily pick functions to use and get help on how to fill out the input values."
'                    onAction="insertFx_Click"/>
'               </gallery>
'               <button id="updateFx"
'                   imageMso = "ConnectedToolSyncMenu"
'                   label = "Update Fx"
'                   screentip="Update Functions"
'                   supertip = "Update all functions in current workbook."
'                   onAction = "updateFx_Click"
'                   size="large"/>
'               <button id="getHelp"
'                   imageMso = "Help"
'                   label = "Help"
'                   screentip="Help"
'                   supertip = "Open link to webpage."
'                   onAction = "getHelp_Click"
'                   size="large"/>
'            </group>
'         </tab>
'      </tabs>
'   </ribbon>
'</customUI>
'*********************************XML CODE*********************************

Private Sub AddInMenuProperties()
    ' Custom function for changing file properties (not used during run time)
    ActiveWorkbook.BuiltinDocumentProperties("Title").Value = "Rx Function Library 1.5"
    ActiveWorkbook.BuiltinDocumentProperties("Comments").Value = "Function library for custom pharmacy equations"
End Sub

Sub Auto_Open()

    ' Populate RxFxList
    RxFxList = Array("Rx_AdjBW()", "Rx_LBW()", _
        "Rx_IBW_Devine()", "Rx_IBW_Baseline()", _
        "Rx_IBW_BMI()", "Rx_IBW_Hume()", _
        "Rx_BMI()", "Rx_BMI_Class()", _
        "Rx_BSA_DuBois()", "Rx_BSA_Mosteller()", _
        "Rx_CrCl_CG()", "Rx_CrCl_SC()", _
        "Rx_GFR_CKDEPI()", "Rx_GFR_MDRD()", _
        "Rx_GFR_Class()", "Rx_DM_CF()", _
        "Rx_DM_CC()", "Rx_PEDS_AdjAge()", _
        "Rx_PEDS_GFR_BS()", "Rx_PEDS_LenAgeInf()", _
        "Rx_PEDS_WtAgeInf()", "Rx_PEDS_HcAgeInf()", _
        "Rx_PEDS_WtLenInf()", "Rx_PEDS_StatAge()", _
        "Rx_PEDS_WtAge()", "Rx_PEDS_BmiAge()", _
        "Rx_PEDS_WtStat()")

End Sub

Sub getHelp_Click(control As IRibbonControl)

    Dim URL As String
    
    URL = "https://github.com/EszopiCoder/excel-rx-fx-library/wiki"
    
    If MsgBox("You are leaving Microsoft Word to the following website: " & URL & _
    vbNewLine & vbNewLine & "Would you like to proceed?", _
    vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    ActiveWorkbook.FollowHyperlink URL

End Sub

Sub updateFx_Click(control As IRibbonControl)

    Call UpdateFx

End Sub

Sub RxFx_getItemCount(control As IRibbonControl, ByRef returnedVal)
    ' Return the number of functions in the array
    returnedVal = UBound(RxFxList) - LBound(RxFxList) + 1
End Sub

Sub RxFx_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error Resume Next
    ' Return the name of the function without arguments
    returnedVal = Left(RxFxList(index), InStr(1, RxFxList(index), "(") - 1)
    On Error GoTo 0
End Sub

Sub RxFx_getItemScreentip(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error Resume Next
    ' Return the name of the function with arguments
    returnedVal = RxFxList(index)
    On Error GoTo 0
End Sub

Sub RxFx_getItemSupertip(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim Supertip As Variant
    Supertip = _
    Array("Return the adjusted body weight of a person (Devine formula).", _
          "Return the lean body weight of a person.", _
          "Return the ideal body weight of a person (Devine formula).", _
          "Return the ideal body weight of a person under 60 inches (Baseline method).", _
          "Return the ideal body weight of a person under 60 inches (BMI method).", _
          "Return the ideal body weight of a person under 60 inches (Hume formula).", _
          "Return the BMI of a person.", "Return the BMI class of a person.", _
          "Return the BSA of a person (Du Bois formula).", "Return the BSA of a person (Mosteller formula).", _
          "Return the creatinine clearance of a person (Cockcroft-Gault formula).", _
          "Return the creatinine clearance of a person (Salazar-Corcoran formula).", _
          "Return the eGFR of a person (CKDEPI formula).", "Return the eGFR of a person (MDRD formula).", "Return the eGFR class of a person.", _
          "Return the correction factor insulin dose.", _
          "Return the carbohydrate counting insulin dose.", _
          "Return the adjusted age (corrected age).", _
          "Return the eGFR of a child (Bedside-Schwartz formula).", _
          "Return the length-for-age percentile of an infant.", "Return the weight-for-age percentile of an infant.", _
          "Return the head circumference-for-age percentile of an infant.", "Return the weight-for-length percentile of an infant.", _
          "Return the stature-for-age percentile of children and adolescents.", "Return the weight-for-age percentile of children and adolescents.", _
          "Return the BMI-for-age percentile of children and adolescents.", "Return the weight-for-stature percentile for preschoolers.")

    On Error Resume Next
    returnedVal = Supertip(index)
    On Error GoTo 0
End Sub

Sub insertFx_Click(control As IRibbonControl)

    ActiveCell.FunctionWizard

End Sub

Sub RxFx_Click(control As IRibbonControl, id As String, index As Integer)
    On Error Resume Next
    ' Insert function into active cell (same as the other built-in functions)
    If InStr(1, ActiveCell.Formula, "=") > 0 Then
        ActiveCell.Formula = ActiveCell.Formula & "+" & RxFxList(index)
    Else
        ActiveCell.Formula = "=" & RxFxList(index)
    End If
    ' Open function wizard dialog. Clear cell if user hits cancel button.
    If Application.Dialogs(xlDialogFunctionWizard).Show = False Then
        ActiveCell.Formula = ""
    End If
    On Error GoTo 0
End Sub
