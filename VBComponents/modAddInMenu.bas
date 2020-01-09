Attribute VB_Name = "modAddInMenu"
Option Explicit
Dim RxFxList(0 To 16) As Variant

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
    ActiveWorkbook.BuiltinDocumentProperties("Title").Value = "Rx Function Library 1.2"
    ActiveWorkbook.BuiltinDocumentProperties("Comments").Value = "Function library for custom pharmacy equations"
End Sub

Sub Auto_Open()

    ' Populate RxFxList
    RxFxList(0) = "Rx_AdjBW()"
    RxFxList(1) = "Rx_IBW_Devine()"
    RxFxList(2) = "Rx_IBW_Intuitive()"
    RxFxList(3) = "Rx_IBW_Baseline()"
    RxFxList(4) = "Rx_IBW_Hume()"
    RxFxList(5) = "Rx_BMI()"
    RxFxList(6) = "Rx_BMI_Class()"
    RxFxList(7) = "Rx_BSA_DuBois()"
    RxFxList(8) = "Rx_BSA_Mosteller()"
    RxFxList(9) = "Rx_CrCl_CG()"
    RxFxList(10) = "Rx_GFR_CKDEPI()"
    RxFxList(11) = "Rx_GFR_MDRD()"
    RxFxList(12) = "Rx_GFR_Class()"
    RxFxList(13) = "Rx_DM_CF()"
    RxFxList(14) = "Rx_DM_CC()"
    RxFxList(15) = "Rx_PEDS_AdjAge()"
    RxFxList(16) = "Rx_PEDS_GFR_BS()"
    
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
          "Return the ideal body weight of a person 60 inches or greater (Devine formula).", _
          "Return the ideal body weight of a person under 60 inches (Intuitive method).", _
          "Return the ideal body weight of a person under 60 inches (Baseline method).", _
          "Return the ideal body weight of a person under 60 inches (Hume formula).", _
          "Return the BMI of a person.", _
          "Return the BMI class of a person.", _
          "Return the BSA of a person (Du Bois formula).", _
          "Return the BSA of a person (Mosteller formula).", _
          "Return the Cockcroft-Gault creatinine clearance of a person.", _
          "Return the eGFR of a person (CKDEPI formula).", _
          "Return the eGFR of a person (MDRD formula).", _
          "Return the eGFR class of a person.", _
          "Return the correction factor insulin dose.", _
          "Return the carb counting insulin dose.", _
          "Return the adjusted age (corrected age).", _
          "Return the eGFR of a child (Bedside-Schwartz formula).")

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
