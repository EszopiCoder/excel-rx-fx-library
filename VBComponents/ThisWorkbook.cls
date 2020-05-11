VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Add macro descriptions
    Call Rx_HtWtMacroArg
    Call Rx_RenalMacroArg
    Call Rx_DMMacroArg
    Call Rx_PedsMacroArg
    Call Rx_PedsGrowthMacroArg
End Sub

Private Sub Workbook_Open()
    Set App = Application
End Sub

Private Sub Workbook_AddinInstall()
    MsgBox "Thank you for installing the Rx Fx Library." & vbNewLine & _
        "This is an open source project created by EszopiCoder." & vbNewLine & _
        "(https://github.com/EszopiCoder/excel-rx-fx-library)", vbInformation
End Sub

Private Sub Workbook_AddinUninstall()
    MsgBox "Sorry to hear you will no longer be using the 'Rx Fx Library'." & vbNewLine & _
        "Please send suggestions for improvement to pharm.coder@gmail.com", vbInformation
End Sub

