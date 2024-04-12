' This VBA macro consists of two subroutines: OnStart and OnEnd.
' OnStart optimizes Excel's performance by disabling various features like screen updating, events, and alerts, while setting calculation mode to manual. 
' OnEnd restores these settings back to their defaults, ensuring normal functionality.

Attribute VB_Name = "Otimizador"
Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    Application.StatusBar = False
    
End Sub

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    ThisWorkbook.Date1904 = False
    ActiveWindow.View = xlNormalView

End Sub

