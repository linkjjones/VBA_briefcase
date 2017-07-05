Attribute VB_Name = "ExcelOpen"
Public Sub SuppressOpenAlerts()
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.ActiveWorkbook.UpdateLinks = xlUpdateLinksAlways
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub

