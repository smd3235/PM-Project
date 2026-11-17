Attribute VB_Name = "PM_Report_Host"
Option Explicit

' ==================== HOST / DIALOG HELPERS ====================
' File/folder pickers (host-specific UI).
' Entry points: PickSourceFile, PickOutputFolder
Public Function PickSourceFile() As Variant
    PickSourceFile = Application.GetOpenFilename( _
        FileFilter:="Excel Reports (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Select the CURRENT raw-data report")
End Function


' Output folder: Excel folder picker ("" = user cancelled).

Public Function PickOutputFolder() As String
    Dim fd As Object

    On Error Resume Next
    Set fd = Application.FileDialog(4)
    If Not fd Is Nothing Then
        fd.Title = "Choose the folder for finished PM Requests reports"
        fd.AllowMultiSelect = False
        If fd.Show = -1 Then PickOutputFolder = fd.SelectedItems(1)
    End If
    On Error GoTo 0
End Function



