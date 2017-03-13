Attribute VB_Name = "Module1"
Global Emp As Boolean
Global Appt As Boolean

Sub Main()
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open "PROVIDER=MSDASQL;dsn=PMngr;uid=Admin;pwd=soni;"
    
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub
