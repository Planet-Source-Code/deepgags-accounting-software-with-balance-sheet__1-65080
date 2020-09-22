Attribute VB_Name = "Module1"
Public compname, address, city, filepath As String
Public dates, datet As Date
Public conn As New ADODB.Connection
Public conn1 As New ADODB.Connection
Public sno, p, l As Currency
Public cmd As New ADODB.Command

Sub Main()
conn1.CursorLocation = adUseClient
conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\company.mdb"
cmd.ActiveConnection = conn1
'ShowCompInfo
'infoform.Left = MDIForm1.Width - (infoform.Width + 100)
'infoform.Top = MDIForm1.Height - (infoform.Height + 100)
selectcomp.Show , MDIForm1
MDIForm1.Show
End Sub

Public Function chkdate(dt As Date)
If Format(dt, "dd/mm/yyyy") < Format(dates, "dd/mm/yyyy") Or Format(dt, "dd/mm/yyyy") > Format(datet, "dd/mm/yyyy") Then
MsgBox "Invalid Date......" & Chr(13) & "Date Must Be Between " & Format(dates, "dd/mm/yyyy") & " to " & Format(datet, "dd/mm/yyyy")
chkdate = "01/01/1900"
Else
chkdate = dt
End If

End Function
