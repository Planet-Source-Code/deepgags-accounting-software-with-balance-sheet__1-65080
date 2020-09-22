VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form group 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6255
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11033
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   100
      BackColorBkg    =   16777215
      GridLinesFixed  =   0
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Group.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Group.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Group.frx":6B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Group.frx":CDCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   11456
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount Rs."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Accounting Heads"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nod As Node
Private Sub Form_Load()
Dim tm As New ADODB.Recordset
Dim temp1 As New ADODB.Recordset
Dim temp2 As New ADODB.Recordset
Dim temp3 As New ADODB.Recordset
tm.Open "select * from groups where childof=1", conn

tv.Nodes.Clear
Set nod = tv.Nodes.Add()
nod.Text = "All Accounts"
nod.Bold = True
nod.Expanded = True
Do While tm.EOF = False
grp1 = tm![nameis]
p = "a" & tm![ID]
Set nod = tv.Nodes.Add(1, tvwChild, p, grp1, 3)
nod.Expanded = True
If temp1.State = 1 Then temp1.Close
temp1.Open "select * from groups where childof=" & tm![ID], conn
Do While temp1.EOF = False
tmp = temp1![nameis]
mtp = "b" & temp1![ID]
Set nod = tv.Nodes.Add("a" & tm![ID], tvwChild, mtp, tmp, 2)
'nod.Expanded = True
If temp2.State = 1 Then temp2.Close
temp2.Open "select * from groups where childof=" & temp1![ID], conn
Do While temp2.EOF = False

tmp = temp2![nameis]
'MsgBox temp2![ID]
Set nod = tv.Nodes.Add(mtp, tvwChild, "g" & temp2![ID], tmp, 1)


'**************Ledger Values
If temp3.State = 1 Then temp3.Close
temp3.Open "select * from ledgers where undergroup=" & temp2![ID], conn
Do While temp3.EOF = False
tmp = temp3![nameis]
'MsgBox temp2![ID]
Set nod = tv.Nodes.Add("g" & temp2![ID], tvwChild, "L" & temp3![ID], tmp, 4)
temp3.MoveNext
Loop





temp2.MoveNext
Loop





temp1.MoveNext
Loop




tm.MoveNext
Loop
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.Caption = x & "     " & y
End Sub

Private Sub Form_Resize()
tv.Left = 0
tv.Width = Me.Width * 0.45
MSHFlexGrid1.Left = tv.Width + 50
MSHFlexGrid1.Width = Me.Width - (tv.Width + 50 + 200)
Label2.Top = 600

Label2.Left = MSHFlexGrid1.Left
Label3.Left = Label1.Left + Label1.Width
MSHFlexGrid1.Height = Me.Height - 1300
tv.Height = Me.Height - 1000
End Sub

Private Sub tv_Click()
If tv.SelectedItem.Key = "" Then Exit Sub
MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 3600
MSHFlexGrid1.ColWidth(2) = 1500
'When User Click on Ledger code
If Left(tv.SelectedItem.Key, 1) = "L" Then
Dim rsmp As New ADODB.Recordset
t = Right(tv.SelectedItem.Key, (Len(tv.SelectedItem.Key) - 1))
rsmp.Open "select * from ledgers where id='" & t & "'", conn
MSHFlexGrid1.Rows = rsmp.RecordCount + 1
MSHFlexGrid1.Cols = 3
R = 0
tot = 0
Do While rsmp.EOF = False
MSHFlexGrid1.TextMatrix(R, 0) = rsmp![ID]
MSHFlexGrid1.TextMatrix(R, 1) = rsmp![nameis]
MSHFlexGrid1.TextMatrix(R, 2) = Format(GetLedgerAmt(rsmp![ID]), "0.00")
tot = tot + MSHFlexGrid1.TextMatrix(R, 2)
R = R + 1
rsmp.MoveNext
Loop
MSHFlexGrid1.TextMatrix(R, 1) = "Total Rs. "
MSHFlexGrid1.TextMatrix(R, 2) = Format(tot, "0.00")

MSHFlexGrid1.Col = 2
MSHFlexGrid1.Row = R
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
Else
'When Click on Group
Dim actp As New ADODB.Recordset
Dim actp1 As New ADODB.Recordset
actp.Open "select * from groups where childof=" & Right(tv.SelectedItem.Key, (Len(tv.SelectedItem.Key) - 1)), conn
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = actp.RecordCount + 1
MSHFlexGrid1.Cols = 3
'When Show group
If actp.EOF = False Then
tot = 0
R = 0
Do While actp.EOF = False
MSHFlexGrid1.TextMatrix(R, 0) = actp![ID]
MSHFlexGrid1.TextMatrix(R, 1) = actp![nameis]
MSHFlexGrid1.TextMatrix(R, 2) = Format(GetHeadAmt(actp![ID]), "0.00")
tot = tot + MSHFlexGrid1.TextMatrix(R, 2)
R = R + 1
actp.MoveNext
Loop
MSHFlexGrid1.TextMatrix(R, 1) = "Total Rs. "
MSHFlexGrid1.TextMatrix(R, 2) = Format(tot, "0.00")

MSHFlexGrid1.Col = 2
MSHFlexGrid1.Row = R
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
Else
'When show Ledgers
actp1.Open "select * from ledgers where undergroup=" & Right(tv.SelectedItem.Key, (Len(tv.SelectedItem.Key) - 1)), conn
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = actp1.RecordCount + 1
MSHFlexGrid1.Cols = 3
R = 0
tot = 0
Do While actp1.EOF = False
MSHFlexGrid1.TextMatrix(R, 0) = actp1![ID]
MSHFlexGrid1.TextMatrix(R, 1) = actp1![nameis]
MSHFlexGrid1.TextMatrix(R, 2) = Format(GetLedgerAmt(actp1![ID]), "0.00")
tot = tot + MSHFlexGrid1.TextMatrix(R, 2)
R = R + 1
actp1.MoveNext
Loop
MSHFlexGrid1.TextMatrix(R, 1) = "Total Rs. "
MSHFlexGrid1.TextMatrix(R, 2) = Format(tot, "0.00")

MSHFlexGrid1.Col = 2
MSHFlexGrid1.Row = R
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
End If
End If








End Sub

Private Sub tv_DblClick()
If Left(tv.SelectedItem.Key, 1) = "L" Then
 tt = LedgerShow(Right(tv.SelectedItem.Key, Len(tv.SelectedItem.Key) - 1), Now(), Now(), True)
 End If
End Sub

Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
tv_Click
End Sub
