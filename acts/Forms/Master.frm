VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form master 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
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
            Picture         =   "Master.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Master.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Master.frx":6B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Master.frx":CDCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   12091
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
Attribute VB_Name = "master"
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
Set nod = tv.Nodes.Add(mtp, tvwChild, "l" & temp2![ID], tmp, 1)


'**************Ledger Values
If temp3.State = 1 Then temp3.Close
temp3.Open "select * from ledgers where undergroup=" & temp2![ID], conn
Do While temp3.EOF = False
tmp = temp3![nameis]
'MsgBox temp2![ID]
Set nod = tv.Nodes.Add("l" & temp2![ID], tvwChild, "Ledger" & temp3![ID], tmp, 4)
temp3.MoveNext
Loop





temp2.MoveNext
Loop





temp1.MoveNext
Loop




tm.MoveNext
Loop
End Sub

