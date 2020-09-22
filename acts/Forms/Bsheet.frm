VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form bsheet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10920
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Assets"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label10 
      Caption         =   "Libilities"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Balance Sheet"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   10935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   10935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rs."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rs."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
End
Attribute VB_Name = "bsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Liability Side
Dim rsm As New ADODB.Recordset
rsm.Open "select * from groups where childof=30", conn
'Set DataGrid1.DataSource = rsm
MSHFlexGrid1.Rows = 100
Dim t1 As New ADODB.Recordset
Dim t2 As New ADODB.Recordset
R = 1
MSHFlexGrid1.ColWidth(0, 0) = 2400
MSHFlexGrid1.ColWidth(1, 0) = 1200
MSHFlexGrid1.ColWidth(2, 0) = 1200

Do While rsm.EOF = False
If t1.State = 1 Then t1.Close
If rsm.EOF = False Then MSHFlexGrid1.TextMatrix(R, 0) = rsm![nameis]
mt = R
t1.Open "select * from groups where childof=" & rsm![ID], conn
'If t1.EOF = False Then MSHFlexGrid1.TextMatrix(R, 0) = t1![nameis]
R = R + 1
'Set DataGrid2.DataSource = t1
'If t1.EOF = True Then MsgBox rsm![ID]
a_tot = 0

Do While t1.EOF = False
't2.Open "select * from ledgers where childof=" & t1![ID], conn
MSHFlexGrid1.TextMatrix(R, 0) = "      " & t1![nameis]
gamt = GetHeadAmt(t1![ID])
MSHFlexGrid1.TextMatrix(R, 1) = gamt
a_tot = a_tot + gamt
R = R + 1
t1.MoveNext
Loop
If rsm![ID] = 5 Then
MSHFlexGrid1.TextMatrix(R, 0) = "Less : Drawings "
MSHFlexGrid1.TextMatrix(R, 1) = GetHeadAmt(41)
a_tot = a_tot + GetHeadAmt(41)
End If

MSHFlexGrid1.TextMatrix(mt, 2) = a_tot
MSHFlexGrid1.Row = mt
MSHFlexGrid1.Col = 0
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.Col = 2
MSHFlexGrid1.CellFontBold = True
R = R + 1
gtota = gtota + a_tot


rsm.MoveNext
Loop
gtota = (gtota - p)
gtota = (gtota + l)
'MsgBox p & " " & l
If Val(p) > 0 Then
MSHFlexGrid1.TextMatrix(R, 0) = "Net Profit Trf from P & L "
MSHFlexGrid1.TextMatrix(R, 2) = p
End If
If Val(Abs(l)) > 0 Then
MSHFlexGrid1.TextMatrix(R, 0) = "Net Loss Trf from P & L "
MSHFlexGrid1.TextMatrix(R, 2) = l
End If

Label1.Caption = (gtota)
Label1.Refresh
End Sub

Private Sub Command2_Click()
MsgBox GetHeadAmt(6)
End Sub

Private Sub Command3_Click()
'Assets Side
Dim rsm As New ADODB.Recordset
rsm.Open "select * from groups where childof=31", conn
'Set DataGrid1.DataSource = rsm
MSHFlexGrid1.Rows = 100
Dim t1 As New ADODB.Recordset
Dim t2 As New ADODB.Recordset
R = 1
MSHFlexGrid1.ColWidth(3, 0) = 2400
MSHFlexGrid1.ColWidth(4, 0) = 1200
MSHFlexGrid1.ColWidth(5, 0) = 1200
gtota = 0
Do While rsm.EOF = False
If t1.State = 1 Then t1.Close
If rsm.EOF = False Then MSHFlexGrid1.TextMatrix(R, 3) = rsm![nameis]
mt = R
t1.Open "select * from groups where childof=" & rsm![ID], conn
'If t1.EOF = True Then MsgBox rsm![ID]
'If t1.EOF = False Then MSHFlexGrid1.TextMatrix(R, 0) = t1![nameis]
R = R + 1
'Set DataGrid2.DataSource = t1
a_tot = 0
Do While t1.EOF = False
't2.Open "select * from ledgers where childof=" & t1![ID], conn
MSHFlexGrid1.TextMatrix(R, 3) = "      " & t1![nameis]
gamt = GetHeadAmt(t1![ID])
MSHFlexGrid1.TextMatrix(R, 4) = gamt
a_tot = a_tot + gamt
R = R + 1
t1.MoveNext
Loop
MSHFlexGrid1.TextMatrix(mt, 5) = a_tot
MSHFlexGrid1.Row = mt
MSHFlexGrid1.Col = 3
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.Col = 5
MSHFlexGrid1.CellFontBold = True
gtota = gtota + a_tot
rsm.MoveNext
Loop
Label2.Caption = (gtota)
End Sub

Private Sub Form_Load()
Command1_Click
Command3_Click
End Sub
