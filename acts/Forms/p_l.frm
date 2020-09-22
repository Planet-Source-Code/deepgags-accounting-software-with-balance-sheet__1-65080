VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form P_L 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Balance Sheet"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "p_l.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridLines       =   0
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
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   240
      Width           =   10935
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount (Rs.)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Income"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount (Rs.)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Expenditures"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Profit && Loss Account"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   720
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
      Left            =   6240
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Loss"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   6240
      Width           =   1335
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
      Left            =   1560
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Profit "
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
      Left            =   2160
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Left            =   6960
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Left            =   2760
      TabIndex        =   3
      Top             =   6240
      Width           =   1575
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
      Left            =   6960
      TabIndex        =   2
      Top             =   6480
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
      Left            =   2760
      TabIndex        =   1
      Top             =   6480
      Width           =   1575
   End
End
Attribute VB_Name = "P_L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fs As New FileSystemObject
Dim st As TextStream
Set st = fs.CreateTextFile(App.Path & "\profitloss.txt")
Dim tp1 As New ADODB.Recordset
Dim tp As New ADODB.Recordset
MSHFlexGrid1.ColWidth(0, 0) = 2600
MSHFlexGrid1.ColWidth(1, 0) = 1200
MSHFlexGrid1.ColWidth(2, 0) = 1200
tp.Open "select groups_1.nameis,groups.nameis,ledgers.nameis,ledgers.id,opbalance from oye where childof=32 order by groups.nameis", conn
tp1.Open "select groups_1.nameis,groups.nameis,ledgers.nameis,ledgers.id,opbalance from oye where childof=33 order by groups.nameis", conn
If tp.RecordCount > tp1.RecordCount Then MSHFlexGrid1.Rows = tp.RecordCount + 10
If tp.RecordCount < tp1.RecordCount Then MSHFlexGrid1.Rows = tp1.RecordCount + 10
If tp.RecordCount = tp1.RecordCount Then MSHFlexGrid1.Rows = tp1.RecordCount + 10
R = 1
Dim mt As Currency
Dim totdr, totcr As Currency
Do While tp.EOF = False
a = tp![groups.nameis]
MSHFlexGrid1.TextMatrix(R, 0) = a
MSHFlexGrid1.Row = R
MSHFlexGrid1.Col = 0
MSHFlexGrid1.CellFontBold = True
mt = R
R = R + 1
tot = 0
'st.WriteLine a
Do While a = tp![groups.nameis]
mm = GetLedgerAmt(tp![ID])

MSHFlexGrid1.TextMatrix(R, 0) = "     " & tp![ledgers.nameis]
MSHFlexGrid1.TextMatrix(R, 1) = mm
tot = tot + mm
tp.MoveNext
R = R + 1
If tp.EOF = True Then Exit Do
Loop
MSHFlexGrid1.TextMatrix(mt, 2) = tot
totdr = totdr + tot
MSHFlexGrid1.Row = mt
MSHFlexGrid1.Col = 2
MSHFlexGrid1.CellFontBold = True

If tp.EOF = True Then Exit Do
'tp.MoveNext
Loop
'st.Close

'Income Side
R = 1
mt = 0

MSHFlexGrid1.ColWidth(3, 0) = 2600
MSHFlexGrid1.ColWidth(4, 0) = 1200
MSHFlexGrid1.ColWidth(5, 0) = 1200
Do While tp1.EOF = False
a = tp1![groups.nameis]
MSHFlexGrid1.TextMatrix(R, 3) = a
MSHFlexGrid1.Row = R
MSHFlexGrid1.Col = 3
MSHFlexGrid1.CellFontBold = True
mt = R
R = R + 1
tot = 0
'st.WriteLine a
Do While a = tp1![groups.nameis]
mm = GetLedgerAmt(tp1![ID])
MSHFlexGrid1.TextMatrix(R, 3) = "     " & tp1![ledgers.nameis]
MSHFlexGrid1.TextMatrix(R, 4) = IIf(tot > 0, -mm, Abs(mm))
tot = tot + mm
tp1.MoveNext
R = R + 1
If tp1.EOF = True Then Exit Do
Loop
MSHFlexGrid1.TextMatrix(mt, 5) = IIf(tot > 0, -tot, Abs(tot))
totcr = totcr + tot
MSHFlexGrid1.Row = mt
MSHFlexGrid1.Col = 5
MSHFlexGrid1.CellFontBold = True

If tp1.EOF = True Then Exit Do
'tp.MoveNext
Loop


Label1.Caption = Abs(totdr)
Label2.Caption = Abs(totcr)
If Abs(totdr) < Abs(totcr) Then
Label1.Caption = Abs(totcr)
Label3.Caption = Abs(totdr) - Abs(totcr)
Else
Label2.Caption = Abs(totdr)
Label4.Caption = Abs(totcr) - Abs(totdr)
End If
'utility.RichTextBox1.Filename = "d:\acts\profitloss.txt"
'utility.Show

End Sub

Private Sub Form_Load()
Command1_Click
End Sub

Private Sub LaVolpeButton1_Click()
p = 0
l = 0
p = Abs(Val(Label3.Caption))
l = Abs(Val(Label4.Caption))
'bsheet.Label3.Caption = Abs(Val(Label3.Caption))
'bsheet.Label4.Caption = Abs(Val(Label4.Caption))

bsheet.Show
End Sub
