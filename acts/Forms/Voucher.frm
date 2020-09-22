VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Voucher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Entry"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Voucher.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   7560
      TabIndex        =   35
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   60817409
      CurrentDate     =   38204
   End
   Begin VB.TextBox dated 
      Height          =   285
      Left            =   5760
      TabIndex        =   34
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Voucher.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   8235
      TabIndex        =   25
      Top             =   0
      Width           =   8295
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5520
         TabIndex        =   33
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Voucher entry"
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
         TabIndex        =   26
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Please make sure you have selected the right Voucher type, you wont be able to change the voucher type after save."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   8175
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Voucher.frx":3B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Voucher.frx":9DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Voucher.frx":FB40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Voucher.frx":15332
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Voucher.frx":1564C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Voucher.frx":1B26E
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox remar 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   5280
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo part 
      Height          =   315
      Left            =   1170
      TabIndex        =   5
      Top             =   1800
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12648384
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox dr_cr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   7080
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "D"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox amt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   6000
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1170
      ScaleHeight     =   585
      ScaleWidth      =   4785
      TabIndex        =   10
      Top             =   2115
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label l1 
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
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.TextBox code 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1060
   End
   Begin VB.TextBox vno 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mf1 
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      BackColorUnpopulated=   192
      GridLinesFixed  =   0
      GridLinesUnpopulated=   1
      SelectionMode   =   1
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin LVbuttons.LaVolpeButton save 
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Voucher.frx":1B28A
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton delete 
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Voucher.frx":1B2A6
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "5"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   4560
      TabIndex        =   32
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Voucher.frx":1B2C2
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Dr/Cr"
      Height          =   255
      Left            =   7080
      TabIndex        =   31
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Amount(Rs.)"
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Ledger Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   29
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label13 
      Caption         =   "A/c Code"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label10 
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
      Left            =   6600
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Rs."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Rs."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Rs."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   5040
      Width           =   1200
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Remarks"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Dated"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Voucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ROWNO, CHK As Integer

Private Sub amt_GotFocus()
amt.SelStart = 0
amt.SelLength = Len(amt.Text)

End Sub

Private Sub amt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
dr_cr.SetFocus
End If
End Sub

Private Sub amt_LostFocus()
p1.Visible = False
End Sub

Private Sub code_GotFocus()
code.SelStart = 0
code.SelLength = Len(code.Text)

p1.Visible = False
End Sub

Private Sub code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
part.SetFocus
End If
End Sub

Private Sub code_LostFocus()
If Len(code.Text) = 0 Then Exit Sub
part.BoundText = code.Text
l1.Caption = "Current Balance For the account : "
p1.Visible = True
If GetLedgerAmt(code.Text) > 0 Then
l2.Caption = "Debit : Rs. " & GetLedgerAmt(code.Text)
l2.ForeColor = &HFF0000
Else
l2.Caption = "Credit : Rs. " & GetLedgerAmt(code.Text)
l2.ForeColor = &HFF&
End If
If GetLedgerAmt(code.Text) = 0 Then l2.Caption = "NIL"

End Sub

Private Sub Command1_Click()
got_Click
End Sub

Private Sub dated_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
code.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub dated_LostFocus()
If Format$(dated.Text, "yyyy/mm/dd") < Format$(dates, "yyyy/mm/dd") Or Format$(dated.Text, "yyyy/mm/dd") > Format$(datet, "yyyy/mm/dd") Then
MsgBox "Invalid Date......" & Chr(13) & "Date Must Be Between " & Format(dates, "dd/mm/yyyy") & " to " & Format(datet, "dd/mm/yyyy") & Chr(13) & "You have entered : " & Format(dated.Text, "dd/mm/yyyy"), vbCritical, "Error"
dated.SetFocus
End If

End Sub

Private Sub dated_Validate(Cancel As Boolean)

'Or Format(dated.Text, "dd/mm/yyyy") > Format(datet, "dd/mm/yyyy") Then
'
'Cancel = True
'Else
'Cancel = fase
'End If


End Sub

Private Sub delete_Click()
mm = PrintVoucher(12, "Cash")

End Sub

Private Sub dr_cr_GotFocus()
dr_cr.SelStart = 0
dr_cr.SelLength = Len(dr_cr.Text)
End Sub

Private Sub dr_cr_KeyPress(KeyAscii As Integer)
If KeyAscii = 49 Then
KeyAscii = 0
dr_cr.Text = "D"
End If
If KeyAscii = 50 Then
dr_cr.Text = "C"
KeyAscii = 0
End If
If KeyAscii = 13 Then
p1.Visible = False
KeyAscii = 0
If code.Text <> Empty Or Val(code.Text) <> 0 Or Val(amt.Text) <> 0 Then
If CHK = 0 Then ROWNO = mf1.Rows - 1
mf1.TextMatrix(ROWNO, 0) = code.Text
mf1.TextMatrix(ROWNO, 1) = part.Text
mf1.TextMatrix(ROWNO, 2) = amt.Text
mf1.TextMatrix(ROWNO, 3) = dr_cr.Text
code.SetFocus
If CHK = 0 Then mf1.Rows = mf1.Rows + 1
CHK = 0
Label3.Caption = "Add Mode"
End If
chkbal
a = 0
dr1 = 0
cr1 = 0
Do While mf1.Rows > a
If mf1.TextMatrix(a, 3) = "D" Then dr1 = Val(dr1) + Val(mf1.TextMatrix(a, 2))
If mf1.TextMatrix(a, 3) = "C" Then cr1 = Val(cr1) + Val(mf1.TextMatrix(a, 2))
a = a + 1
Loop

bal1 = Val(dr1) - Val(cr1)
Label8.Caption = Format(dr1, "0.00")
Label9.Caption = Format(cr1, "0.00")
Label10.Caption = Format(bal1, "0.00")
If dr1 > cr1 Then
amt.Text = dr1 - cr1
dr_cr = "C"
ElseIf cr1 > dr1 Then
amt.Text = cr1 - dr1
dr_cr = "D"
End If
End If
If bal1 = 0 Then
remar.SetFocus
save.Enabled = True
Else
save.Enabled = False
End If
End Sub



Private Sub dr_cr_LostFocus()
p1.Visible = False
code.Text = Empty
End Sub

Private Sub dr_cr_Validate(Cancel As Boolean)
If dr_cr.Text = "D" Or dr_cr.Text = "C" Then
Cancel = False
Else
Cancel = True
End If


End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
DTPicker1_Click
End Sub

Private Sub DTPicker1_Change()
DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
dated.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
dated.SetFocus
End Sub

Private Sub Form_Load()
Combo2.AddItem "New Record"
Combo2.AddItem "First Record"
Combo2.AddItem "Last Record"
Combo2.AddItem "Next Record"
Combo2.AddItem "Previous Record"
Combo2.AddItem "Find"
Combo2.ListIndex = 0

Combo1.AddItem "Cash"
Combo1.ListIndex = 0


CHK = 0
Label3.Caption = "Add Mode"
 mf1.ColWidth(0, 0) = code.Width
 mf1.ColWidth(1, 0) = part.Width
  mf1.ColWidth(2, 0) = amt.Width
 mf1.ColWidth(3, 0) = dr_cr.Width
 
Dim rt As New ADODB.Recordset
rt.Open "select * from ledgers order by nameis", conn
Set part.DataSource = rt
Set part.RowSource = rt
part.ListField = "NAMEIS"
part.BoundColumn = "ID"
End Sub

Private Sub Label11_Click()
MsgBox "D"
End Sub

Private Sub Label3_Click()
If Label3.Caption = "Add Mode" Then
Label3.Caption = "Edit Mode"
ROWNO = mf1.Row
CHK = 1
Else
CHK = 0
Label3.Caption = "Add Mode"
End If
End Sub

Private Sub LaVolpeButton1_Click()
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
vno_Change
End Sub

Private Sub got_Click()
If Combo2.Text = "New Record" Then
Dim rsave As New ADODB.Recordset
rsave.Open "select max(vno) as vno from vouchmst where vtype='" & Combo1.Text & "'", conn
If rsave.EOF = False Then
vno.Text = IIf(IsNull(rsave![vno]), 1, rsave![vno] + 1)
Else
vno.Text = 1
End If
dated.Text = Format(Now(), "dd-mm-yyyy")
dated.SetFocus
End If
Dim frec As New ADODB.Recordset
If Combo2.Text = "First Record" Then

frec.Open "select min(vno) as vno from vouchmst where vtype='" & Combo1.Text & "'", conn
If IsNull(frec![vno]) = False Then
vno.Text = frec![vno]
vno_Change
End If
End If

If Combo2.Text = "Last Record" Then
frec.Open "select max(vno) as vno from vouchmst where vtype='" & Combo1.Text & "'", conn
If IsNull(frec![vno]) = False Then
vno.Text = frec![vno]
vno_Change
End If
End If

If Combo2.Text = "Previous Record" Then
frec.Open "select max(vno) as vno from vouchmst where vtype='" & Combo1.Text & "' and vno<" & vno.Text, conn
If IsNull(frec![vno]) = False Then
vno.Text = frec![vno]
vno_Change
End If
End If
If Combo2.Text = "Next Record" Then

frec.Open "select min(vno) as vno from vouchmst where vtype='" & Combo1.Text & "' and vno>" & Val(vno.Text), conn
If IsNull(frec![vno]) = False Then
vno.Text = frec![vno]
vno_Change
End If
End If
If Combo2.Text = "Find" Then
tt = InputBox("Enter Dated in DD/MM/YYYY Format", "Find Voucher", 0)
frec.Open "select vno from vouchmst where vtype='" & Combo1.Text & "' and dated=#" & Format(tt, "mm/dd/yyyy") & "#", conn
If IsNull(frec![vno]) = False Then
If frec.EOF = False Then
vno.Text = frec![vno]
vno_Change
End If
End If
End If
End Sub

Private Sub mf1_Click()
p1.Visible = False
End Sub

Private Sub mf1_DblClick()
ROWNO = mf1.Row
MsgBox "You are in Edit Mode", vbInformation, "Mode Change"
CHK = 1
Label3.Caption = "Edit Mode"
End Sub

Private Sub part_Click(Area As Integer)
p1.Visible = False
End Sub

Private Sub part_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then amt.SetFocus
End Sub

Private Sub part_LostFocus()
code.Text = part.BoundText
If p1.Visible = False Then
code_LostFocus
End If
End Sub

Private Sub Picture2_Click()
got_Click
End Sub

Private Sub remar_Change()
If KeyAscii = 13 Then save.SetFocus
End Sub

Private Sub save_Click()
Dim rfind As New ADODB.Recordset
rfind.Open "select * from vouchmst where vno=" & vno.Text & " and vtype='" & Combo1.Text & "'", conn
If rfind.EOF = False Then
Dim rsave1 As New ADODB.Recordset
'Inserting in Vouchmast
rsave1.Open "update vouchmst set dated='" & _
Format(dated.Text, "dd/mm/yyyy") & "',vtype='" & _
Combo1.Text & "',remarks='" & Mid(remar.Text, 1, 149) & "' where vno=" & vno.Text & " and vtype='" & Combo1.Text & "'", conn

rsave1.Open "delete from vouchdat where vno=" & vno.Text & " and vtype='" & Combo1.Text & "'", conn

'Inserting data in Vouchdat file
a = 0
Do While a <> mf1.Rows
If Len(mf1.TextMatrix(a, 0)) > 0 And Val(mf1.TextMatrix(a, 2)) > 0 Then
rsave1.Open "insert into vouchdat values(" & vno.Text & ",'" & Format(dated.Text, "dd/mm/yyyy") & "','" & _
mf1.TextMatrix(a, 0) & "'," & mf1.TextMatrix(a, 2) & ",'" & mf1.TextMatrix(a, 3) & "','" & Combo1.Text & "')", conn
End If
a = a + 1
Loop

save.Enabled = False



Exit Sub
End If


Dim rsave As New ADODB.Recordset
'Inserting in Vouchmast
rsave.Open "insert into vouchmst values(" & vno.Text & ",'" & _
Format(dated.Text, "dd/mm/yyyy") & "','" & _
Combo1.Text & "','" & remar.Text & "')", conn

'Inserting data in Vouchdat file
a = 0
Do While a <> mf1.Rows
If Len(mf1.TextMatrix(a, 0)) > 0 And Val(mf1.TextMatrix(a, 2)) > 0 Then
rsave.Open "insert into vouchdat values(" & vno.Text & ",'" & Format(dated.Text, "dd/mm/yyyy") & "','" & _
mf1.TextMatrix(a, 0) & "'," & mf1.TextMatrix(a, 2) & ",'" & mf1.TextMatrix(a, 3) & "','" & Combo1.Text & "')", conn
End If
a = a + 1
Loop
rsave.Open "select max(vno) as vno from vouchmst where vtype='" & Combo1.Text & "'", conn
If rsave.EOF = False Then
vno.Text = rsave![vno] + 1
Else
vno.Text = 1
End If
dated.SetFocus
save.Enabled = False
End Sub

Private Sub vno_Change()
vno_KeyPress 13
End Sub

Private Sub vno_GotFocus()
p1.Visible = False
End Sub

Public Sub chkbal()

End Sub

Private Sub vno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
mf1.Rows = 0
code.Text = Empty
amt.Text = Empty
remar.Text = ""
dr_cr = "D"
Dim rmt As New ADODB.Recordset
rmt.Open "select idno,ledgers.nameis,amount,dr_cr,dated from vouchdat,ledgers where vouchdat.idno=ledgers.id and vno=" & Val(vno.Text) & " and vtype='" & Combo1.Text & "'", conn
a = 0
mf1.Rows = rmt.RecordCount + 1
If rmt.EOF = False Then dated.Text = Format(rmt![dated], "dd-mm-yyyy")
Do While rmt.EOF = False
mf1.TextMatrix(a, 0) = rmt![idno]
mf1.TextMatrix(a, 1) = rmt![nameis]
mf1.TextMatrix(a, 2) = rmt![amount]
mf1.TextMatrix(a, 3) = rmt![dr_cr]
a = a + 1
rmt.MoveNext
Loop
If rmt.State = 1 Then rmt.Close
rmt.Open "select * from vouchmst where vtype='" & Combo1.Text & "' and vno=" & Val(vno.Text), conn
If rmt.EOF = False Then remar.Text = rmt![remarks]
End If
End Sub
