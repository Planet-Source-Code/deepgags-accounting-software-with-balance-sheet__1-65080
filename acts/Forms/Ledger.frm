VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Ledger 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7275
   ClientLeft      =   345
   ClientTop       =   45
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include Naration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   2175
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Exit"
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
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Ledger.frx":0000
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ledger.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ledger.frx":5C3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Dates"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox idno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Default         =   -1  'True
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Show"
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
         MICON           =   "Ledger.frx":BED8
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin MSMask.MaskEdBox edate 
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yyyy"
         Mask            =   "##-##-####"
         PromptChar      =   "#"
      End
      Begin MSMask.MaskEdBox sdate 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mm-yyyy"
         Mask            =   "##-##-####"
         PromptChar      =   "#"
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Hide"
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
         MICON           =   "Ledger.frx":BEF4
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "2"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ledger ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   10610
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      Text            =   ""
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Select Ledger Name and Press Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Ledgers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
idno.Text = DataCombo1.BoundText
Frame1.Visible = True
sdate.SetFocus
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
idno.Text = DataCombo1.BoundText
Frame1.Visible = True
sdate.SetFocus
End If
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub edate_GotFocus()
edate.SelStart = 0
edate.SelLength = 10
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Dim rms As New ADODB.Recordset
rms.Open "select id,nameis from ledgers order by nameis", conn
Set DataCombo1.DataSource = rms
Set DataCombo1.RowSource = rms
 DataCombo1.ListField = "nameis"
 DataCombo1.BoundColumn = "id"
 sdate.Text = Format(dates, "DD-MM-yyyy")
 edate.Text = Format(datet, "DD-MM-yyyy")
End Sub

Private Sub LaVolpeButton1_Click()
tt = LedgerShow(DataCombo1.BoundText, Format(sdate.Text, "DD/MM/YYYY"), Format(edate.Text, "DD/MM/YYYY"), Check1.Value)
Frame1.Visible = False
End Sub

Private Sub LaVolpeButton2_Click()
Frame1.Visible = False
End Sub

Private Sub LaVolpeButton3_Click()
Unload Me
End Sub

Private Sub sdate_GotFocus()
sdate.SelStart = 0
sdate.SelLength = 10
End Sub
