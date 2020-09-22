VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form utility 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reports"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Utility.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   8955
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utility.frx":6F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utility.frx":D1DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utility.frx":12DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utility.frx":18A20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Open With Notepad"
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Utility.frx":1ECBA
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
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10186
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   15000
      TextRTF         =   $"Utility.frx":1ECD6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Send to Printer"
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Utility.frx":1ED56
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   615
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Close Ledger"
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Utility.frx":1ED72
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "Send to Floppy"
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Utility.frx":1ED8E
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
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
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "General Ledger"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
tt = Ledger("1", Date, Date)
pp = GetLedgerAmt(1)
MsgBox pp
End Sub

Private Sub Form_Resize()
If Me.Width < 5000 Then Exit Sub
RichTextBox1.Top = 700
RichTextBox1.Left = 0
RichTextBox1.Width = Me.Width - 200
RichTextBox1.Height = Me.Height - 1100
LaVolpeButton4.Left = Me.Width - (200 + LaVolpeButton4.Width + 10)
Label1.Width = Me.Width - (2580 + LaVolpeButton4.Width + 10)
Label2.Width = Me.Width - (2580 + LaVolpeButton4.Width + 10)
End Sub

Private Sub LaVolpeButton1_Click()
Shell ("notepad.exe " & RichTextBox1.Filename), vbMaximizedFocus
End Sub

Private Sub LaVolpeButton4_Click()
Unload Me
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
