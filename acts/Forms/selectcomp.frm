VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form selectcomp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton6 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Shut Down"
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
      MICON           =   "selectcomp.frx":0000
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
      Left            =   120
      Top             =   6720
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
            Picture         =   "selectcomp.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "selectcomp.frx":62B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "selectcomp.frx":BAA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1560
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "selectcomp.frx":11D42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   615
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      MICON           =   "selectcomp.frx":17FDC
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
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6990
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   9172
            MinWidth        =   9172
            Object.ToolTipText     =   "Data File Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            TextSave        =   "4/21/2006"
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "View Large Icons"
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      MICON           =   "selectcomp.frx":17FF8
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
   Begin MSComctlLib.ListView lv 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Financial Year"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Company Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Select"
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
      MICON           =   "selectcomp.frx":18014
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "View Small Icons"
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      MICON           =   "selectcomp.frx":18030
      ALIGN           =   1
      IMGLST          =   "ImageList2"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "View List"
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      MICON           =   "selectcomp.frx":1804C
      ALIGN           =   1
      IMGLST          =   "ImageList3"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   4335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Names"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Select Company"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "selectcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rsp As New ADODB.Recordset
rsp.Open "select * from companyinfo order by companyname", conn1
Dim tm As ListItem
lv.ListItems.Clear

Do While rsp.EOF = False
Set tm = lv.ListItems.Add(, , rsp![CompanyName], 1, 1)
tm.SubItems(1) = " - " & Format(rsp![sdate], "dd-mm-yyyy") & " - " & Format(rsp![edate], "dd-mm-yyyy")
tm.SubItems(2) = " " & rsp![ID]
rsp.MoveNext
Loop


End Sub

Private Sub LaVolpeButton1_Click()
If conn.State = 1 Then conn.Close
Label3.Caption = Trim(lv.SelectedItem.Text)
Label3.Refresh
Dim pt As New ADODB.Recordset
pt.Open "select * from companyinfo where id=" & Trim(lv.SelectedItem.SubItems(2)), conn1
DoEvents
compname = pt![CompanyName]
address = pt![address]
city = pt![city]
dates = pt![sdate]
datet = pt![edate]
filepath = IIf(IsNull(pt![Filename]), "", pt![Filename])
MDIForm1.StatusBar1.Panels(1).Text = compname
MDIForm1.StatusBar1.Panels(1).ToolTipText = compname
MDIForm1.StatusBar1.Panels(2).Text = "FY : " & Format(dates, "yyyy") & "-" & Format(datet, "yyyy")
MDIForm1.StatusBar1.Panels(2).ToolTipText = "FY : " & Format(dates, "dd-mmm-yyyy") & "-" & Format(datet, "dd-mmm-yyyy")
MDIForm1.StatusBar1.Panels(3).Text = filepath
MDIForm1.StatusBar1.Panels(3).ToolTipText = filepath

conn.CursorLocation = adUseClient
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\sonnet.mdb"
cmd.ActiveConnection = conn

Unload Me
'ShowCompInfo
End Sub

Private Sub LaVolpeButton2_Click()
lv.View = lvwIcon
End Sub

Private Sub LaVolpeButton3_Click()
lv.View = lvwSmallIcon
End Sub

Private Sub LaVolpeButton4_Click()
lv.View = lvwReport
End Sub

Private Sub LaVolpeButton5_Click()
Form3.Show
End Sub

Private Sub LaVolpeButton6_Click()
Unload Me
End Sub

Private Sub lv_Click()
Dim pt As New ADODB.Recordset
pt.Open "select * from companyinfo where id=" & Trim(lv.SelectedItem.SubItems(2)), conn1
sb.Panels(1).Text = IIf(IsNull(pt![Filename]), "", pt![Filename])
sb.Panels(1).ToolTipText = IIf(IsNull(pt![Filename]), "", pt![Filename])
End Sub

Private Sub lv_DblClick()
LaVolpeButton1_Click
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then LaVolpeButton1_Click
End Sub
