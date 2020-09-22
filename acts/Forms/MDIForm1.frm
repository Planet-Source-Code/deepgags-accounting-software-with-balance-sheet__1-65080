VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00DAC2A5&
   Caption         =   "MDIForm1"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8340
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":626A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":92EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F586
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15820
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BF14
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C1A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C438
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":226D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25754
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2B376
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30B68
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu master12 
      Caption         =   "Pogramme"
      Begin VB.Menu master2 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Programme|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:195|Gradient}"
      End
      Begin VB.Menu accountmaster 
         Caption         =   "{img:7}Account Master"
         Begin VB.Menu dd 
            Caption         =   "{SIDEBAR:TEXT|CAPTION:Programme|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:195|Gradient}"
         End
         Begin VB.Menu acledcr 
            Caption         =   "{img:8}Account Ledger Creation"
         End
         Begin VB.Menu vacled 
            Caption         =   "{img:9}View Account Ledgers"
         End
      End
      Begin VB.Menu logout 
         Caption         =   "{img:3}Louout"
      End
      Begin VB.Menu createcomp 
         Caption         =   "{img:4}Create Company"
      End
      Begin VB.Menu changecompany 
         Caption         =   "{img:4}Change Company"
      End
      Begin VB.Menu shutcompany 
         Caption         =   "{Img:5}Shut Company"
      End
      Begin VB.Menu exit 
         Caption         =   "{img:6}Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu vouchers 
      Caption         =   "Vouchers"
      Begin VB.Menu c 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Vouchers|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:195|Gradient}"
      End
      Begin VB.Menu ve 
         Caption         =   "{img:13}Voucher Entry"
      End
      Begin VB.Menu voucherprint 
         Caption         =   "{img:11}Voucher Printing"
      End
      Begin VB.Menu searchVoucher 
         Caption         =   "{img:12}Search Voucher"
      End
      Begin VB.Menu typeofvouchers 
         Caption         =   "{img:10}Type of Vouchers"
      End
   End
   Begin VB.Menu reports 
      Caption         =   "Reports"
      Begin VB.Menu g1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Reports|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:195|Gradient}"
      End
      Begin VB.Menu generalledger 
         Caption         =   "General Ledger"
      End
      Begin VB.Menu cashbook 
         Caption         =   "Cash Book"
      End
      Begin VB.Menu bankbook 
         Caption         =   "Bank Book"
      End
      Begin VB.Menu combinedbook 
         Caption         =   "Combined Book"
      End
      Begin VB.Menu trialbalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu bsheet 
         Caption         =   "Profit && Loss Account"
      End
      Begin VB.Menu balancesheet 
         Caption         =   "Balance Sheet"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "Tools"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub createcomp_Click()
company.Show
End Sub

Private Sub generalledger_Click()
Ledger.Show
End Sub

Private Sub MDIForm_Load()
SetMenus hwnd, ImageList1

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'ReleaseMenus hwnd, ImageList1
End Sub

Private Sub searchVoucher_Click()
searchvoucher1.Show
End Sub

Private Sub ve_Click()
Voucher.Show
End Sub
