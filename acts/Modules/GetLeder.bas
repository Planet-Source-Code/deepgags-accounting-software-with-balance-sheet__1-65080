Attribute VB_Name = "Gethead"


Public Function GetHeadAmt(mcode As Integer)

Dim rsmp As New ADODB.Recordset
Dim rsum As New ADODB.Recordset
Dim rsum1 As New ADODB.Recordset
rsmp.Open "select * from groups where id=" & mcode, conn
Dim t As Integer
Dim ID, opbal, dr, cr, closbal As Currency
Dim a As Integer
a = 0

Dim mt As New ADODB.Recordset

Do While rsmp.EOF = False
If mt.State = 1 Then mt.Close

mt.Open "select * from ledgers where undergroup=" & rsmp![ID], conn
Do While mt.EOF = False
If rsum.State = 1 Then rsum.Close
If rsum1.State = 1 Then rsum1.Close
rsum.Open "select (sum(amount)) as cr from vouchdat where dr_cr='C' and idno='" & mt![ID] & "'", conn
rsum1.Open "select (sum(amount)) as dr from vouchdat where dr_cr='D' and idno='" & mt![ID] & "'", conn
DoEvents
closbal = closbal + ((mt![opbalance] + IIf(IsNull(rsum1![dr]), 0, rsum1![dr]) - IIf(IsNull(rsum![cr]), 0, rsum![cr])))
a = a + 1
mt.MoveNext
Loop
rsmp.MoveNext
Loop


GetHeadAmt = closbal
a = 0


End Function

Public Function GetLedgerAmt(mcode As String)
Dim rsmp As New ADODB.Recordset
Dim rsum As New ADODB.Recordset
Dim rsum1 As New ADODB.Recordset
rsmp.Open "select ledgers.id as id,opbalance from oye where ledgers.id='" & mcode & "'", conn
Dim t As Integer
Dim ID, opbal, dr, cr, closbal As Currency
Dim a As Integer
a = 0
Do While rsmp.EOF = False
If rsum.State = 1 Then rsum.Close
rsum.Open "select (sum(amount)) as cr from vouchdat where dr_cr='C' and idno='" & rsmp![ID] & "'", conn
rsum1.Open "select (sum(amount)) as dr from vouchdat where dr_cr='D' and idno='" & rsmp![ID] & "'", conn
DoEvents
closbal = closbal + ((rsmp![opbalance] + IIf(IsNull(rsum1![dr]), 0, rsum1![dr]) - IIf(IsNull(rsum![cr]), 0, rsum![cr])))
a = a + 1
rsmp.MoveNext
Loop
GetLedgerAmt = closbal
a = 0
End Function
Public Function LedgerShow(mcode As String, sdate As Date, edate As Date, narration As Boolean)
Dim t As New ADODB.Recordset
Dim ln As Currency
t.Open "select nameis from ledgers where id='" & mcode & "'", conn
h1 = "GENERAL LEDGER"
h2 = "For : " & t![nameis]
h3 = "From : " & Format(sdate, "dd-mmm-yyyy") & " To : " & Format(edate, "dd-mmm-yyyy")
h4 = String(89, "-")
h5 = "DATED      VNO.  PARTICULARS                     DEBIT RS.   CREDIT RS.     BALANCE DR/CR"
Dim fs As New FileSystemObject
Dim st As TextStream
Set st = fs.CreateTextFile(App.Path & "\ledger.txt", True)
Dim rtotal As Currency
'Calculating Opening Balance
Dim opbl As New ADODB.Recordset
opbl.Open "select sum(opbalance) as opb1,(select sum(amount) from vouchdat where idno='" & _
mcode & "' and dated<#" & sdate & "# and dr_cr='D') as d,(select sum(amount) from vouchdat where idno='" & _
mcode & "' and dated<#" & sdate & "# and dr_cr='C') as c from ledgers where id='" & mcode & "'", conn
opb = IIf(IsNull(opbl![opb1]), 0, opbl![opb1]) + IIf(IsNull(opbl![d]), 0, opbl![d]) - IIf(IsNull(opbl![C]), 0, opbl![C])
'Gathering Ledger Values

Dim rtm As New ADODB.Recordset
rtm.Open "SELECT a.VNO, a.DATED,remarks,a.IDNO, a.vtype,amount,dr_cr FROM vouchdat AS a, vouchmst AS b " & _
" WHERE a.VNO=b.vno AND a.IDNO='" & mcode & "' AND a.vtype=b.vtype and a.dated >=#" & sdate & "# and a.dated <=#" & edate & "# order by a.dated,dr_cr desc", conn


Dim rtm1 As New ADODB.Recordset
'rtm1.Open "select nameis from ledgers where id=(select top " & _
'" 1 idno from vouchdat where vno=a.vno and " & _
'" dr_cr<>rtm!dr_cr and vtype=a.vtype)", conn
'MsgBox "Three"


pb1.ProgressBar1.Max = rtm.RecordCount + 1
pb1.ProgressBar1.Min = 0
'Printing Opening Balances
st.WriteLine h1
st.WriteLine h2
st.WriteLine h3
st.WriteLine h4
st.WriteLine h5
st.WriteLine h4
ln = ln + 6
If opb > 0 Then st.WriteLine Format(sdate, "dd-mm-yyyy") & Space(7) & "Opening Balance" & Space(30 - Len("Opening Balance")) & Space(12 - Len(Format(opb, "0.00"))) & Format(opb, "0.00") & Space(12) & Space(12 - Len(Format(opb, "0.00"))) & Format(opb, "0.00") & IIf(opb > 0, " Dr", " Cr")
If opb < 0 Then st.WriteLine Format(sdate, "dd-mm-yyyy") & Space(7) & "Opening Balance" & Space(30 - Len("Opening Balance")) & Space(12) & Space(12 - Len(Format(opb, "0.00"))) & Format(opb, "0.00") & Space(12 - Len(Format(opb, "0.00"))) & Format(opb, "0.00") & IIf(opb > 0, " Dr", " Cr")
If opb = 0 Then st.WriteLine Format(sdate, "dd-mm-yyyy") & Space(7) & "Opening Balance"
pb1.Show
ln = ln + 1
rtotal = opb
Do While rtm.EOF = False
If rtm1.State = 1 Then rtm1.Close
rtm1.Open "select nameis as particular from ledgers where id=(select top 1 idno from vouchdat where vno=" & rtm![vno] & " and dr_cr<>'" & rtm![dr_cr] & "' and vtype='" & rtm![vtype] & "')", conn

pb1.ProgressBar1.Value = rtm.AbsolutePosition
DoEvents
pp = Format(rtm![dated], "dd-mm-yyyy") & Space(6 - Len(rtm![vno])) & rtm![vno] & " " & Mid(rtm1![particular], 1, 30) & Space(30 - Len(Mid(rtm1![particular], 1, 30)))
If rtm![dr_cr] = "D" Then
pp = pp + Space(12 - Len(Format(rtm![amount], "0.00"))) & Format(rtm![amount], "0.00")
rtotal = rtotal + rtm![amount]
pp = pp + Space(12)
Else
pp = pp + Space(12)
pp = pp + Space(12 - Len(Format(rtm![amount], "0.00"))) & Format(rtm![amount], "0.00")
rtotal = rtotal - rtm![amount]
End If
DoEvents
pp = pp + Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
If rtotal > 0 Then pp = pp + " Dr"
If rtotal < 0 Then pp = pp + " Cr"
If rtotal = 0 Then pp = pp + "   "
st.WriteLine pp

If narration = True Then
rmks = Split(rtm![remarks], Chr(13))
a = 0
Do While a < UBound(rmks) + 1
If Asc(Left(rmks(a), 1)) = 10 Then
st.WriteLine Space(17) & Right(rmks(a), Len(rmks(a)) - 1)
Else
st.WriteLine Space(17) & rmks(a)
End If
a = a + 1
Loop
st.WriteBlankLines 1
End If

rtm.MoveNext

ln = ln + 1
If ln >= 60 Then
st.WriteLine Chr(12)
ln = 0
st.WriteLine h1
st.WriteLine h2
st.WriteLine h3
st.WriteLine h4
st.WriteLine h5
st.WriteLine h4
ln = 6
End If

Loop
Unload pb1
st.Close
utility.Label2.Caption = t![nameis]
utility.RichTextBox1.Filename = App.Path & "\ledger.txt"
utility.Show
End Function

Public Function PrintVoucher(vno As Currency, vtype As String)
Dim rstm As New ADODB.Recordset
rstm.Open "select a.remarks,a.dated,a.vno,a.vtype,b.idno,c.nameis,b.amount," & _
"b.dr_cr from vouchdat b,vouchmst a,ledgers c where a.vno=b.vno and a.vtype='" & vtype & "' and b.idno=c.id and a.vno=" & vno & " order by dr_cr desc", conn

Dim fs As New FileSystemObject
Dim st As TextStream
Set st = fs.CreateTextFile("d:\acts\ledger.txt", True)
st.WriteLine Space((79 - Len(compname)) / 2) & compname
st.WriteLine Space((79 - Len(address)) / 2) & address
st.WriteLine Space((79 - Len(city)) / 2) & city
st.WriteBlankLines 1
st.WriteLine Space((79 - Len(rstm![vtype] & " Voucher")) / 2) & rstm![vtype] & " Voucher"
st.WriteLine "Voucher No. : " & rstm![vno] & Space(19 - Len(rstm![vno])) & Space(20) & "Voucher Dated : " & Format(rstm![dated], "dd-mm-yyyy")
st.WriteBlankLines 1
st.WriteLine String(79, "-")
st.WriteLine "SrNo. Code       Ledger Head                               Dr.        Cr.        "
st.WriteLine String(79, "-")
a = 1
Dim dr, cr As Currency
remark = rstm![remarks]
Do While rstm.EOF = False
pp = Space(3 - Len(a)) & a & "   " & rstm![idno] & Space(10 - Len(rstm![idno])) & " " & rstm![nameis] & Space(35 - Len(rstm![nameis]))
If rstm![dr_cr] = "D" Then
pp = pp & Space(13 - Len(Format(rstm![amount], "0.00"))) & Format(rstm![amount], "0.00")
dr = dr + rstm![amount]
End If

If rstm![dr_cr] = "C" Then
pp = pp & Space(13) & Space(14 - Len(Format(rstm![amount], "0.00"))) & Format(rstm![amount], "0.00")
cr = cr + rstm![amount]
End If

st.WriteLine pp
rstm.MoveNext
a = a + 1
Loop


st.WriteLine Space(41) & String(79 - 41, "-")
st.WriteLine Space(41) & "Total Rs. :" & Space(13 - Len(Format(dr, "0.00"))) & Format(dr, "0.00") & Space(14 - Len(Format(cr, "0.00"))) & Format(cr, "0.00")
st.WriteLine Space(41) & String(79 - 41, "-")
st.WriteLine remark
st.WriteLine Space(64) & "Auth. Signatory"
st.Close
rstm.MoveFirst
'utility.Label2.Caption = rstm![nameis]
utility.RichTextBox1.Filename = "d:\acts\ledger.txt"
utility.Show


End Function



Public Function cashbook(mcode As String, sdate As Date, edate As Date)
Dim t As New ADODB.Recordset
Dim ln As Currency
t.Open "select nameis from ledgers where id='" & mcode & "'", conn
h1 = "CASH BOOK"
h2 = "" '= "For : " & t![nameis]
h3 = "From : " & Format(sdate, "dd-mmm-yyyy") & " To : " & Format(edate, "dd-mmm-yyyy")
h4 = String(79, "-")
h5 = "DATED      VNO.  PARTICULARS                           Receipts     Payments"
Dim fs As New FileSystemObject
Dim st As TextStream
Set st = fs.CreateTextFile(App.Path & "\cashbook.txt", True)
Dim rtotal As Currency
'Calculating Opening Balance
Dim opbl As New ADODB.Recordset
opbl.Open "select sum(opbalance) as opb1,(select sum(amount) from vouchdat where idno='" & _
mcode & "' and dated<#" & Format(sdate, "mm/dd/yyyy") & "# and dr_cr='D') as d,(select sum(amount) from vouchdat where idno='" & _
mcode & "' and dated<#" & Format(sdate, "mm/dd/yyyy") & "# and dr_cr='C') as c from ledgers where id='" & mcode & "'", conn
opb = IIf(IsNull(opbl![opb1]), 0, opbl![opb1]) + IIf(IsNull(opbl![d]), 0, opbl![d]) - IIf(IsNull(opbl![C]), 0, opbl![C])

'Gathering Ledger Values
Dim rtm As New ADODB.Recordset
Dim rtm1 As New ADODB.Recordset
'rtm.Open "select nameis from ledgers where id=(select top 1  idno from vouchdat where vno=a.vno and dr_cr<>a.dr_cr and dated=a.dated and vtype=a.vtype)) AS particular FROM vouchdat AS a, vouchmst AS b WHERE (((a.VNO)=[b].[vno]) AND ((a.DATED)=[b].[dated]) AND ((a.IDNO)='" & mcode & "') AND ((a.vtype)=[b].[vtype])) and a.dated >=#" & Format(sdate, "mm/dd/yyyy") & "# and a.dated <=#" & Format(edate, "mm/dd/yyyy") & "#", conn

rtm.Open "SELECT a.VNO, a.DATED,remarks,a.IDNO, a.vtype,amount,dr_cr FROM vouchdat AS a, vouchmst AS b " & _
" WHERE a.VNO=b.vno AND a.IDNO='" & mcode & "' AND a.vtype=b.vtype and a.dated >=#" & sdate & "# and a.dated <=#" & edate & "# order by a.dated,dr_cr desc", conn

'rtm.Open "SELECT a.VNO, a.DATED,b.remarks, (select nameis from ledgers where id=(select top 1  idno from vouchdat where vno=a.vno and dr_cr<>a.dr_cr and dated=a.dated and vtype=a.vtype)) AS particular, a.IDNO, a.vtype,amount,dr_cr FROM vouchdat AS a, vouchmst AS b WHERE (((a.VNO)=[b].[vno]) AND ((a.DATED)=[b].[dated]) AND ((a.IDNO)='" & mcode & "') AND ((a.vtype)=[b].[vtype])) and a.dated >=#" & Format(sdate, "mm/dd/yyyy") & "# and a.dated <=#" & Format(edate, "mm/dd/yyyy") & "#", conn

pb1.ProgressBar1.Max = rtm.RecordCount + 1
pb1.ProgressBar1.Min = 0
'Printing Opening Balances
st.WriteLine h1
st.WriteLine h2
st.WriteLine h3
st.WriteLine h4
st.WriteLine h5
st.WriteLine h4
ln = ln + 6
If opb > 0 Then st.WriteLine Format(sdate, "dd-mm-yyyy") & Space(7) & "Balance B/F    " & Space(30 - Len("Opening Balance")) & Space(17 - Len(Format(opb, "0.00"))) & Format(opb, "0.00") & Space(12)
If opb < 0 Then st.WriteLine Format(sdate, "dd-mm-yyyy") & Space(7) & "Balance B/F    " & Space(30 - Len("Opening Balance")) & Space(17) & Space(12 - Len(Format(opb, "0.00"))) & Format(opb, "0.00")
If opb = 0 Then st.WriteLine Format(sdate, "dd-mm-yyyy") & Space(7) & "Balance B/F    "
pb1.Show
ln = ln + 1
rtotal = opb
drdatetot = 0
crdatetot = 0
Do While rtm.EOF = False
dt = rtm![dated]
If rtm.EOF = True Then Exit Do
st.WriteLine Format(dt, "dd-mm-yyyy")
Do While dt = rtm![dated]
If rtm1.State = 1 Then rtm1.Close
rtm1.Open "select nameis as particular from ledgers where id=(select top 1 idno from vouchdat where vno=" & rtm![vno] & " and dr_cr<>'" & rtm![dr_cr] & "' and vtype='" & rtm![vtype] & "')", conn


pb1.ProgressBar1.Value = rtm.AbsolutePosition
DoEvents
pp = Space(6 - Len(rtm![vno])) & rtm![vno] & " " & rtm1![particular] & Space(45 - Len(rtm1![particular]))
If rtm![dr_cr] = "D" Then
pp = pp & Space(12 - Len(Format(rtm![amount], "0.00"))) & Format(rtm![amount], "0.00")
rtotal = rtotal + rtm![amount]
drdatetot = drdatetot + rtm![amount]
pp = pp & Space(12)
Else
pp = pp & Space(12)
pp = pp & Space(12 - Len(Format(rtm![amount], "0.00"))) & Format(rtm![amount], "0.00")
crdatetot = crdatetot + rtm![amount]
rtotal = rtotal - rtm![amount]
End If
DoEvents
'pp = pp '+ Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
'If rtotal > 0 Then pp = pp + " Dr"
'If rtotal < 0 Then pp = pp + " Cr"
'If rtotal = 0 Then pp = pp + "   "
st.WriteLine pp
rtm.MoveNext
If rtm.EOF = True Then Exit Do



rmks = Split(rtm![remarks], Chr(13))
a = 0
Do While a < UBound(rmks) + 1
If Asc(Left(rmks(a), 1)) = 10 Then
st.WriteLine Space(7) & Right(rmks(a), Len(rmks(a)) - 1)
Else
st.WriteLine Space(7) & rmks(a)
End If
a = a + 1
Loop

pp = Empty
ln = ln + 1
If ln >= 60 Then
st.WriteLine Chr(12)
ln = 0
st.WriteLine h1
st.WriteLine h2
st.WriteLine h3
st.WriteLine h4
st.WriteLine h5
st.WriteLine h4
ln = 6
End If
Loop
st.WriteLine Space(41) & String(38, "-")
st.WriteLine Space(41) & "Total Rs.  " & Space(12 - Len(Format(drdatetot, "0.00"))) & Format(drdatetot, "0.00") & Space(12 - Len(Format(drdatetot, "0.00"))) & Format(drdatetot, "0.00")
st.WriteLine Space(41) & String(38, "-")
'ap = Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
ap = ""
If rtotal > 0 Then
ap = ap & Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
ap = ap + " Dr"
End If
If rtotal < 0 Then
ap = ap & Space(12) & Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
ap = ap + " Cr"
End If
If rtotal = 0 Then
ap = ap & Space(12) & Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
ap = ap + "  "
End If
st.WriteLine Space(10) & "Balance C/F as on " & Format(dt, "dd-mmm-yyyy") & Space(13) & ap
ap = ""
Loop

Unload pb1
st.Close
utility.Label2.Caption = t![nameis]
utility.RichTextBox1.Filename = App.Path & "\cashbook.txt"
utility.Show
End Function


Public Function JournalBook(sdate As Date, edate As Date)
Dim t As New ADODB.Recordset
Dim ln As Currency
t.Open "select nameis from ledgers where id='" & mcode & "'", conn
h1 = "Journal BOOK"
h2 = "" '= "For : " & t![nameis]
h3 = "From : " & Format(sdate, "dd-mmm-yyyy") & " To : " & Format(edate, "dd-mmm-yyyy")
h4 = String(79, "-")
h5 = "DATED      VNO.  PARTICULARS                           Debit Rs.      Credit Rs."
Dim fs As New FileSystemObject
Dim st As TextStream
Set st = fs.CreateTextFile(App.Path & "\ledger.txt", True)
Dim rtotal As Currency
'Calculating Opening Balance
'Gathering Ledger Values
Dim rtm As New ADODB.Recordset
rtm.Open "SELECT a.VNO, a.DATED,b.remarks," & _
" (select nameis from ledgers where id=a.idno) AS particular, a.IDNO, a.vtype" & _
",amount,dr_cr FROM vouchdat AS a, vouchmst AS b WHERE a.VNO=b.vno AND a.DATED=b.dated AND a.vtype=b.vtype and a.dated >=#" & Format(sdate, "mm/dd/yyyy") & "# and a.dated <=#" & Format(edate, "mm/dd/yyyy") & "#", conn
pb1.ProgressBar1.Max = rtm.RecordCount + 1
pb1.ProgressBar1.Min = 0
'Printing Opening Balances
st.WriteLine h1
st.WriteLine h2
st.WriteLine h3
st.WriteLine h4
st.WriteLine h5
st.WriteLine h4
ln = ln + 6
pb1.Show
ln = ln + 1
rtotal = opb
drdatetot = 0
crdatetot = 0
Do While rtm.EOF = False
dt = rtm![vno]
'If rtm.EOF = True Then Exit Do
st.WriteLine Format(rtm![dated], "dd-mm-yyyy") & Space(10 - Len(rtm![vno])) & rtm![vno] & "  " & rtm![vtype]
Do While dt = rtm![vno]
pb1.ProgressBar1.Value = rtm.AbsolutePosition
DoEvents
pp = "   " & rtm![particular] & Space(47 - Len(rtm![particular]))
If rtm![dr_cr] = "D" Then
pp = pp & Space(14 - Len(Format(rtm![amount], "0.00"))) & Format(rtm![amount], "0.00")
rtotal = rtotal + rtm![amount]
drdatetot = drdatetot + rtm![amount]
pp = pp & Space(14)
Else
pp = pp & Space(14)
pp = pp & Space(14 - Len(Format(rtm![amount], "0.00"))) & Format(rtm![amount], "0.00")
crdatetot = crdatetot + rtm![amount]
rtotal = rtotal - rtm![amount]
End If
DoEvents
'pp = pp '+ Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
'If rtotal > 0 Then pp = pp + " Dr"
'If rtotal < 0 Then pp = pp + " Cr"
'If rtotal = 0 Then pp = pp + "   "
st.WriteLine pp
rmks = Split(rtm![remarks], Chr(13))
rtm.MoveNext
If rtm.EOF = True Then Exit Do
pp = Empty
ln = ln + 1
If ln >= 60 Then
st.WriteLine Chr(12)
ln = 0
st.WriteLine h1
st.WriteLine h2
st.WriteLine h3
st.WriteLine h4
st.WriteLine h5
st.WriteLine h4
ln = 6
End If

Loop

a = 0
Do While a < UBound(rmks) + 1
If Asc(Left(rmks(a), 1)) = 10 Then
st.WriteLine Space(3) & Right(rmks(a), Len(rmks(a)) - 1)
Else
st.WriteLine Space(3) & rmks(a)
End If
a = a + 1
Loop
st.WriteLine String(79, "-")
'ap = Space(12 - Len(Format(rtotal, "0.00"))) & Format(rtotal, "0.00")
Loop
st.WriteLine Space(41) & String(38, "-")
st.WriteLine Space(41) & "Total Rs.  " & Space(12 - Len(Format(drdatetot, "0.00"))) & Format(drdatetot, "0.00") & Space(14 - Len(Format(drdatetot, "0.00"))) & Format(drdatetot, "0.00")
st.WriteLine Space(41) & String(38, "-")

Unload pb1
st.Close
'utility.Label2.Caption = t![nameis]
utility.RichTextBox1.Filename = App.Path & "\ledger.txt"
utility.Show
End Function



Public Function ShowCompInfo()
infoform.Left = MDIForm1.Width - (infoform.Width + 100)
infoform.Top = MDIForm1.Height - (infoform.Height + 100)
infoform.Label4.Caption = IIf(IsNull(address), "", address)
infoform.Label2.Caption = compname
infoform.Label6.Caption = IIf(IsNull(city), "", city)
infoform.Label8.Caption = Format(dates, "dd-mmm-yyyy")
infoform.Label9.Caption = Format(datet, "dd-mmm-yyyy")
infoform.Label10.Caption = filepath
'infoform.Show , MDIForm1
infoform.Label10.ToolTipText = filepath
End Function
