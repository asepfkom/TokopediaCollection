VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetailPayment 
   Appearance      =   0  'Flat
   BackColor       =   &H00ABE18E&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Detail Payment"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstPayment 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmDetailPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim RS As New ADODB.Recordset
Dim Listitem As Listitem

TXT_X = 70
LstPayment.ColumnHeaders.ADD 1, , "No", 10 * TXT_X
LstPayment.ColumnHeaders.ADD 2, , "Tahun", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 3, , "Jan", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 4, , "Feb", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 5, , "Mar", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 6, , "Apr", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 7, , "Mei", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 8, , "Jun", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 9, , "Jul", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 10, , "Aug", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 11, , "Sep", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 12, , "Okt", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 13, , "Nop", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 14, , "Des", 15 * TXT_X

'Cmdsql = " SELECT m.custid, m.tahun, COALESCE(m.""1"", 0) AS ""Jan"", COALESCE(m.""2"", 0) AS ""Feb"", COALESCE(m.""3"", 0) AS ""Mar"", COALESCE(m.""4"", 0) AS ""Apr"", COALESCE(m.""5"", 0) AS ""Mei"", COALESCE(m.""6"", 0) AS ""Jun"", COALESCE(m.""7"", 0) AS ""Jul"", COALESCE(m.""8"", 0) AS ""Aug"", COALESCE(m.""9"", 0) AS ""Sep"", COALESCE(m.""10"", 0) AS ""Okt"", COALESCE(m.""11"", 0) AS ""Nop"", COALESCE(m.""12"", 0) AS ""Des"""
'Cmdsql = Cmdsql + "  FROM crosstab('select custid, date_part(''year'',paydate)as tahun, date_part(''month'',paydate) as bulan, sum(payment) as payment from tbllunas"
'Cmdsql = Cmdsql + "   where custid=''" + FrmCC_Colection.lblCustId.Caption + "'' group by custid, tahun,bulan order by  tahun,custid'::text, 'select m from generate_series(1,12) m'::text) m(custid text, ""tahun"" text,  ""1"" numeric, ""2"" numeric, ""3"" numeric, ""4"" numeric, ""5"" numeric, ""6"" numeric, ""7"" numeric, ""8"" numeric, ""9"" numeric, ""10"" numeric, ""11"" numeric, ""12"" numeric);"

Cmdsql = "SELECT m.custid, m.tahun, COALESCE(m.""1"", 0) AS ""Jan"","
Cmdsql = Cmdsql + " COALESCE(m.""2"", 0) AS ""Feb"", COALESCE(m.""3"", 0) AS ""Mar"", COALESCE(m.""4"", 0) AS ""Apr"", COALESCE(m.""5"", 0) AS ""Mei"","
Cmdsql = Cmdsql + " COALESCE(m.""6"", 0) AS ""Jun"", COALESCE(m.""7"", 0) AS ""Jul"", COALESCE(m.""8"", 0) AS ""Aug"", COALESCE(m.""9"", 0) AS ""Sep"","
Cmdsql = Cmdsql + " COALESCE(m.""10"", 0) AS ""Okt"", COALESCE(m.""11"", 0) AS ""Nop"", COALESCE(m.""12"", 0) AS ""Des""  "
Cmdsql = Cmdsql + " FROM crosstab('select date_part(''year'',paydate)as tahun,custid,date_part(''month'',paydate) as bulan, "
Cmdsql = Cmdsql + " sum(payment) as payment from tbllunas where custid=''" + FrmCC_Colection.lblcustid.Caption + "''"
Cmdsql = Cmdsql + " group by tahun,custid,bulan order by  tahun,custid'::text, 'select m from generate_series(1,12) m'::text) m(custid text, ""tahun"" text,  ""1"" numeric, ""2"" numeric, ""3"" numeric, ""4"" numeric, ""5"" numeric, ""6"" numeric, ""7"" numeric, ""8"" numeric, ""9"" numeric, ""10"" numeric, ""11"" numeric, ""12"" numeric);"

Set RS = New ADODB.Recordset
RS.CursorLocation = adUseClient
RS.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not RS.EOF
 Set Listitem = LstPayment.ListItems.ADD(, , RS.Bookmark)
     Listitem.SubItems(1) = IIf(IsNull(RS!Tahun), "", RS!Tahun)
     Listitem.SubItems(2) = Format(IIf(IsNull(RS!Jan), "", RS!Jan), "#,###,###")
     Listitem.SubItems(3) = Format(IIf(IsNull(RS!feb), "", RS!feb), "#,###,###")
     Listitem.SubItems(4) = Format(IIf(IsNull(RS!Mar), "", RS!Mar), "#,###,###")
     Listitem.SubItems(5) = Format(IIf(IsNull(RS!Apr), "", RS!Apr), "#,###,###")
     Listitem.SubItems(6) = Format(IIf(IsNull(RS!Mei), "", RS!Mei), "#,###,###")
     Listitem.SubItems(7) = Format(IIf(IsNull(RS!Jun), "", RS!Jun), "#,###,###")
     Listitem.SubItems(8) = Format(IIf(IsNull(RS!Jul), "", RS!Jul), "#,###,###")
     Listitem.SubItems(9) = Format(IIf(IsNull(RS!Aug), "", RS!Aug), "#,###,###")
     Listitem.SubItems(10) = Format(IIf(IsNull(RS!Sep), "", RS!Sep), "#,###,###")
     Listitem.SubItems(11) = Format(IIf(IsNull(RS!Okt), "", RS!Okt), "#,###,###")
     Listitem.SubItems(12) = Format(IIf(IsNull(RS!Nop), "", RS!Nop), "#,###,###")
     Listitem.SubItems(13) = Format(IIf(IsNull(RS!des), "", RS!des), "#,###,###")
     RS.MoveNext
Wend

Set RS = Nothing
End Sub
