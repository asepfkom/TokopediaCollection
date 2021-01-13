VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frmdeltbllunas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Deletion"
   ClientHeight    =   6180
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   10065
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10065
   Begin VB.Frame Frame4 
      BackColor       =   &H80000014&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4440
      TabIndex        =   13
      Top             =   1920
      Width           =   5535
      Begin MSComctlLib.ListView ListView1 
         Height          =   3030
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   5345
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame form_trade 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   9855
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000FF&
         Caption         =   "Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   4215
      Begin MSComctlLib.ListView LvPayment 
         Height          =   3030
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   5345
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   320
         Width           =   1935
      End
      Begin TDBDate6Ctl.TDBDate tgl_mulai1 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "Frmdeltbllunas.frx":0000
         Caption         =   "Frmdeltbllunas.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frmdeltbllunas.frx":0184
         Keys            =   "Frmdeltbllunas.frx":01A2
         Spin            =   "Frmdeltbllunas.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd, mmm yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tgl_akhir1 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   1200
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "Frmdeltbllunas.frx":0228
         Caption         =   "Frmdeltbllunas.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frmdeltbllunas.frx":03AC
         Keys            =   "Frmdeltbllunas.frx":03CA
         Spin            =   "Frmdeltbllunas.frx":0428
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd, mmm yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__, ___ ____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   765
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frmdeltbllunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    tgl_mulai1.Value = Null
    tgl_akhir1.Value = Null
    Text1.Enabled = True
    Text2.Enabled = True
    tgl_mulai1.Enabled = True
    tgl_akhir1.Enabled = True
    Command2.Enabled = True
    LvPayment.ListItems.CLEAR
    Command5.Enabled = False
End Sub

Private Sub search()
    LvPayment.ListItems.CLEAR
    q = "select * from tbllunas where 1=1"
    If Text1.Text <> "" Then
        q = q + " and custid = '" + Text1.Text + "' "
    End If
    a = cnull(Format(tgl_mulai1.Value, "yyyy-mm-dd"))
    B = cnull(Format(tgl_akhir1.Value, "yyyy-mm-dd"))
    If a <> "" And a <> "" Then
        q = q + " and date(paydate) between '" + a + "' and '" + B + "'"
    End If
    If Text2.Text <> "" Then
        q = q + " and payment = " & Text2.Text & ""
    End If
    
    Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
            Set listItem = LvPayment.ListItems.ADD(, , cnull(r("custid")))
                 listItem.SubItems(1) = Format(cnull(r("paydate")), "yyyy-mm-dd")
                 listItem.SubItems(2) = cnull(r("payment"))
            r.MoveNext
        Wend
    If r.RecordCount = 0 Then
        MsgBox "Data Tidak Ditemukan"
    Else
        Command5.Enabled = True
    End If
End Sub

Private Sub headerpluslog()
    LvPayment.ColumnHeaders.CLEAR
    ListView1.ColumnHeaders.CLEAR

    LvPayment.ColumnHeaders.ADD 1, , "Customer ID", 10 * 150
    LvPayment.ColumnHeaders.ADD 2, , "Pay Date", 20 * 60
    LvPayment.ColumnHeaders.ADD 3, , "Payment", 20 * 60
       
    ListView1.ColumnHeaders.ADD 1, , "Customer ID", 10 * 150
    ListView1.ColumnHeaders.ADD 2, , "Pay Date", 20 * 60
    ListView1.ColumnHeaders.ADD 3, , "Payment", 20 * 60
    ListView1.ColumnHeaders.ADD 4, , "Tanggal Delete", 20 * 60
    
    ListView1.ListItems.CLEAR
    q = "select * from backup_tbllunas_app order by delete_date desc"
    Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not r.EOF
            Set listItem = ListView1.ListItems.ADD(, , cnull(r("custid")))
                 listItem.SubItems(1) = Format(cnull(r("paydate")), "yyyy-mm-dd")
                 listItem.SubItems(2) = cnull(r("payment"))
                 listItem.SubItems(3) = cnull(r("delete_date"))
            r.MoveNext
        Wend
End Sub

Private Sub Command2_Click()
    Text1.Enabled = False
    Text2.Enabled = False
    tgl_mulai1.Enabled = False
    tgl_akhir1.Enabled = False
    Command2.Enabled = False
    
    a = cnull(Format(tgl_mulai1.Value, "yyyy-mm-dd"))
    B = cnull(Format(tgl_akhir1.Value, "yyyy-mm-dd"))
    
    If Text1.Text = "" And (a = "" And B = "") And Text2.Text = "" Then
        MsgBox "Harap Filter dengan benar"
        Exit Sub
    End If
    Call search
    
End Sub

Private Sub Command3_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If LvPayment.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To LvPayment.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LvPayment.ColumnHeaders(col)
        Next
     
        For Row = 2 To LvPayment.ListItems.Count + 1
            For col = 1 To LvPayment.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = "'" + LvPayment.ListItems(Row - 1).Text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = LvPayment.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
zzz:
        MsgBox "No data to export", vbInformation, Me.Caption
    End If
End Sub

Private Sub Command5_Click()
    m_msgbox = MsgBox("Yakin Akan di Delete ??", vbYesNo + vbExclamation, "Aplikasi")
    If m_msgbox = vbNo Then
        Exit Sub
    End If
    
    Call deletepayment
    
    MsgBox "Delete berhasil"
    Call Command1_Click
    Call Form_Load
End Sub

Private Sub Form_Load()
    Call headerpluslog
End Sub

Private Sub deletepayment()
    q = "insert into backup_tbllunas_app (custid,paydate,payment,agent,fieldname,datafrom,id,sts,tglinsert,curr_bal,curr_pri,id_negoptp) " & vbCrLf
    q = q + "select * from tbllunas where 1=1"
    If Text1.Text <> "" Then
        q = q + " and custid = '" + Text1.Text + "' "
    End If
    a = cnull(Format(tgl_mulai1.Value, "yyyy-mm-dd"))
    B = cnull(Format(tgl_akhir1.Value, "yyyy-mm-dd"))
    If a <> "" And a <> "" Then
        q = q + " and date(paydate) between '" + a + "' and '" + B + "'"
    End If
    If Text2.Text <> "" Then
        q = q + " and payment = " & Text2.Text & ""
    End If
    q = q + " ;" & vbCrLf
    
    q = q + " update backup_tbllunas_app set delete_date = now() where delete_date is null;" & vbCrLf
    
    q = q + "Delete from tbllunas where 1=1"
    If Text1.Text <> "" Then
        q = q + " and custid = '" + Text1.Text + "' "
    End If
    a = cnull(Format(tgl_mulai1.Value, "yyyy-mm-dd"))
    B = cnull(Format(tgl_akhir1.Value, "yyyy-mm-dd"))
    If a <> "" And a <> "" Then
        q = q + " and date(paydate) between '" + a + "' and '" + B + "'"
    End If
    If Text2.Text <> "" Then
        q = q + " and payment = " & Text2.Text & ""
    End If
    q = q + ";"
    
    M_OBJCONN.Execute q
End Sub
