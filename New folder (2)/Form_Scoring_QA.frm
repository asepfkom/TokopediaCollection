VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Scoring_QA 
   Caption         =   "Form Scoring QA"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame9 
      BackColor       =   &H80000003&
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.ComboBox cbEdit 
         Height          =   315
         ItemData        =   "Form_Scoring_QA.frx":0000
         Left            =   90
         List            =   "Form_Scoring_QA.frx":000A
         TabIndex        =   8
         Top             =   870
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   90
         TabIndex        =   7
         Top             =   1230
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   45
         TabIndex        =   1
         Top             =   4785
         Width           =   12525
         Begin VB.ComboBox CmbActionQA 
            BackColor       =   &H00FFC0FF&
            Height          =   315
            ItemData        =   "Form_Scoring_QA.frx":0017
            Left            =   4410
            List            =   "Form_Scoring_QA.frx":0021
            TabIndex        =   2
            Top             =   -360
            Visible         =   0   'False
            Width           =   1590
         End
         Begin Threed.SSCommand cmdSave 
            Height          =   405
            Left            =   135
            TabIndex        =   3
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   714
            _Version        =   196610
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Save"
         End
         Begin Threed.SSCommand cmdExport 
            Height          =   405
            Left            =   1230
            TabIndex        =   4
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   714
            _Version        =   196610
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Export"
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Action"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3135
            TabIndex        =   5
            Top             =   -345
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fg 
         Height          =   4650
         Left            =   30
         TabIndex        =   6
         Top             =   135
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   8202
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Form_Scoring_QA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim old_Col, old_Row As Integer
'
'Private Sub cmdSave_Click()
'    cmdSave.Enabled = False
'    Call save_all_data
'    'Call isi_excel
'    cmdSave.Enabled = True
'End Sub
'
'Private Sub save_all_data()
'    Dim rs As ADODB.Recordset
'    Dim STSRQL, MHWERE, jenis_produk, action, f_action, tgl_action As String
'
'    Strsql = "select * from tbl_submit_qa limit 1"
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    M_OBJCONN.execute "delete from tbl_submit_qa where custid = " + CStr(Val(Form_customer.txtidmgm.text)) '+ "'"
'    With fg
'        For i = 1 To 9
'            If .TextMatrix(i, 5) <> Empty Then
'                rs.AddNew
'                rs!CustId = Form_customer.txtidmgm.text
'                rs!question = .TextMatrix(i, 1)
'                rs!result = .TextMatrix(i, 2)
'                rs!keterangan = .TextMatrix(i, 3)
'                rs!Note = .TextMatrix(i, 4)
'                rs!Nilai = .TextMatrix(i, 5)
'                rs.update
'            End If
'        Next i
'    End With
'    msg ("done")
'    Set rs = Nothing
'End Sub
'
'Private Sub cbEdit_LostFocus()
'    Dim rs As New ADODB.Recordset
'    Dim CMDSQL As String
'
'    With fg
'        .TextMatrix(old_Row, old_Col) = cbEdit.text
'        If old_Col = 2 Then
'            rs.CursorLocation = adUseClient
'            CMDSQL = " select keterangan,nilai from scoring where question='" + .TextMatrix(old_Row, 1) + "' AND result='" + cbEdit.text + "'"
'            rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If rs.RecordCount > 0 Then
'                .TextMatrix(old_Row, 3) = cnull(rs!keterangan)
'                .TextMatrix(old_Row, 5) = cnull(rs!Nilai)
'                If cbEdit.text = "2" Then
'                    .TextMatrix(old_Row, 4) = "Sudah sesuai dengan Parameter QM"
'                End If
'            End If
'        End If
'    End With
'End Sub
'
'Public Sub SetFlex()
'    Dim Baris As Integer
'    Dim r As Single
'    Dim i As Single
'    Dim rs As New ADODB.Recordset
'    Dim CMDSQL As String
'    rs.CursorLocation = adUseClient
'
'    CMDSQL = "select"
'    CMDSQL = CMDSQL + vbCrLf + " a.question,"
'    CMDSQL = CMDSQL + vbCrLf + " b.result,"
'    CMDSQL = CMDSQL + vbCrLf + " b.keterangan,"
'    CMDSQL = CMDSQL + vbCrLf + " b.nilai,"
'    CMDSQL = CMDSQL + vbCrLf + " b.note"
'    CMDSQL = CMDSQL + vbCrLf + " from ("
'    CMDSQL = CMDSQL + vbCrLf + " select max(id),question from scoring group by question order by max"
'    CMDSQL = CMDSQL + vbCrLf + " ) as a left join ("
'    CMDSQL = CMDSQL + vbCrLf + " select * from tbl_submit_qa where custid = " + CStr(Val(FrmCC_Colection.lblCustId))
'    CMDSQL = CMDSQL + vbCrLf + " ) b on (b.question=a.question) order by max"
'
'    rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    With fg
'        .TextMatrix(0, 0) = "No"
'        .TextMatrix(0, 1) = "Question"
'        .TextMatrix(0, 2) = "Result"
'        .TextMatrix(0, 3) = "Keterangan"
'        .TextMatrix(0, 4) = "Note"
'        .TextMatrix(0, 5) = "Nilai"
'
'        .ColWidth(0) = 450
'        .ColWidth(1) = 6200
'        .ColWidth(2) = 800
'        .ColWidth(3) = 6200
'        .ColWidth(4) = 600
'        .ColWidth(5) = 500
'
'        For Baris = 1 To rs.RecordCount
'            DoEvents
'            .Rows = Baris + 1
'            .TextMatrix(Baris, 0) = Baris
'            .TextMatrix(Baris, 1) = cnull(rs!question)
'            .TextMatrix(Baris, 2) = cnull(rs!result)
'            .TextMatrix(Baris, 3) = cnull(rs!keterangan)
'            .TextMatrix(Baris, 4) = cnull(rs!Note)
'            .TextMatrix(Baris, 5) = cnull(rs!Nilai)
'            rs.MoveNext
'        Next Baris
'    End With
'    Set rs = Nothing
'End Sub
'
'Private Sub fg_Click()
'    With fg
'        txtEdit.Visible = False: cbEdit.Visible = False
'        old_Col = .col: old_Row = .row
'        Select Case fg.col
'        Case 4
'            txtEdit.Left = fg.CellLeft + fg.Left
'            txtEdit.Top = fg.CellTop + fg.Top
'            txtEdit.Width = fg.CellWidth
'            txtEdit.Height = fg.CellHeight
'            txtEdit.text = fg.text
'            txtEdit.SelStart = 0
'            txtEdit.SelLength = Len(txtEdit.text)
'            txtEdit.Visible = True
'            txtEdit.SetFocus
'        Case 2
'            cbEdit.Left = fg.CellLeft + fg.Left
'            cbEdit.Top = fg.CellTop + fg.Top
'            cbEdit.Width = fg.CellWidth
'            cbEdit.text = fg.text
'            cbEdit.SelStart = 0
'            cbEdit.SelLength = Len(txtEdit.text)
'            cbEdit.Visible = True
'            cbEdit.SetFocus
'        End Select
'
'        If (fg.row = 4 And fg.col = 2) Or (fg.row = 5 And fg.col = 2) Then
'            cbEdit.Left = fg.CellLeft + fg.Left
'            cbEdit.Top = fg.CellTop + fg.Top
'            cbEdit.Width = fg.CellWidth
'            cbEdit.SelStart = 0
'            cbEdit.SelLength = Len(txtEdit.text)
'            cbEdit.Visible = True
'            cbEdit.clear
'            cbEdit.AddItem "Yes"
'            cbEdit.AddItem "Half"
'            cbEdit.AddItem "No"
'            cbEdit.text = fg.text
'            cbEdit.SetFocus
'        ElseIf (fg.row = 1 And fg.col = 2) Or (fg.row = 2 And fg.col = 2) Or (fg.row = 3 And fg.col = 2) Or (fg.row = 6 And fg.col = 2) Or (fg.row = 7 And fg.col = 2) Or (fg.row = 8 And fg.col = 2) Or (fg.row = 9 And fg.col = 2) Then
'            cbEdit.Left = fg.CellLeft + fg.Left
'            cbEdit.Top = fg.CellTop + fg.Top
'            cbEdit.Width = fg.CellWidth
'            cbEdit.SelStart = 0
'            cbEdit.SelLength = Len(txtEdit.text)
'            cbEdit.Visible = True
'            cbEdit.clear
'            cbEdit.AddItem "Yes"
'            cbEdit.AddItem "No"
'            cbEdit.text = fg.text
'            cbEdit.SetFocus
'        End If
'    End With
'End Sub
'
'Private Sub Form_Load()
'    With fg
'    .Rows = 9
'    .Cols = 6
'    End With
'    SetFlex
'End Sub
'
'Private Sub txtEdit_LostFocus()
'    With fg
'        .TextMatrix(old_Row, old_Col) = txtEdit.text
'    End With
'End Sub
'
'Private Sub CmdExport_Click()
'    Call isi_excel
'End Sub
'Private Sub isi_excel()
'   Dim Filename_000 As String
'    Dim objExcel As New Excel.Application
'    Dim objWorkbook As New Excel.Workbook
'    Dim objWorkSheet As New Excel.Worksheet
'    Dim str As String
'    Dim resultset As Variant
'    Dim result As Variant
'    Dim leads As String
'    Dim nmleads As String
'    objExcel.Visible = True
'
'    Dim rs3 As ADODB.Recordset
'
'    Set rs3 = New ADODB.Recordset
'    rs3.CursorLocation = adUseClient
'    rs3.Open "SELECT tbluser_groupspvcode, tbluser_ketgroupspv from tbluser where tbluser_name='" & Form_customer.txtagentname.text & "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdTex
'    If rs3.EOF = False Then
'        leads = rs3!tbluser_groupspvcode
'        nmleads = rs3!tbluser_ketgroupspv
'    End If
'
'    'Set objWorkbook = objExcel.Workbooks.open("http://192.168.20.2/crm/script/Template%20QA-DNN.xlsx")
'    Set objWorkbook = objExcel.Workbooks.Open("C:\Template\Template QA-DNN.xlsx")
'
'    Set objWorkbook = objExcel.Workbooks(1)
'    objWorkbook.Activate
'    objExcel.DisplayAlerts = False
'
'    Set objWorkSheet = objWorkbook.Worksheets("QA Scoring Sample")
'    objWorkSheet.Activate
'
'    objWorkSheet.Application.Range("E2").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = Form_customer.txtagentname.text
'    objWorkSheet.Application.Range("E3").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = MDIForm1.TxtNama.text
'    objWorkSheet.Application.Range("E4").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = Format(FungsiWaktuServer, "DD-MM-YYYY")
'    objWorkSheet.Application.Range("E5").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = nmleads
'    objWorkSheet.Application.Range("E6").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = Form_customer.txt1(21).text
'    objWorkSheet.Application.Range("E7").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = "http://192.168.20.1/vl/?token=7c2c35ab736f3c061ce78b663fe534b5&uid=" & Form_customer.LVHistoryCall.SelectedItem.SubItems(17) & ""
'    objWorkSheet.Application.Range("E8").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = Form_customer.txt_shopname.text
'
'    objWorkSheet.Application.Range("E13").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(1, 2)
'    objWorkSheet.Application.Range("E15").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(2, 2)
'    objWorkSheet.Application.Range("E17").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(3, 2)
'    objWorkSheet.Application.Range("E19").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(4, 2)
'    objWorkSheet.Application.Range("E21").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(5, 2)
'    objWorkSheet.Application.Range("E23").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(6, 2)
'    objWorkSheet.Application.Range("E25").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(7, 2)
'    objWorkSheet.Application.Range("E27").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(8, 2)
'    objWorkSheet.Application.Range("E29").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(9, 2)
'
'    objWorkSheet.Application.Range("E14").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(1, 4)
'    objWorkSheet.Application.Range("E16").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(2, 4)
'    objWorkSheet.Application.Range("E18").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(3, 4)
'    objWorkSheet.Application.Range("E20").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(4, 4)
'    objWorkSheet.Application.Range("E22").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(5, 4)
'    objWorkSheet.Application.Range("E24").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(6, 4)
'    objWorkSheet.Application.Range("E26").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(7, 4)
'    objWorkSheet.Application.Range("E28").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(8, 4)
'    objWorkSheet.Application.Range("E30").Select
'    objWorkSheet.Application.ActiveCell.FormulaR1C1 = fg.TextMatrix(9, 4)
'
'Filename_000 = "D:\Sample_QA_" & Form_customer.txtagentname.text & "_" & Form_customer.txt_shopname & "_" & Form_customer.txtidmgm.text & ".xlsx"
'objWorkbook.SaveAs FileName:=Filename_000, FileFormat:=51, ReadOnlyRecommended:=True, CreateBackup:=False
''objWorkbook.Close
''objExcel.Quit
'
'Exit Sub
'a:
'If err.Number = 1004 Then
'    msg "Nama File Sudah ada, Mohon Dihapus data terlebih dahulu"
'Else
'    MsgBox err.Description
'End If
'
'Set rs3 = Nothing
'End Sub
'
