VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form_trade 
   BackColor       =   &H80000002&
   Caption         =   "Trading Form"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form5"
   ScaleHeight     =   6555
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data on TL"
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox cmb_team 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "form_trade.frx":0000
         Left            =   1200
         List            =   "form_trade.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "View"
         Height          =   375
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmd_trade 
         BackColor       =   &H0000FF00&
         Caption         =   "Trade"
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CheckBox cek_all_ptp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4350
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7673
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin VB.Label lbl_team 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Team  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbldata 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   5880
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data on Trade"
      Height          =   6495
      Left            =   9480
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "Clear"
         Height          =   375
         Left            =   6000
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_trade.frx":0004
         Left            =   1080
         List            =   "form_trade.frx":0006
         TabIndex        =   16
         Text            =   "Combo2"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Filter"
         Height          =   375
         Left            =   7200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_trade.frx":0008
         Left            =   1080
         List            =   "form_trade.frx":000A
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5880
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4350
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   7673
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
      Begin VB.CheckBox cek_all_payment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate tgl_mulai1 
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   360
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":000C
         Caption         =   "form_trade.frx":0124
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":0190
         Keys            =   "form_trade.frx":01AE
         Spin            =   "form_trade.frx":020C
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
         Left            =   6120
         TabIndex        =   19
         Top             =   360
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   661
         Calendar        =   "form_trade.frx":0234
         Caption         =   "form_trade.frx":034C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_trade.frx":03B8
         Keys            =   "form_trade.frx":03D6
         Spin            =   "form_trade.frx":0434
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
      Begin VB.Label Label5 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Date   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Status  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Team   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   5880
         Width           =   2655
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
Attribute VB_Name = "form_trade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cek_all_ptp_Click()
    Dim r As Integer
        
    If cek_all_ptp.Value = vbChecked Then
        If ListView1.ListItems.Count = 0 Then
            MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
            Exit Sub
        End If
        
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = True
        Next r
    Else
        For r = 1 To ListView1.ListItems.Count
            ListView1.ListItems(r).Checked = False
        Next r
    End If
End Sub

Private Sub isitrade()
    Dim query As String
    Dim rs As ADODB.Recordset

    ListView2.ListItems.CLEAR
    
    query = "select mgm.custid,mgm.name,mgm.agentlama,mgm.f_cek_new,temp_trade.tanggal_trader from mgm,temp_trade where mgm.custid = temp_trade.custid order by custid"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , rs("custid"))
             listItem.SubItems(1) = rs("name")
             listItem.SubItems(2) = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
             listItem.SubItems(3) = rs("f_cek_new")
             listItem.SubItems(4) = Format(rs("tanggal_trader"), "YYYY-MM-DD")
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
End Sub

Private Sub filtertrade()
    Dim query As String
    Dim rs As ADODB.Recordset

    ListView2.ListItems.CLEAR
    
    Tgl1 = Format(tgl_mulai1.Value, "YYYY-MM-DD")
    Tgl2 = Format(tgl_akhir1.Value, "YYYY-MM-DD")
    
    query = "select mgm.custid,mgm.name,mgm.agentlama,mgm.f_cek_new,temp_trade.tanggal_trader from mgm,temp_trade where mgm.custid = temp_trade.custid "
    If Combo1.Text <> "" Then
        query = query + " and mgm.agentlama = '" + Combo1.Text + "' "
    End If
    If Combo2.Text <> "" Then
        query = query + " and mgm.f_cek_new = '" + Combo2.Text + "' "
    End If
    If Tgl1 <> "" And Tgl2 <> "" Then
        query = query + " and date(temp_trade.tanggal_trader) between '" + Tgl1 + "' and '" + Tgl2 + "' "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query + " order by custid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , rs("custid"))
             listItem.SubItems(1) = rs("name")
             listItem.SubItems(2) = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
             listItem.SubItems(3) = rs("f_cek_new")
             listItem.SubItems(4) = Format(rs("tanggal_trader"), "YYYY-MM-DD")
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
End Sub

Private Sub combo1dah()
    Dim query As String
    Dim rs As ADODB.Recordset

    Combo1.CLEAR
    
    query = "select distinct (agentlama) from mgm,temp_trade where mgm.custid = temp_trade.custid order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Combo1.AddItem rs!agentlama
        rs.MoveNext
    Wend
    
    Combo2.CLEAR
    
    query = "select distinct (f_cek_new) from mgm,temp_trade where mgm.custid = temp_trade.custid order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Combo2.AddItem rs!f_cek_new
        rs.MoveNext
    Wend
End Sub

Private Sub cmd_trade_Click()
    Dim K, w, cek As Integer
    Dim query As String
    Dim rs As ADODB.Recordset
    If cmb_team.Text = "" Then
        MsgBox "Pilih team!"
        Exit Sub
    End If
    
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    If cek = 0 Then
        MsgBox "You Must Select a Data!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If

    For w = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(w).Checked = True Then
            CustId = ListView1.ListItems(w).Text
            Nama = ListView1.ListItems(w).ListSubItems(1)
            agent = ListView1.ListItems(w).ListSubItems(3)
            STATUS = ListView1.ListItems(w).ListSubItems(2)
            
            query = "INSERT INTO temp_trade (custid,name,agent,status_account) values ('" + CustId + "','" + Nama + "','" + agent + "','" + STATUS + "') ;" & vbCrLf
            query = query + "UPDATE mgm set agent = 'TRADE', agentlama = agent where custid = '" + CustId + "' ;"
            M_OBJCONN.Execute query
        End If
    Next w
    
    
    MsgBox "Data Traded"
    
    
    
    ListView2.ListItems.CLEAR
    
    query = "select mgm.custid,mgm.name,mgm.agentlama,mgm.f_cek_new,temp_trade.tanggal_trader from mgm,temp_trade where mgm.custid = temp_trade.custid order by custid"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView2.ListItems.ADD(, , rs("custid"))
             listItem.SubItems(1) = rs("name")
             listItem.SubItems(2) = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
             listItem.SubItems(3) = rs("f_cek_new")
             listItem.SubItems(4) = rs("tanggal_trader")
        rs.MoveNext
    Wend
    
    Label1.Caption = "Jumlah Data  : " & rs.RecordCount
    
    
    cek_all_ptp.Value = 0
    Call showmgmtl
    Call showtl
End Sub

Private Sub Command1_Click()
    If cmb_team.Text = "" Then
        MsgBox "Pilih team!"
        Exit Sub
    End If
    Call showmgmtl
End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    On Error GoTo zzz
    If ListView2.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView2.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView2.ColumnHeaders(col)
        Next
     
        For Row = 2 To ListView2.ListItems.Count + 1
            For col = 1 To ListView2.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = ListView2.ListItems(Row - 1).Text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + ListView2.ListItems(Row - 1).SubItems(col - 1)
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

Private Sub Command3_Click()
    Call filtertrade
End Sub

Private Sub Command4_Click()
    Combo1.Text = ""
    Combo2.Text = ""
    tgl_mulai1.Value = Null
    tgl_akhir1.Value = Null
End Sub

Private Sub Form_Load()
    Call showtl
    Call header
    Call isitrade
    Call combo1dah
End Sub

Private Sub showtl()
    Dim query As String
    Dim rs As ADODB.Recordset
    
    cmb_team.CLEAR
    
    query = "select distinct agent from mgm where agent ilike 'TL%' order by 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        cmb_team.AddItem rs!agent
        rs.MoveNext
    Wend
End Sub

Private Sub showmgmtl()
    Dim query As String
    Dim rs As ADODB.Recordset
    
    ListView1.ListItems.CLEAR
    
    query = "select * from mgm where agent = '" + cmb_team.Text + "' order by custid"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    While Not rs.EOF
        Set listItem = ListView1.ListItems.ADD(, , rs("custid"))
             listItem.SubItems(1) = rs("name")
             listItem.SubItems(2) = rs("f_cek_new")
             listItem.SubItems(3) = IIf(IsNull(rs("agent")), "", rs("agent"))
        rs.MoveNext
    Wend
    
    lbldata.Caption = "Jumlah Data  : " & rs.RecordCount
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "Status Account", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "Agent", 8 * 120
       
    ListView2.ColumnHeaders.ADD 1, , "Customer ID", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "CH Name", 20 * 120
    ListView2.ColumnHeaders.ADD 3, , "Agent", 8 * 120
    ListView2.ColumnHeaders.ADD 4, , "Status Account", 20 * 120
    ListView2.ColumnHeaders.ADD 5, , "Tanggal Trade", 20 * 120
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).Text = ListView1.SelectedItem.Text
        form_trade.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
    If ListView2.ListItems.Count > 0 Then
        VIEW_MGMDATA.Text1(2).Text = ListView2.SelectedItem.Text
        form_trade.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub
