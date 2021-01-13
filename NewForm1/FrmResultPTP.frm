VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form FrmResultPTP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Result PTP"
   ClientHeight    =   4485
   ClientLeft      =   840
   ClientTop       =   5715
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Ubah PTP"
      Enabled         =   0   'False
      Height          =   2595
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   4455
      Begin VB.TextBox txtStatusAcc 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtID 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin TDBDate6Ctl.TDBDate TxtTglPTP 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmResultPTP.frx":0000
         Caption         =   "FrmResultPTP.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmResultPTP.frx":0184
         Keys            =   "FrmResultPTP.frx":01A2
         Spin            =   "FrmResultPTP.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12648384
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
      Begin TDBNumber6Ctl.TDBNumber txtpembayaran 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1440
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   450
         Calculator      =   "FrmResultPTP.frx":0228
         Caption         =   "FrmResultPTP.frx":0248
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmResultPTP.frx":02B4
         Keys            =   "FrmResultPTP.frx":02D2
         Spin            =   "FrmResultPTP.frx":031C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12648384
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate TxtTglPTP1 
         Height          =   285
         Left            =   2820
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmResultPTP.frx":0344
         Caption         =   "FrmResultPTP.frx":045C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmResultPTP.frx":04C8
         Keys            =   "FrmResultPTP.frx":04E6
         Spin            =   "FrmResultPTP.frx":0544
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   65535
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
         Format          =   "dd/mm/yyyy"
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   6815745
         Value           =   39876
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber txtpembayaran1 
         Height          =   255
         Left            =   2820
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   450
         Calculator      =   "FrmResultPTP.frx":056C
         Caption         =   "FrmResultPTP.frx":058C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmResultPTP.frx":05F8
         Keys            =   "FrmResultPTP.frx":0616
         Spin            =   "FrmResultPTP.frx":0660
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   65535
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin VB.Label Label6 
         Caption         =   "Status Account: "
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "ID"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Pay/month:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tgl.PTP:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   675
      End
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3300
      Width           =   1275
   End
   Begin VB.ComboBox CmbResultPTP 
      Height          =   315
      ItemData        =   "FrmResultPTP.frx":0688
      Left            =   1140
      List            =   "FrmResultPTP.frx":068A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Catatlah hasil negosiasi account PTP anda!"
      Height          =   615
      Left            =   60
      TabIndex        =   3
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Result PTP:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "FrmResultPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tgltagih As Date

Private Sub IsiData()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbResultPTP.CLEAR
    
        cmdsql = "select * from tbl_desc_result_ptp "
        cmdsql = cmdsql + " where aktif='1' order by desc_result_ptp asc "

    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "SELECT * FROM enabledptp"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If FrmCC_Colection.LstPayment.ListItems.Count > 0 And FrmCC_Colection.ListView1(0).ListItems.Count > 0 Then
        a = FrmCC_Colection.LstPayment.ListItems(1).SubItems(2)
        B = FrmCC_Colection.ListView1(0).ListItems(1).text
    End If
    
    If rs!Enabled = 0 Then
        If M_Objrs.RecordCount > 0 Then
            If a > B Then
            
                sqla = "SELECT * FROM tblnegoptp_temp_app where custid = '" + CStr(FrmCC_Colection.lblCustId.Caption) + "' "
                Set rsa = New ADODB.Recordset
                rsa.CursorLocation = adUseClient
                rsa.Open sqla, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If rsa.RecordCount = 0 Then
                    CmbResultPTP.AddItem "CHANGE PTP (Approval)"
                End If
                
                While Not M_Objrs.EOF
                    If M_Objrs("desc_result_ptp") <> "Change PTP" Then
                            CmbResultPTP.AddItem IIf(IsNull(M_Objrs("desc_result_ptp")), "", M_Objrs("desc_result_ptp"))
                    End If
                    M_Objrs.MoveNext
                Wend
            Else
                While Not M_Objrs.EOF
                    CmbResultPTP.AddItem IIf(IsNull(M_Objrs("desc_result_ptp")), "", M_Objrs("desc_result_ptp"))
                    M_Objrs.MoveNext
                Wend
            End If
        End If
    Else
        If M_Objrs.RecordCount > 0 Then
                sqla = "SELECT * FROM tblnegoptp_temp_app where custid = '" + CStr(FrmCC_Colection.lblCustId.Caption) + "' "
                Set rsa = New ADODB.Recordset
                rsa.CursorLocation = adUseClient
                rsa.Open sqla, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
                If rsa.RecordCount = 0 Then
                    CmbResultPTP.AddItem "CHANGE PTP (Approval)"
                End If
            While Not M_Objrs.EOF
                If M_Objrs("desc_result_ptp") <> "Change PTP" Then
                        CmbResultPTP.AddItem IIf(IsNull(M_Objrs("desc_result_ptp")), "", M_Objrs("desc_result_ptp"))
                End If
                M_Objrs.MoveNext
            Wend
        End If
    End If
    
    Set M_Objrs = Nothing
End Sub



Private Sub CmbResultPTP_Click()
    If UCase(CmbResultPTP.text) = "CHANGE PTP" Then
        Frame1.Enabled = True
    ElseIf CmbResultPTP.text = "CHANGE PTP (Approval)" Then
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
End Sub

Private Sub CmdSimpan_Click()
    Dim cmdsql As String
    Dim TanggalPTP As String
    Dim TanggalTagih As String
    Dim M_Objrs_Cek_CPA As ADODB.Recordset
    Dim AmountDeal As Double
     
    TanggalPTP = Format(TxtTglPTP.Value, "yyyy-mm-dd")
     
    If CmbResultPTP.text = "" Then
        MsgBox "Result PTP tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        'Unload Me
        Exit Sub
    End If
             
             
    '@@27 Juni 2012, Jika Result PTP=CHANGE PTP maka user harus mengubah isi ptp dulu
    If UCase(CmbResultPTP.text) = "CHANGE PTP" Then
        
        TGL = Format(TxtTglPTP.Value, "YYYY-MM-DD")
        
            SqlWaktu = "select now() as tgl"
            Set m_waktuserver = New ADODB.Recordset
            m_waktuserver.CursorLocation = adUseClient
            m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If TGL < Format(m_waktuserver!TGL, "yyyy-mm-dd") Then
            MsgBox "Tidak Boleh Back Date"
            Exit Sub
        End If

        
        If TanggalPTP = TxtTglPTP1.Value And TxtPembayaran.Value = txtpembayaran1.Value Then
            MsgBox "Anda memilih result PTP: Change PTP! Anda harus mengubah PTP pada kolom ubah ptp!", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        Else
                
            cmdsql = "select * from tblcpa where vcustid='"
            cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "' order by nid desc limit 1"
            Set M_Objrs_Cek_CPA = New ADODB.Recordset
            M_Objrs_Cek_CPA.CursorLocation = adUseClient
            M_Objrs_Cek_CPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Cek_CPA.RecordCount > 0 Then
                AmountDeal = M_Objrs_Cek_CPA("nttlpayment")
            Else
                MsgBox "CPA belum tersedia! Tekan Send PTP untuk membuat CPA dan PTP di Form Customer!", vbOKOnly + vbInformation, "Informasi"
                Unload Me
                Exit Sub
            End If
            
            Set M_Objrs_Cek_CPA = Nothing
                
                
            If TxtTglPTP.ValueIsNull = True Or _
               TxtPembayaran.ValueIsNull = True Or _
               TxtPembayaran.Value = 0 Then
                MsgBox "Anda memilih Result PTP: Change PTP! Tgl.PTP dan Pay/Month tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
                Exit Sub
            End If
            
            Call CariTanggalTagih
            
            TanggalPTP = Format(TxtTglPTP.Value, "yyyy-mm-dd")
            TanggalTagih = Format(FrmCC_Colection.TdbTglTagih.Value, "yyyy-mm-dd")
            
            cmdsql = "update tblnegoptp set promisedate='"
            cmdsql = cmdsql + TanggalPTP + "',promisepay='"
            cmdsql = cmdsql + CStr(TxtPembayaran.Value) + "'"
            cmdsql = cmdsql + " where id='"
            cmdsql = cmdsql + CStr(TxtID.text) + "'"
            M_OBJCONN.Execute cmdsql
            
            Call BikinStatusPTP
            
'            'Update juga di tabel mgm
'            CMDSQL = "update mgm set dateptp='"
'            CMDSQL = CMDSQL + TanggalPTP + "',tgl_tagih='"
'            CMDSQL = CMDSQL + TanggalTagih + "' where custid='"
'            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Caption) + "'"
'            M_OBJCONN.Execute CMDSQL
            
            'Update Via PTP
            If Trim(FrmCC_Colection.CmbViaPtp.text) = "" Then
                cmdsql = "update mgm set ptpvia='ATM LAINNYA' where custid='"
                cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "'"
                M_OBJCONN.Execute cmdsql
            End If
            
        End If
    End If
             
    If CmbResultPTP.text = "CHANGE PTP (Approval)" Then
        TGL = Format(TxtTglPTP.Value, "YYYY-MM-DD")
        
            SqlWaktu = "select now() as tgl"
            Set m_waktuserver = New ADODB.Recordset
            m_waktuserver.CursorLocation = adUseClient
            m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If TGL < Format(m_waktuserver!TGL, "yyyy-mm-dd") Then
            MsgBox "Tidak Boleh Back Date"
            Exit Sub
        End If

        
        
        If TanggalPTP = TxtTglPTP1.Value And TxtPembayaran.Value = txtpembayaran1.Value Then
            MsgBox "Anda memilih result PTP: Change PTP! Anda harus mengubah PTP pada kolom ubah ptp!", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        Else
                
            cmdsql = "select * from tblcpa where vcustid='"
            cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "' order by nid desc limit 1"
            Set M_Objrs_Cek_CPA = New ADODB.Recordset
            M_Objrs_Cek_CPA.CursorLocation = adUseClient
            M_Objrs_Cek_CPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Cek_CPA.RecordCount > 0 Then
                AmountDeal = M_Objrs_Cek_CPA("nttlpayment")
            Else
                MsgBox "CPA belum tersedia! Tekan Send PTP untuk membuat CPA dan PTP di Form Customer!", vbOKOnly + vbInformation, "Informasi"
                Unload Me
                Exit Sub
            End If
            
            Set M_Objrs_Cek_CPA = Nothing
                
                
            If TxtTglPTP.ValueIsNull = True Or _
               TxtPembayaran.ValueIsNull = True Or _
               TxtPembayaran.Value = 0 Then
                MsgBox "Anda memilih Result PTP: Change PTP! Tgl.PTP dan Pay/Month tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
                Exit Sub
            End If
            
            TanggalPTP = Format(TxtTglPTP.Value, "yyyy-mm-dd")
            TanggalTagih = Format(FrmCC_Colection.TdbTglTagih.Value, "yyyy-mm-dd")
            Tenor = FrmCC_Colection.txttenor.Value
            
            Call CariTanggalTagihChangePtpApp
            
            quins = "INSERT INTO tblnegoptp_temp_app (custid,promisedate,promisepay,statusptp,via,agent,tgltagih,tenor) values "
            quins = quins + "('" + FrmCC_Colection.lblCustId.Caption + "', "
            quins = quins + "'" + TanggalPTP + "',"
            quins = quins + "'" & TxtPembayaran.Value & "',"
                If FrmCC_Colection.ListView1(0).ListItems.Count > 0 Then
                    quins = quins + "'PTP-POP',"
                Else
                    quins = quins + "'PTP-NEW',"
                End If
            
                If Trim(FrmCC_Colection.CmbViaPtp.text) = "" Then
                    quins = quins + "'ATM LAINNYA',"
                Else
                    quins = quins + "'" + FrmCC_Colection.CmbViaPtp.text + "',"
                End If
            quins = quins + "'" + MDIForm1.Text1.text + "', '" + Format(tgltagih, "yyyy-mm-dd") + "', '" & Tenor & "')"
            M_OBJCONN.Execute quins
            MsgBox "Data terkirim untuk di Approve SPV"
        End If
    End If
                 
    '@@27-06-2012 Jika dia Pay maka cek paymentnya
    If UCase(CmbResultPTP.text) = "PAY" Then
        cmdsql = "update mgm set tglstatus= now() ,KETHSLKERJA_NEW='POP-PROGRESS OF PAYMENT',"
        cmdsql = cmdsql + "KETHSLKERJADESC_NEW='POP-PROGRESS OF PAYMENT',F_CEK_NEW='POP',"
        cmdsql = cmdsql + "F_CEK='POP',LASTSTATUS='POP',KETHSLKERJA='POP',"
        cmdsql = cmdsql + "REMARKS = 'POP',RECSTATUS='C',OTO='Y'  where f_cek_new like 'PTP%' "
        cmdsql = cmdsql + " and custid='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "' "
        cmdsql = cmdsql + " and custid in( select custid from vwwlunas where custid='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "')"
        M_OBJCONN.Execute cmdsql
    End If
                 
    'Simpan ke history result ptp
    cmdsql = "insert into tbl_hst_resultptp (tgl,hst,user_log,user_handle,custid) values "
    cmdsql = cmdsql + "(now(),'"
    cmdsql = cmdsql + CmbResultPTP.text + "','"
    cmdsql = cmdsql + MDIForm1.Text1.text + "','"
    cmdsql = cmdsql + FrmCC_Colection.lblaoc.Caption + "','"
    cmdsql = cmdsql + FrmCC_Colection.lblCustId.Caption + "')"
    M_OBJCONN.Execute cmdsql
    
    'Simpan ke tabel mgm
    cmdsql = "update mgm set result_ptp='"
    cmdsql = cmdsql + CmbResultPTP.text + "' where custid='"
    cmdsql = cmdsql + FrmCC_Colection.lblCustId.Caption + "'"
    M_OBJCONN.Execute cmdsql
    
    If TxtID.text = Empty Or TxtID.text = "" Or IsNull(TxtID.text) = True Then
        Unload Me
        Exit Sub
    Else
        '@@ 27-06-2012 Catat juga di negoptp
        If TxtID.text <> Empty Or TxtID.text <> "" Or IsNull(TxtID.text) = False Then
            'Update status PTP di tabel negoptp
            cmdsql = "update tblnegoptp set result_ptp='" + CmbResultPTP.text + "' where id='"
            cmdsql = cmdsql + CStr(TxtID.text) + "'"
            M_OBJCONN.Execute cmdsql
        End If
    End If
    
'    If TxtID.Text <> Empty And TxtTglPTP.ValueIsNull = False Then
'        Call CariTanggalTagih
'        'Catet ulang status PTP
'        Call BikinStatusPTP
'    End If
    
'    'Update Via PTP
'    If Trim(FrmCC_Colection.CmbViaPtp.Text) = "" Then
'        CMDSQL = "update mgm set ptpvia='ATM LAINNYA' where custid='"
'        CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Caption) + "'"
'        M_OBJCONN.Execute CMDSQL
'    End If
    
    MsgBox "Result PTP Berhasil ditambahkan!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub Form_Load()
    Call IsiData
    Call CariPTPTerakhir
End Sub

'@@27-06-2012 Promisedate dan Promisepay harus diisi
Private Sub CariPTPTerakhir()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "select * from tblnegoptp where custid='"
    cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "' order by promisedate desc limit 1"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        TxtTglPTP.Value = Format(M_Objrs("promisedate"), "dd/mm/yyyy")
        TxtTglPTP1.Value = Format(M_Objrs("promisedate"), "dd/mm/yyyy")
        TxtPembayaran.Value = IIf(IsNull(M_Objrs("promisepay")), 0, M_Objrs("promisepay"))
        txtpembayaran1.Value = IIf(IsNull(M_Objrs("promisepay")), 0, M_Objrs("promisepay"))
        TxtID.text = M_Objrs("id")
    End If
    Set M_Objrs = Nothing
End Sub

Private Sub Frame1_Click()
    If UCase(CmbResultPTP.text) <> "CHANGE PTP" And CmbResultPTP.text <> "CHANGE PTP (Approval)" Then
        MsgBox "Anda hanya dapat mengedit PTP jika anda memilih Change PTP!", vbOKOnly + vbInformation
    End If
End Sub


Private Sub CariTanggalTagih()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim TglPaymentEffective As String
    
    
    TglPaymentEffective = Format(TxtTglPTP.Value, "yyyy-mm-dd")
    
    cmdsql = "Select  date('" + TglPaymentEffective + "')-"
    If UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "HSBC" Then
        cmdsql = cmdsql + "1"
    ElseIf UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "BERSAMA" Then
        cmdsql = cmdsql + "1"
    ElseIf UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "KANTOR POS" Then
        cmdsql = cmdsql + "3"
    ElseIf UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "PUM" Then
        cmdsql = cmdsql + "1"
    Else
        cmdsql = cmdsql + "3"
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    On Error GoTo SALAH
    FrmCC_Colection.TdbTglTagih.Value = Format(M_Objrs(0), "mm/dd/yyyy")
    
    Set M_Objrs = Nothing
    Exit Sub
SALAH:
    MsgBox "Ada Error: " & err.Description
End Sub

Private Sub CariTanggalTagihChangePtpApp()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim TglPaymentEffective As String
    
    
    TglPaymentEffective = Format(TxtTglPTP.Value, "yyyy-mm-dd")
    
    cmdsql = "Select  date('" + TglPaymentEffective + "')-"
    If UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "HSBC" Then
        cmdsql = cmdsql + "1"
    ElseIf UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "BERSAMA" Then
        cmdsql = cmdsql + "1"
    ElseIf UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "KANTOR POS" Then
        cmdsql = cmdsql + "3"
    ElseIf UCase(Trim(FrmCC_Colection.CmbViaPtp.text)) = "PUM" Then
        cmdsql = cmdsql + "1"
    Else
        cmdsql = cmdsql + "3"
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    On Error GoTo SALAH
    tgltagih = Format(M_Objrs(0), "yyyy-mm-dd")
    
    Set M_Objrs = Nothing
    Exit Sub
SALAH:
    MsgBox "Ada Error: " & err.Description
End Sub


Private Sub BikinStatusPTP()
    Dim cmdsql As String
    Dim Cmdsql_Cek_status As String
    Dim M_Objrs As ADODB.Recordset
    Dim TglPTPNew As String
    Dim StatusPTP As String
    Dim M_Objrs_Cek_CPA As ADODB.Recordset
    Dim AmountDeal As Double
    
    
    cmdsql = "select * from tblcpa where vcustid='"
    cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "' order by nid desc limit 1"
    Set M_Objrs_Cek_CPA = New ADODB.Recordset
    M_Objrs_Cek_CPA.CursorLocation = adUseClient
    M_Objrs_Cek_CPA.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_CPA.RecordCount > 0 Then
        If Val(M_Objrs_Cek_CPA("nttlpayment")) > 0 Then
            AmountDeal = M_Objrs_Cek_CPA("nttlpayment")
        End If
    End If
    
    Set M_Objrs_Cek_CPA = Nothing
    
    
    
    If FrmCC_Colection.ListView1(0).ListItems.Count > 0 Then
        StatusPTP = "PTP-POP"
    Else
        StatusPTP = "PTP-NEW"
    End If
     
   If StatusPTP = "PTP-NEW" Then
        'Tapi jika status sebelumnya bukan ptp new maka update tglptpnew=now
        Cmdsql_Cek_status = "select * from mgm where custid='"
        Cmdsql_Cek_status = Cmdsql_Cek_status + CStr(FrmCC_Colection.lblCustId.Caption) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open Cmdsql_Cek_status, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_Status.RecordCount > 0 Then
            If M_Objrs_Cek_Status("tglptpnew") = "" Or IsNull(M_Objrs_Cek_Status("tglptpnew")) = True _
               Or M_Objrs_Cek_Status("tglptpnew") = Empty Then
                TglPTPNew = "now()"
             Else
                TglPTPNew = "'" + CStr(Format(M_Objrs_Cek_Status("tglptpnew"), "yyyy-mm-dd")) + "'"
             End If
        End If
        
        Set M_Objrs_Cek_Status = Nothing
    
        cmdsql = "update mgm set dateptpnew='"
        cmdsql = cmdsql + Format(TxtTglPTP.Value, "yyyy-mm-dd") + "',tgl_tagih='"
        cmdsql = cmdsql + Format(FrmCC_Colection.TdbTglTagih.Value, "yyyy-mm-dd") + "', "
        
        
        '@@20062012, amountnew ambil dari negoptp terakhir aja deh....
        cmdsql = cmdsql + " tglallptp='"
        cmdsql = cmdsql + Format(TxtTglPTP.Value, "yyyy-mm-dd") + "',f_cek_new='PTP-NE',"
        cmdsql = cmdsql + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.CmbViaPtp.text) + "',ptpdesc='PTP-NEW', dateptp='"
        cmdsql = cmdsql + Format(TxtTglPTP.Value, "yyyy-mm-dd") + "',tglptpnew=" + TglPTPNew
        cmdsql = cmdsql + ",tenor='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.txttenor.Value) + "',ttlptp='"
        cmdsql = cmdsql + CStr(AmountDeal) + "' "
        cmdsql = cmdsql + "where custid='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "'"
        DoEvents
        M_OBJCONN.Execute cmdsql
    End If
    
    If StatusPTP = "PTP-POP" Then
        cmdsql = "update mgm set dateptp='"
        cmdsql = cmdsql + Format(TxtTglPTP.Value, "yyyy-mm-dd") + "',tgl_tagih='"
        cmdsql = cmdsql + Format(FrmCC_Colection.TdbTglTagih.Value, "yyyy-mm-dd") + "',tglallptp='"
        cmdsql = cmdsql + Format(TxtTglPTP.Value, "yyyy-mm-dd") + "',f_cek_new='PTP-PO',"
        
        cmdsql = cmdsql + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.CmbViaPtp.text) + "',ptpdesc='PTP-POP',"
        cmdsql = cmdsql + "tenor='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.txttenor.Value) + "',ttlptp='"
        cmdsql = cmdsql + CStr(AmountDeal) + "'"
        cmdsql = cmdsql + " where custid='"
        cmdsql = cmdsql + CStr(FrmCC_Colection.lblCustId.Caption) + "'"
        M_OBJCONN.Execute cmdsql
    End If
End Sub

