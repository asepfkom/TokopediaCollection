VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Begin VB.Form FrmReqTelepon 
   Caption         =   "Req.Num.Telp"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNoTelp 
      Height          =   315
      Left            =   1680
      TabIndex        =   26
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox CmbRequestDi 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmReqTelepon.frx":0000
      Left            =   1680
      List            =   "FrmReqTelepon.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox CmbKategori 
      Height          =   315
      ItemData        =   "FrmReqTelepon.frx":0004
      Left            =   1680
      List            =   "FrmReqTelepon.frx":001D
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   480
      Width           =   4635
   End
   Begin VB.TextBox TxtKeterangan 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   5220
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   2760
      Width           =   1035
   End
   Begin TDBMask6Ctl.TDBMask TxtHome1 
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":007F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":00EB
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin TDBMask6Ctl.TDBMask TxtHome2 
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":012D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":0199
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtOffice1 
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":01DB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":0247
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtOffice2 
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0289
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":02F5
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtMobile1 
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0337
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":03A3
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtMobile2 
      Height          =   315
      Left            =   1620
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":03E5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":0451
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtEcPhone 
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":04FF
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtNoTelp1 
      Height          =   315
      Left            =   4680
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0541
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":05AD
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "               "
      Value           =   ""
   End
   Begin VB.Label Label12 
      Caption         =   "Request di:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label11 
      Caption         =   "No.Telepon"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label Label10 
      Caption         =   "Kategori"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label9 
      Caption         =   "Keterangan"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "EC Phone"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label7 
      Caption         =   "Additional Mobile 2"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   6540
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label6 
      Caption         =   "Additional Mobile 1"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   6180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "Additional Officeno 2"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   5820
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Additional Officeno 1"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   5460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Additional Home 2"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   5100
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Additional Home 1"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Custid:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "FrmReqTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdKeluar_Click()
    Me.Hide
End Sub

Private Sub cmbreqdi()
    CmbRequestDi.clear
    If Len(TxtNoTelp.text) >= 2 Then
        If Left(TxtNoTelp.text, 2) <> "08" Then
            CmbRequestDi.AddItem "AddHome1"
            CmbRequestDi.AddItem "AddHome2"
            CmbRequestDi.AddItem "AddOffice1"
            CmbRequestDi.AddItem "AddOffice2"
        ElseIf Left(TxtNoTelp.text, 2) = "08" Then
            CmbRequestDi.AddItem "AddMobile1"
            CmbRequestDi.AddItem "AddMobile2"
        End If
        'CmbRequestDi.AddItem "AddOther"
        CmbRequestDi.Enabled = True
    ElseIf Len(TxtNoTelp.text) = 0 Then
        CmbRequestDi.Enabled = False
    End If
        'AddHome2
        'AddOffice2
        'AddMobile2
        'AddOther
End Sub

Private Sub CmdSimpan_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Strsql As String
    
    Dim home1 As String
    Dim home2 As String
    Dim office1 As String
    Dim office2 As String
    Dim mobile1 As String
    Dim mobile2 As String
    Dim ec As String
    
'    If TxtHome1.Value = "" Then
'        home1 = "null"
'    Else
'        home1 = "'" + TxtHome1.Value + "'"
'    End If
'
'    If TxtHome2.Value = "" Then
'        home2 = "null"
'    Else
'        home2 = "'" + TxtHome2.Value + "'"
'    End If
'
'    If TxtOffice1.Value = "" Then
'        office1 = "null"
'    Else
'        office1 = "'" + TxtOffice1.Value + "'"
'    End If
'
'    If TxtOffice2.Value = "" Then
'        office2 = "null"
'    Else
'        office2 = "'" + TxtOffice2.Value + "'"
'    End If
'
'    If TxtMobile1.Value = "" Then
'        mobile1 = "null"
'    Else
'        mobile1 = "'" + TxtMobile1.Value + "'"
'    End If
'
'    If TxtMobile2.Value = "" Then
'        mobile2 = "null"
'    Else
'        mobile2 = "'" + TxtMobile2.Value + "'"
'    End If
'
'    If TxtEcPhone.Value = "" Then
'        ec = "null"
'    Else
'        ec = "'" + TxtEcPhone.Value + "'"
'    End If
'
'
'    'Ambil waktu server
'    STRSQL = "select now() as waktu "
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    M_OBJRS.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
''    CMDSQL = "insert into tblrequestadditionalphone (custid,"
''    CMDSQL = CMDSQL + "home1,home2,office1,office2,mobile1,mobile2,agent,tglreq) values ('"
''    CMDSQL = CMDSQL + TxtCustid.Text + "',"
''    CMDSQL = CMDSQL + IIf(IsNull(TxtHome1.Value), "null", CStr("'" + TxtHome1.Value + "'")) + ","
''    CMDSQL = CMDSQL + IIf(IsNull(TxtHome2.Value), "null", CStr("'" + TxtHome2.Value + "'")) + ","
''    CMDSQL = CMDSQL + IIf(IsNull(TxtOffice1.Value), "null", CStr("'" + TxtOffice1.Value + "'")) + ","
''    CMDSQL = CMDSQL + IIf(IsNull(TxtOffice2.Value), "null", CStr("'" + TxtOffice2.Value + "'")) + ","
''    CMDSQL = CMDSQL + IIf(IsNull(TxtMobile1.Value), "null", CStr("'" + TxtMobile1.Value + "'")) + ","
''    CMDSQL = CMDSQL + IIf(IsNull(TxtMobile2.Value), "null", CStr("'" + TxtMobile2.Value + "'")) + ",'"
''    CMDSQL = CMDSQL + MDIForm1.Text1.Text + "','"
''    CMDSQL = CMDSQL + Format(M_OBJRS(0), "yyyy-mm-dd hh:mm:ss") + "')"
'
'
'    CMDSQL = "insert into tblrequestadditionalphone (custid,"
'    CMDSQL = CMDSQL + "home1,home2,office1,office2,mobile1,mobile2,agent,tglreq,ecphone,keterangan) values ('"
'    CMDSQL = CMDSQL + TxtCustid.Text + "',"
'    CMDSQL = CMDSQL + home1 + ","
'    CMDSQL = CMDSQL + home2 + ","
'    CMDSQL = CMDSQL + office1 + ","
'    CMDSQL = CMDSQL + office2 + ","
'    CMDSQL = CMDSQL + mobile1 + ","
'    CMDSQL = CMDSQL + mobile2 + ",'"
'    CMDSQL = CMDSQL + MDIForm1.Text1.Text + "','"
'    CMDSQL = CMDSQL + Format(M_OBJRS(0), "yyyy-mm-dd hh:mm:ss") + "',"
'    CMDSQL = CMDSQL + ec + ",'"
'    CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "')"
'
'    M_OBJCONN.Execute CMDSQL
'    Set M_OBJRS = Nothing
    
    '@@17042012, Di Remarks dulu diganti dengan kategori
    If CmbKategori.text = "" Then
        MsgBox "Kategori tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtNoTelp.text = "" Then
        MsgBox "Nomor Telepon tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbRequestDi.text = "" Then
        MsgBox "Jenis Request Number tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If txtketerangan.text = "" Then
        MsgBox "Keterangan Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Update Nomor Telepon
    cmdsql = "insert into tblrequestadditionalphone (custid,"
    cmdsql = cmdsql + " request_number,kategori,keterangan,agent,tglreq,jenis) values ('"
    cmdsql = cmdsql + TxtCustid.text + "','"
    cmdsql = cmdsql + IIf(IsNull(TxtNoTelp.text), "", CStr(TxtNoTelp.text)) + "','"
    cmdsql = cmdsql + IIf(IsNull(CmbKategori.text), "", CStr(CmbKategori.text)) + "','"
    cmdsql = cmdsql + IIf(IsNull(txtketerangan.text), "", CStr(txtketerangan.text)) + "','"
    cmdsql = cmdsql + IIf(IsNull(MDIForm1.Text1.text), "", MDIForm1.Text1.text) + "','"
    cmdsql = cmdsql + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss")) + "','"
    cmdsql = cmdsql + Trim(CmbRequestDi.text) + "')"
    M_OBJCONN.execute cmdsql
    
    'jejaktian06042016
    If CmbRequestDi.text = "AddHome1" Then
        FrmCC_Colection.txtHomeAdd1.text = TxtNoTelp.text
        FrmCC_Colection.txtHomeAdd1.ReadOnly = True
    ElseIf CmbRequestDi.text = "AddHome2" Then
        FrmCC_Colection.txtHomeAdd2.text = TxtNoTelp.text
        FrmCC_Colection.txtHomeAdd2.ReadOnly = True
    ElseIf CmbRequestDi.text = "AddOffice1" Then
        FrmCC_Colection.txtOfficeAdd1.text = TxtNoTelp.text
        FrmCC_Colection.txtOfficeAdd1.ReadOnly = True
    ElseIf CmbRequestDi.text = "AddOffice2" Then
        FrmCC_Colection.txtOfficeAdd2.text = TxtNoTelp.text
        FrmCC_Colection.txtOfficeAdd2.ReadOnly = True
    ElseIf CmbRequestDi.text = "AddMobile1" Then
        FrmCC_Colection.txtMobileAdd1.text = TxtNoTelp.text
        FrmCC_Colection.txtMobileAdd1.ReadOnly = True
    ElseIf CmbRequestDi.text = "AddMobile2" Then
        FrmCC_Colection.txtMobileAdd2.text = TxtNoTelp.text
        FrmCC_Colection.txtMobileAdd2.ReadOnly = True
    End If
    '======================================================
    
'    'Update buat ngasih tanda ke TL/SPV/Admin
'    CMDSQL = "update usertbl set f_req_number='1' where userid in ("
'    CMDSQL = CMDSQL + "select team from usertbl where userid='"
'    CMDSQL = CMDSQL + MDIForm1.Text1.Text + "') or userid in (select userid from usertbl where "
'    CMDSQL = CMDSQL + "usertype='20' or usertype='25' or usertype='11') "
'    M_OBJCONN.Execute CMDSQL

    '@@07-08-2012 Kirim Via Form Pesan
    'Kirim Ke Semua TL/SPV
    'CMDSQL = "select userid from usertbl where usertype in ('6','11','25','20') "
    cmdsql = "select userid from usertbl where userid in ("
    cmdsql = cmdsql + "select team from usertbl where userid='"
    cmdsql = cmdsql + MDIForm1.Text1.text + "') "
    cmdsql = cmdsql + " and userid is not null "
    Set M_Objrs_CariTL = New ADODB.Recordset
    M_Objrs_CariTL.CursorLocation = adUseClient
    M_Objrs_CariTL.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_CariTL.RecordCount > 0 Then
        Remarks = "Ada Request Number!" & vbCrLf
        Remarks = Remarks + "-------------------------------------------------" & vbCrLf
        Remarks = Remarks + " Custid: " & TxtCustid.text & vbCrLf
        Remarks = Remarks + " Agent:  " & MDIForm1.Text1.text & vbCrLf
        Remarks = Remarks + " Kategori: " & CmbKategori.text & vbCrLf
        Remarks = Remarks + " Tgl.Request: " & CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss"))
        
        
        
        While Not M_Objrs_CariTL.EOF
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + CStr(Trim(M_Objrs_CariTL("userid"))) + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.Text1.text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            M_OBJCONN.execute cmdsql
            M_Objrs_CariTL.MoveNext
        Wend
    End If
    
    MsgBox "Request berhasil dikirim!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub TxtNoTelp_Change()
    textval = TxtNoTelp.text
    If IsNumeric(textval) Then
      numval = textval
    Else
      TxtNoTelp.text = CStr(numval)
    End If
    
    Call cmbreqdi
End Sub

Private Sub TxtNoTelp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub
