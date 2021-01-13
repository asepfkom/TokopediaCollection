VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmJadwalPembayaranCpa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jadwal Pembayaran CPA"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPenawaranvia 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2100
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   5340
      Visible         =   0   'False
      Width           =   6075
   End
   Begin VB.ComboBox CmbJawabanCH 
      Height          =   315
      ItemData        =   "FrmJadwalPembayaranCPA.frx":0000
      Left            =   6360
      List            =   "FrmJadwalPembayaranCPA.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   4740
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox CmbPermohonanMelalui 
      Height          =   315
      ItemData        =   "FrmJadwalPembayaranCPA.frx":0021
      Left            =   7620
      List            =   "FrmJadwalPembayaranCPA.frx":002E
      TabIndex        =   33
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   435
      Left            =   9300
      TabIndex        =   19
      Top             =   4020
      Width           =   1395
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   435
      Left            =   9300
      TabIndex        =   31
      Top             =   3600
      Width           =   1395
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "&Tambah Schedule Manual"
      Height          =   435
      Left            =   4740
      TabIndex        =   18
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton CmdTambahOtomatis 
      Caption         =   "&Tambah Schedule Otomatis"
      Height          =   435
      Left            =   1980
      TabIndex        =   28
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox TxtFromOs 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1740
      TabIndex        =   27
      Top             =   2580
      Width           =   975
   End
   Begin VB.TextBox TxtNoSurat 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   2580
      TabIndex        =   25
      Top             =   120
      Width           =   1155
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   11040
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox TxtNoTelp 
      Height          =   315
      ItemData        =   "FrmJadwalPembayaranCPA.frx":0046
      Left            =   5460
      List            =   "FrmJadwalPembayaranCPA.frx":0048
      TabIndex        =   24
      Top             =   1440
      Width           =   2715
   End
   Begin VB.TextBox TxtIdCPA 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1740
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   435
      Left            =   9300
      TabIndex        =   21
      Top             =   4860
      Width           =   1395
   End
   Begin VB.CommandButton CmdCetak 
      Caption         =   "&Cetak"
      Height          =   435
      Left            =   7680
      TabIndex        =   20
      Top             =   3600
      Width           =   1395
   End
   Begin VB.TextBox txtTL 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1740
      TabIndex        =   16
      Top             =   1020
      Width           =   975
   End
   Begin VB.TextBox TxtAgent 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1740
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1740
      TabIndex        =   6
      Top             =   420
      Width           =   1995
   End
   Begin VB.TextBox TxtAlamat 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5460
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   4995
   End
   Begin VB.TextBox TxtNama 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1740
      TabIndex        =   1
      Top             =   720
      Width           =   1995
   End
   Begin TDBNumber6Ctl.TDBNumber TxtBalance 
      Height          =   255
      Left            =   1740
      TabIndex        =   10
      Top             =   1680
      Width           =   1980
      _Version        =   65536
      _ExtentX        =   3492
      _ExtentY        =   450
      Calculator      =   "FrmJadwalPembayaranCPA.frx":004A
      Caption         =   "FrmJadwalPembayaranCPA.frx":006A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmJadwalPembayaranCPA.frx":00D6
      Keys            =   "FrmJadwalPembayaranCPA.frx":00F4
      Spin            =   "FrmJadwalPembayaranCPA.frx":013E
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0"
      EditMode        =   0
      Enabled         =   0
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
   Begin TDBNumber6Ctl.TDBNumber TxtPayment 
      Height          =   255
      Left            =   1740
      TabIndex        =   12
      Top             =   1980
      Width           =   1980
      _Version        =   65536
      _ExtentX        =   3492
      _ExtentY        =   450
      Calculator      =   "FrmJadwalPembayaranCPA.frx":0166
      Caption         =   "FrmJadwalPembayaranCPA.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmJadwalPembayaranCPA.frx":01F2
      Keys            =   "FrmJadwalPembayaranCPA.frx":0210
      Spin            =   "FrmJadwalPembayaranCPA.frx":025A
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0"
      EditMode        =   0
      Enabled         =   0
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
   Begin TDBNumber6Ctl.TDBNumber TxtInstallment 
      Height          =   255
      Left            =   1740
      TabIndex        =   14
      Top             =   2280
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   450
      Calculator      =   "FrmJadwalPembayaranCPA.frx":0282
      Caption         =   "FrmJadwalPembayaranCPA.frx":02A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmJadwalPembayaranCPA.frx":030E
      Keys            =   "FrmJadwalPembayaranCPA.frx":032C
      Spin            =   "FrmJadwalPembayaranCPA.frx":0376
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0"
      EditMode        =   0
      Enabled         =   0
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
   Begin MSComctlLib.ListView LvJadwal 
      Height          =   5160
      Left            =   60
      TabIndex        =   17
      Top             =   4140
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   9102
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin TDBDate6Ctl.TDBDate TxtTglBayar 
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   3660
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "FrmJadwalPembayaranCPA.frx":039E
      Caption         =   "FrmJadwalPembayaranCPA.frx":04B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmJadwalPembayaranCPA.frx":0522
      Keys            =   "FrmJadwalPembayaranCPA.frx":0540
      Spin            =   "FrmJadwalPembayaranCPA.frx":059E
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
   Begin TDBDate6Ctl.TDBDate TxtTglPengajuan 
      Height          =   285
      Left            =   6900
      TabIndex        =   35
      Top             =   480
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "FrmJadwalPembayaranCPA.frx":05C6
      Caption         =   "FrmJadwalPembayaranCPA.frx":06DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmJadwalPembayaranCPA.frx":074A
      Keys            =   "FrmJadwalPembayaranCPA.frx":0768
      Spin            =   "FrmJadwalPembayaranCPA.frx":07C6
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
   Begin VB.Label Label17 
      Caption         =   "Ketik alamat ch (jika tertulis) / no telepon CH  (jika penawaran via telepon)"
      Height          =   255
      Left            =   2100
      TabIndex        =   39
      Top             =   5100
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label16 
      Caption         =   "CH memberikan jawaban atas penawaran ini melalui?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   36
      Top             =   4800
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label14 
      Caption         =   "Tanggal permohonan diajukan?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   34
      Top             =   540
      Width           =   4575
   End
   Begin VB.Label Label13 
      Caption         =   "Permohonan yang anda ajukan melalui?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   32
      Top             =   180
      Width           =   4575
   End
   Begin VB.Label Label12 
      Caption         =   "Tgl.Bayar:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label Label11 
      Caption         =   "From O/S balance:"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label Label10 
      Caption         =   "Id CPA:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label9 
      Caption         =   "TL:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "Installment period:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label7 
      Caption         =   "Total Payment:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label Label6 
      Caption         =   "Balance:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Agent:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Custid:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "No.Telepon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Alamat CH:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   900
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Nama CH:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "FrmJadwalPembayaranCpa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JawabanCh As String
Private Sub HeaderListJadwal()
    LvJadwal.ColumnHeaders.ADD 1, , "Id", 700
    LvJadwal.ColumnHeaders.ADD 2, , "Tahun", 800
    LvJadwal.ColumnHeaders.ADD 3, , "Bulan", 800
    LvJadwal.ColumnHeaders.ADD 4, , "Tunggakan", 1000
    LvJadwal.ColumnHeaders.ADD 5, , "Pembayaran", 1000
    LvJadwal.ColumnHeaders.ADD 6, , "Hutang Pokok", 2000
    LvJadwal.ColumnHeaders.ADD 7, , "Interest", 2000
    LvJadwal.ColumnHeaders.ADD 8, , "Balance Terakhir", 2000
    LvJadwal.ColumnHeaders.ADD 9, , "Tanggal Bayar", 2000
End Sub

Private Sub CmbJawabanCH_Click()
    JawabanCh = ""
    If Trim(UCase(CmbJawabanCH.Text)) = "TELEPON" Then
        JawabanCh = "via telepon ke nomor " & Trim(TxtPenawaranvia.Text)
    End If
    If Trim(UCase(CmbJawabanCH.Text)) = "TERTULIS" Then
        JawabanCh = "ke alamat " & Trim(TxtPenawaranvia.Text)
    End If
End Sub

Private Sub CmdCetak_Click()
    Dim Cmdsql As String
    Dim JenisKartu_1 As String
    Dim JenisKartu_2 As String
    
    If TxtAlamat.Text = Empty Then
        MsgBox "Alamat tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    If TxtNoTelp.Text = Empty Then
        MsgBox "Nomor telepon tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglPengajuan.ValueIsNull = True Then
        MsgBox "Tanggal pengajuan tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

'    If CmbJawabanCH.Text = "" Or TxtPenawaranvia.Text = "" Then
'        MsgBox "Jawaban CH tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
'        Exit Sub
'    End If

    If CmbPermohonanMelalui.Text = "" Then
        MsgBox "Permohonan melalui tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LvJadwal.ListItems.Count = 0 Then
        MsgBox "List pembayaran tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    IsiJadwal
    
    If Trim(UCase(FrmCC_Colection.lbltype.Caption)) = "PIL" Or Trim(UCase(FrmCC_Colection.lbltype.Caption)) = "GRF" Then
        JenisKartu_1 = "KTA"
        JenisKartu_2 = "KARTU TANPA AGUNAN"
    Else
        JenisKartu_1 = "Kartu Kredit Utama HSBC"
        JenisKartu_2 = "KARTU KREDIT UTAMA HSBC"
    End If
    
    '@@13-10-2011, Log pembayaran
    Cmdsql = "insert into log_jadwal_pembayaran (custid,no_surat,nama_ch,"
    Cmdsql = Cmdsql + "alamat_ch,telpon_ch,tglsurat,agent,balance,payment,cetak_by) values ('"
    Cmdsql = Cmdsql + TxtCustid.Text + "','"
    Cmdsql = Cmdsql + TxtNoSurat.Text + "','"
    Cmdsql = Cmdsql + TxtNama.Text + "','"
    Cmdsql = Cmdsql + TxtAlamat.Text + "','"
    Cmdsql = Cmdsql + TxtNoTelp.Text + "','"
    Cmdsql = Cmdsql + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "','"
    Cmdsql = Cmdsql + TxtAgent.Text + "','"
    Cmdsql = Cmdsql + CStr(txtbalance.Value) + "','"
    Cmdsql = Cmdsql + CStr(txtPayment.Value) + "','"
    Cmdsql = Cmdsql + MDIForm1.Text2.Text + "')"
    M_OBJCONN.Execute Cmdsql
    
    
    WaitSecs (2)
    RPT.Reset
    RPT.Formulas(1) = "@nama = totext('" + CStr(Trim(TxtNama.Text)) + "')"
    RPT.Formulas(2) = "@alamat = totext('" + CStr(Trim(TxtAlamat.Text)) + "')"
    RPT.Formulas(3) = "@telpon = totext('" + CStr(Trim(TxtNoTelp.Text)) + "')"
    RPT.Formulas(4) = "@custid = totext('" + CStr(Trim(TxtCustid.Text)) + "')"
    RPT.Formulas(5) = "@balance = totext('" + "Rp. " + CStr(Trim(txtbalance.Text)) + "')"
    RPT.Formulas(6) = "@payment = totext('" + "Rp. " + CStr(Trim(txtPayment.Text)) + "')"
    RPT.Formulas(7) = "@no_surat = totext('" + CStr(Trim(TxtNoSurat.Text)) + "')"
    RPT.Formulas(8) = "@jenis_kartu_1 = totext('" + CStr(Trim(JenisKartu_1)) + "')"
    RPT.Formulas(9) = "@jenis_kartu_2 = totext('" + CStr(Trim(JenisKartu_2)) + "')"
    'RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptJadwalBayar.rpt"
    
    '@@20112012 Tambahan
    JawabanCh = ""
    If Trim(UCase(CmbJawabanCH.Text)) = "TELEPON" Then
        JawabanCh = "via telepon ke nomor " & Trim(TxtPenawaranvia.Text)
    End If
    If Trim(UCase(CmbJawabanCH.Text)) = "TERTULIS" Then
        JawabanCh = "ke alamat " & Trim(TxtPenawaranvia.Text)
    End If
    RPT.Formulas(10) = "@pengajuan_melalui= totext('" + CStr(Trim(CmbPermohonanMelalui.Text)) + "')"
    RPT.Formulas(11) = "@tanggal_pengajuan= totext('" + CStr(Trim(Format(TxtTglPengajuan.Value, "dd-Mmm-yyyy"))) + "')"
    RPT.Formulas(12) = "@jawabanch= totext('" + CStr(Trim(JawabanCh)) + "')"
    
    '@@ 201112, Lihat apakah pembayaran hanya satu kali/lebih
    If LvJadwal.ListItems.Count > 1 Then
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptJadwalBayarnew.rpt"
    Else
        RPT.Formulas(13) = "@jatuhtempo= totext('" + CStr(LvJadwal.ListItems(1).SubItems(2) & "-" & LvJadwal.ListItems(1).SubItems(1)) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptJadwalBayarnew2.rpt"
    End If
    Call SHOW_PRN
End Sub

Private Sub IsiJadwal()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    
    'Hapus Data
    Cmdsql = "delete from ReportJadwalBayar "
    M_RPTCONN.Execute Cmdsql
    
    Cmdsql = "select * from tbljadwalpembayaran where nid='"
    Cmdsql = Cmdsql + TxtIdCpa.Text + "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    While Not M_Objrs.EOF
        Cmdsql = "insert into reportjadwalbayar (tahun,bulan,tunggakan,pembayaran,"
        Cmdsql = Cmdsql + "hutang_pokok,interest,balance_terakhir,tgl_bayar) values ('"
        Cmdsql = Cmdsql + CStr(M_Objrs("tahun")) + "','"
        Cmdsql = Cmdsql + CStr(M_Objrs("bulan")) + "','"
        Cmdsql = Cmdsql + CStr(M_Objrs("tunggakan")) + "','"
        Cmdsql = Cmdsql + CStr(M_Objrs("pembayaran")) + "','"
        Cmdsql = Cmdsql + CStr(M_Objrs("hutang_pokok")) + "','"
        Cmdsql = Cmdsql + CStr(M_Objrs("interest")) + "','"
        Cmdsql = Cmdsql + CStr(M_Objrs("balance_terakhir")) + "','"
        Cmdsql = Cmdsql + CStr(Format(M_Objrs("tglbayar"), "yyyy-mm-dd")) + "')"
        M_RPTCONN.Execute Cmdsql
        M_Objrs.MoveNext
    Wend
        
    Set M_Objrs = Nothing
End Sub

Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub

Private Sub CmdEdit_Click()
    Dim W As Integer
    Dim Pembayaran As Double
    
    Pembayaran = 0
    If LvJadwal.ListItems.Count = 0 Then
        'FrmTambahJadwalPembayaran.TxtTunggakan.Value = TxtPayment.Value
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    Else
'        For W = 1 To LvJadwal.ListItems.Count
'            Pembayaran = Pembayaran + Val(Replace(LvJadwal.ListItems(W).SubItems(4), ",", ""))
'        Next W
        'FrmTambahJadwalPembayaran.TxtTunggakan.Value = TxtPayment.Value - Pembayaran
        FrmTambahJadwalPembayaran.Caption = "EDIT"
        FrmTambahJadwalPembayaran.CmdTambah.Caption = "Update"
        FrmTambahJadwalPembayaran.TxtTglBayar.Value = Format(LvJadwal.SelectedItem.SubItems(8), "dd/mm/yyyy")
        FrmTambahJadwalPembayaran.txtpembayaran.Value = LvJadwal.SelectedItem.SubItems(4)
        FrmTambahJadwalPembayaran.TxtID.Text = LvJadwal.SelectedItem.Text
        FrmTambahJadwalPembayaran.TxtTunggakan.Value = LvJadwal.SelectedItem.SubItems(3)
        FrmTambahJadwalPembayaran.TxtHutangPokok.Value = LvJadwal.SelectedItem.SubItems(5)
        FrmTambahJadwalPembayaran.TxtInterest.Value = LvJadwal.SelectedItem.SubItems(6)
        FrmTambahJadwalPembayaran.txtBalanceTerakhir.Value = IIf(IsNull(LvJadwal.SelectedItem.SubItems(7)), "0", LvJadwal.SelectedItem.SubItems(7))
        FrmTambahJadwalPembayaran.txtpembayaran.Enabled = False
    End If
    FrmTambahJadwalPembayaran.Show vbModal
    Call IsiListJadwal
End Sub

Private Sub CmdHapus_Click()
    Dim a As String
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    If LvJadwal.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        Cmdsql = "delete from tbljadwalpembayaran where id='"
        Cmdsql = Cmdsql + LvJadwal.SelectedItem.Text + "'"
        M_OBJCONN.Execute Cmdsql
        MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
        Call IsiListJadwal
    End If
    
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub CmdTambah_Click()
    Dim W As Integer
    Dim Pembayaran As Double
    
    Pembayaran = 0
    If LvJadwal.ListItems.Count = 0 Then
        FrmTambahJadwalPembayaran.TxtTunggakan.Value = txtPayment.Value
    Else
        For W = 1 To LvJadwal.ListItems.Count
            Pembayaran = Pembayaran + Val(Replace(LvJadwal.ListItems(W).SubItems(4), ",", ""))
        Next W
        FrmTambahJadwalPembayaran.TxtTunggakan.Value = txtPayment.Value - Pembayaran
    End If
    FrmTambahJadwalPembayaran.Show vbModal
    Call IsiListJadwal
End Sub

Private Sub CmdTambahOtomatis_Click()
    Dim W As Integer
    Dim Cmdsql As String
    Dim Tunggakan As Double
    Dim Cicilan As Double
    Dim SisaBalance As Double
    Dim Tanggal As String
    Dim a As String
    
    a = MsgBox("Anda yakin akan membuat schedule otomatis? (Jika data schedule sebelumnya ada, maka data tersebut akan dihapus!", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        Cmdsql = "delete from tbljadwalpembayaran where nid='"
        Cmdsql = Cmdsql + TxtIdCpa.Text + "'"
        M_OBJCONN.Execute Cmdsql
        LvJadwal.ListItems.CLEAR
    Else
        Exit Sub
    End If
    
    If TxtInstallment.Value = 0 Then
        MsgBox "Instalment period tidak boleh 0(nol)!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtInstallment.Value < 0 Then
        MsgBox "Instalment period tidak boleh lebih kecil dari 0 (nol)!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglBayar.ValueIsNull = True Then
        MsgBox "Tanggal bayar tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    Cicilan = txtPayment.Value / TxtInstallment.Value
    Tunggakan = txtPayment.Value
    SisaBalance = txtPayment.Value
    
    n = 0
    For W = 1 To TxtInstallment.Value
        'Tunggakan = (TxtPayment.Value - (TxtPayment.Value / TxtInstallment.Value))
        
        SisaBalance = SisaBalance - Cicilan
        
        
        Tanggal = DateAdd("m", n, Format(TxtTglBayar.Value, "yyyy-mm-dd"))
        
        Tahun = Format(Tanggal, "yyyy")
        Bulan = Format(Tanggal, "dd-mmm")
        
        Cmdsql = "insert into tbljadwalpembayaran (nid,tahun,bulan,tunggakan,"
        Cmdsql = Cmdsql + "pembayaran,hutang_pokok,interest,balance_terakhir,tglbayar) values ('"
        Cmdsql = Cmdsql + FrmJadwalPembayaranCpa.TxtIdCpa.Text + "','"
        Cmdsql = Cmdsql + Tahun + "','"
        Cmdsql = Cmdsql + Bulan + "','"
        Cmdsql = Cmdsql + CStr(Tunggakan) + "','"
        Cmdsql = Cmdsql + CStr(Cicilan) + "','"
        Cmdsql = Cmdsql + CStr("0") + "','"
        Cmdsql = Cmdsql + CStr("0") + "','"
        Cmdsql = Cmdsql + CStr(SisaBalance) + "','"
        Cmdsql = Cmdsql + CStr(Format(Tanggal, "yyyy-mm-dd")) + "')"
        M_OBJCONN.Execute Cmdsql
        
        
        Tunggakan = Tunggakan - Cicilan
        n = n + 1
    Next W
    MsgBox "Schedule berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    Call IsiListJadwal
End Sub

Private Sub Form_Activate()
    Call NoSurat
End Sub

Private Sub Form_Load()
    Call HeaderListJadwal
    Call IsiListJadwal
    
End Sub

Private Sub IsiListJadwal()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listitem As listitem
    
    LvJadwal.ListItems.CLEAR
    
    Cmdsql = "select * from tbljadwalpembayaran where nid='"
    Cmdsql = Cmdsql + frmcpanew.IdCPA + "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    While Not M_Objrs.EOF
        Set listitem = LvJadwal.ListItems.ADD(, , M_Objrs("id"))
        listitem.SubItems(1) = M_Objrs("tahun")
        listitem.SubItems(2) = M_Objrs("bulan")
        listitem.SubItems(3) = Format(M_Objrs("tunggakan"), "##,###")
        listitem.SubItems(4) = Format(M_Objrs("pembayaran"), "##,###")
        listitem.SubItems(5) = Format(M_Objrs("hutang_pokok"), "##,###")
        listitem.SubItems(6) = Format(M_Objrs("interest"), "##,###")
        listitem.SubItems(7) = Format(M_Objrs("balance_terakhir"), "##,###")
        listitem.SubItems(8) = Format(M_Objrs("tglbayar"), "yyyy-mm-dd")
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub NoSurat()
    Dim TglSurat As String
    
    TglSurat = Format(MDIForm1.TDBDate1.Value, "ddmmyy")
    TxtNoSurat.Text = TglSurat & " " & TxtAgent.Text & TxtTL.Text & "/" & CStr(TxtInstallment.Value) & "-" & TxtFromOs.Text
    
End Sub
