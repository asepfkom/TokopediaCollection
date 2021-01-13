VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11820
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rpt 
      Left            =   5160
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox TxtLpa 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   24
      Top             =   7260
      Width           =   1515
   End
   Begin VB.TextBox TxtTotalData 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1380
      TabIndex        =   22
      Text            =   "0"
      Top             =   7260
      Width           =   795
   End
   Begin VB.CheckBox CekBulan 
      Caption         =   "04-Apr"
      Height          =   315
      Index           =   3
      Left            =   5040
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Data"
      Height          =   1695
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   11595
      Begin VB.CommandButton CmdKeluar 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   9960
         TabIndex        =   26
         Top             =   900
         Width           =   1455
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "&Export"
         Height          =   375
         Left            =   9960
         TabIndex        =   25
         Top             =   480
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1320
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.CommandButton CmdLoad 
         Caption         =   "&Load data"
         Height          =   375
         Left            =   9960
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdUncekAll 
         Caption         =   "UnCek All"
         Height          =   315
         Left            =   1020
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "Cek All"
         Height          =   315
         Left            =   1020
         TabIndex        =   17
         Top             =   180
         Width           =   975
      End
      Begin TDBNumber6Ctl.TDBNumber TxtTahun 
         Height          =   255
         Left            =   1020
         TabIndex        =   16
         Top             =   900
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   450
         Calculator      =   "FrmBalance.frx":0000
         Caption         =   "FrmBalance.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmBalance.frx":008C
         Keys            =   "FrmBalance.frx":00AA
         Spin            =   "FrmBalance.frx":00F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1179649
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "12-Des"
         Height          =   315
         Index           =   11
         Left            =   6660
         TabIndex        =   15
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "11-Nov"
         Height          =   315
         Index           =   10
         Left            =   5760
         TabIndex        =   14
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "10-Okt"
         Height          =   315
         Index           =   9
         Left            =   4860
         TabIndex        =   13
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "09-Sep"
         Height          =   315
         Index           =   8
         Left            =   3960
         TabIndex        =   12
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "08-Ags"
         Height          =   315
         Index           =   7
         Left            =   3000
         TabIndex        =   11
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "07-Jul"
         Height          =   315
         Index           =   6
         Left            =   2100
         TabIndex        =   10
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "06-Jun"
         Height          =   315
         Index           =   5
         Left            =   6660
         TabIndex        =   9
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "05-Mei"
         Height          =   315
         Index           =   4
         Left            =   5760
         TabIndex        =   8
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "03-Mar"
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   6
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "02-Feb"
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox CekBulan 
         Caption         =   "01-Jan"
         Height          =   315
         Index           =   0
         Left            =   2100
         TabIndex        =   4
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Tahun:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Pilih Bulan:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   795
      End
   End
   Begin MSComctlLib.ListView LvBalance 
      Height          =   5280
      Left            =   180
      TabIndex        =   0
      Top             =   1920
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   9313
      View            =   3
      LabelEdit       =   1
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Total LPA:"
      Height          =   195
      Left            =   2340
      TabIndex        =   23
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah Data:"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   7320
      Width           =   1095
   End
End
Attribute VB_Name = "FrmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StringBulan As String

Private Sub IsiHeader()
    LvBalance.ColumnHeaders.ADD 1, , "No.", 700
    LvBalance.ColumnHeaders.ADD 2, , "Custid", 2000
    LvBalance.ColumnHeaders.ADD 3, , "LPD", 2000
    LvBalance.ColumnHeaders.ADD 4, , "LPA", 2000
    LvBalance.ColumnHeaders.ADD 5, , "Bulan", 1000
    LvBalance.ColumnHeaders.ADD 6, , "Tahun", 1000
End Sub

Private Sub CmdCekAll_Click()
    Dim w As Integer
    
    For w = 0 To 11
        CekBulan(w).Value = vbChecked
    Next w
End Sub

Private Sub CmdExport_Click()
    Dim m_objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim w As Integer
    
    If LvBalance.ListItems.Count = 0 Then
        MsgBox "Data belum tersedia di listview balance! Klik load data terlebih dahulu!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    Pb1.Max = LvBalance.ListItems.Count
    
    M_RPTCONN.Execute "delete from tblbalance"
    For w = 1 To LvBalance.ListItems.Count
        Pb1.Value = w
        cmdsql = "insert into tblbalance (custid,lpd,lpa) values ('"
        cmdsql = cmdsql + Trim(LvBalance.ListItems(w).SubItems(1)) + "','"
        cmdsql = cmdsql + CStr(Format(LvBalance.ListItems(w).SubItems(2), "yyyy-mm-dd")) + "','"
        cmdsql = cmdsql + CStr(Format(LvBalance.ListItems(w).SubItems(3), "##############")) + "')"
        M_RPTCONN.Execute cmdsql
    Next w
    
    WaitSecs (2)
    rpt.Reset
    rpt.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptBalance.rpt"
    Call SHOW_PRN
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub CmdLoad_Click()
    Dim m_objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim listitem As listitem
    
    LvBalance.ListItems.CLEAR
    
    AmbilBulan
    'Jika belum memilih bulan sama sekali
    If StringBulan = "" Then
        MsgBox "Anda belum memilih salah satu bulan yang akan ditampilkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Jika belum memilih tahun
    If IsNull(TxtTahun.Value) = True Then
        MsgBox "Anda belum memilih tahun!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
'    'Ambil Custidnya dulu
'    cmdsql = "select distinct custid from tbllunas where  date_part('year',paydate)='"
'    cmdsql = cmdsql + CStr(TxtTahun.Value) + "' and date_part('month',paydate) in ("
'    cmdsql = cmdsql + StringBulan + ") order by custid asc"
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If m_objrs.RecordCount > 0 Then
'        PB1.Max = m_objrs.RecordCount
'        While Not m_objrs.EOF
'            PB1.Value = m_objrs.Bookmark
'            Set listitem = LvBalance.ListItems.ADD(, , m_objrs.Bookmark)
'                listitem.SubItems(1) = m_objrs("custid")
'            m_objrs.MoveNext
'        Wend
'    Else
'        TxtLpa.Text = "0"
'        TxtTotalData.Text = "0"
'        MsgBox "Data Tidak Tersedia!", vbOKOnly + vbInformation, "Informasi"
'    End If
'
'    TxtTotalData.Text = Format(m_objrs.RecordCount, "##,###")
'    Set m_objrs = Nothing
'
'    'Ambil data LPD dan LPA
'    Call LpdLpa

    Pb1.Max = 11
    For w = 0 To 11
        Pb1.Value = w
        If CekBulan(w).Value Then
            'Ambil Custidnya dulu
            bulan = w + 1
            cmdsql = "select distinct custid from tbllunas where  date_part('year',paydate)='"
            cmdsql = cmdsql + CStr(TxtTahun.Value) + "' and date_part('month',paydate)='" + CStr(bulan) + "' "
            cmdsql = cmdsql + " order by custid asc"
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
            If m_objrs.RecordCount > 0 Then
            'Pb1.Max = m_objrs.RecordCount
                While Not m_objrs.EOF
                    'Pb1.Value = m_objrs.Bookmark
                    Set listitem = LvBalance.ListItems.ADD(, , m_objrs.Bookmark)
                        listitem.SubItems(1) = m_objrs("custid")
                        listitem.SubItems(4) = w + 1
                        listitem.SubItems(5) = TxtTahun.Value
                    m_objrs.MoveNext
                Wend
            End If
        End If
    Next w
    
    TxtTotalData.Text = LvBalance.ListItems.Count
    Call LpdLpa
End Sub

Private Sub CmdUnCekAll_Click()
    Dim k As Integer
    
    For k = 0 To 11
        CekBulan(k).Value = vbUnchecked
    Next k
End Sub

Private Sub Form_Load()
    Call IsiHeader
    TxtTahun.Value = Format(Now, "yyyy")
End Sub

Private Sub AmbilBulan()
    Dim w As Integer
    Dim cmdsql As String
    Dim m_objrs As ADODB.Recordset
    Dim bulan As Integer
    Dim listitem As listitem
    
    StringBulan = ""
    
    For w = 0 To 11
        If CekBulan(w).Value Then
            If StringBulan = "" Then
                StringBulan = CStr(w + 1)
            Else
                StringBulan = StringBulan + "," + CStr(w + 1)
            End If
        End If
    Next w
    
    
End Sub

Private Sub LpdLpa()
    Dim m_objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim w As Integer
    Dim TotalLPA As Double
    
        
    TotalLPA = 0
    If LvBalance.ListItems.Count = 0 Then
        Exit Sub
    End If

    Pb1.Max = LvBalance.ListItems.Count
    For w = 1 To LvBalance.ListItems.Count
        Pb1.Value = w
        cmdsql = "select * from tbllunas where custid='"
        cmdsql = cmdsql + Trim(LvBalance.ListItems(w).SubItems(1)) + "' and date_part('year',paydate)='"
        cmdsql = cmdsql + Trim(LvBalance.ListItems(w).SubItems(5)) + "' and date_part('month',paydate)='"
        cmdsql = cmdsql + Trim(LvBalance.ListItems(w).SubItems(4)) + "' "
        cmdsql = cmdsql + " order by paydate desc limit 1"
        Set m_objrs = New ADODB.Recordset
        m_objrs.CursorLocation = adUseClient
        m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_objrs.RecordCount > 0 Then
            TotalLPA = TotalLPA + Val(m_objrs("payment"))
            LvBalance.ListItems(w).SubItems(2) = Format(m_objrs("paydate"), "yyyy-mm-dd")
            LvBalance.ListItems(w).SubItems(3) = Format(m_objrs("payment"), "##,###")
        End If
        Set m_objrs = Nothing
    Next w
    
   TxtLpa.Text = Format(TotalLPA, "##,###")
End Sub

Private Sub SHOW_PRN()
    rpt.RetrieveDataFiles
    rpt.WindowLeft = 0
    rpt.WindowTop = 0
    rpt.WindowState = crptMaximized
    rpt.WindowShowPrintBtn = True
    rpt.WindowShowRefreshBtn = True
    rpt.WindowShowSearchBtn = True
    rpt.WindowShowPrintSetupBtn = True
    rpt.WindowControls = True
    rpt.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub

Private Sub LvBalance_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvBalance.SortKey = ColumnHeader.Index - 1
    LvBalance.Sorted = True
End Sub
