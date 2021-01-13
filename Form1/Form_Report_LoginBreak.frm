VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_Report_LoginBreak 
   Caption         =   "Form Report Login & Break"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7710
   LinkTopic       =   "Form3"
   ScaleHeight     =   8235
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      Height          =   2190
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox cmb_jenis 
         Height          =   315
         Left            =   1515
         TabIndex        =   17
         Top             =   1035
         Width           =   4065
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3180
         TabIndex        =   11
         Top             =   705
         Width           =   2400
      End
      Begin VB.ComboBox cboagentname 
         Height          =   315
         Left            =   1515
         TabIndex        =   10
         Top             =   705
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E87211&
         Caption         =   "Show Phone Number"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   14040
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton SSCommand2 
         BackColor       =   &H00F1E5DB&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2415
         Picture         =   "Form_Report_LoginBreak.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1395
         Width           =   1530
      End
      Begin VB.CommandButton cmdCari 
         Height          =   360
         Left            =   4020
         Picture         =   "Form_Report_LoginBreak.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1380
         Width           =   1530
      End
      Begin VB.CommandButton SSCommand1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   810
         Left            =   5625
         Picture         =   "Form_Report_LoginBreak.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   930
         Width           =   1545
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   13320
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   13800
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xls"
      End
      Begin TDBDate6Ctl.TDBDate tgl_call 
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   12
         Top             =   375
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   556
         Calendar        =   "Form_Report_LoginBreak.frx":139A
         Caption         =   "Form_Report_LoginBreak.frx":14B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_Report_LoginBreak.frx":151E
         Keys            =   "Form_Report_LoginBreak.frx":153C
         Spin            =   "Form_Report_LoginBreak.frx":159A
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mmm-yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   0
         Format          =   "dd-mm-yyyy"
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__-__-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tgl_call 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   13
         Top             =   375
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         Calendar        =   "Form_Report_LoginBreak.frx":15C2
         Caption         =   "Form_Report_LoginBreak.frx":16DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_Report_LoginBreak.frx":1746
         Keys            =   "Form_Report_LoginBreak.frx":1764
         Spin            =   "Form_Report_LoginBreak.frx":17C2
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mmm-yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   0
         Format          =   "dd-mm-yyyy"
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__-__-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
      End
      Begin Crystal.CrystalReport RPT 
         Left            =   13320
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jenis Report"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1065
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Telecollection"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   315
         Left            =   2970
         TabIndex        =   15
         Top             =   420
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Call"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   375
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6060
      Left            =   0
      TabIndex        =   0
      Top             =   2190
      Width           =   7695
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6660
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   255
         Width           =   915
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5280
         Left            =   45
         TabIndex        =   2
         Top             =   720
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   9313
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   75
         X2              =   7605
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Lead"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5805
         TabIndex        =   3
         Top             =   315
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form_Report_LoginBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sGetSPV As String
Dim strsql_temp As String
Private Sub cboagentname_Click()
    cboagentname_LostFocus
End Sub

Private Sub cboagentname_DropDown()
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    strsql = "select * from usertbl where usertype='1'"
    
'    If MDIForm1.Txtlevel = "Supervisor" Then
'        mwhere = " and team in ('" & MDIForm1.Text1.text & "') "
'    End If
    
    M_objrs.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cboagentname.clear
    While Not M_objrs.EOF
        cboagentname.AddItem IIf(IsNull(M_objrs!Userid), "", M_objrs!Userid)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing


End Sub

Private Sub cboagentname_LostFocus()
Dim MOBJR As New ADODB.Recordset
Set MOBJR = New ADODB.Recordset
MOBJR.CursorLocation = adUseClient

    strsql = "select * from usertbl where userid='" + cboagentname.text + "'"
    MOBJR.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not MOBJR.EOF Then
        Combo3.text = IIf(IsNull(MOBJR!agent), "", MOBJR!agent)
    End If
    
Set MOBJR = Nothing

End Sub

Private Sub cmb_jenis_DropDown()
    cmb_jenis.clear
    cmb_jenis.AddItem "Login & Logout"
    cmb_jenis.AddItem "Agent Break"
End Sub

Private Sub cmdCari_Click()
    Dim MOBJ As ADODB.Recordset
    Dim JML As Double
    Dim getSpvcode As String
    Dim getSpv_name As String
    Dim getUserid As String
    Dim getCampaign_code As String
    Dim getCampaign_name As String
    
    If cboagentname.text <> "" Then
        intvrl = InStr(1, cboagentname.text, "!", vbTextCompare)
        If intvrl <> 0 Then
            ArrayString = Split(cboagentname.text, "!", 2, vbTextCompare)
            getUserid = ArrayString(0)
            getUser_name = ArrayString(1)
        End If
    End If
    
    If cmb_jenis.text = "" Then
        MsgBox "Harap Isi Jenis Report", vbInformation, "Information"
        Exit Sub
    End If
    
    If cmb_jenis = "Login & Logout" Then
        strsql = "select username,status,waktu_login,waktu_logout, (waktu_logout::timestamp(0) - waktu_login::timestamp(0)) as ""Durasi"" from usertbl_log"
        mwhere = " where coalesce(waktu_logout, '') <> ''"
        
        If Not (tgl_call(0).ValueIsNull) And Not (tgl_call(1).ValueIsNull) Then
            If Len(mwhere) = 0 Then
                mwhere = " where  date(session_login) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
                mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
            Else
                mwhere = mwhere + "  and date(session_login) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
                mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
            End If
        Else
            MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
            Exit Sub
        End If
        
        If cboagentname.text <> Empty Then
            If Len(mwhere) = 0 Then
                mwhere = mwhere + " where username ='" + cboagentname.text + "'"
            Else
                mwhere = mwhere + " and username ='" + cboagentname.text + "'"
            End If
        End If
    Else
        strsql = " select agent,status_break,waktu_start,waktu_end, (waktu_end::timestamp(0) - waktu_start::timestamp(0)) as ""Durasi Break"" from tbl_autodialer_agent_break"
        mwhere = " where coalesce(waktu_start::varchar, '') <> '' and coalesce(waktu_end::varchar, '') <> '' and status_break not in ('ManualDial','start_autodialer','AutoDial','form break show')"
        
        If Not (tgl_call(0).ValueIsNull) And Not (tgl_call(1).ValueIsNull) Then
            If Len(mwhere) = 0 Then
                mwhere = " where  date(waktu_start) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
                mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
            Else
                mwhere = mwhere + "  and date(waktu_start) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
                mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
            End If
        Else
            MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
            Exit Sub
        End If
        
        If cboagentname.text <> Empty Then
            If Len(mwhere) = 0 Then
                mwhere = mwhere + " where agent ='" + cboagentname.text + "'"
            Else
                mwhere = mwhere + " and agent ='" + cboagentname.text + "'"
            End If
        End If
    End If
    Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    MOBJ.Open strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockOptimistic
    strsql_temp = strsql + mwhere
    
    txtlead.text = MOBJ.RecordCount
    Set DataGrid1.DATASOURCE = MOBJ
    cmdCari.Enabled = True
End Sub

Private Sub Combo3_Click()
Combo3_LostFocus
End Sub

Private Sub Combo3_DropDown()
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

    strsql = "select * from usertbl where usertype='1'"
        
'    If MDIForm1.Txtlevel = "Supervisor" Then
'        mwhere = " and  usertbl_groupspvcode in ('" & MDIForm1.txtUserName.text & "') "
'    End If

    M_objrs.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo3.clear
    While Not M_objrs.EOF
      Combo3.AddItem IIf(IsNull(M_objrs!agent), "", M_objrs!agent)
        M_objrs.MoveNext
    Wend
 Set M_objrs = Nothing

End Sub

Private Sub Combo3_LostFocus()
Dim MOBJR As New ADODB.Recordset
Set MOBJR = New ADODB.Recordset
   MOBJR.CursorLocation = adUseClient
   
strsql = "select * from usertbl where agent='" + Combo3.text + "'"
MOBJR.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not MOBJR.EOF Then
    cboagentname.text = IIf(IsNull(MOBJR!Userid), "", MOBJR!Userid)
End If
Set MOBJR = Nothing

End Sub

Private Sub Form_Load()
    tgl_call(0).Value = Format(FungsiWaktuServer, "MM - DD - YYYY")
    tgl_call(1).Value = Format(FungsiWaktuServer, "MM-DD-YYYY")
End Sub

Private Sub SSCommand1_Click()
    Call isi_data(strsql_temp)
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub

Private Sub isi_data(strsql As String)
On Error GoTo Salah
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
'    Dim objExcel        As Excel.Application
'    Dim objBook         As Excel.Workbook
'    Dim objSheet        As Excel.Worksheet

 '==== 13/06/2019 ===='
    Dim objExcel        As Object
    Dim objBook         As Object
    Dim objSheet        As Object
 '==== 13/06/2019 ===='
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    If M_objrs.RecordCount = 0 Then
        MsgBox "Data  tidak ada!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

   
Form_Save:
    CD_save.ShowSave
    Txtpath.text = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If

 Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Integer
        If M_objrs.state = 1 Then
            x = 0
            Y = M_objrs.fields().Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
    
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs Txtpath.text + ".csv", xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing

    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub

