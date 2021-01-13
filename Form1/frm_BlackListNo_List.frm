VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_BlackListNo_List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklist No.Telepon"
   ClientHeight    =   7140
   ClientLeft      =   1320
   ClientTop       =   315
   ClientWidth     =   10380
   Icon            =   "frm_BlackListNo_List.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10380
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search By"
      Height          =   3015
      Left            =   10440
      TabIndex        =   21
      Top             =   600
      Width           =   5055
      Begin VB.CommandButton Command5 
         Caption         =   "&Search"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox apajadah 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate StartDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   360
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "frm_BlackListNo_List.frx":058A
         Caption         =   "frm_BlackListNo_List.frx":06A2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_BlackListNo_List.frx":070E
         Keys            =   "frm_BlackListNo_List.frx":072C
         Spin            =   "frm_BlackListNo_List.frx":078A
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate EndDate 
         Height          =   315
         Left            =   3315
         TabIndex        =   23
         Top             =   360
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "frm_BlackListNo_List.frx":07B2
         Caption         =   "frm_BlackListNo_List.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_BlackListNo_List.frx":0936
         Keys            =   "frm_BlackListNo_List.frx":0954
         Spin            =   "frm_BlackListNo_List.frx":09B2
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Blacklist By"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2925
         TabIndex        =   24
         Top             =   435
         Width           =   255
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   20
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export"
      Height          =   495
      Left            =   8580
      TabIndex        =   19
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   11040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Upload Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   10095
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtListLocation 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1815
         TabIndex        =   12
         Top             =   270
         Width           =   3510
      End
      Begin VB.ComboBox cboSheet 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frm_BlackListNo_List.frx":09DA
         Left            =   1815
         List            =   "frm_BlackListNo_List.frx":09DC
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   645
         Width           =   3525
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Upload"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8580
         TabIndex        =   10
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label11 
         Caption         =   "* File Excel (.xls)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6060
         TabIndex        =   16
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Sheet"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Excel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari"
      Height          =   315
      Left            =   4020
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   2715
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   8580
      TabIndex        =   4
      Top             =   2370
      Width           =   1455
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   8580
      TabIndex        =   3
      Top             =   1770
      Width           =   1455
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   8580
      TabIndex        =   2
      Top             =   1170
      Width           =   1455
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   8580
      TabIndex        =   1
      Top             =   570
      Width           =   1455
   End
   Begin MSComctlLib.ListView LVBlackList 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   990
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8440
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD_Brows 
      Left            =   9720
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3765
      Left            =   120
      TabIndex        =   17
      Top             =   7200
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6641
      _Version        =   393216
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
   Begin VB.Label Label1 
      Caption         =   "Cari No.Telp:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "BLACKLIST NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   60
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   30
      Picture         =   "frm_BlackListNo_List.frx":09DE
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "frm_BlackListNo_List.frx":14E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frm_BlackListNo_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New ADODB.Recordset
Dim M_XLSCONN As New ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub apajadah_DropDown()
    Dim noinc As Double
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim listItem As listItem
    
    cmdsql = "select distinct userinput from tblblacklist where 1=1 and userinput <> '' order by 1"
        
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    noinc = 0
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_Objrs.EOF
        apajadah.AddItem M_Objrs!userinput
         M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub apajadah_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbosheet_Click()
    Dim OBJRECORD As New ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient

    ssql = "SELECT * FROM [" & cboSheet.text & "] "
        rsTemp.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set rsTemp = Nothing
        Set OBJRECORD = New ADODB.Recordset
        OBJRECORD.CursorLocation = adUseClient
        
    ssql = "SELECT * FROM [" & cboSheet.text & "] "
        DoEvents
        OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid1.DATASOURCE = OBJRECORD
        'txtcount.Text = OBJRECORD.RecordCount
        
    frm_BlackListNo_List.Height = 12000
End Sub

Private Sub CmdBrowse_Click()
    With CD_Brows
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
    End With
        
    txtListLocation.text = CD_Brows.FileName
    
    If CD_Brows.FileName = "" Then Exit Sub
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtListLocation.text & ";Extended Properties=Excel 8.0;"
    Set rsTemp = M_XLSCONN.OpenSchema(adSchemaTables)
    cboSheet.clear
    If rsTemp.EOF And rsTemp.BOF Then Exit Sub
        
    While Not rsTemp.EOF
        cboSheet.AddItem IIf(IsNull(rsTemp!table_name), "", rsTemp!table_name)
        rsTemp.MoveNext
    Wend
    
    Set rsTemp = Nothing
End Sub

Private Sub CmdCari_Click()
    Call isi_data
End Sub

Private Sub CmdEdit_Click()
    
    Dim cmdsql As String
    
    If LVBlackList.ListItems.Count = 0 Then
        Exit Sub
     Else
        With frm_blacklist
            .Caption = "Edit Data No.Telepon BlackList"
            .TxtNoTelp.text = Trim(LVBlackList.SelectedItem.SubItems(1))
            .txtketerangan.text = LVBlackList.SelectedItem.SubItems(2)
            .TxtID.text = LVBlackList.SelectedItem.SubItems(4)
            .Show vbModal
            If .ok Then
                cmdsql = "update tblblacklist set no_telp='"
                cmdsql = cmdsql + CStr(Trim(.TxtNoTelp.text)) + "', keterangan='"
                cmdsql = cmdsql + IIf(IsNull(.txtketerangan.text), "", Trim(.txtketerangan.text)) + "',tglinput='" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "',userinput='" + MDIForm1.Text7.text + "' where id='"
                cmdsql = cmdsql + Trim(.TxtID.text) + "'"
                
                M_OBJCONN.Execute cmdsql
                
                'Update flag di mgm
                Call update_flag_1
                
                LVBlackList.SelectedItem.SubItems(1) = .TxtNoTelp.text
                LVBlackList.SelectedItem.SubItems(2) = .txtketerangan.text
                LVBlackList.SelectedItem.SubItems(3) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
                LVBlackList.SelectedItem.SubItems(5) = MDIForm1.Text7.text
            End If
        End With
     End If
End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdHapus_Click()
    Dim Cmdsql_Cek As String
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim m_msgbox As String
    
    If LVBlackList.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    m_msgbox = MsgBox("Anda yakin akan menghapus no telepon:" & Trim(LVBlackList.SelectedItem.SubItems(1)), vbYesNo + vbQuestion, "Konfirmasi")
    
    If m_msgbox = vbNo Then
     Exit Sub
    End If
    
    cmdsql = "delete from tblblacklist where no_telp='"
    cmdsql = cmdsql + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    
    M_OBJCONN.Execute cmdsql
     
    'Update ke flag 0
    Call update_flag_0
    
    LVBlackList.ListItems.Remove LVBlackList.SelectedItem.Index
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdTambah_Click()
Dim noinc As Double
    Dim m_msgbox As Variant
    Dim listItem As listItem
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim Cmdsql_Cek As String
    Dim ADD_OK As Boolean
    Dim cmdsql_update
    
    With frm_blacklist
                .Caption = "Tambah Data Black List"
                .Show vbModal
                If .ok Then
                    cmdsql = "insert into tblblacklist (no_telp,keterangan,tglinput,userinput) values ('"
                    cmdsql = cmdsql + Trim(.TxtNoTelp.text) + "','"
                    cmdsql = cmdsql + IIf(IsNull(.txtketerangan.text), "", Trim(.txtketerangan.text)) + "',"
                    cmdsql = cmdsql + "'" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "','" + MDIForm1.Text7.text + "') "
                    'Cek data no telepon yang sama
                    Set M_Objrs = New ADODB.Recordset
                    M_Objrs.CursorLocation = adUseClient
                        Cmdsql_Cek = "select * from tblblacklist where no_telp='"
                        Cmdsql_Cek = Cmdsql_Cek + CStr(Trim(.TxtNoTelp.text)) + "'"
                    M_Objrs.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If M_Objrs.RecordCount <> 0 Then
                        m_msgbox = MsgBox("No Telepon sudah ada. Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan")
                        Exit Sub
                    End If
                    Set M_Objrs = Nothing
                    
                    M_OBJCONN.Execute cmdsql
                    
                    'Update flag ke tabel mgm
                    Call update_flag_1
                    
                    noinc = LVBlackList.ListItems.Count
                    noinc = noinc + 1
                    Set listItem = LVBlackList.ListItems.ADD(, , CStr(noinc))
                        listItem.SubItems(1) = .TxtNoTelp.text
                        listItem.SubItems(2) = .txtketerangan.text
                        listItem.SubItems(3) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
                        listItem.SubItems(5) = MDIForm1.Text7.text
                        
                        
                End If
    End With
End Sub

Private Sub Command1_Click()
    frm_BlackListNo_List.Height = 7590
End Sub

Private Sub Command2_Click()
    On Error GoTo abc
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, Row As Integer
    Dim a As String
    If LVBlackList.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To LVBlackList.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LVBlackList.ColumnHeaders(col)
        Next
     
        For Row = 2 To LVBlackList.ListItems.Count + 1
            For col = 1 To LVBlackList.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(Row, col).Value = LVBlackList.ListItems(Row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + LVBlackList.ListItems(Row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(Row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_Brows.ShowOpen
        a = CD_Brows.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
abc:
        MsgBox "No data to export", vbInformation, Me.Caption
    End If
End Sub

Private Sub Command3_Click()
    Dim rs As New ADODB.Recordset
    Dim temp_rs As ADODB.Recordset
    Dim str_sql As String
    Dim CustId As String
    Dim segment As String
    Dim action As String
    Dim asg_date As String
    Dim tag As String
    
    If CD_Brows.FileName = "" Then
        MsgBox "Browse Data Excel Terlebih Dahulu", vbInformation + vbOKOnly, "Information"
        Exit Sub
    End If
    
    If cboSheet.text = "" Then
       MsgBox "Pilih Sheet", vbInformation + vbOKOnly, "Information"
       cboSheet.SetFocus
       Exit Sub
    End If
        
    ssql = "SELECT * FROM [" & cboSheet.text & "]   "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
        
    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    rsTemporary.CursorType = adOpenDynamic
    rsTemporary.ActiveConnection = M_OBJCONN
    rsTemporary.LockType = adLockOptimistic
        
    rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic

    While Not rs.EOF
        no_telp = IIf(IsNull(rs("NOTELP")), "", rs("NOTELP"))
        
            str_sql = "INSERT INTO tblblacklist(no_telp,keterangan, tglinput, userinput)"
            str_sql = str_sql + " VALUES ('" & no_telp & "', 'UPLOAD', now(),    "
            str_sql = str_sql + " '" + MDIForm1.Text1.text + "') "
        
        M_OBJCONN.Execute str_sql
        rs.MoveNext
    Wend
        
    Set rs = Nothing
    MsgBox "Data Berhasil Di - Upload!"
    DataGrid1.Row = 0
    Unload Me
    Me.Show vbModal
End Sub

Private Sub Command4_Click()
    If frm_BlackListNo_List.Width < 15765 Then
        frm_BlackListNo_List.Width = 15765
        Command4.Caption = "<"
    Else
        frm_BlackListNo_List.Width = 10470
        Command4.Caption = ">"
    End If
End Sub

Private Sub Command5_Click()
    Call isi_data
End Sub

Private Sub Form_Load()
    
    Call header_lvblacklist
    Call isi_data
End Sub

Private Sub header_lvblacklist()
    'Membuat Header ListView Program
    LVBlackList.ColumnHeaders.ADD 1, , "No", 10 * 200
    LVBlackList.ColumnHeaders.ADD 2, , "No. Telepon Black List", 10 * 200
    LVBlackList.ColumnHeaders.ADD 3, , "Keterangan", 20 * 100
    LVBlackList.ColumnHeaders.ADD 4, , "Tgl Input", 20 * 100
    LVBlackList.ColumnHeaders.ADD 5, , "Id", 0
    LVBlackList.ColumnHeaders.ADD 6, , "Coding", 20 * 100
End Sub

Private Sub isi_data()
    Dim noinc As Double
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim listItem As listItem
    
    cmdsql = "select * from tblblacklist where 1=1 "
    If TxtCari.text <> Empty Then
        cmdsql = cmdsql + " and no_telp like '%"
        cmdsql = cmdsql + TxtCari.text + "%' "
    End If
    If StartDate.Value <> "" And EndDate.Value <> "" Then
    
        a = Format(StartDate.Value, "YYYY-MM-DD")
        B = Format(EndDate.Value, "YYYY-MM-DD")
        cmdsql = cmdsql + " and date(tglinput) between '" & a & "' and '" & B & "'"
    End If
    
    If apajadah.text <> Empty Then
        cmdsql = cmdsql + " and userinput = '" + apajadah.text + "' "
    End If
    
    cmdsql = cmdsql + " order by no_telp asc"
    
    LVBlackList.ListItems.clear
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    noinc = 0
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_Objrs.EOF
        noinc = noinc + 1
         Set listItem = LVBlackList.ListItems.ADD(, , CStr(noinc))
           listItem.SubItems(1) = IIf(IsNull(M_Objrs("no_telp")), "", M_Objrs("no_telp"))
           listItem.SubItems(2) = IIf(IsNull(M_Objrs("keterangan")), "", M_Objrs("keterangan"))
           listItem.SubItems(3) = Format(IIf(IsNull(M_Objrs("tglinput")), "", M_Objrs("tglinput")), "dd/mm/yyyy")
           listItem.SubItems(4) = IIf(IsNull(M_Objrs("id")), "", M_Objrs("id"))
           listItem.SubItems(5) = IIf(IsNull(M_Objrs("userinput")), "", M_Objrs("userinput"))
         M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub LVBlackList_DblClick()
    CmdEdit_Click
End Sub

Private Sub update_flag_1()
    Dim cmdsql_homeno As String
    Dim cmdsql_homeno2 As String
    
    Dim cmdsql_mobileno As String
    Dim cmdsql_mobileno2 As String
    
    Dim cmdsql_officeno As String
    Dim cmdsql_officeno2 As String
    
    Dim cmdsql_homenoadd1 As String
    Dim cmdsql_homenoadd2 As String
    
    Dim cmdsql_officenoadd1 As String
    Dim cmdsql_officenoadd2 As String
    
    Dim cmdsql_mobilenoadd1 As String
    Dim cmdsql_mobilenoadd2 As String
    
    Dim cmdsql_ec_telp As String
    
    '@@22062010 Update ke flag di mgm, supaya tanda no merah di agent tidak berat
    'Update flag telepon rumah
    cmdsql_homeno = "update mgm set f_homeno='1' where homeno='"
    cmdsql_homeno = cmdsql_homeno + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_homeno
    
    cmdsql_homeno2 = "update mgm set f_homeno2='1' where homeno2='"
    cmdsql_homeno2 = cmdsql_homeno2 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_homeno2
    
    'Update flag ke telepon hp
    cmdsql_mobileno = "update mgm set f_mobileno='1' where mobileno='"
    cmdsql_mobileno = cmdsql_mobileno + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_mobileno
    
    cmdsql_mobileno2 = "update mgm set f_mobileno2='1' where mobileno2='"
    cmdsql_mobileno2 = cmdsql_mobileno2 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_mobileno2
    
    'Update flag ke telepon office
    cmdsql_officeno = "update mgm set f_officeno='1' where officeno='"
    cmdsql_officeno = cmdsql_officeno + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_officeno
    
    cmdsql_officeno2 = "update mgm set f_officeno2='1' where officeno2='"
    cmdsql_officeno2 = cmdsql_officeno2 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_officeno2
    
    'Update flag ke telepon home add
    cmdsql_homenoadd1 = "update mgm set f_homenoadd1='1' where homenoadd1='"
    cmdsql_homenoadd1 = cmdsql_homenoadd1 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd1
    
    cmdsql_homenoadd2 = "update mgm set f_homenoadd2='1' where homenoadd2='"
    cmdsql_homenoadd2 = cmdsql_homenoadd2 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd2
    
    
    'Update flag ke telepon office add
    cmdsql_officenoadd1 = "update mgm set f_officenoadd1='1' where officenoadd1='"
    cmdsql_officenoadd1 = cmdsql_officenoadd1 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    cmdsql_officenoadd2 = "update mgm set f_officenoadd2='1' where officenoadd2='"
    cmdsql_officenoadd2 = cmdsql_officenoadd2 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    'Update flag ke telepon mobileno add
    cmdsql_mobilenoadd1 = "update mgm set f_mobilenoadd1='1' where mobilenoadd1='"
    cmdsql_mobilenoadd1 = cmdsql_mobilenoadd1 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd1
    
    cmdsql_mobilenoadd2 = "update mgm set f_mobilenoadd2='1' where mobilenoadd2='"
    cmdsql_mobilenoadd2 = cmdsql_mobilenoadd2 + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd2
    
    'Update flag ke telepon ec_telp
    cmdsql_ec_telp = "update mgm set f_ec_telp='1' where ec_telp='"
    cmdsql_ec_telp = cmdsql_ec_telp + Trim(frm_blacklist.TxtNoTelp.text) + "'"
    M_OBJCONN.Execute cmdsql_ec_telp
    
End Sub
Private Sub update_flag_0()
    Dim cmdsql_homeno As String
    Dim cmdsql_homeno2 As String
    
    Dim cmdsql_mobileno As String
    Dim cmdsql_mobileno2 As String
    
    Dim cmdsql_officeno As String
    Dim cmdsql_officeno2 As String
    
    Dim cmdsql_homenoadd1 As String
    Dim cmdsql_homenoadd2 As String
    
    Dim cmdsql_officenoadd1 As String
    Dim cmdsql_officenoadd2 As String
    
    Dim cmdsql_mobilenoadd1 As String
    Dim cmdsql_mobilenoadd2 As String
    
    Dim cmdsql_ec_telp As String
    
    '@@22062010 Update ke flag di mgm, supaya tanda no merah di agent tidak berat
    'Update flag telepon rumah
    cmdsql_homeno = "update mgm set f_homeno='0' where homeno='"
    cmdsql_homeno = cmdsql_homeno + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homeno
    
    cmdsql_homeno2 = "update mgm set f_homeno2='0' where homeno2='"
    cmdsql_homeno2 = cmdsql_homeno2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homeno2
    
    'Update flag ke telepon hp
    cmdsql_mobileno = "update mgm set f_mobileno='0' where mobileno='"
    cmdsql_mobileno = cmdsql_mobileno + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobileno
    
    cmdsql_mobileno2 = "update mgm set f_mobileno2='0' where mobileno2='"
    cmdsql_mobileno2 = cmdsql_mobileno2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobileno2
    
    'Update flag ke telepon office
    cmdsql_officeno = "update mgm set f_officeno='0' where officeno='"
    cmdsql_officeno = cmdsql_officeno + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officeno
    
    cmdsql_officeno2 = "update mgm set f_officeno2='0' where officeno2='"
    cmdsql_officeno2 = cmdsql_officeno2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officeno2
    
    'Update flag ke telepon home add
    cmdsql_homenoadd1 = "update mgm set f_homenoadd1='0' where homenoadd1='"
    cmdsql_homenoadd1 = cmdsql_homenoadd1 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd1
    
    cmdsql_homenoadd2 = "update mgm set f_homenoadd2='0' where homenoadd2='"
    cmdsql_homenoadd2 = cmdsql_homenoadd2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd2
    
    
    'Update flag ke telepon office add
    cmdsql_officenoadd1 = "update mgm set f_officenoadd1='0' where officenoadd1='"
    cmdsql_officenoadd1 = cmdsql_officenoadd1 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    cmdsql_officenoadd2 = "update mgm set f_officenoadd2='0' where officenoadd2='"
    cmdsql_officenoadd2 = cmdsql_officenoadd2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    'Update flag ke telepon mobileno add
    cmdsql_mobilenoadd1 = "update mgm set f_mobilenoadd1='0' where mobilenoadd1='"
    cmdsql_mobilenoadd1 = cmdsql_mobilenoadd1 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd1
    
    cmdsql_mobilenoadd2 = "update mgm set f_mobilenoadd2='0' where mobilenoadd2='"
    cmdsql_mobilenoadd2 = cmdsql_mobilenoadd2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd2
    
    'Update flag ke telepon ec_telp
    cmdsql_ec_telp = "update mgm set f_ec_telp='0' where ec_telp='"
    cmdsql_ec_telp = cmdsql_ec_telp + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_ec_telp

End Sub

