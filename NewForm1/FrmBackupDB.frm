VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmBackupDbToExcel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Backup Database Ke Excel"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   6060
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdUncekall 
      Caption         =   "&UncekAll"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   375
      Left            =   5340
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TxtJumlah 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "0"
      Top             =   6660
      Width           =   1215
   End
   Begin VB.CommandButton Vb 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   7860
      TabIndex        =   6
      Top             =   6600
      Width           =   1635
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus tabel backup"
      Height          =   375
      Left            =   5340
      TabIndex        =   5
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton CmbBackup 
      Caption         =   "&Backup ke excel..."
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton CmdDirektori 
      Caption         =   "..."
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   675
   End
   Begin VB.TextBox TxtDirektori 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   300
      Width           =   3135
   End
   Begin MSComctlLib.ListView LvTblBck 
      Height          =   4620
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8149
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Pilih tabel yang akan di backup:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "Pilih direktori untuk menyimpan file backup"
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
      Left            =   1080
      TabIndex        =   11
      Top             =   660
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah Tabel:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6660
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Direktori:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   795
   End
End
Attribute VB_Name = "FrmBackupDbToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HeaderListBackup()
    LvTblBck.ColumnHeaders.ADD 1, , "Id ", 1000
    LvTblBck.ColumnHeaders.ADD 2, , "Tgl.Backup", 3000
    LvTblBck.ColumnHeaders.ADD 3, , "Nama Tabel Backup", 5000
    LvTblBck.ColumnHeaders.ADD 4, , "User eksekusi", 4000
    LvTblBck.ColumnHeaders.ADD 5, , "Path backup excel", 5000
End Sub


Private Sub Isi_Tabel_Backup()
    Dim M_OBJRS As ADODB.Recordset
    Dim Cmdsql As String
    Dim listitem As listitem
    
    Cmdsql = "select distinct tbl_hst_upload_del.id, tbl_hst_upload_del.path_didb,"
    Cmdsql = Cmdsql + "tbl_hst_upload_del.path_excel, "
    Cmdsql = Cmdsql + "tbl_hst_upload_del.tgl_execute, tbl_hst_upload_del.user_excecute "
    Cmdsql = Cmdsql + "from information_schema.columns as ic,tbl_hst_upload_del "
    Cmdsql = Cmdsql + " where ic.table_schema='public' and ic.table_name=tbl_hst_upload_del.path_didb "
    Cmdsql = Cmdsql + " order by tbl_hst_upload_del.tgl_execute desc "
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJumlah.Text = M_OBJRS.RecordCount
    
    LvTblBck.ListItems.CLEAR
    If M_OBJRS.RecordCount = 0 Then
        Exit Sub
        Set M_OBJRS = Nothing
    End If
    
    While Not M_OBJRS.EOF
        Set listitem = LvTblBck.ListItems.ADD(, , M_OBJRS("id"))
            listitem.SubItems(1) = IIf(IsNull(M_OBJRS("tgl_execute")), "", Format(M_OBJRS("tgl_execute"), "yyyy-mm-dd hh:mm:ss"))
            listitem.SubItems(2) = IIf(IsNull(M_OBJRS("path_didb")), "", M_OBJRS("path_didb"))
            listitem.SubItems(3) = IIf(IsNull(M_OBJRS("user_excecute")), "", M_OBJRS("user_excecute"))
            listitem.SubItems(4) = IIf(IsNull(M_OBJRS("path_excel")), "", M_OBJRS("path_excel"))
        M_OBJRS.MoveNext
    Wend
    
    Set M_OBJRS = Nothing
End Sub

Private Sub isi_data(sPath As String, ssql)
    Dim M_OBJRS As ADODB.Recordset
    Dim Cmdsql As String
    Dim listitem As listitem
    Dim cmdsql_update As String
    Dim objExcel As Excel.Application
    Dim objBook  As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
    
   
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    
    If M_OBJRS.RecordCount = 0 Then
        'MsgBox "Data backup tidak ada !", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
    
    On Error GoTo salah
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Integer
        If M_OBJRS.state = 1 Then
            x = 0
            Y = M_OBJRS.fields().Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_OBJRS.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_OBJRS '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs sPath, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    'MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_OBJRS = Nothing
salah:
    Exit Sub
End Sub

Private Sub CmbBackup_Click()
    Dim K, W As Integer
    
    If LvTblBck.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang akan dibackup!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtDirektori.Text = "" Then
        MsgBox "Direktori belum diisi!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    K = 0
    
    'Cek data, apakah ada data yang dicentang?
    For W = 1 To LvTblBck.ListItems.Count
        If LvTblBck.ListItems(W).Checked = True Then
            K = K + 1
        End If
    Next W
    
    If K = 0 Then
        MsgBox "Tidak ada tabel yang dicentang!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Proses Backup
    Dim NamaFile, NamaTabel, Tanggal, Cmdsql As String
    PB1.Max = LvTblBck.ListItems.Count
    If Len(TxtDirektori.Text) > 3 Then
        TxtDirektori.Text = TxtDirektori.Text & "\"
    End If
    On Error GoTo salah
    For W = 1 To LvTblBck.ListItems.Count
        PB1.Value = W
        NamaTabel = IIf(IsNull(LvTblBck.ListItems(W).SubItems(2)), "", Trim(LvTblBck.ListItems(W).SubItems(2)))
        Tanggal = Format(MDIForm1.TDBDate1.Value, "ddmmyyyy") & "_" & Format(Now, "hhmmss")
        NamaFile = "Backup_" & Tanggal & "_" & NamaTabel & ".xls"
        
        Cmdsql = "select * from " & NamaTabel
        
        isi_data TxtDirektori.Text & NamaFile, Cmdsql
    Next W
    TxtDirektori.Text = ""
    MsgBox "Proses backup berhasil! Jika tabel tersebut kosong, maka file backup tidak akan di buat!", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
salah:
    MsgBox "Ada error: " & Err.Description, vbOKOnly + vbExclamation, "Peringatan"
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvTblBck.ListItems.Count = 0 Then
        MsgBox "Data tabel tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTblBck.ListItems.Count
        LvTblBck.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdDirektori_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    'Ganti 'This Is My Title' dengan judul yang ingin Anda 'letakkan pada kotak dialog "Browse For Folders" 'tersebut.
    szTitle = "This Is My Title"
    With tBrowseInfo
     .hWndOwner = Me.hwnd
     .lpszTitle = lstrcat(szTitle, "")
     .ulFlags = BIF_RETURNONLYFSDIRS + _
                BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
     sBuffer = Space(MAX_PATH)
     SHGetPathFromIDList lpIDList, sBuffer
     'Nilai sBuffer adalah directori yang dipilih oleh
     'user pada kotak dialog.
     sBuffer = Left(sBuffer, InStr(sBuffer, _
               vbNullChar) - 1)
     TxtDirektori.Text = sBuffer
    End If
End Sub

Private Sub CmdHapus_Click()
    Dim W As Integer
    Dim A As String
    
    If LvTblBck.ListItems.Count = 0 Then
        MsgBox "Tidak ada tabel yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    A = MsgBox("Yakin anda akan menghapus tabel backup?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If A = vbNo Then
        MsgBox "Penghapusan tabel dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTblBck.ListItems.Count
        On Error GoTo salah
        Cmdsql = "drop table " & IIf(IsNull(LvTblBck.ListItems(W).SubItems(2)), "", LvTblBck.ListItems(W).SubItems(2))
        M_OBJCONN.Execute Cmdsql
    Next W
    
    Call Isi_Tabel_Backup
    
    MsgBox "Penghapusan tabel berhasil", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
salah:
    MsgBox "Ada error: " & Err.Description, "Informasi"
    
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub CmdUncekall_Click()
    Dim W As Integer
    
    If LvTblBck.ListItems.Count = 0 Then
        MsgBox "Data tabel tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTblBck.ListItems.Count
        LvTblBck.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Call HeaderListBackup
    Call Isi_Tabel_Backup
End Sub
