VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListReportTelepon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Report Telepon"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12585
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "&UnCekAll"
      Height          =   435
      Left            =   1260
      TabIndex        =   4
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdFollowUp 
      Caption         =   "&Follow up"
      Height          =   435
      Left            =   3900
      TabIndex        =   3
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   435
      Left            =   2580
      TabIndex        =   2
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdLoadData 
      Caption         =   "&Load data"
      Height          =   435
      Left            =   5040
      TabIndex        =   0
      Top             =   7740
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   315
      Left            =   8820
      TabIndex        =   1
      Top             =   7800
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView LvListProblemTelepon 
      Height          =   7620
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   13441
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
End
Attribute VB_Name = "FrmListReportTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HeaderList()
    LvListProblemTelepon.ColumnHeaders.ADD 1, , "ID", 900
    LvListProblemTelepon.ColumnHeaders.ADD 2, , "Status", 1200
    LvListProblemTelepon.ColumnHeaders.ADD 3, , "Tgl.Pengajuan", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 4, , "Userid", 1000
    LvListProblemTelepon.ColumnHeaders.ADD 5, , "Nama", 2000
    LvListProblemTelepon.ColumnHeaders.ADD 6, , "Telepon Masalah", 2000
    LvListProblemTelepon.ColumnHeaders.ADD 7, , "Jenis Kerusakan", 5000
    LvListProblemTelepon.ColumnHeaders.ADD 8, , "Keterangan", 4500
    
    '@@18012012 Tambahan
    LvListProblemTelepon.ColumnHeaders.ADD 9, , "Tanggal Solusi", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 10, , "Solusi Oleh", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 11, , "Keterangan", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 12, , "Jenis Telepon", 1500
End Sub


Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvListProblemTelepon.ListItems.Count
        LvListProblemTelepon.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdFollowUp_Click()
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If UCase(LvListProblemTelepon.SelectedItem.SubItems(1)) = "FIXED" Then
        MsgBox "Masalah sudah fix! tidak dapat di edit lagi!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    With FrmFollowUpProblemTelepon
        .TxtID.Text = LvListProblemTelepon.SelectedItem.Text
        .TxtTglPengajuan.Text = LvListProblemTelepon.SelectedItem.SubItems(2)
        .TxtUserid.Text = LvListProblemTelepon.SelectedItem.SubItems(3)
        .TxtNama.Text = LvListProblemTelepon.SelectedItem.SubItems(4)
        .txtNotelp.Text = LvListProblemTelepon.SelectedItem.SubItems(5)
        .TxtJenisKerusakan.Text = LvListProblemTelepon.SelectedItem.SubItems(6)
        .txtketerangan.Text = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(7)), "", LvListProblemTelepon.SelectedItem.SubItems(7))
        
        .TxtTglSolusi.Value = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(8)), Format(Now, "dd/mm/yyyy"), Format(LvListProblemTelepon.SelectedItem.SubItems(8), "dd/mm/yyyy"))
        .TxtSolusiOleh.Text = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(9)), "", LvListProblemTelepon.SelectedItem.SubItems(9))
        .TxtKetSolusi.Text = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(10)), "", LvListProblemTelepon.SelectedItem.SubItems(10))
        .CmbStatusSolusi.Text = IIf(UCase(LvListProblemTelepon.SelectedItem.SubItems(1)) = "NOT FOLLOW UP", "Follow Up", LvListProblemTelepon.SelectedItem.SubItems(1))
        .Show vbModal
    End With
        
End Sub

Private Sub CmdHapus_Click()
    Dim Cmdsql As String
    Dim a As String
    Dim W As Integer
    
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan menghapus data?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    If a = vbYes Then
        For W = 1 To LvListProblemTelepon.ListItems.Count
            If LvListProblemTelepon.ListItems(W).Checked = True Then
                Cmdsql = "delete from tbl_problem_telepon where id='"
                Cmdsql = Cmdsql + CStr(LvListProblemTelepon.ListItems(W).Text) + "'"
                M_OBJCONN.Execute Cmdsql
            End If
        Next W
    End If
    
    MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Infromasi"
    Call IsiData
End Sub

Private Sub CmdLoadData_Click()
    Call IsiData
End Sub

Private Sub CmdUncekAll_Click()
    Dim W As Integer
    
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvListProblemTelepon.ListItems.Count
        LvListProblemTelepon.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Call HeaderList
End Sub

Public Sub IsiData()
    Dim M_Objrs As ADODB.Recordset
    Dim Cmdsql As String
    Dim listitem As listitem
    Dim K As Integer
    
    Cmdsql = "select * from tbl_problem_telepon order by status_solusi,tgl_pengajuan asc"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvListProblemTelepon.ListItems.CLEAR
    
    If M_Objrs.RecordCount > 0 Then
        PB1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            Set listitem = LvListProblemTelepon.ListItems.ADD(, , M_Objrs("id"))
                listitem.SubItems(1) = IIf(IsNull(M_Objrs("status_solusi")), "NOT FOLLOW UP", M_Objrs("status_solusi"))
                listitem.SubItems(2) = Format(M_Objrs("tgl_pengajuan"), "yyyy-mm-dd hh:nn:ss")
                listitem.SubItems(3) = M_Objrs("userid")
                listitem.SubItems(4) = M_Objrs("nama")
                listitem.SubItems(5) = M_Objrs("no_telepon")
                listitem.SubItems(6) = M_Objrs("jenis_kerusakan")
                listitem.SubItems(7) = IIf(IsNull(M_Objrs("keterangan")), "", M_Objrs("keterangan"))
                
                '@@18012013 Tambahan
                listitem.SubItems(8) = IIf(IsNull(M_Objrs("tgl_solusi")), "", Format(M_Objrs("tgl_solusi"), "yyyy-mm-dd"))
                listitem.SubItems(9) = IIf(IsNull(M_Objrs("solusi_by")), "", M_Objrs("solusi_by"))
                
                listitem.SubItems(10) = IIf(IsNull(M_Objrs("solusi")), "", M_Objrs("solusi"))
                listitem.SubItems(11) = IIf(IsNull(M_Objrs("jenis_telepon")), "", M_Objrs("jenis_telepon"))
                
                K = 1
                
                If IsNull(M_Objrs("status_solusi")) = True Or M_Objrs("status_solusi") = "" Then
                     LvListProblemTelepon.ForeColor = vbRed
                     For K = 1 To 11
                        listitem.ListSubItems(K).ForeColor = vbRed
                     Next K
                End If
                
                If UCase(M_Objrs("status_solusi")) = "FOLLOW UP" Then
                     LvListProblemTelepon.ForeColor = vbYellow
                     For K = 1 To 11
                        listitem.ListSubItems(K).ForeColor = vbYellow
                     Next K
                End If
                
                If UCase(M_Objrs("status_solusi")) = "FIXED" Then
                     LvListProblemTelepon.ForeColor = vbGreen
                     For K = 1 To 11
                        listitem.ListSubItems(K).ForeColor = vbGreen
                     Next K
                End If
                
            M_Objrs.MoveNext
        Wend
    Else
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    Set M_Objrs = Nothing
End Sub


Private Sub LvListProblemTelepon_DblClick()
    CmdFollowUp_Click
End Sub

