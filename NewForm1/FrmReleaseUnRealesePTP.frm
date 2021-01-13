VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReleaseUnRealesePTP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Realese / Unrealese PTP"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   3975
      Begin VB.Label Label2 
         Caption         =   $"FrmReleaseUnRealesePTP.frx":0000
         Height          =   1095
         Left            =   180
         TabIndex        =   8
         Top             =   1260
         Width           =   3675
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmReleaseUnRealesePTP.frx":00D9
         Height          =   795
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   2220
      TabIndex        =   5
      Top             =   5220
      Width           =   1995
   End
   Begin VB.CommandButton CmdProses 
      Caption         =   "&Proses"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5220
      Width           =   1995
   End
   Begin VB.ComboBox CmbPilihGroup 
      Height          =   315
      ItemData        =   "FrmReleaseUnRealesePTP.frx":017B
      Left            =   240
      List            =   "FrmReleaseUnRealesePTP.frx":017D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4680
      Width           =   1635
   End
   Begin VB.CommandButton CmdUncek 
      Caption         =   "UNCEK"
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   4620
      Width           =   1155
   End
   Begin VB.CommandButton CmdCek 
      Caption         =   "CEK"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4620
      Width           =   1155
   End
   Begin MSComctlLib.ListView LvUser 
      Height          =   4335
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
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
End
Attribute VB_Name = "FrmReleaseUnRealesePTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IsiCombo()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbPilihGroup.CLEAR
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        CmbPilihGroup.AddItem Trim(MDIForm1.Text1.Text)
    Else
        CmbPilihGroup.AddItem "ALL"
        cmdsql = "select team from usertbl where usertype='6' order by team asc"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            While Not M_Objrs.EOF
                CmbPilihGroup.AddItem IIf(IsNull(M_Objrs("team")), "", M_Objrs("team"))
                M_Objrs.MoveNext
            Wend
        End If
        
        Set M_Objrs = Nothing
    End If
    
End Sub

Private Sub header()
    LvUser.ColumnHeaders.ADD 1, , "Userid", 1500
    LvUser.ColumnHeaders.ADD 2, , "Nama", 5000
    LvUser.ColumnHeaders.ADD 3, , "Team", 4000
End Sub

Private Sub IsiData()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    Dim a As Integer
    
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        cmdsql = " select * from  usertbl where "
        cmdsql = cmdsql + " team='"
        cmdsql = cmdsql + Trim(MDIForm1.Text1.Text) + "' and usertype='1' order by team,userid asc "
    Else
        cmdsql = " select * from  usertbl where usertype='1' order by team,userid asc "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            Set listItem = LvUser.ListItems.ADD(, , M_Objrs("userid"))
                listItem.SubItems(1) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("team")), "", M_Objrs("team"))
                
                If M_Objrs("f_status_ptp") = "ALL" Then
                    listItem.Checked = True
                End If
            M_Objrs.MoveNext
        Wend
        
        
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub CmdCek_Click()
    Dim w As Integer
    
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Tidak ada data user!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbPilihGroup.Text = "" Then
        MsgBox "Pilih kriteria data yang akan diceklist!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Jika pilihan= ALL
    If Trim(CmbPilihGroup.Text) = "ALL" Then
        For w = 1 To LvUser.ListItems.Count
            LvUser.ListItems(w).Checked = True
        Next w
    Else
        For w = 1 To LvUser.ListItems.Count
            If Trim(LvUser.ListItems(w).SubItems(2)) = Trim(CmbPilihGroup.Text) Then
                LvUser.ListItems(w).Checked = True
            End If
        Next w
    End If
    
    MsgBox "Data berhasil di ceklist!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdProses_Click()
    Dim K As Integer
    Dim cmdsql As String
    Dim a As String
    Dim Remarks As String
    
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Data user tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan memproses data ini?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    For K = 1 To LvUser.ListItems.Count
    
        'Jika dicentang maka semua data tanggal PTP di tampilkan
        If LvUser.ListItems(K).Checked = True Then
            cmdsql = "update usertbl set f_status_ptp='ALL' where userid='"
            cmdsql = cmdsql + Trim(LvUser.ListItems(K).Text) + "'"
            M_OBJCONN.Execute cmdsql
            
            'Informasikan ke agent melalui pesan
            Remarks = "Informasi : " + vbCrLf
            Remarks = Remarks + "---------------------------------------" + vbCrLf
            Remarks = Remarks + "Data PTP anda dapat ditampilkan untuk semua tanggal Tagih di list PTP," + vbCrLf
            Remarks = Remarks + "dengan mengklik tombol search Tgl.Tagih!"
            
            
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + Trim(LvUser.ListItems(K).Text) + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.Text1.Text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            
            M_OBJCONN.Execute cmdsql
        End If
        
        'Jika tidak dicentang, maka data ptp yang tampil hanya hari ini dan 3 hari berikutnya
        If LvUser.ListItems(K).Checked = False Then
            cmdsql = "update usertbl set f_status_ptp=null where userid='"
            cmdsql = cmdsql + Trim(LvUser.ListItems(K).Text) + "'"
            M_OBJCONN.Execute cmdsql
            
             'Informasikan ke agent melalui pesan
            Remarks = "Informasi : " + vbCrLf
            Remarks = Remarks + "---------------------------------------" + vbCrLf
            Remarks = Remarks + "Data PTP anda dapat ditampilkan untuk tanggal tagih hari ini," + vbCrLf
            Remarks = Remarks + " dan tiga hari berikutnya dari hari ini dengan mengklik tombol search tgl.tagih!"
            
            
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + Trim(LvUser.ListItems(K).Text) + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.Text1.Text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            
            M_OBJCONN.Execute cmdsql
        End If
        
    Next K
    
    MsgBox "Data berhasil di proses!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub CmdUncek_Click()
    Dim w As Integer
    
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Tidak ada data user!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbPilihGroup.Text = "" Then
        MsgBox "Pilih kriteria data yang akan diceklist!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Jika pilihan= ALL
    If Trim(CmbPilihGroup.Text) = "ALL" Then
        For w = 1 To LvUser.ListItems.Count
            LvUser.ListItems(w).Checked = False
        Next w
    Else
        For w = 1 To LvUser.ListItems.Count
            If Trim(LvUser.ListItems(w).SubItems(2)) = Trim(CmbPilihGroup.Text) Then
                LvUser.ListItems(w).Checked = False
            End If
        Next w
    End If
    
    MsgBox "Data berhasil di uncek!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub Form_Load()
    Call header
    Call IsiCombo
    Call IsiData
End Sub
