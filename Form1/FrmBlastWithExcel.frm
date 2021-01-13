VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSmsBlastExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send SMS Blast With Excel"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Upload"
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin MSComDlg.CommonDialog Cdupdate 
         Left            =   6360
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdSendSMS 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Send SMS"
         Enabled         =   0   'False
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox TxtJmlData 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1050
         Width           =   1095
      End
      Begin VB.CommandButton CmdBrowse 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Browse..."
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox CmbSheet 
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2445
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2190
         TabIndex        =   1
         Top             =   210
         Width           =   6015
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   1470
         Visible         =   0   'False
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jumlah data :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File excel:"
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
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pilih Sheet Excel :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmSmsBlastExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBrowse_Click()
form_save:
    With Cdupdate
    .CancelError = False
    .DialogTitle = "Inputkan data SMS"
    'On Error GoTo X
    .Filter = "Ms. Excel 9|*.xls"
    .ShowOpen
    Txtpath.Text = .FileName
    End With
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses Send SMS
        m_msgbox = MsgBox("Anda ingin Send SMS dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses send sms, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Send SMS dibatalkan!", vbOKOnly + vbInformation, "Informasi"
              CmdSendSMS.Enabled = False
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses Send Sms
          GoTo form_save        '-> maka goto form_save
        End If
    End If
 Call isi_sheet
 CmdSendSMS.Enabled = True
End Sub

Private Sub CmdSendSMS_Click()
 Dim mobj As New ADODB.Recordset
 Dim M_Objrs As ADODB.Recordset
 Dim Strsql As String
 Dim cmdsql As String
 Dim textsms As String
 Dim Nohp As String
 Dim NoAcc As String
 Dim koneksi_excel As New ADODB.Connection
 Dim WaktuServer As Date
 
 
 'Ambil Waktu server sekarang
 cmdsql = "select now()"
 Set M_Objrs = New ADODB.Recordset
 M_Objrs.CursorLocation = adUseClient
 M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 WaktuServer = Format(M_Objrs(0), "yyyy-mm-dd hh:mm:ss")
 Set M_Objrs = Nothing
 
 Set koneksi_excel = New ADODB.Connection
     koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & Trim(Txtpath.Text) & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
   
   Set mobj = New ADODB.Recordset
   mobj.CursorLocation = adUseClient
   
    '-> Membuka recordset Ms.Excel dengan status=gagal
    mobj.Open "Select * FROM [" & CmbSheet.Text & "]", _
                         koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText
    TxtJmlData.Text = mobj.RecordCount
    ProgressBar1.Max = mobj.RecordCount + 1
    
    While Not mobj.EOF
        ProgressBar1.Value = mobj.Bookmark
        DoEvents
        
        If Len(mobj(2)) > 160 Or Len(mobj(2)) = 0 Then
            MsgBox "Maaf! Cek data excel anda kembali, karena ada text sms yang kosong atau lebih besar dari 160 karakter!"
            Exit Sub
        End If
        
        If (mobj(0) = Empty Or mobj(0) = "") Or (mobj(1) = Empty Or mobj(1) = "") Then
            MsgBox "Maaf! Cek data excel anda kembali, karena ada no.telepon dan no.acc yang masih kosong!"
            Exit Sub
        End If
        mobj.MoveNext
    Wend
    
    mobj.MoveFirst
    While Not mobj.EOF
       
        textsms = Trim(Replace(mobj(2), "'", ""))
        Nohp = Trim(mobj(1))
        NoAcc = Trim(mobj(0))
        
        'Simpan data ke tabel smsblastexcel
        Strsql = "insert into smsblastexcel (noacc,nohp,textsms,tglupload) values ('"
        Strsql = Strsql + Trim(mobj(0)) + "','"
        Strsql = Strsql + Trim(mobj(1)) + "','"
        Strsql = Strsql + textsms + "','"
        Strsql = Strsql + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "')"
        M_OBJCONN.Execute Strsql
        
        'SimpanData di tabel outbox sms untuk dikirim
        Strsql = "insert into outbox (destinationnumber,textdecoded,senderid,creatorid) values ('"
        Strsql = Strsql + Nohp + "','"
        Strsql = Strsql + textsms + "','phone2','"
        Strsql = Strsql + Trim(NoAcc) + "-BlastExcelCard" + "')"
        M_OBJCONN1.Execute Strsql
        
        mobj.MoveNext
    Wend
    
    MsgBox "Data telah berhasil di inputkan, dan sekarang sedang proses pengiriman sms!", vbInformation + vbOKOnly, "Pesan"
    CmdSendSMS.Enabled = False
End Sub
Private Sub isi_sheet()
    Set koneksi_excel = CreateObject("ADODB.Connection")
    Set recordsetexcel = CreateObject("ADODB.Recordset")

    '-> Koneksi ke Ms.Excel
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & Txtpath.Text & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
                       
    '-> Membuka recordset Ms.Excel dengan status=gagal
    Set recordsetexcel = koneksi_excel.OpenSchema(adSchemaTables)
       
       
                       
                         
    'Mengsisi sheet pada CmbSheet
    CmbSheet.CLEAR
    CmbSheet.AddItem ""
    
    While Not recordsetexcel.EOF
       If Left(recordsetexcel.fields("Table_Name").Value, 4) <> "MSys" And Left(recordsetexcel.fields("Table_Name").Value, 3) <> "Sys" Then
        CmbSheet.AddItem recordsetexcel.fields("Table_Name")
       End If
       recordsetexcel.MoveNext
    Wend
                       
End Sub




