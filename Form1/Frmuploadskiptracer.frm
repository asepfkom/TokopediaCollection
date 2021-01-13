VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmuploadskiptarcer 
   Caption         =   "Form2"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Upload"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4920
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2190
         TabIndex        =   5
         Top             =   210
         Width           =   6015
      End
      Begin VB.ComboBox CmbSheet 
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton CmdBrowse 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Browse..."
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox TxtJmlData 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   1050
         Width           =   1095
      End
      Begin VB.CommandButton CmdUpdateStatus 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Upload..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pilih Sheet Excel :"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   8
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File excel:"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jumlah data :"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   6
         Top             =   1020
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmuploadskiptarcer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBrowse_Click()
form_save:
    With Cdupdate
    .CancelError = False
    .DialogTitle = "Cari data masukan Upload data"
    'On Error GoTo X
    .Filter = "Ms. Excel 9|*.xls"
    .ShowOpen
    Txtpath.Text = .FileName
    End With
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Update dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Update dibatalkan!", vbOKOnly + vbInformation, "Informasi"
              CmdUpdateStatus.Enabled = False
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo form_save        '-> maka goto form_save
        End If
    End If
    
    'Aktifkan tombol updatestatus
    CmdUpdateStatus.Enabled = True
    Call isi_sheet
    
    
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

Private Sub CmdUpdateStatus_Click()
Dim mobj As New ADODB.Recordset
Dim koneksi_excel As New ADODB.Connection
Dim m_msgbox As String
Dim mobjNEW As New ADODB.Recordset
    'Konfirmasi dulu ke user, apakah akan melanjutkan submit data??
    m_msgbox = MsgBox("Anda yakin akan melanjutkan proses upload?", vbYesNo + vbQuestion, "Konfirmasi")
    
    '->Jika membatalkan proses update
    If m_msgbox = vbNo Then
      Txtpath.Text = ""
      CmdUpdateStatus.Enabled = False
      Exit Sub
    End If
     
    'Jika tidak ada sheet yang dipilih
    If CmbSheet.Text = "" Then
      MsgBox "Pilih sheet dari data yang akan diupdate!", vbOKOnly + vbInformation, "Informasi"
      Exit Sub
    End If
    Set koneksi_excel = New ADODB.Connection
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & Txtpath.Text & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
   
   Set mobj = New ADODB.Recordset
   mobj.CursorLocation = adUseClient
   
    '-> Membuka recordset Ms.Excel dengan status=gagal
    mobj.Open "Select * FROM [" & CmbSheet.Text & "]", _
                         koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText
TxtJmlData.Text = mobj.RecordCount
Set mobjNEW = New ADODB.Recordset
mobjNEW.CursorLocation = adUseClient
mobjNEW.Open "select * from opening_screen where name=''", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not mobj.EOF
    Strsql = "insert into  opening_screen(name,personal_alamat,personal_telp, "
    Strsql = Strsql + " personal_hp,office_alamat,office_telp,office_hp,familiy1_name ,familiy1_alamat,"
    Strsql = Strsql + " familiy1_telp,familiy1_hp ,familiy2_name,familiy2_alamat,familiy2_telp ,familiy2_hp ,"
    Strsql = Strsql + " familiy3_name,familiy3_alamat,familiy3_telp,familiy3_hp ,"
    Strsql = Strsql + " friend1_name,friend1_alamat,friend1_telp,friend1_hp,"
    Strsql = Strsql + " friend2_name,friend2_alamat,friend2_telp ,friend2_hp,"
    Strsql = Strsql + " friend3_name,friend3_alamat,friend3_telp ,friend3_hp,"
    Strsql = Strsql + " tglupdate,tglinsert,f_dl,tgldownload,REMARKS,stsaccount ,hp,flagtarik) values ( "
    Strsql = Strsql + "'" + IIf(IsNull(mobj!Name), "", mobj!Name) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!personal_alamat), "", mobj!personal_alamat) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!personal_telp), "", mobj!personal_telp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!personal_hp), "", mobj!personal_hp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!office_alamat), "", mobj!office_alamat) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!office_telp), "", mobj!office_telp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!office_hp), "", mobj!office_hp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy1_name), "", mobj!familiy1_name) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy1_alamat), "", mobj!familiy1_alamat) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy1_telp), "", mobj!familiy1_telp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy1_hp), "", mobj!familiy1_hp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy2_name), "", mobj!familiy2_name) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy2_alamat), "", mobj!familiy2_alamat) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy2_telp), "", mobj!familiy2_telp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy2_hp), "", mobj!familiy2_hp) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy3_name), "", mobj!familiy3_name) + "',"
    Strsql = Strsql + "'" + IIf(IsNull(mobj!familiy3_alamat), "", mobj!familiy3_alamat) + "',"
    
mobjNEW!familiy3_alamat = IIf(IsNull(mobj!familiy3_alamat), "", mobj!familiy3_alamat)
mobjNEW!familiy3_telp = IIf(IsNull(mobj!familiy3_telp), "", mobj!familiy3_telp)
mobjNEW!familiy3_hp = IIf(IsNull(mobj!familiy3_hp), "", mobj!familiy3_hp)
mobjNEW!friend1_name = IIf(IsNull(mobj!friend1_name), "", mobj!friend1_name)
mobjNEW!friend1_alamat = IIf(IsNull(mobj!friend1_alamat), "", mobj!friend1_alamat)
mobjNEW!friend1_telp = IIf(IsNull(mobj!friend1_telp), "", mobj!friend1_telp)
mobjNEW!friend1_hp = IIf(IsNull(mobj!friend1_hp), "", mobj!friend1_hp)
mobjNEW!friend2_name = IIf(IsNull(mobj!friend2_name), "", mobj!friend2_name)
mobjNEW!friend2_alamat = IIf(IsNull(mobj!friend2_alamat), "", mobj!friend2_alamat)
mobjNEW!friend2_telp = IIf(IsNull(mobj!friend2_telp), "", mobj!friend2_telp)
mobjNEW!friend2_hp = IIf(IsNull(mobj!friend2_hp), "", mobj!friend2_hp)
mobjNEW!friend3_name = IIf(IsNull(mobj!friend3_name), "", mobj!friend3_name)
mobjNEW!friend3_alamat = IIf(IsNull(mobj!friend3_alamat), "", mobj!friend3_alamat)
mobjNEW!friend3_telp = IIf(IsNull(mobj!friend3_telp), "", mobj!friend3_telp)
mobjNEW!friend3_hp = IIf(IsNull(mobj!friend3_hp), "", mobj!friend3_hp)
mobjNEW!tglupdate = IIf(IsNull(mobj!tglupdate), Null, mobj!tglupdate)
mobjNEW!tglinsert = IIf(IsNull(mobj!tglinsert), "", mobj!tglinsert)
mobjNEW!f_dl = IIf(IsNull(mobj!f_dl), "", mobj!f_dl)
mobjNEW!tgldownload = IIf(IsNull(mobj!tgldownload), Null, mobj!tgldownload)
mobjNEW!Remarks = IIf(IsNull(mobj!Remarks), "", mobj!Remarks)
'mobjNEW!REMARKS =
mobjNEW!stsaccount = IIf(IsNull(mobj!stsaccount), "", mobj!stsaccount)
mobjNEW!hp = IIf(IsNull(mobj!hp), "", mobj!hp)
mobjNEW!flagtarik = IIf(IsNull(mobj!flagtarik), 0, mobj!flagtarik)
mobjNEW.update
mobj.MoveNext
Wend

CmdUpdateStatus.Enabled = False
MsgBox "data telah diupload"

End Sub


