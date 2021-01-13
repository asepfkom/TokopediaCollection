VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmUploadCPAPTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload CPA dan PTP"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Upload"
      Height          =   1845
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9525
      Begin MSComDlg.CommonDialog CDUpload 
         Left            =   7560
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdUpload 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Upload..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1065
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
         Width           =   1065
      End
      Begin VB.ComboBox CmbSheet 
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4335
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
         Left            =   150
         TabIndex        =   6
         Top             =   1500
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
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
         Left            =   120
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
Attribute VB_Name = "FrmUploadCPAPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBrowse_Click()
form_save:
    With CDUpload
    .CancelError = False
    .DialogTitle = "Cari data masukan Upload data"
    
    .Filter = "Ms. Excel 9|*.xls"
    .ShowOpen
    TxtPath.Text = .FileName
    End With
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Upload dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Upload CPA dan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
              cmdupload.Enabled = False
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo form_save        '-> maka goto form_save
        End If
    End If
 Call isi_sheet
 cmdupload.Enabled = True
End Sub

Private Sub isi_sheet()
    Set koneksi_excel = CreateObject("ADODB.Connection")
    Set recordsetexcel = CreateObject("ADODB.Recordset")

    '-> Koneksi ke Ms.Excel
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & TxtPath.Text & _
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

Private Sub CmdUpload_Click()
    Dim mobj As New ADODB.Recordset
    Dim koneksi_excel As New ADODB.Connection
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Vrdate As String
    Dim JmlPay As Double
    Dim w As Integer

    Set koneksi_excel = New ADODB.Connection
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & TxtPath.Text & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"

    Set mobj = New ADODB.Recordset
    mobj.CursorLocation = adUseClient

     '-> Membuka recordset Ms.Excel dengan status=gagal
     mobj.Open "Select * FROM [" & CmbSheet.Text & "]", _
                          koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText

     If mobj.RecordCount = 0 Then
        MsgBox "Tidak ada data yang diupload!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
     End If


     TxtJmlData.Text = mobj.RecordCount
     ProgressBar1.Max = mobj.RecordCount + 1
     While Not mobj.EOF
        ProgressBar1.Value = mobj.Bookmark
        DoEvents

        'Menginputkan data ke dalam tabel tblcpa
        cmdsql = "insert into tblcpa(vcustid,nttlpayment,nbalance,nprincipal,nperiod,vjust,dpropsal)"
        cmdsql = cmdsql + " values ('"
        cmdsql = cmdsql + IIf(IsNull(mobj(0).Value), "", mobj(0).Value) + "','"
        cmdsql = cmdsql + IIf(IsNull(mobj(1).Value), "", CStr(mobj(1).Value)) + "','"
        cmdsql = cmdsql + IIf(IsNull(mobj(2).Value), "", CStr(mobj(2).Value)) + "','"
        cmdsql = cmdsql + IIf(IsNull(mobj(3).Value), "", CStr(mobj(3).Value)) + "','"
        cmdsql = cmdsql + IIf(IsNull(mobj(4).Value), "", CStr(mobj(4).Value)) + "','"
        cmdsql = cmdsql + IIf(IsNull(mobj(5).Value), "", CStr(mobj(5).Value)) + "','"
        cmdsql = cmdsql + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "')"
        M_OBJCONN.Execute cmdsql


        'Ambil data payment terakhir dulu
        cmdsql = "select * from tbllunas where custid ='"
        cmdsql = cmdsql + Trim(mobj(0).Value) + "' order by paydate desc limit 1"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

        'Jika Data Payment Di temukan maka status Account ubah ke PTP-PO, Jika tidak maka PTP-NE
        If M_Objrs.RecordCount > 0 Then
            'Status Account PTP-PO
            Vrdate = DateAdd("m", 1, Format(M_Objrs("paydate"), "yyyy-mm-dd"))
            cmdsql = "update mgm set amountptp='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', ttlptp='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', f_cek_new='PTP-PO',"
            cmdsql = cmdsql + "dateptp='"
            cmdsql = cmdsql + Format(Vrdate, "yyyy-mm-dd") + "', tenor='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(4).Value), "0", Round(mobj(4).Value))) + "', kethslkerja_new='PTP-POP',  kethslkerjadesc_new='PTP-POP' "
            cmdsql = cmdsql + " where custid='"
            cmdsql = cmdsql + mobj(0).Value + "'"
            M_OBJCONN.Execute cmdsql
        Else
            'Status Account PTP-NE
            Vrdate = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")
            cmdsql = "update mgm set amountnew='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', "
            cmdsql = cmdsql + " f_cek_new='PTP-NE',"
            cmdsql = cmdsql + "tglptpnew='"
            cmdsql = cmdsql + Format(Vrdate, "yyyy-mm-dd") + "', tenor='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(4).Value), "0", Round(mobj(4).Value))) + "', dateptp='"
            cmdsql = cmdsql + Format(Vrdate, "yyyy-mm-dd") + "',amountptp='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', ttlptp='"
            cmdsql = cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', kethslkerja_new='PTP-NEW', kethslkerjadesc_new='PTP-NEW' "
            cmdsql = cmdsql + " where custid='"
            cmdsql = cmdsql + mobj(0).Value + "'"
            M_OBJCONN.Execute cmdsql
        End If

        'Inputkan data ke tabel negoptp dan reserve
        'Jika tenor = 0 atau 1
        If Round(Val(mobj(4).Value)) = 0 Or Round(Val(mobj(4).Value)) = 1 Then
            '@@ 12 Januari 2012, hapus data di tblnegoptp
            cmdsql = "delete from tblnegoptp where custid='"
            cmdsql = cmdsql + Trim(mobj(0).Value) + "'"
            M_OBJCONN.Execute cmdsql

            'Inputkan ke tblnegoptp
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + mobj(0).Value + "', "
            cmdsql = cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.Execute cmdsql
            ' isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + mobj(0).Value + "', "
            cmdsql = cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'" + MDIForm1.Text1.Text + "','P')"
            M_OBJCONN.Execute cmdsql
        Else
            'Jika tenor lebih besar dari 1

            '@@ 12 Januari 2012, hapus data di tblnegoptp
            cmdsql = "delete from tblnegoptp where custid='"
            cmdsql = cmdsql + Trim(mobj(0).Value) + "'"
            M_OBJCONN.Execute cmdsql

            '@@ 12 Januari 2012, hapus data di tblreserve
            cmdsql = "delete from tblreserve where custid='"
            cmdsql = cmdsql + Trim(mobj(0).Value) + "'"
            M_OBJCONN.Execute cmdsql

            JmlPay = Val(mobj(1).Value) / Round(Val(mobj(4).Value))

            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + mobj(0).Value + "', "
            cmdsql = cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(IIf(IsNull(JmlPay), "0", JmlPay)) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.Execute cmdsql
            ' isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + mobj(0).Value + "', "
            cmdsql = cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(IIf(IsNull(JmlPay), "0", JmlPay)) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'" + MDIForm1.Text1.Text + "','P')"
            M_OBJCONN.Execute cmdsql

            For w = 1 To Round(Val(mobj(4).Value)) - 1
                Vrdate = DateAdd("m", 1, Format(Vrdate, "yyyy-mm-dd"))
                cmdsql = "INSERT INTO tblreserve "
                cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                cmdsql = cmdsql + "VALUES "
                cmdsql = cmdsql + "('" + mobj(0).Value + "', "
                cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
                cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "'IPO')"
                M_OBJCONN.Execute cmdsql

                cmdsql = "INSERT INTO TblNegoptp_log "
                cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
                cmdsql = cmdsql + "VALUES "
                cmdsql = cmdsql + "('" + mobj(0).Value + "', "
                cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
                cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "'" + MDIForm1.Text1.Text + "','R')"
                M_OBJCONN.Execute cmdsql
            Next w
        End If

        Set M_Objrs = Nothing
        mobj.MoveNext
     Wend
     MsgBox "Data telah di upload!", vbInformation + vbOKOnly, "Pesan"
     cmdupload.Enabled = False
End Sub



''@@ 12 Januari 2012, Force Gaby Payment dikurangi 3 hari
'Private Sub ForceGaby()
'    Dim mobj As New ADODB.Recordset
'    Dim koneksi_excel As New ADODB.Connection
'    Dim Cmdsql As String
'    Dim M_Objrs As ADODB.Recordset
'    Dim Vrdate As String
'    Dim JmlPay As Double
'    Dim W As Integer
'
'    Set koneksi_excel = New ADODB.Connection
'    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                       "Data Source=" & TxtPath.Text & _
'                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
'
'    Set mobj = New ADODB.Recordset
'    mobj.CursorLocation = adUseClient
'
'     '-> Membuka recordset Ms.Excel dengan status=gagal
'     mobj.Open "Select * FROM [" & CmbSheet.Text & "]", _
'                          koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText
'
'     If mobj.RecordCount = 0 Then
'        MsgBox "Tidak ada data yang diupload!", vbOKOnly + vbExclamation, "Peringatan"
'        Exit Sub
'     End If
'
'
'     TxtJmlData.Text = mobj.RecordCount
'     ProgressBar1.Max = mobj.RecordCount + 1
'     While Not mobj.EOF
'        ProgressBar1.Value = mobj.Bookmark
'        DoEvents
'
'        'Menginputkan data ke dalam tabel tblcpa
'        Cmdsql = "insert into tblcpa(vcustid,nttlpayment,nbalance,nprincipal,nperiod,vjust,dpropsal)"
'        Cmdsql = Cmdsql + " values ('"
'        Cmdsql = Cmdsql + IIf(IsNull(mobj(0).Value), "", mobj(0).Value) + "','"
'        Cmdsql = Cmdsql + IIf(IsNull(mobj(1).Value), "", CStr(mobj(1).Value)) + "','"
'        Cmdsql = Cmdsql + IIf(IsNull(mobj(2).Value), "", CStr(mobj(2).Value)) + "','"
'        Cmdsql = Cmdsql + IIf(IsNull(mobj(3).Value), "", CStr(mobj(3).Value)) + "','"
'        Cmdsql = Cmdsql + IIf(IsNull(mobj(4).Value), "", CStr(mobj(4).Value)) + "','"
'        Cmdsql = Cmdsql + IIf(IsNull(mobj(5).Value), "", CStr(mobj(5).Value)) + "','"
'        Cmdsql = Cmdsql + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "')"
'        M_OBJCONN.Execute Cmdsql
'
'
'        'Ambil data payment terakhir dulu
'        Cmdsql = "select date(paydate)-3 as tglbayar,* from tbllunas where custid ='"
'        Cmdsql = Cmdsql + Trim(mobj(0).Value) + "' order by paydate desc limit 1"
'        Set M_Objrs = New ADODB.Recordset
'        M_Objrs.CursorLocation = adUseClient
'        M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        'Jika Data Payment Di temukan maka status Account ubah ke PTP-PO, Jika tidak maka PTP-NE
'        If M_Objrs.RecordCount > 0 Then
'            'Status Account PTP-PO
'            Vrdate = DateAdd("m", 1, Format(M_Objrs("tglbayar"), "yyyy-mm-dd"))
'            Cmdsql = "update mgm set amountptp='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', ttlptp='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', f_cek_new='PTP-PO',"
'            Cmdsql = Cmdsql + "dateptp='"
'            Cmdsql = Cmdsql + Format(Vrdate, "yyyy-mm-dd") + "', tenor='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(4).Value), "0", Round(mobj(4).Value))) + "', kethslkerja_new='PTP-POP',  kethslkerjadesc_new='PTP-POP' "
'            Cmdsql = Cmdsql + " where custid='"
'            Cmdsql = Cmdsql + mobj(0).Value + "'"
'            M_OBJCONN.Execute Cmdsql
'        Else
'            'Status Account PTP-NE
'            Vrdate = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")
'            Cmdsql = "update mgm set amountnew='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', "
'            Cmdsql = Cmdsql + " f_cek_new='PTP-NE',"
'            Cmdsql = Cmdsql + "tglptpnew='"
'            Cmdsql = Cmdsql + Format(Vrdate, "yyyy-mm-dd") + "', tenor='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(4).Value), "0", Round(mobj(4).Value))) + "', dateptp='"
'            Cmdsql = Cmdsql + Format(Vrdate, "yyyy-mm-dd") + "',amountptp='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', ttlptp='"
'            Cmdsql = Cmdsql + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + "', kethslkerja_new='PTP-NEW', kethslkerjadesc_new='PTP-NEW' "
'            Cmdsql = Cmdsql + " where custid='"
'            Cmdsql = Cmdsql + mobj(0).Value + "'"
'            M_OBJCONN.Execute Cmdsql
'        End If
'
'        'Inputkan data ke tabel negoptp dan reserve
'        'Jika tenor = 0 atau 1
'        If Round(Val(mobj(4).Value)) = 0 Or Round(Val(mobj(4).Value)) = 1 Then
'            '@@ 12 Januari 2012, hapus data di tblnegoptp
'            Cmdsql = "delete from tblnegoptp where custid='"
'            Cmdsql = Cmdsql + Trim(mobj(0).Value) + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            'Inputkan ke tblnegoptp
'            Cmdsql = "INSERT INTO TblNegoPTP "
'            Cmdsql = Cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
'            Cmdsql = Cmdsql + "VALUES "
'            Cmdsql = Cmdsql + "('" + mobj(0).Value + "', "
'            Cmdsql = Cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
'            Cmdsql = Cmdsql + "" + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + " , "
'            Cmdsql = Cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
'            Cmdsql = Cmdsql + "'IPO')"
'            M_OBJCONN.Execute Cmdsql
'            ' isi ke tbl log_ptp
'            Cmdsql = "INSERT INTO tblnegoptp_log "
'            Cmdsql = Cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
'            Cmdsql = Cmdsql + "VALUES "
'            Cmdsql = Cmdsql + "('" + mobj(0).Value + "', "
'            Cmdsql = Cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
'            Cmdsql = Cmdsql + "" + CStr(IIf(IsNull(mobj(1).Value), "0", mobj(1).Value)) + " , "
'            Cmdsql = Cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
'            Cmdsql = Cmdsql + "'" + MDIForm1.Text1.Text + "','P')"
'            M_OBJCONN.Execute Cmdsql
'        Else
'            'Jika tenor lebih besar dari 1
'
'            '@@ 12 Januari 2012, hapus data di tblnegoptp
'            Cmdsql = "delete from tblnegoptp where custid='"
'            Cmdsql = Cmdsql + Trim(mobj(0).Value) + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            '@@ 12 Januari 2012, hapus data di tblreserve
'            Cmdsql = "delete from tblreserve where custid='"
'            Cmdsql = Cmdsql + Trim(mobj(0).Value) + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            JmlPay = Val(mobj(1).Value) / Round(Val(mobj(4).Value))
'
'            Cmdsql = "INSERT INTO TblNegoPTP "
'            Cmdsql = Cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
'            Cmdsql = Cmdsql + "VALUES "
'            Cmdsql = Cmdsql + "('" + mobj(0).Value + "', "
'            Cmdsql = Cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
'            Cmdsql = Cmdsql + "" + CStr(IIf(IsNull(JmlPay), "0", JmlPay)) + " , "
'            Cmdsql = Cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
'            Cmdsql = Cmdsql + "'IPO')"
'            M_OBJCONN.Execute Cmdsql
'            ' isi ke tbl log_ptp
'            Cmdsql = "INSERT INTO tblnegoptp_log "
'            Cmdsql = Cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
'            Cmdsql = Cmdsql + "VALUES "
'            Cmdsql = Cmdsql + "('" + mobj(0).Value + "', "
'            Cmdsql = Cmdsql + "'" + Format(Vrdate, "yyyy-mm-dd") + "', "
'            Cmdsql = Cmdsql + "" + CStr(IIf(IsNull(JmlPay), "0", JmlPay)) + " , "
'            Cmdsql = Cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
'            Cmdsql = Cmdsql + "'" + MDIForm1.Text1.Text + "','P')"
'            M_OBJCONN.Execute Cmdsql
'
'            For W = 1 To Round(Val(mobj(4).Value)) - 1
'                Vrdate = DateAdd("m", 1, Format(Vrdate, "yyyy-mm-dd"))
'                Cmdsql = "INSERT INTO tblreserve "
'                Cmdsql = Cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
'                Cmdsql = Cmdsql + "VALUES "
'                Cmdsql = Cmdsql + "('" + mobj(0).Value + "', "
'                Cmdsql = Cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
'                Cmdsql = Cmdsql + "" + CStr(JmlPay) + " , "
'                Cmdsql = Cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
'                Cmdsql = Cmdsql + "'IPO')"
'                M_OBJCONN.Execute Cmdsql
'
'                Cmdsql = "INSERT INTO TblNegoptp_log "
'                Cmdsql = Cmdsql + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
'                Cmdsql = Cmdsql + "VALUES "
'                Cmdsql = Cmdsql + "('" + mobj(0).Value + "', "
'                Cmdsql = Cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
'                Cmdsql = Cmdsql + "" + CStr(JmlPay) + " , "
'                Cmdsql = Cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
'                Cmdsql = Cmdsql + "'" + MDIForm1.Text1.Text + "','R')"
'                M_OBJCONN.Execute Cmdsql
'            Next W
'        End If
'
'        Set M_Objrs = Nothing
'        mobj.MoveNext
'     Wend
'     MsgBox "Data telah di upload!", vbInformation + vbOKOnly, "Pesan"
'     cmdupload.Enabled = False
'End Sub

