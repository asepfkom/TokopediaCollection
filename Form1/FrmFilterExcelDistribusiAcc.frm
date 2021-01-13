VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmFilterExcelDistribusiAcc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter Account Berdasarkan File Excel - Manage Distribusi Account"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9045
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check_decease 
      Caption         =   "Include Account Decease [ 835 ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtlocation 
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6300
      TabIndex        =   4
      Top             =   1320
      Width           =   1275
   End
   Begin VB.ComboBox cbosheet 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   4425
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Label LblLokasiFile 
      Caption         =   "(Pilih lokasi file excel. Format .xls)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2940
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pilih Sheet:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama File Excel(.xls):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "FrmFilterExcelDistribusiAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_XLSCONN As New ADODB.Connection
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdBrowse_Click()
    With CommonDialog1
            .DialogTitle = "Pilih file excel"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
        End With
        
        txtlocation.Text = CommonDialog1.FileName
        LblLokasiFile.Caption = CommonDialog1.FileName
        
        If CommonDialog1.FileName = "" Then Exit Sub
        If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
                M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtlocation.Text & ";Extended Properties=Excel 8.0;"
        Set rsTemp = M_XLSCONN.OpenSchema(adSchemaTables)
        cbosheet.CLEAR
        If rsTemp.EOF And rsTemp.BOF Then Exit Sub
        While Not rsTemp.EOF
            cbosheet.AddItem IIf(IsNull(rsTemp!table_name), "", rsTemp!table_name)
            rsTemp.MoveNext
        Wend
        Set rsTemp = Nothing
End Sub

Private Sub CmdOK_Click()
    Dim OBJRECORD As New ADODB.Recordset
    Dim ssql As String
    Dim CustId As String
    Dim cmdsql As String
    Dim listItem As listItem
    Dim M_Objrs As ADODB.Recordset
    
    On Error GoTo SALAH
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    ssql = "SELECT * FROM [" & cbosheet.Text & "] where [custid] is not null"
    DoEvents
    OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        
    CustId = ""
    
    If OBJRECORD.RecordCount > 0 Then
        While Not OBJRECORD.EOF
            If CustId = "" Then
                CustId = "'" & IIf(IsNull(OBJRECORD(0)), "", OBJRECORD(0)) & "'"
            Else
                CustId = CustId & ",'" & IIf(IsNull(OBJRECORD(0)), "", OBJRECORD(0)) & "'"
            End If
            OBJRECORD.MoveNext
        Wend
    Else
        MsgBox "Data di file excel anda kosong! Cek file excel anda atau mungkin anda salah memilih sheet!", vbOKOnly + vbExclamation, "Peringatan"
        Set OBJRECORD = Nothing
        Exit Sub
    End If
    Set OBJRECORD = Nothing
    
    cmdsql = "SELECT * FROM mgm WHERE custid in (" & CustId & ") "
    cmdsql = cmdsql & " AND agent NOT IN ('LUNAS','COMPLAIN','CLAIM','AKSESALL','REVIEW','REVIEW1','REVIEW2','REVIEW3','REVIEW4','REVIEW5','REVIEW6','REVIEW7','REVIEW8','REVIEW9','REVIEW10') AND coalesce(agent,'')<>'' "
    cmdsql = cmdsql & " AND custid NOT IN (select distinct custid from tblsendptp ) "
    ' TAMBAHAN AGAR CLASS 835 TIDAK KENA AKSES ALL
    ' DIGANTI 23 FEB 2015
    If Check_decease.Value = 0 Then
        cmdsql = cmdsql & " AND coalesce(cust_class,'')<>'835' "
    End If
    ' -------------------------------------------
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        With FrmDistribusiAcc
            .LvAcc.ListItems.CLEAR
            .PB1.Max = M_Objrs.RecordCount
            While Not M_Objrs.EOF
                .PB1.Value = M_Objrs.Bookmark
                Set listItem = .LvAcc.ListItems.ADD(, , M_Objrs("custid"))
                listItem.SubItems(1) = M_Objrs("name")
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
                listItem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
                listItem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
                listItem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
                If UCase(M_Objrs("agent")) = "AKSESALL" Then
                    listItem.ForeColor = vbRed
                    listItem.ListSubItems(1).ForeColor = vbRed
                    listItem.ListSubItems(2).ForeColor = vbRed
                    listItem.ListSubItems(3).ForeColor = vbRed
                    listItem.ListSubItems(4).ForeColor = vbRed
                    listItem.ListSubItems(5).ForeColor = vbRed
                    listItem.ListSubItems(6).ForeColor = vbRed
                End If
            
                If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                    listItem.ForeColor = vbBlue
                    listItem.ListSubItems(1).ForeColor = vbBlue
                    listItem.ListSubItems(2).ForeColor = vbBlue
                    listItem.ListSubItems(3).ForeColor = vbBlue
                    listItem.ListSubItems(4).ForeColor = vbBlue
                    listItem.ListSubItems(5).ForeColor = vbBlue
                    listItem.ListSubItems(6).ForeColor = vbBlue
                End If
                M_Objrs.MoveNext
            Wend
        End With
        MsgBox "Data berhasil di load!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs = Nothing
        Unload Me
    Else
        MsgBox "Data tidak ditemukan!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    Exit Sub
SALAH:
    MsgBox "Maaf ada kesalahan! " & err.Description, vbOKOnly + vbExclamation, "Peringatan"
    
End Sub

