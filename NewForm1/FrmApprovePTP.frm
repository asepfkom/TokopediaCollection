VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmApprovePTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Approve PTP"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "&UnCek All"
      Height          =   435
      Left            =   10080
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   435
      Left            =   10080
      TabIndex        =   5
      Top             =   1260
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reject PTP"
      Height          =   435
      Left            =   10080
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdApprovePTP 
      Caption         =   "&Approve PTP"
      Height          =   435
      Left            =   10080
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TxtJmlhData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "0"
      Top             =   4920
      Width           =   1155
   End
   Begin MSComctlLib.ListView LvApprovePTP 
      Height          =   4680
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   8255
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Label Label1 
      Caption         =   "Jumlah Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "FrmApprovePTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApprovePTP_Click()
    Dim CMDSQL As String
    Dim Strsql As String
    Dim M_OBJRS As ADODB.Recordset
    Dim W As Integer
    
    If LvApprovePTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvApprovePTP.ListItems.Count
        If LvApprovePTP.ListItems(W).Checked = True Then
            '1.-- BUAT CPAnya Dulu
            CMDSQL = "insert into tblcpa(vcustid,nttlpayment,nbalance,nprincipal,nperiod,vjust,dpropsal)"
            CMDSQL = CMDSQL + " values ('"
            CMDSQL = CMDSQL + IIf(IsNull(LvApprovePTP.ListItems(W).SubItems(2)), "", CStr(LvApprovePTP.ListItems(W).SubItems(2))) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(LvApprovePTP.ListItems(W).SubItems(4)), "", CStr(LvApprovePTP.ListItems(W).SubItems(4))) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(LvApprovePTP.ListItems(W).SubItems(9)), "", CStr(LvApprovePTP.ListItems(W).SubItems(9))) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(LvApprovePTP.ListItems(W).SubItems(10)), "", CStr(LvApprovePTP.ListItems(W).SubItems(10))) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(LvApprovePTP.ListItems(W).SubItems(5)), "", CStr(LvApprovePTP.ListItems(W).SubItems(5))) + "','"
            CMDSQL = CMDSQL + "Create Otomatic By System (Sending Request PTP)','"
            CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "')"
            M_OBJCONN.Execute CMDSQL
       End If
            
            '2.-- Cek Apakah memiliki Payment
    Next
    
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvApprovePTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvApprovePTP.ListItems.Count
        LvApprovePTP.ListItems(W).Checked = True
    Next W
    
End Sub

Private Sub CmdUncekall_Click()
    Dim W As Integer
    
    If LvApprovePTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvApprovePTP.ListItems.Count
        LvApprovePTP.ListItems(W).Checked = False
    Next W
End Sub

Private Sub header()
    With LvApprovePTP.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Tgl.Payment Effective", 2500
        .ADD 5, , "Total Amount", 1000
        .ADD 6, , "Tenor", 700
        .ADD 7, , "Pembayaran Via", 2000
        .ADD 8, , "Tgl.Tagih", 1500
        .ADD 9, , "Status", 1000
        .ADD 10, , "Balance", 1000
        .ADD 11, , "Principal", 1000
    End With
End Sub

Private Sub Form_Load()
    Call header
    Call IsiData
End Sub

Private Sub IsiData()
    Dim CMDSQL As String
    Dim M_OBJRS As ADODB.Recordset
    Dim listitem As listitem
    
    CMDSQL = "select * from tblsendptp where agent in "
    'Jika yang akses TeamLeader, Maka data yang ditampilkan anaknya saja
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + " (select userid from usertbl where usertype='1' and "
        CMDSQL = CMDSQL + " team='"
        CMDSQL = CMDSQL + MDIForm1.Text1.Text + "') "
    End If
    'Jika yang akses Supervisor/Admin/Administrator
    If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or _
       UCase(MDIForm1.Text2.Text) = "ADMIN" Or _
       UCase(MDIForm1.Text2.Text) = "ADMINISTRATOR" Then
        
        CMDSQL = CMDSQL + " (select userid from usertbl where usertype='1')"
    End If
    CMDSQL = CMDSQL + " and status='0'"
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvApprovePTP.ListItems.CLEAR
    TxtJumlah.Text = M_OBJRS.RecordCount
    If M_OBJRS.RecordCount > 0 Then
        Dim STATUS As String
        While Not M_OBJRS.EOF
            Set listitem = LvApprovePTP.ListItems.ADD(, , M_OBJRS("id"))
                listitem.SubItems(1) = IIf(IsNull(M_OBJRS("jenis_ptp")), "", M_OBJRS("jenis_ptp"))
                listitem.SubItems(2) = IIf(IsNull(M_OBJRS("custid")), "", M_OBJRS("custid"))
                listitem.SubItems(3) = IIf(IsNull(M_OBJRS("date_payment_effective")), "", Format(M_OBJRS("date_payment_effective"), "yyyy-mm-dd"))
                listitem.SubItems(4) = IIf(IsNull(M_OBJRS("total_amount_deal")), "", Format(M_OBJRS("total_amount_deal"), "##,###"))
                listitem.SubItems(5) = IIf(IsNull(M_OBJRS("tenor")), "", Format(M_OBJRS("tenor"), "##,###"))
                listitem.SubItems(6) = IIf(IsNull(M_OBJRS("pembayaran_via")), "", M_OBJRS("pembayaran_via"))
                listitem.SubItems(7) = IIf(IsNull(M_OBJRS("tgl_tagih")), "", Format(M_OBJRS("tgl_tagih"), "yyyy-mm-dd"))
                
                If M_OBJRS("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_OBJRS("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_OBJRS("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                listitem.SubItems(8) = IIf(IsNull(STATUS), "", STATUS)
                listitem.SubItems(9) = IIf(IsNull(M_OBJRS("balance")), "", Format(M_OBJRS("balance"), "##,###"))
                listitem.SubItems(10) = IIf(IsNull(M_OBJRS("principal")), "", Format(M_OBJRS("principal"), "##,###"))
            M_OBJRS.MoveNext
        Wend
    End If
    Set M_OBJRS = Nothing
End Sub
