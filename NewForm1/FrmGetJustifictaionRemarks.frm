VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGetJustifictaionRemarks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Get Justification From Remarks"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   11040
      TabIndex        =   2
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   9900
      TabIndex        =   1
      Top             =   4140
      Width           =   1095
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   4080
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7197
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   10147522
      BorderStyle     =   1
      Appearance      =   0
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
Attribute VB_Name = "FrmGetJustifictaionRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub HEADER_HISTORY()
    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "History", 70 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "User Log", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 4, , "Handle By", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 5, , "Sts Account", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 6, , "Sts Call", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 7, , "Sts Telp With", 25 * TXT
    ListView1(1).ColumnHeaders.ADD 8, , "Id", 25 * TXT
End Sub

Private Sub AmbilHistory()
    Dim CMDSQL As String
    Dim M_OBJRS As ADODB.Recordset
    Dim listitem As listitem
    
    CMDSQL = "select * from mgm_hst where custid='"
    CMDSQL = CMDSQL + Trim(frmcpanew.txtcardno.Text) + "' order by tgl desc"
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_OBJRS.RecordCount > 0 Then
        While Not M_OBJRS.EOF
            Set listitem = ListView1(1).ListItems.ADD(, , Format(IIf(IsNull(M_OBJRS("TGL")), "", M_OBJRS!TGL), "mm-dd-yyyy hh:mm:ss"))
                listitem.SubItems(1) = IIf(IsNull(M_OBJRS("HST")), "", M_OBJRS("HST"))
                listitem.SubItems(2) = IIf(IsNull(M_OBJRS("user_log")), "", M_OBJRS("user_log"))
                listitem.SubItems(3) = IIf(IsNull(M_OBJRS("AGENT")), "", M_OBJRS("AGENT"))
                listitem.SubItems(4) = IIf(IsNull(M_OBJRS("KodeDs")), "", M_OBJRS("KodeDs"))
                listitem.SubItems(5) = IIf(IsNull(M_OBJRS("statuscall")), "", M_OBJRS("statuscall"))
                listitem.SubItems(6) = IIf(IsNull(M_OBJRS("ststelpwith")), "", M_OBJRS("ststelpwith"))
                listitem.SubItems(7) = IIf(IsNull(M_OBJRS("id")), "", M_OBJRS("id"))
            M_OBJRS.MoveNext
        Wend
    End If
    Set M_OBJRS = Nothing
End Sub

Private Sub CmdOk_Click()
     If ListView1(1).ListItems.Count = 0 Then
        MsgBox "Data Remarks belum tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    If frmcpanew.txtjust.Text = "" Then
        frmcpanew.txtjust.Text = ListView1(1).SelectedItem.SubItems(1)
    Else
        frmcpanew.txtjust.Text = frmcpanew.txtjust.Text + vbCrLf + vbCrLf + "-" + ListView1(1).SelectedItem.SubItems(1)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call HEADER_HISTORY
    Call AmbilHistory
End Sub



Private Sub ListView1_DblClick(Index As Integer)
    If ListView1(1).ListItems.Count = 0 Then
        MsgBox "Data Remarks belum tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    If frmcpanew.txtjust.Text = "" Then
        frmcpanew.txtjust.Text = ListView1(1).SelectedItem.SubItems(1)
    Else
        frmcpanew.txtjust.Text = frmcpanew.txtjust.Text + vbCrLf + vbCrLf + "-" + ListView1(1).SelectedItem.SubItems(1)
    End If
    
    
    Unload Me
End Sub
