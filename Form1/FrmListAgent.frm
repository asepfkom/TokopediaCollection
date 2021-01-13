VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListAgent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Agent"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6480
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   435
      Left            =   4980
      TabIndex        =   8
      Top             =   6300
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   3600
      TabIndex        =   7
      Top             =   6300
      Width           =   1335
   End
   Begin VB.TextBox TxtJmlhAgent 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   6
      Text            =   "0"
      Top             =   5940
      Width           =   915
   End
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "UnCek All"
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton CmdCekAlll 
      Caption         =   "Cek All"
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
   Begin VB.ComboBox CmbFilterAgent 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin MSComctlLib.ListView LvAgent 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah Agent:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Filter Agent:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "FrmListAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HeaderAgent()
    LvAgent.ColumnHeaders.ADD 1, , "Agent", 2000
    LvAgent.ColumnHeaders.ADD 2, , "Nama Agent", 3000
    LvAgent.ColumnHeaders.ADD 3, , "TL", 3000
End Sub

Private Sub isicombo()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbFilterAgent.clear
    CmbFilterAgent.AddItem "ALL"
    
    If Left(MDIForm1.Text2.text, 2) = "AM" Then
        cmdsql = "select * from usertbl where usertype='6' and spvcode is not null and userid in (select tl from tblsettingam where am = '" & MDIForm1.Text1.text & "') order by spvcode asc "
    Else
        cmdsql = "select * from usertbl where usertype='6' and spvcode is not null order by spvcode asc "
    End If
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbFilterAgent.AddItem M_Objrs("spvcode")
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub



Private Sub CmbFilterAgent_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    If CmbFilterAgent.text = "ALL" Then
        If Left(MDIForm1.Text2.text, 2) = "AM" Then
            cmdsql = "select * from usertbl where usertype in ('1','6') "
            cmdsql = cmdsql & " and userid not in ('COMPLAIN','LUNAS','AKSESALL','#KOSONG#','CLAIM')  "
            cmdsql = cmdsql & " and userid not in (select userid from usertbl where spvcode='RESERVED') "
            cmdsql = cmdsql & " and userid in (select userid from usertbl where userid ilike 'D%' and team in (select tl from tblsettingam where am ='" & MDIForm1.Text1.text & "')) "
            cmdsql = cmdsql & " order by spvcode,userid asc "
        Else
            cmdsql = "select * from usertbl where usertype in ('1','6') "
            cmdsql = cmdsql & " and userid not in ('COMPLAIN','LUNAS','AKSESALL','#KOSONG#','CLAIM')  "
            cmdsql = cmdsql & " and userid not in (select userid from usertbl where spvcode='RESERVED') "
            cmdsql = cmdsql & " order by spvcode,userid asc "
        End If
    Else
        cmdsql = "select * from usertbl where usertype in ('1','6') and spvcode='"
        cmdsql = cmdsql + CmbFilterAgent.text + "' and userid not in ('COMPLAIN','LUNAS','AKSESALL','#KOSONG#') "
        cmdsql = cmdsql & " and userid not in (select userid from usertbl where spvcode='RESERVED') "
        cmdsql = cmdsql + " order by userid asc "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAgent.ListItems.clear
    TxtJmlhAgent.text = M_Objrs.RecordCount
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            Set listItem = LvAgent.ListItems.ADD(, , M_Objrs("userid"))
                listItem.SubItems(1) = M_Objrs("agent")
                listItem.SubItems(2) = cnull(M_Objrs("spvcode"))
            M_Objrs.MoveNext
        Wend
    End If
    
    Set M_Objrs = Nothing
End Sub
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdCekAlll_Click()
    Dim i As Integer
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For i = 1 To LvAgent.ListItems.Count
        LvAgent.ListItems(i).Checked = True
    Next i
End Sub

Private Sub CmdOK_Click()
    Dim a, NamaAgent As String
    Dim w, K, S As Integer
    
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan menambahkan agent yang di ceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    S = 0
    For w = 1 To LvAgent.ListItems.Count
       If LvAgent.ListItems(w).Checked = True Then
        S = S + 1
       End If
    Next w
    
    If S = 0 Then
        MsgBox "Anda belum memilih agent yang akan ditambahkan!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    NamaAgent = ""
    FrmDistribusiAcc.TxtAgent.text = ""
    For K = 1 To LvAgent.ListItems.Count
        If LvAgent.ListItems(K).Checked = True Then
            If NamaAgent = "" Then
                NamaAgent = "'" & LvAgent.ListItems(K).text & "'"
            Else
                NamaAgent = NamaAgent & ",'" & LvAgent.ListItems(K).text & "'"
            End If
        End If
    Next K
    
    FrmDistribusiAcc.TxtAgent.text = NamaAgent
    Unload Me
    
End Sub

Private Sub CmdUnCekAll_Click()
    Dim i As Integer
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For i = 1 To LvAgent.ListItems.Count
        LvAgent.ListItems(i).Checked = False
    Next i
End Sub

Private Sub Form_Load()
    Call HeaderAgent
    Call isicombo
End Sub

Private Sub LvAgent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAgent.SortKey = ColumnHeader.Index - 1
    LvAgent.Sorted = True
End Sub
