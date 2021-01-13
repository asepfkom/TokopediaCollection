VERSION 5.00
Begin VB.Form FRMUNTUK 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5715
      Left            =   15
      TabIndex        =   4
      Top             =   -90
      Width           =   5760
      Begin VB.CommandButton CmdServer5 
         Caption         =   "Server 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton CmdServer4 
         Caption         =   "Server 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   4260
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Team"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Group 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Spv"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   4350
         TabIndex        =   6
         Top             =   2340
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Per Team"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   5
         Top             =   2715
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Kel&uar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   1740
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5520
         ItemData        =   "FRMUNTUK.frx":0000
         Left            =   45
         List            =   "FRMUNTUK.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   0
         Top             =   135
         Width           =   4230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ambil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Semua"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   2
         Top             =   615
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRMUNTUK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim agen As String * 10

Private Sub CmdServer4_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim i As Integer
    
    List1.clear
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        ' UPDATE 29 MEI 2013 BY IZUDDIN
        cmdsql = "SELECT TEAM as USERID,SPVNAME as agent FROM SPVTBL WHERE TEAM IN ("
        cmdsql = cmdsql + " SELECT TEAM FROM usertbl WHERE USERID='" & MDIForm1.Text1.text & "')"
    Else
        'CMDSQL = "Select * from usertbl where spvcode ='" + Combo2(0).Text + "'"
        cmdsql = "select usertbl.agent,usertbl.userid from tbl_ip,usertbl where tbl_ip.ip_addr in "
        cmdsql = cmdsql + " (select ip from tbl_ip_icentra where ip_icentra='192.168.10.4') "
        cmdsql = cmdsql + " and usertbl.userid=tbl_ip.agent "
    End If
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    
    For i = 1 To M_Objrs.RecordCount
        agen = M_Objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(M_Objrs("AGENT")), "", M_Objrs("AGENT"))
        M_Objrs.MoveNext
    Next i
    Set M_Objrs = Nothing
End Sub

Private Sub CmdServer5_Click()
        Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim i As Integer
    
    List1.clear
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        ' UPDATE 29 MEI 2013 BY IZUDDIN
        cmdsql = "SELECT TEAM as USERID,SPVNAME as agent FROM SPVTBL WHERE TEAM IN ("
        cmdsql = cmdsql + " SELECT TEAM FROM usertbl WHERE USERID='" & MDIForm1.Text1.text & "')"
    Else
        'CMDSQL = "Select * from usertbl where spvcode ='" + Combo2(0).Text + "'"
        cmdsql = "select usertbl.agent,usertbl.userid from tbl_ip,usertbl where tbl_ip.ip_addr in "
        cmdsql = cmdsql + " (select ip from tbl_ip_icentra where ip_icentra='192.168.10.5') "
        cmdsql = cmdsql + " and usertbl.userid=tbl_ip.agent "
    End If
    
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    
    For i = 1 To M_Objrs.RecordCount
        agen = M_Objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(M_Objrs("AGENT")), "", M_Objrs("AGENT"))
        M_Objrs.MoveNext
    Next i
    Set M_Objrs = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim M_Objrs As ADODB.Recordset
Dim i As Integer
Select Case Index
Case 0
    If FRMSENDMSG.Text1.text = Empty Then
         For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                FRMSENDMSG.Text1.text = FRMSENDMSG.Text1.text & List1.list(i) & ";"
            End If
         Next i
    Else
         For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                FRMSENDMSG.Text1.text = FRMSENDMSG.Text1.text & List1.list(i) & ";"
            End If
         Next i
    End If
    Set M_Objrs = Nothing
    Unload Me
Case 1
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        M_Objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 AND SPVCODE ='" + MDIForm1.Text1.text + "' AND USERTYPE =1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_Objrs.RecordCount
                agen = M_Objrs("USERID")
                FRMSENDMSG.Text1.text = FRMSENDMSG.Text1.text + agen & ";"
                M_Objrs.MoveNext
            Next i
        Set M_Objrs = Nothing
    Else
        M_Objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_Objrs.RecordCount
                agen = M_Objrs("USERID")
                FRMSENDMSG.Text1.text = FRMSENDMSG.Text1.text + agen & ";"
                M_Objrs.MoveNext
            Next i
        Set M_Objrs = Nothing
    End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
Case 2
    Dim cmdsql As String
    List1.clear
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    If Left(MDIForm1.Text2.text, 2) = "AM" And Left(Combo2(0).text, 2) = "AM" Then
        cmdsql = "Select * from usertbl where spvcode in (select 'SPV'|| "
        cmdsql = cmdsql + " case when length(tl)=3 then right(tl,1)"
        cmdsql = cmdsql + " when length(tl)=4 then right(tl,2) end as SPV from tblsettingam where am = '" + MDIForm1.Text1.text + "' ) order by spvcode,userid"
    Else
        cmdsql = "Select * from usertbl where spvcode ='" + Combo2(0).text + "' order by spvcode,userid"
    End If
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    For i = 1 To M_Objrs.RecordCount
        agen = M_Objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(M_Objrs("AGENT")), "", M_Objrs("AGENT"))
        M_Objrs.MoveNext
    Next i
    Set M_Objrs = Nothing
End Select
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim M_Objrs As ADODB.Recordset
Dim i As Integer

    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        M_Objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 AND  USERTYPE =20 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_Objrs.RecordCount
                agen = M_Objrs("USERID")
                FRMSENDMSG.Text1.text = FRMSENDMSG.Text1.text + agen & ";"
                M_Objrs.MoveNext
            Next i
        Set M_Objrs = Nothing
    'Else
     '   m_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
      '      For i = 1 To m_objrs.RecordCount
       '         agen = m_objrs("USERID")
        '        FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
         '       m_objrs.MoveNext
          '  Next i
        'Set m_objrs = Nothing
   ' End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Dim M_Objrs As ADODB.Recordset
Dim i As Integer
Dim ssql As String
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        M_Objrs.Open "SELECT USERID,AGENT FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        For i = 1 To M_Objrs.RecordCount
           agen = IIf(IsNull(M_Objrs("USERID")), "", M_Objrs("USERID"))
           List1.AddItem agen & "!" & IIf(IsNull(M_Objrs("AGENT")), "", M_Objrs("AGENT"))
            M_Objrs.MoveNext
        Next i
    Else
        If UCase(MDIForm1.Text2.text) = "AGENT" Then
            ssql = "SELECT TEAM,SPVNAME FROM SPVTBL WHERE TEAM IN ("
            ssql = ssql + " SELECT TEAM FROM usertbl WHERE USERID='" & MDIForm1.Text1.text & "')"
            'Command1(1).Visible = False
            'M_OBJRS.Open "SELECT SPVCODE,SPVNAME FROM SPVTBL order by SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            M_Objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                For i = 1 To M_Objrs.RecordCount
                   agen = IIf(IsNull(M_Objrs("TEAM")), "", M_Objrs("TEAM"))
                   List1.AddItem agen & "!" & IIf(IsNull(M_Objrs("SPVNAME")), "", M_Objrs("SPVNAME"))
                    M_Objrs.MoveNext
                Next i
                'Command1(1).Visible = False
                Command1(2).Visible = False
                Combo2(0).Visible = False
                Command3.Visible = False
                Group.Visible = False
                Command1(1).Visible = False
        Else
            M_Objrs.Open "SELECT USERID,AGENT FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                For i = 1 To M_Objrs.RecordCount
                    agen = IIf(IsNull(M_Objrs("USERID")), "", M_Objrs("USERID"))
                    List1.AddItem agen & "!" & IIf(IsNull(M_Objrs("AGENT")), "", M_Objrs("AGENT"))
                    M_Objrs.MoveNext
                Next i
        End If
    End If
Set M_Objrs = Nothing
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open "Select * from spvtbl", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_Objrs.EOF
    Combo2(0).AddItem M_Objrs!SPVCODE
    M_Objrs.MoveNext
Wend
If Left(MDIForm1.Text2.text, 2) = "AM" Then
    Combo2(0).AddItem MDIForm1.Text2.text
End If
Set M_Objrs = Nothing
End Sub


Private Sub Group_Click()
Dim M_Objrs As ADODB.Recordset
Dim i As Integer

    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        M_Objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 AND  USERTYPE =6 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_Objrs.RecordCount
                agen = M_Objrs("USERID")
                FRMSENDMSG.Text1.text = FRMSENDMSG.Text1.text + agen & ";"
                M_Objrs.MoveNext
            Next i
        Set M_Objrs = Nothing
    'Else
     '   m_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
      '      For i = 1 To m_objrs.RecordCount
       '         agen = m_objrs("USERID")
        '        FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
         '       m_objrs.MoveNext
          '  Next i
        'Set m_objrs = Nothing
   ' End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
End Sub


