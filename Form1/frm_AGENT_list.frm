VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_AGENT_LIST 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "H&st"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1820
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5505
      Left            =   -15
      TabIndex        =   1
      Top             =   -15
      Width           =   9660
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Set &AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1425
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1020
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   630
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4755
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   8387
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   8640
         TabIndex        =   7
         Top             =   2640
         Width           =   615
      End
   End
End
Attribute VB_Name = "FRM_AGENT_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "User Id", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "Nama Agent", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "Kode TeamLeader", 15 * 120
    ListView1.ColumnHeaders.ADD 4, , "Nama TeamLeader", 20 * 120 '
    ListView1.ColumnHeaders.ADD 5, , "Unit", 15 * 120
    ListView1.ColumnHeaders.ADD 6, , "Team", 15 * 120
    ListView1.ColumnHeaders.ADD 7, , "Level", 14 * 120
    ListView1.ColumnHeaders.ADD 8, , "Status Agent", 14 * 120
    ListView1.ColumnHeaders.ADD 9, , "AM", 14 * 120
    
End Sub

Private Sub Form_Load()
    If MDIForm1.Text2.text <> "Supervisor" Then
        Command1(5).Visible = False
    End If
    
    Dim M_Objrs As ADODB.Recordset
    Dim M_DATA As New CLSSPV_AGENT
    Dim listItem As listItem
    Dim cek As Integer
    Dim M_WHERE As String
    Call header
    If UCase(MDIForm1.Text2) = "TEAMLEADER" Then
    M_WHERE = "TEAM='" + MDIForm1.Text1 + "'"
    ElseIf UCase(MDIForm1.Text2) = "SUPERVISOR" Or UCase(MDIForm1.Text2) = "ADMIN" Then
    M_WHERE = ""
    End If
    Set M_Objrs = M_DATA.QUERY_AGENT(M_OBJCONN, M_WHERE)
    While Not M_Objrs.EOF
         Set listItem = ListView1.ListItems.ADD(, , M_Objrs("userid"))
             listItem.SubItems(1) = M_Objrs("AGENT")
             listItem.SubItems(2) = IIf(IsNull(M_Objrs("SPVCODE")), "", M_Objrs("SPVCODE"))
             listItem.SubItems(3) = IIf(IsNull(M_Objrs("teamleader")), "", M_Objrs("teamleader"))
             listItem.SubItems(4) = IIf(IsNull(M_Objrs("UNIT")), "", M_Objrs("UNIT"))
             listItem.SubItems(5) = IIf(IsNull(M_Objrs("TEAM")), "", M_Objrs("TEAM"))
             listItem.SubItems(6) = IIf(IsNull(M_Objrs("LVL")), "", M_Objrs("LVL"))
             cek = IIf(IsNull(M_Objrs("AKTIF")), 0, M_Objrs("AKTIF"))
             If cek = 0 Then
                listItem.SubItems(7) = "Works"
             Else
                listItem.SubItems(7) = "Resign"
             End If
             listItem.SubItems(8) = IIf(IsNull(M_Objrs("AM")), "", M_Objrs("AM"))
        M_Objrs.MoveNext
    Wend
        M_Objrs.Close
        Set M_Objrs = Nothing

End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim STATUS As String
Dim gaji As Currency
Dim gaji1 As String
Dim listItem As listItem
Dim M_DATA As New CLSSPV_AGENT
Dim sId As Integer
Dim listdo As String


Select Case Index
    Case 5
        frmsettingam.Show vbModal
    Case 0
            With frm_AGENT
                .Caption = "Tambah Data Agent"
                .Option1(0).Value = True
                .TDBNumber1.Value = 0
                .Show vbModal
                If .ok Then
                If .Option1(0).Value Then
                    STATUS = "0"
                Else
                    STATUS = "1"
                End If
                    Dim M_Objrs_x As ADODB.Recordset
                    Set M_Objrs_x = New ADODB.Recordset
                    M_Objrs_x.Open "SELECT max(id) as id_x  FROM usertbl", M_OBJCONN, adOpenStatic, adLockOptimistic
                    If M_Objrs_x.RecordCount > 0 Then
                        sId = IIf(IsNull(M_Objrs_x!id_x), 0, M_Objrs_x!id_x) + 1
                    End If
                    M_DATA.ADD_AGENT M_OBJCONN, .Text1.text, .Text2.text, .Combo1(0).text, CStr(.TDBNumber1.Value), .Text4.text, STATUS, .Combo2.text, .Text5.text, .Text3.text, sId, Label1.Caption
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        Set listItem = ListView1.ListItems.ADD(, , .Text1.text)
                            listItem.SubItems(1) = .Text2.text
                            listItem.SubItems(2) = .Combo1(0).text
                            listItem.SubItems(3) = .Combo1(1).text
                            listItem.SubItems(4) = .Text5.text
                            listItem.SubItems(5) = .Text4.text
                            listItem.SubItems(6) = .Combo2.text
                            If .Option1(0).Value Then
                                listItem.SubItems(7) = "Works"
                            Else
                                listItem.SubItems(7) = "Resign"
                            End If
                            listItem.SubItems(8) = .Text3.text
                            '==asep=='
                            'listItem.SubItems(9) = .Combo3.text
                            '========='
                    On Error GoTo 0
                    End If
                End If
                Unload frm_AGENT
            End With
            'listdo = "ADD"
            
        Exit Sub
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
            With frm_AGENT
                .Caption = "Ubah Data Agent"
                .Text1.text = ListView1.SelectedItem.text
                .Text2.text = ListView1.SelectedItem.SubItems(1)
                .Text6.text = ListView1.SelectedItem.SubItems(1)
                .Combo1(0).text = ListView1.SelectedItem.SubItems(2)
                .Combo1(1).text = ListView1.SelectedItem.SubItems(3)
                .Text5.text = ListView1.SelectedItem.SubItems(4)
                .Text4.text = ListView1.SelectedItem.SubItems(5)
                .Combo2.text = ListView1.SelectedItem.SubItems(6)
                
                If ListView1.SelectedItem.SubItems(7) = "Works" Then
                    .Option1(0).Value = True
                Else
                    .Option1(1).Value = True
                End If
                .Text3.text = ListView1.SelectedItem.SubItems(8)
                '.Combo3.text = ListView1.SelectedItem.SubItems(9)
                .Text1.Locked = True
                .Text1.TabStop = False
                .Text1.BackColor = &H8000000F
                .Text1.Appearance = 0
                .Show vbModal
                If .ok Then
                    If .Option1(0).Value Then
                        STATUS = "0"
                    Else
                        STATUS = "1"
                    End If
                
                    M_DATA.UPDATE_AGENT M_OBJCONN, .Text1.text, .Text2.text, .Combo1(0).text, CStr(.TDBNumber1.Value), .Text4.text, STATUS, .Combo2.text, .Text5.text, .Text3.text
                    
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        ListView1.SelectedItem.SubItems(1) = .Text2.text
                        ListView1.SelectedItem.SubItems(2) = .Combo1(0).text
                        ListView1.SelectedItem.SubItems(3) = .Combo1(1).text
                        ListView1.SelectedItem.SubItems(4) = .Text5.text
                        ListView1.SelectedItem.SubItems(5) = .Text4.text
                        ListView1.SelectedItem.SubItems(6) = .Combo2.text
                    If .Option1(0).Value Then
                        ListView1.SelectedItem.SubItems(7) = "Works"
                    Else
                        ListView1.SelectedItem.SubItems(7) = "Resign"
                    End If
                        ListView1.SelectedItem.SubItems(8) = .Text3.text
                        ListView1.SelectedItem.SubItems(9) = .Combo3.text
                    On Error GoTo 0
                    End If
                    
                End If
                Unload frm_AGENT
            End With
        Exit Sub
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
            M_DATA.DELETE_AGENT M_OBJCONN, ListView1.SelectedItem.text
            If M_DATA.ADD_OK Then
                ListView1.ListItems.Remove ListView1.SelectedItem.Index
            End If
        End If
        Exit Sub
    Case 3
        Unload Me
        Exit Sub
    Case 4
       Formhsttelecolection.Show vbModal
End Select
add_error:
End Sub

Private Sub ListView1_DblClick()
    Call Command1_Click(1)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click(1)
    End If
End Sub

'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   ListView1.SortKey = ColumnHeader.Index - 1
'   ListView1.Sorted = True
'End Sub
