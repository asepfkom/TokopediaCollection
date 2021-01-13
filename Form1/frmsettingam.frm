VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmsettingam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting AM"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Group View"
      Height          =   2415
      Left            =   0
      TabIndex        =   19
      Top             =   2280
      Width           =   7335
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "Choose TL"
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   7335
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   18
         Top             =   1560
         Width           =   7215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL12"
         Height          =   255
         Index           =   11
         Left            =   4320
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL11"
         Height          =   255
         Index           =   10
         Left            =   4320
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL10"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL9"
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL8"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL7"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL6"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL5"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL4"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL3"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TL1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "DEL"
         Height          =   255
         Left            =   6480
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3720
         TabIndex        =   22
         Top             =   200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SET"
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Text            =   "AM"
         Top             =   200
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   165
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AM Caption"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose AM"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AM Caption"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmsettingam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    uncek
    cekcheck
    klikcombo
End Sub

Private Sub Command1_Click()
    qd = "delete from tblsettingam where am = '" & Combo1.text & "'"
    M_OBJCONN.Execute qd
    MsgBox "Deleted"
    uncek
    cekcheck
    isilist
End Sub

Private Sub uncek()
    For i = 0 To 11
        Check1(i).Value = 0
    Next i
End Sub

Private Sub cekcheck()
    qs = "select *,case when length(tl) = 3 then right(tl,1) when length(tl) = 4 then right(tl,2) end::int urut  from tblsettingam"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Check1(rs!urut - 1).Enabled = False
            rs.MoveNext
        Next i
    End If
End Sub

Private Sub klikcombo()
    qs = "select *,case when length(tl) = 3 then right(tl,1) when length(tl) = 4 then right(tl,2) end::int urut  from tblsettingam where am = '" & Combo1.text & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Check1(rs!urut - 1).Enabled = True
            Check1(rs!urut - 1).Value = 1
            rs.MoveNext
        Next i
    End If
    
End Sub

Private Sub comboam()
    qs = " select userid," & vbCrLf
    qs = qs + " case when length(userid) = 3 then right(userid,1) " & vbCrLf
    qs = qs + " when length(userid) = 4 then right(userid,2) end::int urut " & vbCrLf
    qs = qs + " from usertbl where userid ilike 'TL%' and aktif = 0 order by 2 "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    For i = 1 To rs.RecordCount
        Combo1.AddItem rs!Userid
        rs.MoveNext
    Next i
End Sub

Private Sub Command2_Click()
    If Combo1.text = "" Then
        MsgBox "Choose AM"
        Exit Sub
    End If
    If Text2.text = "" Then
        MsgBox "Please create AM Caption"
        Exit Sub
    End If
    
    a = 0
    For i = 0 To 11
        If Check1(i).Value = 1 Then
            a = a + 1
        End If
    Next i
    
    If a = 0 Then
        MsgBox "Please Check TeamLeader"
        Exit Sub
    End If
    
    For i = 0 To 11
        If Check1(i).Value = 1 Then
            tlc = "TL" & (i + 1)
            qi = "insert into tblsettingam (am,amcaption,tl) values ('" & Combo1.text & "','AM-" & Text2.text & "','" & tlc & "')"
            M_OBJCONN.Execute qi
        End If
    Next i
    MsgBox "Created"
    isilist
End Sub

Private Sub Form_Load()
    Call comboam
    Call isilist
    Call cekcheck
End Sub

Private Sub isilist()
    ListView1.ListItems.clear
    ListView1.ColumnHeaders.clear
 
    ListView1.ColumnHeaders.ADD 1, , "Id", 0 * 120
    ListView1.ColumnHeaders.ADD 2, , "AM", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "AM CAPTION", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "TeamLeader", 20 * 120 '
    
    qs = "select * from tblsettingam"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Set listItem = ListView1.ListItems.ADD(, , rs("id"))
             listItem.SubItems(1) = rs("am")
             listItem.SubItems(2) = IIf(IsNull(rs("amcaption")), "", rs("amcaption"))
             listItem.SubItems(3) = IIf(IsNull(rs("TL")), "", rs("TL"))
             rs.MoveNext
        Next i
    End If
End Sub
