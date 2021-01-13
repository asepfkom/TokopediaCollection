VERSION 5.00
Begin VB.Form formsystemtrainingagents 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   7860
   ClientLeft      =   3180
   ClientTop       =   1230
   ClientWidth     =   13605
   LinkTopic       =   "Form5"
   ScaleHeight     =   7860
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7080
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   7560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   360
      ScaleHeight     =   7035
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   360
      Width           =   12855
   End
End
Attribute VB_Name = "formsystemtrainingagents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ii As Integer
Dim train As String
Dim ID As Integer

Private Sub Command1_Click()
    Dim aaa As String
    Dim bbb As String

    'getpath
    'aaa = "C:\" & train & "\"
    aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & train & "\"
        
    'load
    If ii >= 1 And ii <= 53 Then
        ii = ii - 1
        bbb = aaa & ii & ".jpg"
        
        If CheckPath(bbb) = True Then
            Picture1.ScaleMode = 3
            Picture1.AutoRedraw = True
            Picture1.Picture = LoadPicture(bbb)
            Picture1.PaintPicture Picture1.Picture, _
            0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
            0, 0, Picture1.Picture.Width / 26.46, _
            Picture1.Picture.Height / 26.46
            Picture1.Picture = Picture1.Image
        End If
    End If

End Sub

Private Sub Command2_Click()
    Dim aaa As String
    Dim bbb As String
    
    
    Timer1.Enabled = True
    Command2.Enabled = False

    ii = ii + 1

    'getpath
    'aaa = "C:\" & train & "\"
    aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & train & "\"
        
    'load
    If ii <= 53 Then
        bbb = aaa & ii & ".jpg"
        
        If CheckPath(bbb) = True Then
            Picture1.ScaleMode = 3
            Picture1.AutoRedraw = True
            Picture1.Picture = LoadPicture(bbb)
            Picture1.PaintPicture Picture1.Picture, _
            0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
            0, 0, Picture1.Picture.Width / 26.46, _
            Picture1.Picture.Height / 26.46
            Picture1.Picture = Picture1.Image
        Else
            MsgBox "Training sudah selesai"
                qupd = "update tblsystemtraining_partisipan set f_done = 1 where agent = '" & MDIForm1.Text1.text & "' and ids = " & ID & ""
                M_OBJCONN.execute qupd
                
            If UCase(MDIForm1.Text2.text) <> "TEAMLEADER" Then
                If signtimes = "time2" Then
                    MDIForm1.Timer2.Enabled = True
                ElseIf signtimes = "time7" Then
                    MDIForm1.Timer7.Enabled = True
                End If
            End If
            
            MDIForm1.Timer11.Enabled = True
            Unload Me
        
        End If
    End If
End Sub

Private Sub Form_Load()
    Call isi
End Sub


Private Sub isi()
    Dim aaa As String
    Dim bbb As String
    

    qsel = "select * from ("
    qsel = qsel & " select nama_file as training, jam_awal, jam_akhir, agent, f_done, b.ids  from tblsystemtraining_schedule a inner join"
    qsel = qsel & " tblsystemtraining_partisipan b on a.ids = b.ids inner join"
    qsel = qsel & " tblsystemtraining c on a.idp = c.id"
    qsel = qsel & " ) a where jam_awal < now() and jam_akhir > now() and agent = '" & MDIForm1.Text1.text & "' and coalesce(f_done,0) <> 1 order by jam_awal, agent limit 1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qsel, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        train = rs!training
        ID = rs!ids
    
'        aaa = "C:\" & train & "\"
        'train = "TRAINING 1"
        aaa = "\\192.168.10.94\pubcard\SYSTEM TRAINING\" & train & "\"
    
        ii = 0
        
        If ii <= 53 Then
            bbb = aaa & ii & ".jpg"
            
            If CheckPath(bbb) = True Then
                Picture1.ScaleMode = 3
                Picture1.AutoRedraw = True
                Picture1.Picture = LoadPicture(bbb)
                Picture1.PaintPicture Picture1.Picture, _
                0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
                0, 0, Picture1.Picture.Width / 26.46, _
                Picture1.Picture.Height / 26.46
                Picture1.Picture = Picture1.Image
            End If
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Command2.Enabled = True
    Timer1.Enabled = False
End Sub
