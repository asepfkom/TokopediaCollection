VERSION 5.00
Begin VB.Form FrmSendSMS 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4905
   ClientLeft      =   7155
   ClientTop       =   3705
   ClientWidth     =   4965
   LinkTopic       =   "Form2"
   ScaleHeight     =   4905
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   2700
      TabIndex        =   18
      Top             =   4980
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ComboBox CmbSubOption 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2100
      Width           =   3405
   End
   Begin VB.ComboBox CmbOption 
      Height          =   315
      ItemData        =   "FrmSendSMS.frx":0000
      Left            =   1140
      List            =   "FrmSendSMS.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1740
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   1200
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1250
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   "Sub option:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Option:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1710
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah :"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Mobile No :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Text :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   4875
      Left            =   0
      Top             =   0
      Width           =   4890
   End
   Begin VB.Label Label3 
      Caption         =   "Agent :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Custid :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Nama :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmSendSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public awal As String
Public btsakhir As Integer
Dim d As Integer
Dim awalk As Integer
Dim akhirk As Integer
Dim AvgMarks(50, 50) As Double
Dim rowaray As Integer

Private Sub CmbOption_Click()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    
    Text1.Text = ""
    CmbSubOption.CLEAR
    
    cmdsql = "select * from tblscriptsms where option='"
    cmdsql = cmdsql + CmbOption.Text + "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    CmbSubOption.CLEAR
    While Not M_Objrs.EOF
        CmbSubOption.AddItem M_Objrs("suboption")
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub CmbSubOption_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    Text1.Text = ""
    cmdsql = "select * from tblscriptsms where option='"
    cmdsql = cmdsql + CmbOption.Text + "' and suboption='"
    cmdsql = cmdsql + CmbSubOption.Text + "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    Text1.Text = Trim(M_Objrs("scriptsms"))
    
    rowaray = 0
    For i = 1 To Len(Text1.Text)
    If Mid(Text1.Text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
            
    ElseIf Mid(Text1.Text, i, 1) = "]" Then
        akhirk = i
         AvgMarks(1, rowaray) = i
         rowaray = rowaray + 1
    End If
    Next i
    Set M_Objrs = Nothing
End Sub

Private Sub Combo1_Click()
If Text5 = "" Then
If Left(Combo1, 1) <> "0" Then
Text5.Text = Text5.Text & "031" & Combo1.Text
Else
Text5.Text = Text5.Text & Combo1.Text
End If
Else
If Left(Combo1, 1) <> "0" Then
Text5.Text = Text5.Text & ",031" & Combo1.Text
Else
Text5.Text = Text5.Text & "," & Combo1.Text
End If
End If
End Sub

Private Sub Command1_Click()
 Dim teks1 As String
    Dim teks2 As String
    
    teks1 = Replace(Text1.Text, "[", "")
    teks2 = Replace(teks1, "]", "")
    teks2 = Replace(teks2, "'", "")

'cek data udah di simpen ke tabel receive apa belum??
Dim fields() As String
fields() = Split(Text5.Text, ",")
''MsgBox (UBound(fields) + 1)
For i = 0 To UBound(fields)
'    List1.AddItem Trim$(Fields(i))
'Next

'isi disini


            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
            TGLw = R_tglsys("tglsystem")
            TGLSERVERc = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
            
            R_tglsys.MoveNext
            Wend
            
            Set R_tglsys = Nothing
            

       cmdsql = "select * from request_sms where agent='" & Trim$(Text2) & "' and custid='" & Trim$(Text3) & "' and notelp='" & Trim$(fields(i)) & "' and status='0'"
       Set M_Objrs = New ADODB.Recordset
       If M_Objrs.state = 1 Then M_Objrs.Close
              
 M_Objrs.CursorLocation = adUseClient
      M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

       If M_Objrs.RecordCount = 0 Then
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'If Text1 <> "" Or Text5 <> "" Then



                cmdsql = "INSERT INTO request_sms "
                cmdsql = cmdsql + " ( agent, custid,name,notelp,pesan,status,tgl_kirim)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "')"
                M_OBJCONN.Execute cmdsql
'Unload Me
'Else
'End If
End If
Next
'MsgBox "SMS terkirim"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Dim teks1 As String
    Dim teks2 As String
    
    teks1 = Replace(Text1.Text, "[", "")
    teks2 = Replace(teks1, "]", "")
    
    MsgBox teks2
End Sub

Private Sub Form_Load()
Dim RSsms_send As ADODB.Recordset
Set RSsms_send = New ADODB.Recordset
Dim lst As listItem


RSsms_send.CursorLocation = adUseClient
cmdsql = "Select * from mgm where custid='" + FrmCC_Colection.lblcustid + "'"
RSsms_send.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 
While Not RSsms_send.EOF
 
    If (IsNull(RSsms_send("mobileno"))) Or RSsms_send("mobileno") = "" Then
    Else
        '@@281010 Cek apakah no telepon masuk dalam blacklist??
        If Trim(RSsms_send("f_mobileno")) = "0" Then
            Combo1.AddItem Replace(Trim(RSsms_send("mobileno")), " ", "")
        End If
    End If

    If (IsNull(RSsms_send("mobileno2"))) Or RSsms_send("mobileno2") = "" Then
    Else
        '@@281010 Cek apakah no telepon masuk dalam blacklist??
        If Trim(RSsms_send("f_mobileno2")) = "0" Then
            Combo1.AddItem Replace(Trim(RSsms_send("mobileno2")), " ", "")
        End If
    End If

    If (IsNull(RSsms_send("mobilenoadd1"))) Or RSsms_send("mobilenoadd1") = "" Then
    Else
        '@@281010 Cek apakah no telepon masuk dalam blacklist??
        If Trim(RSsms_send("f_mobilenoadd1")) = "0" Then
            Combo1.AddItem Replace(Trim(RSsms_send("mobilenoadd1")), " ", "")
        End If
    End If

    If (IsNull(RSsms_send("mobilenoadd2"))) Or RSsms_send("mobilenoadd2") = "" Then
    Else
        '@@281010 Cek apakah no telepon masuk dalam blacklist??
        If Trim(RSsms_send("f_mobilenoadd2")) = "0" Then
            Combo1.AddItem Replace(Trim(RSsms_send("mobilenoadd2")), " ", "")
        End If
    End If

    RSsms_send.MoveNext
Wend
Set RSsms_send = Nothing

Text3 = FrmCC_Colection.lblcustid
Text4 = FrmCC_Colection.LblNama
Text2 = MDIForm1.Text1

Load_Data_Option_SMSScript
End Sub


Private Sub Text1_Change()
Label6 = "Jumlah : " & Len(Text1)

If Len(Text1) > 160 Then
MsgBox "Hanya dapat mengirim sms sebanyak 160 Karakter"
End If
End Sub

Private Sub Load_Data_Option_SMSScript()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    cmdsql = "select distinct option from tblscriptsms"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    CmbOption.CLEAR
    While Not M_Objrs.EOF
        CmbOption.AddItem M_Objrs("option")
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim cek As Boolean
    cek = False
    For K = 0 To rowaray - 1
        Debug.Print Text1.SelStart
        update
        If Text1.SelStart >= AvgMarks(0, K) And Text1.SelStart < AvgMarks(1, K) Then
  
            If KeyAscii = vbKeyBack Then
                a = Mid(Text1.Text, Text1.SelStart, 1)
                If a = "[" Or a = "]" Then
                    KeyAscii = 0
                End If
            End If
            cek = True
            Exit For
        End If
    Next K

    If cek = False Then
        KeyAscii = 0
    End If
End Sub

Public Sub update()
    Dim i As Integer
    rowaray = 0
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
        ElseIf Mid(Text1.Text, i, 1) = "]" Then
            akhirk = i
            AvgMarks(1, rowaray) = i
            rowaray = rowaray + 1
        End If
    Next i
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        MsgBox "Anda tidak dapat menggunakan klik kanan!", vbCritical + vbOKOnly, "Peringatan"
        Text1.Text = ""
    End If
End Sub
