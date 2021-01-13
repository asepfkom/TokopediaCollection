VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form formnote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note"
   ClientHeight    =   10785
   ClientLeft      =   9840
   ClientTop       =   690
   ClientWidth     =   10440
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   10440
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Text            =   "1"
      Top             =   0
      Width           =   615
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   10320
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   18203
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"formnote.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8400
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   255
      Left            =   9840
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Save Dulu Sebelum Pindah Page"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Page"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   50
      Width           =   495
   End
   Begin VB.Menu mn 
      Caption         =   "Menu"
      Begin VB.Menu mnsave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnrd 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnud 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "formnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim saves As Boolean
Dim a As Integer
Dim M_Objrs As ADODB.Recordset
Dim notes As String

Private Sub Combo1_Click()
    If Combo1.text = "1" Then
        notes = "note"
    ElseIf Combo1.text = "2" Then
        notes = "note2"
    ElseIf Combo1.text = "3" Then
        notes = "note4"
    ElseIf Combo1.text = "4" Then
        notes = "note6"
    ElseIf Combo1.text = "5" Then
        notes = "note8"
    End If
    
    query = "select " & notes & " from tblnote where agent = '" & Text1.text & "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    RichTextBox1.text = cnull(M_Objrs(0))
    
    saves = False
End Sub

Private Sub Combo1_DropDown()
    Combo1.clear
    
    For i = 1 To 5
        Combo1.AddItem i
    Next i
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
        Text1.Visible = True
        Command1.Visible = True
    End If
    
    If UCase(MDIForm1.Text2.text) = "AGENT" Or UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        Text1.text = MDIForm1.Text1.text
    End If
'    Else
'        RichTextBox1.Top = 360
'        If Text1.text = "" Then
'            Text1.text = MDIForm1.Text1.text
'        End If
'    End If
    
    a = 0

    query = "select * from information_schema.columns  where table_name = 'tblnote'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        qcre = "create table tblnote (agent varchar, note text, note1 text, note2 text, note3 text, note4 text, note5 text, note6 text, note7 text, note8 text, note9 text);"
        M_OBJCONN.Execute qcre
    End If
    
    Set M_Objrs = Nothing
    
    saves = True
    
    query = "select note from tblnote where agent = '" & Text1.text & "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        query = "insert into tblnote values ('" & Text1.text & "', '','','','','','','','','','');"
        M_OBJCONN.Execute query
    Else
        RichTextBox1.text = cnull(M_Objrs!note)
    End If
    saves = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If saves = False Then
        m_msgbox = MsgBox("Anda Belum Save Perubahan, Apa Anda Mau Save??", vbYesNo + vbExclamation, "Aplikasi")
        If m_msgbox = vbYes Then
            Call Save
        End If
    End If
End Sub

Private Sub mnrd_Click()
'    a = a
    
'    If a > 0 Then
'        a = a - 1
'
'        If a = 0 Then
'            Dim B As String
'            B = ""
'        Else
'            B = a
'        End If
'
'        query = "select note" & B & " from tblnote where agent = '" & Text1.text & "'"
'        Set M_Objrs = New ADODB.Recordset
'        M_Objrs.CursorLocation = adUseClient
'        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If M_Objrs.RecordCount > 0 Then
'            RichTextBox1.text = cnull(M_Objrs(0))
'        End If
'
'        saves = True
'    End If

    If Combo1.text = "1" Then
        query = "select note from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "2" Then
        query = "select note2 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "3" Then
        query = "select note4 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "4" Then
        query = "select note6 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "5" Then
        query = "select note8 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    End If
End Sub

Private Sub mnsave_Click()
    If saves = False Then
        Call Save
    End If
End Sub

Private Sub Save()
Dim strText As String
Dim newList As ListItems
Dim n  As Integer
     
    If UCase(MDIForm1.Text1.text) = UCase(Text1.text) Then
        
         For n = 1 To Len(RichTextBox1.text)
              If Asc(Mid(RichTextBox1.text, n, 1)) = 13 Then
                   countcr = countcr + 1
              End If
         Next n
         
        If countct > 500 Then
            MsgBox "Harap hapus data yang lama atau tidak diperlukan, data sudah melebihi 500 line."
            GoTo bawah:
        End If
        
        m_msgbox = MsgBox("Data akan disave diPage " & Combo1.text & "?", vbYesNo + vbExclamation, "Aplikasi")
        If m_msgbox = vbNo Then
            GoTo bawah:
        End If
    
        strText = RichTextBox1.text
        
        If Combo1.text = "1" Then
            query = "update tblnote set note1 = note where agent = '" & MDIForm1.Text1.text & "';"
            query = query + "update tblnote set note = '" & Replace(strText, "'", "") & "' where agent = '" & MDIForm1.Text1.text & "';"
            M_OBJCONN.Execute query
        ElseIf Combo1.text = "2" Then
            query = "update tblnote set note3 = note2 where agent = '" & MDIForm1.Text1.text & "';"
            query = query + "update tblnote set note2 = '" & Replace(strText, "'", "") & "' where agent = '" & MDIForm1.Text1.text & "';"
            M_OBJCONN.Execute query
        ElseIf Combo1.text = "3" Then
            query = "update tblnote set note5 = note4 where agent = '" & MDIForm1.Text1.text & "';"
            query = query + "update tblnote set note4 = '" & Replace(strText, "'", "") & "' where agent = '" & MDIForm1.Text1.text & "';"
            M_OBJCONN.Execute query
        ElseIf Combo1.text = "4" Then
            query = "update tblnote set note7 = note6 where agent = '" & MDIForm1.Text1.text & "';"
            query = query + "update tblnote set note6 = '" & Replace(strText, "'", "") & "' where agent = '" & MDIForm1.Text1.text & "';"
            M_OBJCONN.Execute query
        ElseIf Combo1.text = "5" Then
            query = "update tblnote set note9 = note8 where agent = '" & MDIForm1.Text1.text & "';"
            query = query + "update tblnote set note8 = '" & Replace(strText, "'", "") & "' where agent = '" & MDIForm1.Text1.text & "';"
            M_OBJCONN.Execute query
        End If
            saves = True
    MsgBox "Data berhasil disimpan"
    End If

bawah:
End Sub

Private Sub mnud_Click()
'    a = a
    
    If Combo1.text = "1" Then
        query = "select note1 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "2" Then
        query = "select note3 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "3" Then
        query = "select note5 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "4" Then
        query = "select note7 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    ElseIf Combo1.text = "5" Then
        query = "select note9 from tblnote where agent = '" & Text1.text & "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount > 0 Then
            RichTextBox1.text = cnull(M_Objrs(0))
        End If
        saves = True
    
    End If
End Sub

Private Sub RichTextBox1_Change()
'    If saves = False Then
'        Call Save
'    End If
    saves = False
End Sub

