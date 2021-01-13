VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInboXSms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form SMS"
   ClientHeight    =   5595
   ClientLeft      =   6960
   ClientTop       =   960
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Detail Message"
      Height          =   1275
      Left            =   1740
      TabIndex        =   10
      Top             =   4260
      Width           =   10155
      Begin VB.TextBox TxtDetailMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   1740
      TabIndex        =   4
      Top             =   0
      Width           =   10155
      Begin VB.OptionButton OptOutbox 
         Caption         =   "Outbox"
         Height          =   195
         Left            =   4800
         TabIndex        =   16
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton OptInbox 
         Caption         =   "Inbox"
         Height          =   195
         Left            =   3960
         TabIndex        =   15
         Top             =   540
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   435
         Left            =   8280
         TabIndex        =   14
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   435
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtSearch 
         Height          =   285
         Left            =   3960
         TabIndex        =   8
         Top             =   180
         Width           =   2835
      End
      Begin VB.CommandButton CmdReply 
         Caption         =   "&Reply"
         Height          =   435
         Left            =   1500
         TabIndex        =   6
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Custid/name:"
         Height          =   255
         Left            =   2940
         TabIndex        =   7
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame FrameInboxOutbox 
      Caption         =   "Inbox"
      Height          =   3435
      Left            =   1740
      TabIndex        =   3
      Top             =   780
      Width           =   10155
      Begin MSComctlLib.ListView LvInboxOutbox 
         Height          =   2760
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   4868
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   3060
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu"
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton CmdOutbox 
         Caption         =   "&Outbox"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Width           =   1395
      End
      Begin VB.CommandButton CmdInbox 
         Caption         =   "&Inbox"
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   1395
      End
   End
End
Attribute VB_Name = "FrmInboXSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClear_Click()
    txtsearch.Text = ""
End Sub

Private Sub CmdInbox_Click()
    
    Dim satu As String
    Dim dua As String
    Dim tiga As String
    Dim empat As String
    Dim M_Objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim lst As listItem
    Dim JmlBelumBaca As Integer
    Dim JmlSudahBaca As Integer

    'On Error Resume Next
    
    LvInboxOutbox.ListItems.CLEAR
    LvInboxOutbox.ColumnHeaders.CLEAR
    Call HeaderInbox
    TxtDetailMsg.Text = ""

    TELPo = "Select receivingdatetime, sendernumber, textdecoded,id,processed  from inbox where sendernumber in ('a',"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    
    'Jika yang login Agent
    If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + Trim(MDIForm1.Text1.Text) + "'"
    End If
    'Jika yang login TL
     If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
        MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua team anda!", vbOKOnly + vbInformation, "Informasi"
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent in ("
        cmdsql34 = cmdsql34 + "select userid from usertbl where team='"
        cmdsql34 = cmdsql34 + Trim(MDIForm1.Text1.Text) + "')"
    End If
    'Jika yang login admin
    If UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.Text2.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.Text)) = "SUPERVISOR" Then
        'MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua AGENT!", vbOKOnly + vbInformation, "Informasi"
        'cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm "
        MsgBox "Login sebagai agent!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    
    M_Objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.EOF = False Then
        If M_Objrs.RecordCount <> 0 Then
            PB1.Max = M_Objrs.RecordCount
        Else
            MsgBox "Tidak ada data customer!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
    End If
    
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        If M_Objrs.EOF = False Then
            If Len(M_Objrs("mobileno")) <> 0 Then
                satu = FindReplace(M_Objrs("mobileno"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(satu, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        End If
    
        If M_Objrs.EOF = False Then
            If Len(M_Objrs("mobileno2")) <> 0 Then
                dua = FindReplace(M_Objrs("mobileno2"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(dua, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        End If
    
        If M_Objrs.EOF = False Then
            If Len(M_Objrs("mobilenoadd1")) <> 0 Then
                tiga = FindReplace(M_Objrs("mobilenoadd1"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(tiga, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        End If
        
        If M_Objrs.EOF = False Then
            If Len(M_Objrs("mobilenoadd2")) <> 0 Then
                empat = FindReplace(M_Objrs("mobilenoadd2"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(empat, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        End If
    
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
    
    TELPo = Left(TELPo, Len(TELPo) - 1)
    Dim TELPo1
    Dim TELPo2
    
    TELPo1 = TELPo + ") and processed='f' order by receivingdatetime desc " 'Ini yang belum pernah di baca
    TELPo2 = TELPo + ") and processed='t' order by receivingdatetime desc " 'Ini yang udah pernah di baca
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic
    
    'Ini buat data inbox yang belum dibaca
    JmlBelumBaca = M_Objrs.RecordCount
    If M_Objrs.RecordCount <> 0 Then
        PB1.Max = JmlBelumBaca
    Else
        Dim Update_Status As String
        MsgBox "Tidak ada sms baru!", vbOKOnly + vbInformation, "Informasi"
        'Update status sms di usertbl jadi null, supaya ga blink
        Update_Status = "update usertbl set status_sms=null where userid='"
        Update_Status = Update_Status + Trim(MDIForm1.Text1.Text) + "'"
        M_OBJCONN.Execute Update_Status
        MDIForm1.TimerBlink.Enabled = False
        MDIForm1.Label9.ForeColor = vbBlack
    End If
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        
        S = Format(M_Objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
        t = Trim(M_Objrs!sendernumber)
        u = M_Objrs!textdecoded
        v = FindReplace(t, "+62", "0")
    
        If (Left(v, 3) = "021") Then
            v = Mid(v, 4, 20)
        End If
    
        Dim showlist As New ADODB.Recordset
        Dim TOTPTP As Currency
        Dim ssql As String
        
        If showlist.state = 1 Then showlist.Close
        ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
        showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      

        
        If showlist.EOF = False Then
            isicustid = showlist!CustId
            isiname = showlist!Name
            Set showlist = Nothing
        End If
    
        Set lst = LvInboxOutbox.ListItems.ADD(, , Trim(isicustid)) 'custid
            lst.SubItems(1) = Trim(isiname)  'Isi nama
            lst.SubItems(2) = Trim(v) 'Telepon
            lst.SubItems(3) = Trim(S) 'Receivingdatetime
            lst.SubItems(4) = Trim(IIf(IsNull(M_Objrs("textdecoded")), "", M_Objrs("textdecoded"))) 'Textsms
            lst.SubItems(5) = M_Objrs("id")
            lst.SubItems(6) = M_Objrs("processed")
            lst.Bold = True
            LvInboxOutbox.SelectedItem.ForeColor = vbRed
            
            lst.ListSubItems(1).ForeColor = vbRed
            lst.ListSubItems(2).ForeColor = vbRed
            lst.ListSubItems(3).ForeColor = vbRed
            lst.ListSubItems(4).ForeColor = vbRed
            lst.ListSubItems(5).ForeColor = vbRed
            lst.ListSubItems(6).ForeColor = vbRed
            M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
    'Ini buat data inbox yang sudah dibaca
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open TELPo2, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    JmlSudahBaca = M_Objrs.RecordCount
     If M_Objrs.RecordCount <> 0 Then
        PB1.Max = M_Objrs.RecordCount
    Else
        MsgBox "Data Inbox Kosong!", vbOKOnly + vbInformation, "Informasi"
    End If
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        S = Format(M_Objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
        t = Trim(M_Objrs!sendernumber)
        u = M_Objrs!textdecoded
        v = FindReplace(t, "+62", "0")
    
        If (Left(v, 3) = "021") Then
            v = Mid(v, 4, 20)
        End If
        
        If showlist.state = 1 Then showlist.Close
        ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
        showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If showlist.EOF = False Then
            isicustid = showlist!CustId
        End If
        If showlist.EOF = False Then
            isiname = showlist!Name
        End If
        Set showlist = Nothing
    
       Set lst = LvInboxOutbox.ListItems.ADD(, , isicustid) 'custid
            lst.SubItems(1) = Trim(isiname)  'isi nama
            lst.SubItems(2) = Trim(v) 'Telepon
            lst.SubItems(3) = Trim(S) 'Receivingdatetime
            lst.SubItems(4) = Trim(IIf(IsNull(M_Objrs("textdecoded")), "", M_Objrs("textdecoded"))) 'Textsms
            lst.SubItems(5) = Trim(M_Objrs("id"))
            lst.SubItems(6) = M_Objrs("processed")
            'lst.ForeColor = vbBlue
            M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
    FrameInboxOutbox.Caption = "Inbox" & "(" & CStr(JmlBelumBaca) & "/" & CStr(JmlSudahBaca + JmlBelumBaca) & ")"
    CmdInbox.Caption = "&Inbox " & "(" & CStr(JmlBelumBaca) & ")"
End Sub

Private Sub CmdNew_Click()
    '@@ 09022011 jika mengklik form sms diluar jendela customer, maka send sms non aktif
    If FrmInboXSms.Caption = "SMS" Then
        MsgBox "Anda hanya dapat mengirim sms, ketika anda membuka salah satu data customer yang akan anda kirim sms!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    FrmSendSmsNew.Show vbModal
End Sub

Private Sub CmdOutbox_Click()
    Dim satu As String
    Dim dua As String
    Dim tiga As String
    Dim empat As String
    Dim M_Objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim lst As listItem
    Dim JmlBelumBaca As Integer
    Dim JmlSudahBaca As Integer

    On Error Resume Next
  
    LvInboxOutbox.ListItems.CLEAR
    LvInboxOutbox.ColumnHeaders.CLEAR
    Call HeaderOutbox
    TxtDetailMsg.Text = ""
    
    TELPo = "Select sendingdatetime, destinationnumber, textdecoded,status,id  from sentitems where destinationnumber in ('a',"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    
    'Jika yang login Agent
    If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + Trim(MDIForm1.Text1.Text) + "'"
    End If
    'Jika yang login TL
     If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
        MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua team anda!", vbOKOnly + vbInformation, "Informasi"
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent in ("
        cmdsql34 = cmdsql34 + "select userid from usertbl where team='"
        cmdsql34 = cmdsql34 + Trim(MDIForm1.Text1.Text) + "')"
    End If
    'Jika yang login admin
    If UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.Text2.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.Text)) = "SUPERVISOR" Then
        'MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua AGENT!", vbOKOnly + vbInformation, "Informasi"
        'cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm "
        MsgBox "Login sebagai agent!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    M_Objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    While Not M_Objrs.EOF

        If Len(M_Objrs("mobileno")) <> 0 Then
            satu = FindReplace(M_Objrs("mobileno"), "0", "0")
            TELPo = TELPo + "'" + Trim(Replace(satu, "\", "")) + "',"
        Else
            TELPo = TELPo
        End If
    
        If Len(M_Objrs("mobileno2")) <> 0 Then
            dua = FindReplace(M_Objrs("mobileno2"), "0", "0")
            TELPo = TELPo + "'" + Trim(Replace(dua, "\", "")) + "',"
        Else
            TELPo = TELPo
        End If
    
        If Len(M_Objrs("mobilenoadd1")) <> 0 Then
            tiga = FindReplace(M_Objrs("mobilenoadd1"), "0", "0")
            TELPo = TELPo + "'" + Trim(Replace(tiga, "\", "")) + "',"
        Else
            TELPo = TELPo
        End If
    
        If Len(M_Objrs("mobilenoadd2")) <> 0 Then
            empat = FindReplace(M_Objrs("mobilenoadd2"), "0", "0")
            TELPo = TELPo + "'" + Trim(Replace(empat, "\", "")) + "',"
        Else
            TELPo = TELPo
        End If
    
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
    
    TELPo = Left(TELPo, Len(TELPo) - 1)
    Dim TELPo1
    
    
    TELPo1 = TELPo + ") order by sendingdatetime desc "
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    JmlBelumBaca = M_Objrs.RecordCount
    If M_Objrs.RecordCount <> 0 Then
        PB1.Max = JmlBelumBaca
    Else
        MsgBox "Tidak ada data outbox!", vbOKOnly + vbInformation, "Informasi"
    End If
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        
        S = Format(M_Objrs!sendingdatetime, "yyyy-mm-dd hh:mm:ss")
        t = Trim(M_Objrs!destinationnumber)
        u = M_Objrs!textdecoded
        v = FindReplace(t, "+62", "0")
    
        If (Left(v, 3) = "021") Then
            v = Mid(v, 4, 20)
        End If
    
        Dim showlist As New ADODB.Recordset
        Dim TOTPTP As Currency
        Dim ssql As String
        
        ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
        showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        isicustid = showlist!CustId
        isiname = showlist!Name
        Set showlist = Nothing
    
    
        Set lst = LvInboxOutbox.ListItems.ADD(, , isicustid) 'custid
            lst.SubItems(1) = Trim(isiname)  'Isi nama
            lst.SubItems(2) = Trim(v) 'Telepon
            lst.SubItems(3) = Trim(S) 'Receivingdatetime
            lst.SubItems(4) = Trim(IIf(IsNull(M_Objrs("textdecoded")), "", M_Objrs("textdecoded"))) 'Textsms
            lst.SubItems(5) = M_Objrs("id")
            lst.SubItems(6) = M_Objrs("status")
            lst.Bold = True
            M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
    
    
    FrameInboxOutbox.Caption = "Outbox " & "(" & CStr(JmlBelumBaca) & ")"
    CmdOutbox.Caption = "&Outbox " & "(" & CStr(JmlBelumBaca) & ")"
End Sub

Private Sub HeaderInbox()
    LvInboxOutbox.ColumnHeaders.ADD , , "Custid", 1700
    LvInboxOutbox.ColumnHeaders.ADD , , "Name", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "No.Handphone", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "Date Time", 1500
    LvInboxOutbox.ColumnHeaders.ADD , , "Message", 3000
    LvInboxOutbox.ColumnHeaders.ADD , , "Id", 0
    LvInboxOutbox.ColumnHeaders.ADD , , "Status", 1000
End Sub

Private Sub HeaderOutbox()
    LvInboxOutbox.ColumnHeaders.ADD , , "Custid", 1700
    LvInboxOutbox.ColumnHeaders.ADD , , "Name", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "No.Handphone", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "Date Time", 1500
    LvInboxOutbox.ColumnHeaders.ADD , , "Message", 3000
    LvInboxOutbox.ColumnHeaders.ADD , , "Id", 0
    LvInboxOutbox.ColumnHeaders.ADD , , "Status", 1000
End Sub


Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
Dim StartLoc
Dim FoundLoc
  If StartLoc = 0 Then StartLoc = 1
  FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
  If FoundLoc <> 0 And FoundLoc < 2 Then
     ReplaceFirstInstance = Left(SourceString, FoundLoc - 1) & Replacestring & Right(SourceString, Len(SourceString) - (FoundLoc - 1) - Len(Searchstring))
     StartLoc = FoundLoc + Len(Replacestring)
  ElseIf FoundLoc > 1 Then
  
      ReplaceFirstInstance = Replacestring & "21" & SourceString

  Else
     StartLoc = 1

    ReplaceFirstInstance = SourceString
  End If
End Function

Function FindReplace(SourceString, Searchstring, Replacestring) As String
  Dim tmpString1
  Dim tmpString2
  tmpString1 = SourceString
 
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      
      FindReplace = tmpString1
End Function

Private Sub CmdReply_Click()
    Dim CustId As String
    Dim Nama As String
    Dim Nohp As String
    Dim agent As String
    
    If LvInboxOutbox.ListItems.Count = 0 Then
        MsgBox "Tidak data yang akan di reply!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    CustId = LvInboxOutbox.SelectedItem.Text
    Nama = LvInboxOutbox.SelectedItem.SubItems(1)
    Nohp = LvInboxOutbox.SelectedItem.SubItems(2)
    
    With FrmSendSmsNew2
        .Text3.Text = Trim(CustId)
        .Text4.Text = Trim(Nama)
        .Text5.Text = Trim(Nohp)
        .Text2.Text = Trim(MDIForm1.Text1.Text)
        .Show vbModal
    End With
    
End Sub

Private Sub cmdsearch_Click()
    Dim satu As String
    Dim dua As String
    Dim tiga As String
    Dim empat As String
    Dim M_Objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim lst As listItem
    Dim JmlBelumBaca As Integer
    Dim JmlSudahBaca As Integer
    
    On Error Resume Next

    If OptInbox.Value Then
        LvInboxOutbox.ListItems.CLEAR
        LvInboxOutbox.ColumnHeaders.CLEAR
        Call HeaderInbox
        TxtDetailMsg.Text = ""
    
        TELPo = "Select receivingdatetime, sendernumber, textdecoded,id,processed  from inbox where sendernumber in ('a',"
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        
        'Jika yang login agent
        If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
            cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + Trim(MDIForm1.Text1.Text) + "' and (custid like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%' or name like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%')"
        End If
        'Jika yang login TeamLeader
        If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
            cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where (custid like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%' or name like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%') and agent in ("
            cmdsql34 = cmdsql34 + "select userid from usertbl where team='"
            cmdsql34 = cmdsql34 + Trim(MDIForm1.Text1.Text) + "')"
        End If
        'Jika yang login Administrator
        If UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.Text2.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.Text)) = "SUPERVISOR" Then
            MsgBox "Login sebagai agent!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        M_Objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        While Not M_Objrs.EOF
    
            If Len(M_Objrs("mobileno")) <> 0 Then
                satu = FindReplace(M_Objrs("mobileno"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(satu, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        
            If Len(M_Objrs("mobileno2")) <> 0 Then
                dua = FindReplace(M_Objrs("mobileno2"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(dua, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        
            If Len(M_Objrs("mobilenoadd1")) <> 0 Then
                tiga = FindReplace(M_Objrs("mobilenoadd1"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(tiga, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        
            If Len(M_Objrs("mobilenoadd2")) <> 0 Then
                empat = FindReplace(M_Objrs("mobilenoadd2"), "0", "+62")
                TELPo = TELPo + "'" + Trim(Replace(empat, "\", "")) + "',"
            Else
                TELPo = TELPo
            End If
        
            M_Objrs.MoveNext
        Wend
        
        Set M_Objrs = Nothing
        
        TELPo = Left(TELPo, Len(TELPo) - 1)
        Dim TELPo1
        Dim TELPo2
        
        TELPo1 = TELPo + ") and processed='f' order by receivingdatetime desc " 'Ini yang belum pernah di baca
        TELPo2 = TELPo + ") and processed='t' order by receivingdatetime desc " 'Ini yang udah pernah di baca
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        
        'Ini buat data inbox yang belum dibaca
        JmlBelumBaca = M_Objrs.RecordCount
        If M_Objrs.RecordCount <> 0 Then
            PB1.Max = JmlBelumBaca
        Else
            MsgBox "Tidak ada sms baru!", vbOKOnly + vbInformation, "Informasi"
        End If
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            
            S = Format(M_Objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
            t = Trim(M_Objrs!sendernumber)
            u = M_Objrs!textdecoded
            v = FindReplace(t, "+62", "0")
        
            If (Left(v, 3) = "021") Then
                v = Mid(v, 4, 20)
            End If
        
            Dim showlist As New ADODB.Recordset
            Dim TOTPTP As Currency
            Dim ssql As String
            
            ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
            showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
            isicustid = showlist!CustId
            isiname = showlist!Name
            Set showlist = Nothing
        
        
            Set lst = LvInboxOutbox.ListItems.ADD(, , Trim(isicustid)) 'custid
                lst.SubItems(1) = Trim(isiname)  'Isi nama
                lst.SubItems(2) = Trim(v) 'Telepon
                lst.SubItems(3) = Trim(S) 'Receivingdatetime
                lst.SubItems(4) = Trim(IIf(IsNull(M_Objrs("textdecoded")), "", M_Objrs("textdecoded"))) 'Textsms
                lst.SubItems(5) = M_Objrs("id")
                lst.SubItems(6) = M_Objrs("processed")
                lst.ForeColor = vbRed
                lst.Bold = True
                M_Objrs.MoveNext
        Wend
        Set M_Objrs = Nothing
        
        'Ini buat data inbox yang sudah dibaca
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open TELPo2, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        JmlSudahBaca = M_Objrs.RecordCount
         If M_Objrs.RecordCount <> 0 Then
            PB1.Max = M_Objrs.RecordCount
        Else
            MsgBox "Data Inbox Kosong!", vbOKOnly + vbInformation, "Informasi"
        End If
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            S = Format(M_Objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
            t = Trim(M_Objrs!sendernumber)
            u = M_Objrs!textdecoded
            v = FindReplace(t, "+62", "0")
        
            If (Left(v, 3) = "021") Then
                v = Mid(v, 4, 20)
            End If
        
            ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
            showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
            isicustid = showlist!CustId
            isiname = showlist!Name
            Set showlist = Nothing
        
           Set lst = LvInboxOutbox.ListItems.ADD(, , isicustid) 'custid
                lst.SubItems(1) = Trim(isiname)  'isi nama
                lst.SubItems(2) = Trim(v) 'Telepon
                lst.SubItems(3) = Trim(S) 'Receivingdatetime
                lst.SubItems(4) = Trim(IIf(IsNull(M_Objrs("textdecoded")), "", M_Objrs("textdecoded"))) 'Textsms
                lst.SubItems(5) = Trim(M_Objrs("id"))
                lst.SubItems(6) = M_Objrs("processed")
                lst.ForeColor = vbBlue
                M_Objrs.MoveNext
        Wend
        Set M_Objrs = Nothing
        
        FrameInboxOutbox.Caption = "Inbox" & "(" & CStr(JmlBelumBaca) & "/" & CStr(JmlSudahBaca + JmlBelumBaca) & ")"
        CmdInbox.Caption = "&Inbox " & "(" & CStr(JmlBelumBaca) & ")"
    End If
    '-------------------------------------------------------------------
    If OptOutbox.Value Then
        LvInboxOutbox.ListItems.CLEAR
        LvInboxOutbox.ColumnHeaders.CLEAR
        Call HeaderOutbox
        TxtDetailMsg.Text = ""
        
        TELPo = "Select sendingdatetime, destinationnumber, textdecoded,status,id  from sentitems where destinationnumber in ('a',"
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        
        'Jika yang login agent
        If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
            cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + Trim(MDIForm1.Text1.Text) + "' and (custid like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%' or name like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%')"
        End If
        'Jika yang login TeamLeader
        If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
            cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where (custid like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%' or name like '%"
            cmdsql34 = cmdsql34 + Trim(txtsearch.Text) + "%') and agent in ("
            cmdsql34 = cmdsql34 + "select userid from usertbl where team='"
            cmdsql34 = cmdsql34 + Trim(MDIForm1.Text1.Text) + "')"
        End If
        'Jika yang login Admin/Administrator/
        If UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.Text2.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.Text)) = "SUPERVISOR" Then
            MsgBox "Login sebagai agent!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        M_Objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        While Not M_Objrs.EOF
    
            If Len(M_Objrs("mobileno")) <> 0 Then
                satu = FindReplace(M_Objrs("mobileno"), "0", "0")
                TELPo = TELPo + "'" + satu + "',"
            Else
                TELPo = TELPo
            End If
        
            If Len(M_Objrs("mobileno2")) <> 0 Then
                dua = FindReplace(M_Objrs("mobileno2"), "0", "0")
                TELPo = TELPo + "'" + dua + "',"
            Else
                TELPo = TELPo
            End If
        
            If Len(M_Objrs("mobilenoadd1")) <> 0 Then
                tiga = FindReplace(M_Objrs("mobilenoadd1"), "0", "0")
                TELPo = TELPo + "'" + tiga + "',"
            Else
                TELPo = TELPo
            End If
        
            If Len(M_Objrs("mobilenoadd2")) <> 0 Then
                empat = FindReplace(M_Objrs("mobilenoadd2"), "0", "0")
                TELPo = TELPo + "'" + empat + "',"
            Else
                TELPo = TELPo
            End If
        
            M_Objrs.MoveNext
        Wend
        
        Set M_Objrs = Nothing
        
        TELPo = Left(TELPo, Len(TELPo) - 1)

        
        
        TELPo1 = TELPo + ") order by sendingdatetime desc "
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        JmlBelumBaca = M_Objrs.RecordCount
        If M_Objrs.RecordCount <> 0 Then
            PB1.Max = JmlBelumBaca
        Else
            MsgBox "Tidak ada data outbox!", vbOKOnly + vbInformation, "Informasi"
        End If
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            
            S = Format(M_Objrs!sendingdatetime, "yyyy-mm-dd hh:mm:ss")
            t = Trim(M_Objrs!destinationnumber)
            u = M_Objrs!textdecoded
            v = FindReplace(t, "+62", "0")
        
            If (Left(v, 3) = "021") Then
                v = Mid(v, 4, 20)
            End If
        
          
            
            ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
            showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
            isicustid = showlist!CustId
            isiname = showlist!Name
            Set showlist = Nothing
        
        
            Set lst = LvInboxOutbox.ListItems.ADD(, , isicustid) 'custid
                lst.SubItems(1) = Trim(isiname)  'Isi nama
                lst.SubItems(2) = Trim(v) 'Telepon
                lst.SubItems(3) = Trim(S) 'Receivingdatetime
                lst.SubItems(4) = Trim(IIf(IsNull(M_Objrs("textdecoded")), "", M_Objrs("textdecoded"))) 'Textsms
                lst.SubItems(5) = M_Objrs("id")
                lst.SubItems(6) = M_Objrs("status")
                lst.Bold = True
                M_Objrs.MoveNext
        Wend
        Set M_Objrs = Nothing
        
        
        
        FrameInboxOutbox.Caption = "Outbox " & "(" & CStr(JmlBelumBaca) & ")"
        CmdOutbox.Caption = "&Outbox " & "(" & CStr(JmlBelumBaca) & ")"
    End If
End Sub

Private Sub Form_Load()
    ' Matikan Timer
    open_sms = True
    'Jika yang login agent
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        'CmdInbox_Click
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    open_sms = False
End Sub

Private Sub LvInboxOutbox_Click()
    Dim SisaSms As Integer
    Dim listItem As listItem
    
    If LvInboxOutbox.ListItems.Count <> 0 Then
        TxtDetailMsg.Text = LvInboxOutbox.SelectedItem.SubItems(4)
        'Update status
        cmdsql = "update inbox set processed='t' where id='"
        cmdsql = cmdsql + Trim(LvInboxOutbox.SelectedItem.SubItems(5)) + "'"
        LvInboxOutbox.SelectedItem.ForeColor = vbBlack
        M_OBJCONN1.Execute cmdsql
        'Ini jika ada sms masuk, baru dibaca
        If LvInboxOutbox.SelectedItem.SubItems(6) = "0" Then
            LvInboxOutbox.SelectedItem.SubItems(6) = "1"
            
            SisaSms = Val(MDIForm1.LblJmlSmsBaru.Caption) - 1
            
            If SisaSms < 0 Then
                MDIForm1.LblJmlSmsBaru.Caption = "0"
            Else
                MDIForm1.LblJmlSmsBaru.Caption = Val(MDIForm1.LblJmlSmsBaru.Caption) - 1
            End If
            
            CmdInbox.Caption = "&Inbox(" & MDIForm1.LblJmlSmsBaru.Caption & ")"
            'update blinknya di usertbl
            cmdsql = "update usertbl set status_sms=null where userid='"
            cmdsql = cmdsql + Trim(MDIForm1.Text1.Text) + "'"
            M_OBJCONN.Execute cmdsql
            MDIForm1.TimerBlink.Enabled = False
            MDIForm1.Label9.ForeColor = vbBlack
            
        End If
    End If
End Sub


Private Sub LvInboxOutbox_DblClick()
    If LvInboxOutbox.ListItems.Count = 0 Then
        Exit Sub
    End If
    VIEW_MGMDATA.Text1(2).Text = LvInboxOutbox.SelectedItem.Text
End Sub

Private Sub OptInbox_Click()
    txtsearch.SetFocus
End Sub

Private Sub OptOutbox_Click()
    txtsearch.SetFocus
End Sub
