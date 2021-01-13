VERSION 5.00
Begin VB.Form frmreqvalidphonesms 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Pilih Nomor Valid"
   ClientHeight    =   765
   ClientLeft      =   7710
   ClientTop       =   4560
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   765
   ScaleWidth      =   4680
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Req"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   200
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   220
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmreqvalidphonesms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    If Text5 = "" Then
        If Left(Combo1, 1) <> "0" Then
            Text5.text = Text5.text & "021" & Combo1.text
        Else
            Text5.text = Text5.text & Combo1.text
        End If
    Else
        If Left(Combo1, 1) <> "0" Then
            Text5.text = Text5.text & ",021" & Combo1.text
        Else
            Text5.text = Text5.text & "," & Combo1.text
        End If
    End If

End Sub

Private Sub Command1_Click()
    CustId = FrmCC_Colection.lblCustId.Caption
    nomorvalid = Text5.text
    agent = MDIForm1.Text1.text
    
    If MDIForm1.Text2.text <> "Agent" Then
        agent = FrmCC_Colection.lblaoc.Caption
    End If
'--
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseClient
    r.Open "select team from usertbl where userid = '" + agent + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'--
    tl = r!TEAM

    q = "INSERT INTO tblvalidtotl values ('" + CustId + "', '" + nomorvalid + "', '" + agent + "', '" + tl + "', now())"
    M_OBJCONN.Execute q
    
    MsgBox "Request Berhasil, harap tunggu untuk diapprove TL dan SPV"
    Unload Me
End Sub

Private Sub Form_Load()
    CustId = FrmCC_Colection.lblCustId.Caption
    
    q = "select * from tblvalidtotl where custid = '" + CustId + "'"
    Set r = New ADODB.Recordset
    r.CursorLocation = adUseClient
    r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If r.RecordCount > 0 Then
        MsgBox "Nomor sudah di Request, dalam tahap Approve TL"
        'FrmCC_Colection.Label25.Caption = "1"
        'Exit Sub
        'Unload Me
    Else
        Set r = Nothing
        
        q = "select * from tblvalidtospv where custid = '" + CustId + "'"
        Set r = New ADODB.Recordset
        r.CursorLocation = adUseClient
        r.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If r.RecordCount > 0 Then
            MsgBox "Nomor sudah di Request, dalam tahap Approve SPV"
            'Unload Me
        Else
            Set RSsms_send = New ADODB.Recordset
            RSsms_send.CursorLocation = adUseClient
            cmdsql = "SELECT btrim as no_tlp FROM ("
            cmdsql = cmdsql + "    SELECT trim(mobileno) FROM mgm WHERE trim(mobileno) not in (select no_telp from tblblacklist) and custid        = '" + FrmCC_Colection.lblCustId + "' "
            cmdsql = cmdsql + "    Union All"
            cmdsql = cmdsql + "    SELECT trim(mobileno2) FROM mgm WHERE trim(mobileno2) not in (select no_telp from tblblacklist) and            custid = '" + FrmCC_Colection.lblCustId + "' "
            cmdsql = cmdsql + "    Union All"
            cmdsql = cmdsql + "    SELECT trim(mobilenoadd1) FROM mgm WHERE trim(mobilenoadd1) not in (select no_telp from tblblacklist) and       custid = '" + FrmCC_Colection.lblCustId + "' "
            cmdsql = cmdsql + "    Union All"
            cmdsql = cmdsql + "    SELECT trim(mobilenoadd2) FROM mgm WHERE trim(mobilenoadd2) not in (select no_telp from tblblacklist) and       custid = '" + FrmCC_Colection.lblCustId + "') a "
            RSsms_send.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
             
            While Not RSsms_send.EOF
                Combo1.AddItem Replace(Trim(RSsms_send("no_tlp")), " ", "")
                RSsms_send.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
