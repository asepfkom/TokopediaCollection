VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListHotProspect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Hot Prospect"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus status hot prospect"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&UnCek All"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TxtJml 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   4020
      Width           =   1035
   End
   Begin MSComctlLib.ListView LvHotPr 
      Height          =   4020
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   7091
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Label Label1 
      Caption         =   "Jumlah Data:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "FrmListHotProspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    LvHotPr.ColumnHeaders.ADD 1, , "Custid", 2500
    LvHotPr.ColumnHeaders.ADD 2, , "Nama", 3000
    LvHotPr.ColumnHeaders.ADD 3, , "Status Kept", 1500
End Sub


Private Sub CmdCekAll_Click()
    Dim w As Integer
    If LvHotPr.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    For w = 1 To LvHotPr.ListItems.Count
        LvHotPr.ListItems(w).Checked = True
    Next w
End Sub

Private Sub CmdHapus_Click()
    Dim w As Integer
    Dim cmdsql As String
    Dim K As String
    
    If LvHotPr.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    K = MsgBox("Anda yakin akan menghapus data hot prospect yang dicentang?", vbQuestion + vbYesNo, "Konfirmasi")
    If K = vbNo Then
        Exit Sub
    End If
    
    For w = 1 To LvHotPr.ListItems.Count
        If LvHotPr.ListItems(w).Checked = True Then
            cmdsql = "update mgm set status_htc=null where custid='"
            cmdsql = cmdsql + LvHotPr.ListItems(w).text + "'"
            M_OBJCONN.Execute cmdsql
        End If
    Next w
    Call IsiData
    MsgBox "Status Hot Prospect berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub Command1_Click()
    Dim w As Integer
    If LvHotPr.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    For w = 1 To LvHotPr.ListItems.Count
        LvHotPr.ListItems(w).Checked = False
    Next w
End Sub

Private Sub Form_Load()
    Call header
    Call IsiData
End Sub

Private Sub IsiData()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
    If UCase(MDIForm1.Text2.text) = "ADMIN" Then
        cmdsql = "select custid,name,status_keep from mgm where status_htc='1' order by name asc"
    End If
    If UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Then
        cmdsql = "select custid,name,status_keep from mgm where status_htc='1' order by name asc"
    End If
    If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Then
        cmdsql = "select custid,name,status_keep from mgm where status_htc='1' order by name asc"
    End If
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        cmdsql = "select custid,name,status_keep from mgm where status_htc='1' and agent in ("
        cmdsql = cmdsql + "select userid from  usertbl where team='"
        cmdsql = cmdsql + MDIForm1.Text1.text + "' and usertype='1')  order by name asc"
    End If
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        cmdsql = "select custid,name,status_keep from mgm where status_htc='1' and agent='"
        cmdsql = cmdsql + MDIForm1.Text1 + "' "
        cmdsql = cmdsql + "order by name asc"
    End If
    If Left(UCase(MDIForm1.Text2.text), 2) = "AM" Then
        cmdsql = "select custid,name,status_keep from mgm where status_htc='1' and agent in ("
        cmdsql = cmdsql + "select userid from  usertbl where team in (select tl from tblsettingam  where am = '" + MDIForm1.Text1.text + "')  and usertype='1')  order by name asc"
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvHotPr.ListItems.clear
    
    txtjml.text = M_Objrs.RecordCount
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data Hot Prospect tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    While Not M_Objrs.EOF
        Set listItem = LvHotPr.ListItems.ADD(, , M_Objrs("custid"))
            listItem.SubItems(1) = M_Objrs("name")
            If M_Objrs("status_keep") = "1" Then
                listItem.SubItems(2) = "KEPT"
            End If
        M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub
