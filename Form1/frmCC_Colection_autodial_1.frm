VERSION 5.00
Begin VB.Form FrmCC_Colection_autodial 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10950
   ClientLeft      =   540
   ClientTop       =   15
   ClientWidth     =   19140
   ControlBox      =   0   'False
   Icon            =   "frmCC_Colection_autodial.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   19140
   Begin VB.PictureBox SSFrame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   11025
      Left            =   0
      ScaleHeight     =   10965
      ScaleWidth      =   19200
      TabIndex        =   28
      Top             =   0
      Width           =   19260
      Begin VB.Frame Frame19 
         BackColor       =   &H00ABE18E&
         Height          =   2205
         Left            =   60
         TabIndex        =   107
         Top             =   8760
         Width           =   6555
         Begin VB.TextBox txtremarks 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   1335
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   196
            Top             =   720
            Width           =   3135
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            ItemData        =   "frmCC_Colection_autodial.frx":000C
            Left            =   4440
            List            =   "frmCC_Colection_autodial.frx":0016
            TabIndex        =   189
            Top             =   180
            Width           =   2055
         End
         Begin VB.ComboBox cboaccount 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   1320
            TabIndex        =   188
            Top             =   180
            Width           =   1905
         End
         Begin VB.ComboBox cbolastcall 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_autodial.frx":002E
            Left            =   4440
            List            =   "frmCC_Colection_autodial.frx":0035
            TabIndex        =   168
            Top             =   540
            Width           =   2055
         End
         Begin VB.PictureBox cmbDateSch 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4425
            ScaleHeight     =   285
            ScaleWidth      =   1230
            TabIndex        =   108
            Top             =   900
            Width           =   1260
         End
         Begin VB.PictureBox cmbTimeSch 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   5700
            ScaleHeight     =   285
            ScaleWidth      =   765
            TabIndex        =   109
            Top             =   900
            Width           =   795
         End
         Begin VB.PictureBox SSCommand1 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   600
            Index           =   2
            Left            =   5040
            ScaleHeight     =   540
            ScaleWidth      =   585
            TabIndex        =   112
            Top             =   1320
            Width           =   645
         End
         Begin VB.PictureBox SSCommand1 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   600
            Index           =   3
            Left            =   5760
            ScaleHeight     =   540
            ScaleWidth      =   585
            TabIndex        =   113
            Top             =   1320
            Width           =   645
         End
         Begin VB.PictureBox SSCommand1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   600
            Index           =   4
            Left            =   3600
            ScaleHeight     =   540
            ScaleWidth      =   585
            TabIndex        =   170
            Top             =   1320
            Width           =   645
         End
         Begin VB.PictureBox CmdKeep 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   600
            Left            =   4320
            ScaleHeight     =   540
            ScaleWidth      =   585
            TabIndex        =   261
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "HOT PR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   4305
            TabIndex        =   260
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            Caption         =   "Select Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   12
            Left            =   60
            TabIndex        =   191
            Top             =   180
            Width           =   1305
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            Caption         =   "Status Call"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   3210
            TabIndex        =   190
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label31 
            BackColor       =   &H00ABE18E&
            Caption         =   "Speak With"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   3210
            TabIndex        =   169
            Top             =   570
            Width           =   1245
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   5040
            TabIndex        =   116
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5760
            TabIndex        =   115
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "CPA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   3615
            TabIndex        =   114
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label39 
            BackColor       =   &H00ABE18E&
            Caption         =   "Tgl Follow up"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3210
            TabIndex        =   111
            Top             =   900
            Width           =   1245
         End
         Begin VB.Label Label31 
            BackColor       =   &H00ABE18E&
            Caption         =   "Remarks: max 80 karakter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   110
            Top             =   480
            Width           =   2475
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         ForeColor       =   &H80000008&
         Height          =   4755
         Left            =   6780
         TabIndex        =   66
         Top             =   6120
         Width           =   12225
         Begin VB.Timer TimerOfferingDiscon 
            Interval        =   1500
            Left            =   3840
            Top             =   1560
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   600
            TabIndex        =   215
            Text            =   "Text6"
            Top             =   1500
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CommandButton CmdHapusRemarks 
            Caption         =   "&Hapus Remarks"
            Height          =   435
            Left            =   2460
            TabIndex        =   211
            Top             =   120
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.Timer TimerCekMapping 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   3420
            Top             =   840
         End
         Begin VB.Timer TimerBlinkSms 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2760
            Top             =   1380
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   7
            Left            =   60
            TabIndex        =   67
            Top             =   120
            Width           =   12135
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "D E C E A S E"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   375
               Index           =   19
               Left            =   9720
               TabIndex        =   298
               Top             =   70
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   7
               Left            =   60
               Picture         =   "frmCC_Colection_autodial.frx":0044
               Stretch         =   -1  'True
               Top             =   30
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "History Remarks"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   7
               Left            =   540
               TabIndex        =   68
               Top             =   60
               Width           =   2115
            End
         End
         Begin VB.PictureBox listview1 
            Appearance      =   0  'Flat
            BackColor       =   &H009AD6C2&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   4080
            Index           =   1
            Left            =   60
            ScaleHeight     =   4050
            ScaleWidth      =   12045
            TabIndex        =   246
            Top             =   600
            Width           =   12075
         End
      End
      Begin VB.Frame Frame1 
         Height          =   930
         Left            =   9690
         TabIndex        =   29
         Top             =   9210
         Width           =   2775
         Begin VB.Label LblStatus 
            Caption         =   "Label42"
            Height          =   255
            Left            =   600
            TabIndex        =   65
            Top             =   360
            Width           =   255
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   75
            TabIndex        =   34
            Top             =   315
            Width           =   60
         End
         Begin VB.Label lblCardNo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2460
            TabIndex        =   33
            Top             =   315
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label CustId 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "# Card"
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
            Height          =   195
            Index           =   0
            Left            =   1905
            TabIndex        =   32
            Top             =   285
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Emergency Contact"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   15195
            TabIndex        =   31
            Top             =   1590
            Width           =   1890
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Telp Tambahan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   10680
            TabIndex        =   30
            Top             =   135
            Width           =   1500
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         ForeColor       =   &H80000008&
         Height          =   10875
         Left            =   6720
         TabIndex        =   60
         Top             =   60
         Width           =   12615
         Begin VB.CommandButton CmdViewRecording 
            BackColor       =   &H000080FF&
            Caption         =   "View Recording"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   295
            Top             =   4920
            Width           =   1635
         End
         Begin VB.CommandButton cmd_logcomplaint 
            Caption         =   "Create Complaint"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2160
            TabIndex        =   294
            Top             =   5400
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H008080FF&
            Caption         =   "Set Decease"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   293
            Top             =   5400
            Width           =   1635
         End
         Begin VB.CommandButton CmdClaimAcc 
            Caption         =   "Claim Account ini"
            Height          =   435
            Left            =   480
            TabIndex        =   291
            Top             =   5400
            Width           =   1635
         End
         Begin VB.TextBox TxtTelpKe 
            BackColor       =   &H0000C0C0&
            Height          =   285
            Left            =   2220
            TabIndex        =   251
            Text            =   "NoPhone"
            Top             =   5400
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CmdRequestNumber 
            Caption         =   "Request Number"
            Height          =   435
            Left            =   3840
            TabIndex        =   245
            Top             =   4920
            Width           =   1635
         End
         Begin VB.CommandButton CmdDataMapping 
            BackColor       =   &H0080FFFF&
            Caption         =   "&Keep Account"
            Height          =   435
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   244
            Top             =   5400
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Timer TimerBlinkDetailMapping 
            Interval        =   1000
            Left            =   5580
            Top             =   5460
         End
         Begin VB.CommandButton CmdRequest 
            BackColor       =   &H0080FFFF&
            Caption         =   "&List Keep Account"
            Height          =   435
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   212
            Top             =   4920
            Width           =   1635
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Inbox/Outbox"
            Height          =   720
            Left            =   3660
            Style           =   1  'Graphical
            TabIndex        =   209
            Top             =   3900
            Width           =   1665
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   4
            Left            =   6000
            TabIndex        =   166
            Top             =   3900
            Width           =   6195
            Begin VB.CommandButton CmddetailPayment 
               BackColor       =   &H0080FF80&
               Caption         =   "Show Payment"
               Height          =   375
               Left            =   2760
               MaskColor       =   &H0080FF80&
               Style           =   1  'Graphical
               TabIndex        =   290
               Top             =   50
               Width           =   1335
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Payment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   180
               TabIndex        =   167
               Top             =   105
               Width           =   2355
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   9330
            TabIndex        =   179
            Top             =   0
            Width           =   2895
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Emergency Contact"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   2
               Left            =   510
               TabIndex        =   180
               Top             =   120
               Width           =   2175
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   2
               Left            =   90
               Picture         =   "frmCC_Colection_autodial.frx":04B8
               Stretch         =   -1  'True
               Top             =   60
               Width           =   375
            End
         End
         Begin VB.Frame FrmPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            ForeColor       =   &H80000008&
            Height          =   1770
            Left            =   6000
            TabIndex        =   171
            Top             =   4260
            Width           =   6195
            Begin VB.PictureBox txtSisaHutang 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1200
               TabIndex        =   172
               Top             =   750
               Width           =   1230
            End
            Begin VB.PictureBox TxtAfterPay 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1200
               TabIndex        =   173
               Top             =   480
               Width           =   1230
            End
            Begin VB.PictureBox TxtPayment2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1200
               TabIndex        =   174
               Top             =   195
               Width           =   1230
            End
            Begin VB.PictureBox listview1 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1530
               Index           =   0
               Left            =   45
               ScaleHeight     =   1500
               ScaleWidth      =   3645
               TabIndex        =   175
               Top             =   180
               Width           =   3675
            End
            Begin VB.PictureBox TxtLPAPayment 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1215
               TabIndex        =   247
               Top             =   1305
               Width           =   1245
            End
            Begin VB.PictureBox TxtLPDPayment 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               IMEMode         =   3  'DISABLE
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1215
               TabIndex        =   248
               Top             =   1020
               Width           =   1245
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "LPD"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   18
               Left            =   3780
               TabIndex        =   250
               Top             =   1020
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "LPA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   17
               Left            =   3780
               TabIndex        =   249
               Top             =   1305
               Width           =   885
            End
            Begin VB.Label Label15 
               BackColor       =   &H00ABE18E&
               Caption         =   "Sisa:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   3795
               TabIndex        =   178
               Top             =   750
               Width           =   1005
            End
            Begin VB.Label Label13 
               BackColor       =   &H00ABE18E&
               Caption         =   "Jml Dibayar:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3795
               TabIndex        =   177
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label10 
               BackColor       =   &H00ABE18E&
               Caption         =   "Jml PTP:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   3795
               TabIndex        =   176
               Top             =   195
               Width           =   1005
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   5
            Left            =   3300
            TabIndex        =   164
            Top             =   0
            Width           =   5955
            Begin VB.Image Image1 
               Height          =   375
               Index           =   5
               Left            =   75
               Picture         =   "frmCC_Colection_autodial.frx":1D52
               Stretch         =   -1  'True
               Top             =   60
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Additional Info"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   540
               TabIndex        =   165
               Top             =   105
               Width           =   1575
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   1
            Left            =   60
            TabIndex        =   140
            Top             =   0
            Width           =   3135
            Begin VB.Image Image1 
               Height          =   375
               Index           =   1
               Left            =   60
               Picture         =   "frmCC_Colection_autodial.frx":2271
               Stretch         =   -1  'True
               Top             =   60
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Phone Information"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   540
               TabIndex        =   141
               Top             =   105
               Width           =   1815
            End
         End
         Begin VB.Frame Frame12 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3210
            Left            =   60
            TabIndex        =   61
            Top             =   510
            Width           =   12165
            Begin VB.Frame Frame20 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               ForeColor       =   &H80000008&
               Height          =   3285
               Left            =   9240
               TabIndex        =   181
               Top             =   -60
               Width           =   2895
               Begin VB.CommandButton CmdOther 
                  Caption         =   "&Other"
                  Height          =   435
                  Left            =   1320
                  TabIndex        =   254
                  Top             =   2820
                  Width           =   1455
               End
               Begin VB.TextBox txtremarkstrace 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   945
                  Left            =   0
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   193
                  Top             =   1860
                  Width           =   2790
               End
               Begin VB.TextBox txtECAdd 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   765
                  Left            =   735
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   182
                  Top             =   720
                  Width           =   2010
               End
               Begin VB.PictureBox txtECnoA 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   720
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   183
                  Top             =   150
                  Width           =   1755
               End
               Begin VB.PictureBox TxtEC 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   720
                  ScaleHeight     =   225
                  ScaleWidth      =   1980
                  TabIndex        =   184
                  Top             =   420
                  Width           =   2010
               End
               Begin VB.PictureBox txtECno 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   720
                  ScaleHeight     =   225
                  ScaleWidth      =   1545
                  TabIndex        =   185
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.Label LblBlackliSTEC 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2520
                  TabIndex        =   288
                  Top             =   150
                  Width           =   195
               End
               Begin VB.Label Label35 
                  BackColor       =   &H003F9E0C&
                  Caption         =   "Addr"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   765
                  Left            =   30
                  TabIndex        =   203
                  Top             =   720
                  Width           =   705
               End
               Begin VB.Label Label34 
                  Alignment       =   2  'Center
                  BackColor       =   &H003F9E0C&
                  Caption         =   "Add. Info From Tracer"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   192
                  Top             =   1560
                  Width           =   2805
               End
               Begin VB.Label Label23 
                  BackColor       =   &H003F9E0C&
                  Caption         =   "Telp "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   30
                  TabIndex        =   187
                  Top             =   150
                  Width           =   1815
               End
               Begin VB.Label Label21 
                  BackColor       =   &H003F9E0C&
                  Caption         =   "Nama"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   30
                  TabIndex        =   186
                  Top             =   420
                  Width           =   660
               End
            End
            Begin VB.Frame Frame17 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               ForeColor       =   &H80000008&
               Height          =   3285
               Left            =   3240
               TabIndex        =   143
               Top             =   -60
               Width           =   5955
               Begin VB.ComboBox CmbStsKatHome1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3B0B
                  Left            =   3300
                  List            =   "frmCC_Colection_autodial.frx":3B27
                  TabIndex        =   275
                  Text            =   "--Pilih Kategori Telepon--"
                  Top             =   120
                  Width           =   2445
               End
               Begin VB.ComboBox CmbStsKatOffice1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3BA5
                  Left            =   3300
                  List            =   "frmCC_Colection_autodial.frx":3BC1
                  TabIndex        =   274
                  Text            =   "--Pilih Kategori Telepon--"
                  Top             =   840
                  Width           =   2445
               End
               Begin VB.ComboBox CmbStsKatOffice2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3C3F
                  Left            =   3300
                  List            =   "frmCC_Colection_autodial.frx":3C5B
                  TabIndex        =   273
                  Text            =   "--Pilih Kategori Telepon--"
                  Top             =   1230
                  Width           =   2445
               End
               Begin VB.ComboBox CmbStsKatHP1 
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3CD9
                  Left            =   3300
                  List            =   "frmCC_Colection_autodial.frx":3CF5
                  TabIndex        =   272
                  Text            =   "--Pilih Kategori Telepon--"
                  Top             =   1620
                  Width           =   2460
               End
               Begin VB.ComboBox CmbStsKatHP2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3D73
                  Left            =   3300
                  List            =   "frmCC_Colection_autodial.frx":3D8F
                  TabIndex        =   271
                  Text            =   "--Pilih Kategori Telepon--"
                  Top             =   1980
                  Width           =   2460
               End
               Begin VB.ComboBox CmbStsKatHome2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3E0D
                  Left            =   3300
                  List            =   "frmCC_Colection_autodial.frx":3E29
                  TabIndex        =   270
                  Text            =   "--Pilih Kategori Telepon--"
                  Top             =   480
                  Width           =   2445
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00ABE18E&
                  Height          =   795
                  Left            =   3060
                  TabIndex        =   264
                  Top             =   2400
                  Width           =   2775
                  Begin VB.PictureBox TxtNoTelpReq 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFC0&
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   255
                     Left            =   720
                     ScaleHeight     =   225
                     ScaleWidth      =   1905
                     TabIndex        =   268
                     Top             =   480
                     Width           =   1935
                  End
                  Begin VB.Label label1 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00ABE18E&
                     Caption         =   "No.Tlp:"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   21
                     Left            =   60
                     TabIndex        =   267
                     Top             =   480
                     Width           =   1455
                  End
                  Begin VB.Label TxtKategori 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFC0&
                     BorderStyle     =   1  'Fixed Single
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   255
                     Left            =   720
                     TabIndex        =   266
                     Top             =   180
                     Width           =   1950
                  End
                  Begin VB.Label label1 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00ABE18E&
                     Caption         =   "Kat.Tlp:"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   15
                     Left            =   60
                     TabIndex        =   265
                     Top             =   180
                     Width           =   1575
                  End
               End
               Begin VB.PictureBox txtOfficeAdd1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   144
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox txtOfficeAdd2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   145
                  Top             =   1230
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox txtOfficeAdd1A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   146
                  Top             =   840
                  Width           =   1815
               End
               Begin VB.PictureBox txtOfficeAdd2A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   147
                  Top             =   1230
                  Width           =   1815
               End
               Begin VB.PictureBox txtMobileAdd1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   148
                  Top             =   1590
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox txtMobileAdd2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   149
                  Top             =   1950
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox txtMobileAdd1A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   150
                  Top             =   1590
                  Width           =   1815
               End
               Begin VB.PictureBox txtMobileAdd2A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   151
                  Top             =   1950
                  Width           =   1815
               End
               Begin VB.PictureBox AddrNow 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   120
                  ScaleHeight     =   705
                  ScaleWidth      =   2865
                  TabIndex        =   152
                  Top             =   2490
                  Width           =   2895
               End
               Begin VB.PictureBox txtHomeAdd1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   153
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox txtHomeAdd2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   154
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox txtHomeAdd1A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   155
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.PictureBox txtHomeAdd2A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Left            =   900
                  ScaleHeight     =   315
                  ScaleWidth      =   1785
                  TabIndex        =   156
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.Label LblBlacklistAddHP2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   287
                  Top             =   1980
                  Width           =   195
               End
               Begin VB.Label LblBlacklistAddHP1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   286
                  Top             =   1620
                  Width           =   195
               End
               Begin VB.Label LblBlacklistAddOffice2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   285
                  Top             =   1260
                  Width           =   195
               End
               Begin VB.Label LblBlacklistAddOffice1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   284
                  Top             =   900
                  Width           =   195
               End
               Begin VB.Label LblBlacklistAddHome2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   283
                  Top             =   540
                  Width           =   195
               End
               Begin VB.Label LblBlacklistAddHome1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   282
                  Top             =   180
                  Width           =   195
               End
               Begin VB.Label Label19 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Add  Adress:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   163
                  Top             =   2280
                  Width           =   795
               End
               Begin VB.Label label1 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "HP I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   14
                  Left            =   120
                  TabIndex        =   162
                  Top             =   1590
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "HP II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   13
                  Left            =   120
                  TabIndex        =   161
                  Top             =   1950
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Rumah I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   250
                  Index           =   20
                  Left            =   120
                  TabIndex        =   160
                  Top             =   120
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Rumah II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   19
                  Left            =   120
                  TabIndex        =   159
                  Top             =   480
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Kantor I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   17
                  Left            =   120
                  TabIndex        =   158
                  Top             =   840
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Kantor II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   16
                  Left            =   120
                  TabIndex        =   157
                  Top             =   1260
                  Width           =   765
               End
            End
            Begin VB.Frame Frame16 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               ForeColor       =   &H80000008&
               Height          =   3285
               Left            =   0
               TabIndex        =   117
               Top             =   -90
               Width           =   3135
               Begin VB.PictureBox TxtAdditional 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   214
                  Top             =   2820
                  Width           =   1755
               End
               Begin VB.ComboBox CmbPhone 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frmCC_Colection_autodial.frx":3EA7
                  Left            =   1140
                  List            =   "frmCC_Colection_autodial.frx":3EAE
                  Locked          =   -1  'True
                  TabIndex        =   118
                  Text            =   "CmbPhone"
                  Top             =   210
                  Width           =   1920
               End
               Begin VB.PictureBox txtHomeNo2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   119
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.PictureBox txtOfficeNo2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   120
                  Top             =   1605
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.PictureBox txtMobileNo1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   121
                  Top             =   1905
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.PictureBox txtMobileNo2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   122
                  Top             =   2175
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.PictureBox txtHomeNo2A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   123
                  Top             =   945
                  Width           =   1755
               End
               Begin VB.PictureBox txtOfficeNo2A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   124
                  Top             =   1605
                  Width           =   1755
               End
               Begin VB.PictureBox txtMobileNo1A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   125
                  Top             =   1905
                  Width           =   1755
               End
               Begin VB.PictureBox txtMobileNo2A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   126
                  Top             =   2175
                  Width           =   1755
               End
               Begin VB.PictureBox txtHomeNo1A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   127
                  Top             =   630
                  Width           =   1755
               End
               Begin VB.PictureBox txtOfficeNo1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   128
                  Top             =   1275
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.PictureBox txtOfficeNo1A 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   129
                  Top             =   1275
                  Width           =   1755
               End
               Begin VB.PictureBox txtHomeNo1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   225
                  ScaleWidth      =   1725
                  TabIndex        =   130
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.PictureBox tdbtelptrace 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1020
                  ScaleHeight     =   255
                  ScaleWidth      =   1695
                  TabIndex        =   194
                  Top             =   2175
                  Visible         =   0   'False
                  Width           =   1695
               End
               Begin VB.Label LblBlacklistHp2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   281
                  Top             =   2175
                  Width           =   195
               End
               Begin VB.Label LblBlacklistHp1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   280
                  Top             =   1905
                  Width           =   195
               End
               Begin VB.Label LblBlacklistOfficeno2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   279
                  Top             =   1620
                  Width           =   195
               End
               Begin VB.Label LblBlacklistOffice1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   278
                  Top             =   1260
                  Width           =   195
               End
               Begin VB.Label LblBlacklistHome2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   277
                  Top             =   960
                  Width           =   195
               End
               Begin VB.Label LblBlakcListHome1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00004000&
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H003F9E0C&
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   276
                  Top             =   630
                  Width           =   195
               End
               Begin VB.Label Label22 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Add."
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   213
                  Top             =   2820
                  Width           =   735
                  WordWrap        =   -1  'True
               End
               Begin VB.Label LblMother 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFC0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1035
                  TabIndex        =   198
                  Top             =   2460
                  Width           =   1755
               End
               Begin VB.Label Label22 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Mother Name"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   197
                  Top             =   2460
                  Width           =   735
                  WordWrap        =   -1  'True
               End
               Begin VB.Label label1 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "No Tujuan :"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Index           =   9
                  Left            =   120
                  TabIndex        =   137
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Kantor II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   136
                  Top             =   1605
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Kantor I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   135
                  Top             =   1275
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Rumah I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   134
                  Top             =   615
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "Rumah II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   133
                  Top             =   945
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "HP I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   132
                  Top             =   1875
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00ABE18E&
                  Caption         =   "HP II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   131
                  Top             =   2175
                  Width           =   735
               End
            End
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            ForeColor       =   &H80000008&
            Height          =   1725
            Left            =   6060
            TabIndex        =   69
            Top             =   8160
            Visible         =   0   'False
            Width           =   5805
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   3720
               TabIndex        =   79
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   3360
               TabIndex        =   78
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.OptionButton Option10 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Send"
               Height          =   255
               Left            =   4710
               TabIndex        =   77
               Top             =   360
               Width           =   735
            End
            Begin VB.OptionButton Option9 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Inbox"
               Height          =   255
               Left            =   4710
               TabIndex        =   76
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   4200
               TabIndex        =   75
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Timer Timer_cek_inbox 
               Enabled         =   0   'False
               Interval        =   30000
               Left            =   4020
               Top             =   420
            End
         End
         Begin VB.Label lbl_agentlama 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Lama"
            Height          =   375
            Left            =   4200
            TabIndex        =   292
            Top             =   5520
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   289
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label LabelSms 
            Alignment       =   2  'Center
            BackColor       =   &H003F9E0C&
            Caption         =   "Label SMS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3660
            TabIndex        =   210
            Top             =   4620
            Width           =   1665
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "Offers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   2640
            TabIndex        =   142
            Top             =   4620
            Width           =   900
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "Call"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   139
            Top             =   4620
            Width           =   900
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H003F9E0C&
            Caption         =   "Hang Up"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   1620
            TabIndex        =   138
            Top             =   4620
            Width           =   900
         End
         Begin VB.Label LBLEXP 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   7980
            TabIndex        =   62
            Top             =   7080
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         ForeColor       =   &H80000008&
         Height          =   10935
         Left            =   -180
         TabIndex        =   35
         Top             =   0
         Width           =   6885
         Begin VB.CommandButton CmdSendPTP 
            Caption         =   "&Send PTP"
            Height          =   435
            Left            =   5280
            TabIndex        =   269
            Top             =   5160
            Width           =   1515
         End
         Begin VB.ComboBox cboPTP 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmCC_Colection_autodial.frx":3EB7
            Left            =   960
            List            =   "frmCC_Colection_autodial.frx":3EB9
            Locked          =   -1  'True
            TabIndex        =   255
            Top             =   5160
            Width           =   1455
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   3
            Left            =   0
            TabIndex        =   96
            Top             =   8280
            Width           =   7035
            Begin VB.Image Image1 
               Height          =   375
               Index           =   3
               Left            =   75
               Picture         =   "frmCC_Colection_autodial.frx":3EBB
               Stretch         =   -1  'True
               Top             =   30
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Call Actvity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   97
               Top             =   60
               Width           =   1455
            End
         End
         Begin VB.PictureBox LstDoubleId 
            Appearance      =   0  'Flat
            BackColor       =   &H009AD6C2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   810
            Left            =   240
            ScaleHeight     =   780
            ScaleWidth      =   6450
            TabIndex        =   59
            Top             =   4380
            Width           =   6480
         End
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3915
            Left            =   240
            TabIndex        =   36
            Top             =   480
            Width           =   6465
            Begin VB.PictureBox lblOfficeAddr 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Left            =   780
               ScaleHeight     =   645
               ScaleWidth      =   2970
               TabIndex        =   37
               Top             =   2160
               Width           =   3000
            End
            Begin VB.PictureBox lblAddr 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   780
               ScaleHeight     =   660
               ScaleWidth      =   2985
               TabIndex        =   38
               Top             =   1425
               Width           =   3015
            End
            Begin VB.PictureBox lblAmount 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   53
               Top             =   45
               Width           =   1545
            End
            Begin VB.PictureBox LblPrompA 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   54
               Top             =   330
               Width           =   1545
            End
            Begin VB.PictureBox tdbmaxad 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4800
               ScaleHeight     =   255
               ScaleWidth      =   1545
               TabIndex        =   72
               Top             =   3990
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.PictureBox tdbminad 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4800
               ScaleHeight     =   255
               ScaleWidth      =   1545
               TabIndex        =   73
               Top             =   4260
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.PictureBox LblMinPayment 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   2040
               ScaleHeight     =   375
               ScaleWidth      =   1740
               TabIndex        =   208
               Top             =   3480
               Width           =   1740
            End
            Begin VB.PictureBox TxtInstallment 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   218
               Top             =   600
               Width           =   1545
            End
            Begin VB.PictureBox lblOpenDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               IMEMode         =   3  'DISABLE
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   221
               Top             =   1500
               Width           =   1545
            End
            Begin VB.PictureBox lblLimit 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   222
               Top             =   1200
               Width           =   1545
            End
            Begin VB.PictureBox lblBD 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               IMEMode         =   3  'DISABLE
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   225
               Top             =   1800
               Width           =   1545
            End
            Begin VB.PictureBox lblLastPay 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   226
               Top             =   2385
               Width           =   1545
            End
            Begin VB.PictureBox lblPayDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               IMEMode         =   3  'DISABLE
               Left            =   4845
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   227
               Top             =   2100
               Width           =   1545
            End
            Begin VB.PictureBox TxtInterest 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   231
               Top             =   2700
               Width           =   1545
            End
            Begin VB.PictureBox Tdbbalance 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   233
               Top             =   3000
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.PictureBox TDB_cur_bal 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   235
               Top             =   3300
               Width           =   1545
            End
            Begin VB.PictureBox TxtCurpri 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               ScaleHeight     =   225
               ScaleWidth      =   1515
               TabIndex        =   237
               Top             =   3600
               Width           =   1545
            End
            Begin VB.Label Txtperiod 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4860
               TabIndex        =   239
               Top             =   900
               Width           =   1545
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Cur  Pri"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   15
               Left            =   3960
               TabIndex        =   238
               Top             =   3600
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Curr Bal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   11
               Left            =   3960
               TabIndex        =   236
               Top             =   3300
               Width           =   885
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   14
               Left            =   3960
               TabIndex        =   234
               Top             =   3000
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Interest"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   16
               Left            =   3960
               TabIndex        =   232
               Top             =   2700
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Wo Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   3960
               TabIndex        =   230
               Top             =   1800
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "LPA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   3960
               TabIndex        =   229
               Top             =   2385
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "LPD"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   228
               Top             =   2100
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Limit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   3960
               TabIndex        =   224
               Top             =   1200
               Width           =   885
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Open Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3960
               TabIndex        =   223
               Top             =   1500
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Instalment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   9
               Left            =   3960
               TabIndex        =   220
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Period"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   10
               Left            =   3960
               TabIndex        =   219
               Top             =   900
               Width           =   885
            End
            Begin VB.Label lblpurge 
               Appearance      =   0  'Flat
               BackColor       =   &H0000C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3240
               TabIndex        =   217
               Top             =   210
               Width           =   540
            End
            Begin VB.Label lbltype 
               Appearance      =   0  'Flat
               BackColor       =   &H00008080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2640
               TabIndex        =   216
               Top             =   210
               Width           =   540
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H003F9E0C&
               Caption         =   "MIN.PAYMENT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   2040
               TabIndex        =   207
               Top             =   3240
               Width           =   1740
            End
            Begin VB.Label LblCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   960
               TabIndex        =   202
               Top             =   3480
               Width           =   960
            End
            Begin VB.Label LblMap 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FFFF&
               Height          =   375
               Left            =   -60
               TabIndex        =   201
               Top             =   3480
               Width           =   960
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H003F9E0C&
               Caption         =   "CYCLE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   960
               TabIndex        =   200
               Top             =   3240
               Width           =   960
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H003F9E0C&
               Caption         =   "MAP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   -60
               TabIndex        =   199
               Top             =   3240
               Width           =   960
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Min A.d"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   3900
               TabIndex        =   71
               Top             =   4260
               Visible         =   0   'False
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Max A.d"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   3900
               TabIndex        =   70
               Top             =   3990
               Visible         =   0   'False
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Principle"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   8
               Left            =   3960
               TabIndex        =   56
               Top             =   330
               Width           =   885
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   3975
               TabIndex        =   55
               Top             =   45
               Width           =   885
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   52
               Top             =   525
               Width           =   720
            End
            Begin VB.Label lblNama 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   51
               Top             =   525
               Width           =   3030
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "ID No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   840
               Width           =   720
            End
            Begin VB.Label lblID 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   49
               Top             =   810
               Width           =   3030
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "DOB"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   1140
               Width           =   720
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   690
               Left            =   0
               TabIndex        =   47
               Top             =   1420
               Width           =   720
            End
            Begin VB.Label Label27 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Office Add"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   675
               Left            =   0
               TabIndex        =   46
               Top             =   2160
               Width           =   720
            End
            Begin VB.Label lblZIP 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2760
               TabIndex        =   45
               Top             =   2880
               Width           =   1020
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "ZipCode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   1980
               TabIndex        =   44
               Top             =   2880
               Width           =   720
            End
            Begin VB.Label LblDOB 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   43
               Top             =   1110
               Width           =   1380
            End
            Begin VB.Label Label37 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "Region"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   42
               Top             =   2880
               Width           =   720
            End
            Begin VB.Label lblregion 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   41
               Top             =   2880
               Width           =   1140
            End
            Begin VB.Label lblCustId 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   40
               Top             =   210
               Width           =   1830
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               Caption         =   "No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   65
               Left            =   0
               TabIndex        =   39
               Top             =   210
               Width           =   720
            End
         End
         Begin VB.Frame Frame18 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            Caption         =   "Reserve PTP"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1365
            Left            =   3600
            TabIndex        =   103
            Top             =   6900
            Width           =   3270
            Begin VB.PictureBox LstReserve 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1035
               Left            =   60
               ScaleHeight     =   1005
               ScaleWidth      =   2295
               TabIndex        =   104
               Top             =   225
               Width           =   2325
            End
            Begin VB.PictureBox SSCommand2 
               AutoSize        =   -1  'True
               Height          =   615
               Index           =   3
               Left            =   2430
               ScaleHeight     =   555
               ScaleWidth      =   555
               TabIndex        =   105
               Top             =   210
               Width           =   615
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H003F9E0C&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   2430
               TabIndex        =   106
               Top             =   810
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            Caption         =   "PTP Jatuh Tempo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1365
            Left            =   240
            TabIndex        =   98
            Top             =   6900
            Width           =   3375
            Begin VB.PictureBox LstPayment 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1005
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   2475
               TabIndex        =   99
               Top             =   240
               Width           =   2505
            End
            Begin VB.PictureBox SSCommand2 
               AutoSize        =   -1  'True
               Height          =   615
               Index           =   2
               Left            =   2670
               ScaleHeight     =   555
               ScaleWidth      =   555
               TabIndex        =   100
               Top             =   270
               Width           =   615
            End
            Begin VB.PictureBox SSCommand2 
               Height          =   735
               Index           =   1
               Left            =   3690
               ScaleHeight     =   675
               ScaleWidth      =   690
               TabIndex        =   101
               Top             =   1710
               Visible         =   0   'False
               Width           =   750
            End
            Begin VB.Label lblhapus 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H003F9E0C&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   2670
               TabIndex        =   102
               Top             =   855
               Width           =   615
            End
         End
         Begin VB.CheckBox C_PTP 
            BackColor       =   &H003F9E0C&
            Caption         =   "PTP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   95
            Top             =   5160
            Width           =   1710
         End
         Begin VB.Frame frmPTP 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            Enabled         =   0   'False
            ForeColor       =   &H003F9E0C&
            Height          =   1500
            Left            =   240
            TabIndex        =   82
            Top             =   5460
            Width           =   6585
            Begin VB.ComboBox CmbViaPtp 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmCC_Colection_autodial.frx":4403
               Left            =   1740
               List            =   "frmCC_Colection_autodial.frx":4416
               TabIndex        =   253
               Top             =   1140
               Width           =   3015
            End
            Begin VB.ComboBox CmbBaseOn 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               ItemData        =   "frmCC_Colection_autodial.frx":4447
               Left            =   2400
               List            =   "frmCC_Colection_autodial.frx":4449
               TabIndex        =   206
               Top             =   1140
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.ComboBox cmbDiscount 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               ItemData        =   "frmCC_Colection_autodial.frx":444B
               Left            =   3960
               List            =   "frmCC_Colection_autodial.frx":444D
               TabIndex        =   205
               Text            =   "0"
               Top             =   1140
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.CheckBox C_Payment 
               Enabled         =   0   'False
               Height          =   255
               Left            =   3060
               TabIndex        =   84
               Top             =   1200
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CheckBox Chktenor 
               BackColor       =   &H00ABE18E&
               Caption         =   "Tenor"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1740
               TabIndex        =   83
               Top             =   480
               Width           =   795
            End
            Begin VB.PictureBox txttenor 
               Appearance      =   0  'Flat
               BackColor       =   &H00004000&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2520
               ScaleHeight     =   225
               ScaleWidth      =   465
               TabIndex        =   85
               Top             =   480
               Width           =   495
            End
            Begin VB.PictureBox TDBDate3 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1740
               ScaleHeight     =   225
               ScaleWidth      =   1425
               TabIndex        =   86
               Top             =   780
               Width           =   1485
            End
            Begin VB.PictureBox txtPayment 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1740
               ScaleHeight     =   225
               ScaleWidth      =   1560
               TabIndex        =   87
               Top             =   180
               Width           =   1590
            End
            Begin VB.PictureBox Tdabamoint 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1740
               ScaleHeight     =   225
               ScaleWidth      =   1380
               TabIndex        =   88
               Top             =   780
               Width           =   1410
            End
            Begin VB.PictureBox tdbptpnew 
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   4560
               ScaleHeight     =   225
               ScaleWidth      =   1425
               TabIndex        =   90
               Top             =   120
               Width           =   1485
            End
            Begin VB.PictureBox TdbTglTagih 
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   4200
               ScaleHeight     =   225
               ScaleWidth      =   1425
               TabIndex        =   257
               Top             =   540
               Width           =   1485
            End
            Begin VB.PictureBox SSCommand2 
               AutoSize        =   -1  'True
               Height          =   615
               Index           =   0
               Left            =   5880
               ScaleHeight     =   555
               ScaleWidth      =   555
               TabIndex        =   89
               Top             =   600
               Width           =   615
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Tgl.Tagih"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   11
               Left            =   3360
               TabIndex        =   256
               Top             =   540
               Width           =   1005
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Pembayaran Via:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   252
               Top             =   1200
               Width           =   1665
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Date PTPNew"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   18
               Left            =   3360
               TabIndex        =   204
               Top             =   180
               Width           =   1245
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Date Payment Effective"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   60
               TabIndex        =   94
               Top             =   780
               Width           =   1605
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Total Amount Deal:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   77
               Left            =   60
               TabIndex        =   93
               Top             =   180
               Width           =   1785
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Installment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   42
               Left            =   1800
               TabIndex        =   92
               Top             =   780
               Width           =   1005
            End
            Begin VB.Label lbltambahedit 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H003F9E0C&
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   5880
               TabIndex        =   91
               Top             =   1200
               Width           =   615
            End
         End
         Begin VB.TextBox TXTRUMUS 
            Height          =   315
            Left            =   300
            TabIndex        =   81
            Top             =   4740
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   0
            Left            =   60
            TabIndex        =   57
            Top             =   15
            Width           =   6795
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   255
               Left            =   1500
               TabIndex        =   74
               Tag             =   "0"
               Top             =   0
               Visible         =   0   'False
               Width           =   135
            End
            Begin VB.Label lblClass 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   297
               Top             =   75
               Width           =   780
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H003F9E0C&
               Caption         =   "CLASS :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   7
               Left            =   5400
               TabIndex        =   296
               Top             =   75
               Width           =   645
            End
            Begin VB.Label Label32 
               BackColor       =   &H003F9E0C&
               Caption         =   "Coding "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   1860
               TabIndex        =   243
               Top             =   60
               Width           =   735
            End
            Begin VB.Label lblaoc 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   2580
               TabIndex        =   242
               Top             =   75
               Width           =   750
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H003F9E0C&
               Caption         =   "Batch"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   80
               Left            =   3420
               TabIndex        =   241
               Tag             =   "0"
               Top             =   75
               Width           =   660
            End
            Begin VB.Label lblRecsource 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4080
               TabIndex        =   240
               Top             =   75
               Width           =   1230
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   0
               Left            =   135
               Picture         =   "frmCC_Colection_autodial.frx":444F
               Stretch         =   -1  'True
               Tag             =   "0"
               Top             =   -30
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Personal Data"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   540
               TabIndex        =   58
               Top             =   -15
               Width           =   1455
            End
         End
         Begin VB.TextBox txthasil 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   3960
            TabIndex        =   195
            Top             =   3840
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label LblPP 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Left            =   6420
            TabIndex        =   263
            Top             =   60
            Width           =   435
         End
         Begin VB.Label LblPop 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Left            =   5400
            TabIndex        =   262
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label LblResultPTP 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3600
            TabIndex        =   259
            Top             =   5220
            Width           =   1440
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00ABE18E&
            Caption         =   "Result PTP:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   2580
            TabIndex        =   258
            Top             =   5220
            Width           =   1245
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Frame FrmPayment1 
      Height          =   1365
      Left            =   1920
      TabIndex        =   22
      Top             =   8295
      Width           =   2085
      Begin VB.CheckBox Check3 
         Caption         =   "Regular to paid Off"
         Height          =   195
         Left            =   75
         TabIndex        =   25
         Top             =   285
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Iregular to Paid Off"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Regular Payment"
         Height          =   195
         Left            =   75
         TabIndex        =   23
         Top             =   870
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox TdbPTP 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   1440
         TabIndex        =   26
         Top             =   585
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.PictureBox TdbDatePTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   285
         TabIndex        =   27
         Top             =   1065
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.Frame Frame9 
      Height          =   3405
      Left            =   75
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   1755
      Begin VB.OptionButton Option8 
         Caption         =   "Tambah"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   2
         Top             =   2070
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Batal"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1395
         TabIndex        =   1
         Top             =   2085
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Frame Frame8 
         ForeColor       =   &H000000FF&
         Height          =   1725
         Left            =   60
         TabIndex        =   3
         Top             =   2145
         Visible         =   0   'False
         Width           =   7560
         Begin VB.TextBox TxtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   540
            Width           =   3135
         End
         Begin VB.TextBox TxtCustid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   3375
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   915
            TabIndex        =   7
            Top             =   225
            Width           =   1815
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Alamat Billing"
            Height          =   195
            Index           =   0
            Left            =   4125
            TabIndex        =   6
            Top             =   855
            Width           =   1440
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Rumah"
            Height          =   195
            Index           =   1
            Left            =   5565
            TabIndex        =   5
            Top             =   855
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Kantor"
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   4
            Top             =   840
            Width           =   840
         End
         Begin VB.PictureBox TDBNumber1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   915
            ScaleHeight     =   285
            ScaleWidth      =   585
            TabIndex        =   10
            Top             =   870
            Width           =   615
         End
         Begin VB.PictureBox TXtDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   4080
            ScaleHeight     =   540
            ScaleWidth      =   3330
            TabIndex        =   11
            Top             =   225
            Width           =   3360
         End
         Begin VB.PictureBox TDBDate2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   915
            ScaleHeight     =   285
            ScaleWidth      =   1425
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.PictureBox TDBDate1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1590
            ScaleHeight     =   285
            ScaleWidth      =   1425
            TabIndex        =   13
            Top             =   870
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.PictureBox TxtAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   4065
            ScaleHeight     =   510
            ScaleWidth      =   3345
            TabIndex        =   14
            Top             =   1065
            Width           =   3375
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Nomor"
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
            Height          =   255
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Note:"
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
            Height          =   255
            Left            =   2925
            TabIndex        =   20
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Date"
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
            Height          =   255
            Left            =   30
            TabIndex        =   19
            Top             =   1245
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Ke"
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
            Height          =   255
            Left            =   30
            TabIndex        =   18
            Top             =   930
            Width           =   810
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Nama"
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
            Height          =   255
            Left            =   30
            TabIndex        =   17
            Top             =   540
            Width           =   810
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Custid"
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
            Height          =   315
            Index           =   2
            Left            =   420
            TabIndex        =   16
            Top             =   3375
            Width           =   1095
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Ke:"
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
            Height          =   270
            Left            =   3390
            TabIndex        =   15
            Top             =   915
            Width           =   615
         End
      End
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   7695
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txtPhoneA 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   7680
      Width           =   1905
   End
   Begin VB.PictureBox TDBNumber2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1395
      TabIndex        =   80
      Top             =   0
      Width           =   1395
   End
End
Attribute VB_Name = "FrmCC_Colection_autodial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cust As ADODB.Recordset
Dim M_update As ADODB.Recordset
Dim M_Objrs As ADODB.Recordset
Dim stscall As Boolean
Dim TYPETELP As String
Dim kontak As Boolean
Dim spend As Boolean
Dim adaSCH As Boolean
Dim adaREG As Boolean
Dim adaPO As Boolean
Dim vrcek As String
Dim vrdateptp As String
Dim vramount As String
Dim vrtdbdateptp As String
Dim vrbaseon As String
Dim vrdiskon As String
Dim vrtenor As String
Dim vrttlptp As String
Dim TglPTPNew As String
Dim vrnewdate As String
Dim KelapKelip As Integer
Dim kelapkelipDetail As Integer
'@@02-05-2012 Tambahan buat Catet Status Kategori
Dim StsKategoriTelepon As String
Dim KelompokKategoriTlp As String
Dim StatusSpeakWith As String
Dim StatusAccount As String

Dim TanggalPaidOff As String


Private Sub C_Contacted_Click()
    
    
If C_Contacted.Value Then
        C_VALID.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
      '  C_POPSP.Value = False
        FrmContacted.Enabled = True
      '  cboPOPSP.Text = ""
   Else
        cmbContacted.text = ""
        cmbDescCon.text = ""
        FrmContacted.Enabled = False
        If cboPOPSP.text = "" Then
            C_Payment.Value = False
        End If
        CmbBaseOn.text = ""
        cmbDiscount.text = 0
        TdbPTP.Value = ""
        txtPayment.Value = 0
End If
End Sub

Private Sub C_NotContacted_Click()
   If C_NotContacted.Value Then
      FrmUnContacted.Enabled = True
      C_Contacted.Value = False
      C_Payment.Value = False
   Else
      FrmUnContacted.Enabled = False
      cmbDescUn.text = ""
      cmbUncontacted = ""
   End If
End Sub

Private Sub C_Payment_Click()
   If C_Payment.Value Then
     ' Frame54.Enabled = True
   Else
     ' Frame54.Enabled = False
     'If cboPOPSP.Text <> "" Then
     'Exit Sub
     'End If
     
      'cmbDiscount.Text = ""
   End If
End Sub
Private Sub C_PTP_Click()
    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim m_objrs_payment As ADODB.Recordset
    
    

If C_PTP.Value Then
       '@@ 29 Desember 2011, Cek terlebih dahulu, apakah ada CPA atau tidak, jika tidak ada CPA maka
        'tidak bisa melakukan PTP

       cmdsql = "select * from tblcpa where vcustid='"
       cmdsql = cmdsql + Trim(lblCustId.Caption) + "' order by nid desc"
       Set M_Objrs = New ADODB.Recordset
       M_Objrs.CursorLocation = adUseClient
       M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

       If M_Objrs.RecordCount = 0 Then
        'C_PTP.Value = vbUnchecked
        'MsgBox "Untuk membuat status account PTP, harus dibuat terlebih dahulu CPA nya!", vbOKOnly + vbInformation, "Informasi"
        'Set M_OBJRS = Nothing
        'Exit Sub
       Else
        'Ambil Nilai Payment di CPA untuk di tempatkan di Total Amount Deal
        txtPayment.Value = IIf(IsNull(M_Objrs("nttlpayment")), "", M_Objrs("nttlpayment"))
        txttenor.Value = IIf(IsNull(M_Objrs("nperiod")), "", M_Objrs("nperiod"))
        Chktenor.Value = vbChecked
       End If

       Set M_Objrs = Nothing
       
 '@@ 11042012 Dinonaktifkan
'       If Left(cboaccount.Text, 3) <> "ON-" Then
'         cboaccount.Text = ""
'       End If
       
        bcekptp = False
 '       C_VALID.Value = False
'        C_SKIP.Value = False
'        C_Contacted.Value = False
        frmPTP.Enabled = True
        FrmPayment.Enabled = True
        'cboPOPSP.Tag = 0
        Label43(2).Visible = True
       ' cboPOPSP.Text = ""
        C_Payment.Value = 1
        If UCase(MDIForm1.Text2) = "AGENT" Then
            SSCommand1(4).Visible = False
            Label43(2).Visible = False
            Else
            SSCommand1(4).Visible = True
            Label43(2).Visible = True
        End If
        
        
        '@@22 Desember 2011 Tambahan, jika tidak ada pembayaran maka status PTP= PTP NEW
'        If listview1(0).ListItems.Count = 0 Then
'            cboPTP.Text = "PTP-NEW"
'        End If
'        If listview1(0).ListItems.Count > 0 Then
'            cboPTP.Text = "PTP-POP"
'        End If
        cmdsql = "SELECT b.custid as custid1, a.CUSTID,a.PayDate,a.Payment,"
        cmdsql = cmdsql + "a.Agent,a.FieldName,a.Id from tbllunas a inner join mgm b "
        cmdsql = cmdsql + "on a.custid=b.custid WHERE a.custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and date(a.Paydate)+1  > b.tglsource "
        Set m_objrs_payment = New ADODB.Recordset
        m_objrs_payment.CursorLocation = adUseClient
        m_objrs_payment.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_objrs_payment.RecordCount = 0 Then
            cboPTP.text = "PTP-NEW"
        Else
            cboPTP.text = "PTP-POP"
        End If
        Set m_objrs_payment = Nothing
        
   Else
       bcekptp = False
       Label43(2).Visible = False
        'C_Payment.Value = 0
       ' CmbBaseOn.Text = ""
       ' cmbDiscount.Text = 0
        'txtPayment.Value = 0
'        TxtPtpAddr.Text = ""
 '       TxtPhonePTP.Text = ""
      '  FrmPayment.Enabled = False
        cboPTP.text = ""
        SSCommand1(4).Visible = False
        frmPTP.Enabled = False
        TdbPTP.Value = ""
        CmbBaseOn.text = ""
        cmbDiscount.text = 0
        TdbPTP.Value = ""
        txtPayment.Value = 0
        'C_Payment = False
        Chktenor.Value = vbUnchecked
        txttenor.Value = 0
        TDBDate3.Value = ""
        CmbViaPtp.text = ""
        tdbptpnew.Value = ""
        TdbTglTagih.Value = ""
End If

End Sub

Private Sub C_SKIP_Click()
If C_SKIP.Value Then
        C_VALID.Value = False
        C_Contacted.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
     
        FrmSKIP.Enabled = True
   Else
        cboskip.text = ""
        cbodescskip.text = ""
        FrmSKIP.Enabled = False
End If

End Sub

Private Sub C_VALID_Click()
If C_VALID.Value Then
        C_Contacted.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
        
        FrMValid.Enabled = True
   Else
        cbovalid.text = ""
        cbodescvalid.text = ""
        FrMValid.Enabled = False
End If

End Sub

Private Sub cbodescskip_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cbodescvalid_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cboaccount_Click()
    Dim M_COL1 As New ADODB.Recordset
    cboaccount.Locked = True
    
'@@ 11-04-2012, Dinonaktifkan
'    If Left(cboaccount, 3) <> "ON-" Then
'        C_Payment.Value = vbUnchecked
'        C_PTP.Value = vbUnchecked
'    End If

 C_Payment.Value = vbUnchecked
 C_PTP.Value = vbUnchecked

If UCase(Left(cboaccount.text, 2)) = "SP" Then
    C_PTP.Value = 0
    CmbBaseOn.text = ""
    cmbDiscount.text = ""
    txtPayment.Value = 0
    Tdabamoint.Value = 0
    TDBDate3.Value = ""
    txttenor.Value = 0
    C_Payment.Value = 1
    FrmPayment.Enabled = True
            Set M_COL1 = New ADODB.Recordset
            cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
            M_COL1.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
            CmbBaseOn.text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(M_COL1!dateptp), "", M_COL1!dateptp))
            txttenor.Value = CStr(IIf(IsNull(M_COL1!Tenor), 0, M_COL1!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(M_COL1!amountptp), 0, M_COL1!amountptp))
End If


End Sub

Private Sub cboaccount_DropDown()
     cboaccount.Locked = False
End Sub

Private Sub cboaccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 And Shift = 1 Then KeyCode = 0 'paste(shift + insert)
    If KeyCode = 45 And Shift = 2 Then KeyCode = 0 'copy(ctrl + insert)
    If KeyCode = 86 And Shift = 2 Then KeyCode = 0 'paste(ctrl + v)
    If KeyCode = 67 And Shift = 2 Then KeyCode = 0 'copy(ctrl + c)
    If KeyCode = 88 And Shift = 2 Then KeyCode = 0 'copy(ctrl + x)
End Sub

Private Sub cboaccount_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbolastcall_Click()
    Select Case UCase(cbolastcall.text)
        Case "CH"
            StatusSpeakWith = "CH"
        Case "RECEPTION/OPERATOR/SEC/OB"
            StatusSpeakWith = "ROSO"
        Case "ATASAN"
            StatusSpeakWith = "BOSS"
        Case "HRD"
            StatusSpeakWith = "HRD"
        Case "TEMAN KANTOR"
            StatusSpeakWith = "FRND"
        Case "ORANG TUA"
            StatusSpeakWith = "PRNT"
        Case "KAKAK/ADIK/ANAK"
            StatusSpeakWith = "BSSD"
        Case "SPOUSE"
            StatusSpeakWith = "SPS"
        Case "KELUARGA DEKAT LAINNYA"
            StatusSpeakWith = "OFAM"
        Case "EX SPOUSE"
            StatusSpeakWith = "ESPS"
        Case "PEMBANTU/SUPIR"
            StatusSpeakWith = "MAID"
        Case "OTHER"
            StatusSpeakWith = "OTH"
        Case "TETANGGA"
            StatusSpeakWith = "NGBR"
        Case "PENGURUS LINGKUNGAN"
            StatusSpeakWith = "RTRW"
        Case "KONTRAKAN"
            StatusSpeakWith = "KNTK"
        Case "LAWYER"
            StatusSpeakWith = "LAW"
        Case "TEMAN"
            StatusSpeakWith = "FRND"
        Case "TEMAN KANTOR"
            StatusSpeakWith = "FRND"
        Case "LAWYER"
            StatusSpeakWith = "LAW"
         Case "UNRECEIVE"
            StatusSpeakWith = "NRCV"
    End Select
End Sub

Private Sub cbolastcall_GotFocus()
'cbolastcall.CLEAR
'Dim M_OBJRS As ADODB.Recordset
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'If UCase(MDIForm1.Text2.Text) = "AGENT" Then
'    If Left(cmbContacted.Text, 2) = "OP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('SP-SETTLE PAYMENT','PTP-PROMISE TO PAY') "
'    ElseIf Left(cboPTP.Text, 3) = "PTP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('OP-ON PROGRESS','SP-SETTLE PAYMENT') "
'    Else
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('SP-SETTLE PAYMENT')"
'    End If
' Else
'    If Left(cmbContacted.Text, 2) = "OP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'PTP-PROMISE TO PAY' "
'    ElseIf Left(cboPTP.Text, 3) = "PTP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'OP-ON PROGRESS' "
'    Else
'    CMDSQL = " Select * from ContactedDesc"
'    End If
' End If
'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
End Sub

Private Sub cbolastcall_KeyDown(KeyCode As Integer, Shift As Integer)

cbolastcall.text = ""
Exit Sub
End Sub

Private Sub cboPOPSP_Click()
Dim M_COL1 As New ADODB.Recordset
If Left(cboPOPSP.text, 2) = "SP" Then
    C_Contacted.Value = 0
    C_SKIP.Value = 0
    C_PTP.Value = 0
    C_VALID.Value = 0
    CmbBaseOn.text = ""
    cmbDiscount.text = ""
    txtPayment.Value = 0
    Tdabamoint.Value = 0
    TDBDate3.Value = ""
    txttenor.Value = 0
    cmbDescCon.Enabled = False
    C_Payment.Value = 1
    FrmPayment.Enabled = True
            Set M_COL1 = New ADODB.Recordset
            cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
            M_COL1.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
            CmbBaseOn.text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(M_COL1!dateptp), "", M_COL1!dateptp))
            txttenor.Value = CStr(IIf(IsNull(M_COL1!Tenor), 0, M_COL1!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(M_COL1!amountptp), 0, M_COL1!amountptp))
End If

'C_Payment.Value = 0



'txtPayment.Value = 0

End Sub

Private Sub cboPOPSP_KeyDown(KeyCode As Integer, Shift As Integer)

cboPOPSP.text = ""
End Sub


Private Sub cboskip_Click()
cbodescskip.clear
If Left(cboskip.text, 2) <> "MV" Then
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   M_Objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For I = 0 To 3
           cbodescskip.AddItem M_Objrs("Description")
           M_Objrs.MoveNext
         Next I
   Set M_Objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
      M_Objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
       While Not M_Objrs.EOF
           cbodescskip.AddItem M_Objrs("Description")
           M_Objrs.MoveNext
       Wend
   Set M_Objrs = Nothing
   C_Payment.Value = 0
End If

End Sub

Private Sub cbovalid_Click()
Dim I As Integer
cbodescvalid.clear
If Left(cbovalid.text, 2) = "NA" Then
        cbodescvalid.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
          M_Objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_Objrs.EOF
            cbodescvalid.AddItem M_Objrs("Description")
            M_Objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_Objrs = Nothing
'        FrmPayment.Enabled = False
Else
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
          M_Objrs.Open "Select * from DescunContacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_Objrs.EOF
            cbodescvalid.AddItem M_Objrs("Description")
            M_Objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_Objrs = Nothing
End If

End Sub

Private Sub cbovalid_KeyDown(KeyCode As Integer, Shift As Integer)

cbovalid.text = ""
Exit Sub
End Sub



Private Sub cbolastcall_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPTP_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Check1_Click()
regnego = False
Check2.Value = 0
Check3.Value = 0
If CmbBaseOn.text = "PRINCIPLE" Then
    MsgBox "Regular payment only TOTAL AMOUNT"
    CmbBaseOn.SetFocus
    Exit Sub
Else
'Call CEKPTP
'If adaSCH Then
'    MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'    Exit Sub
'Else
    Call ISIJMLPAYMENT
    If Check1.Value = 1 Then
        frmregpayment.Show
    End If
End If
End Sub

Sub CEKPTP()
Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select TYPE from TBLNEGOPTP where custid='" & lblCustId.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
If rs.BOF And rs.EOF Then
Else
    While Not rs.EOF
        If rs!Type = "SCH" Then
            adaSCH = True
        ElseIf rs!Type = "REG" Then
            adaREG = True
        ElseIf rs!Type = "PO" Then
            adaPO = True
        End If
        rs.MoveNext
    Wend
End If
Set rs = Nothing
End Sub


Private Sub Check2_Click()
Check1.Value = 0
Check3.Value = 0
If Check2.Value = 1 Then
'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        MsgBox "Regular payment only TOTAL AMOUNT"
'        CmbBaseOn.SetFocus
'        Exit Sub
'    Else
'        Call CEKPTP
'        If adaREG Then
'            MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'            Exit Sub
'        Else
            'Call ISIJMLPAYMENT
            regnego = True
            FrmNegoPTP.Show
'        End If
End If
'End If
End Sub

Private Sub Check3_Click()
regnego = False
Check1.Value = 0
Check2.Value = 0

'Call CEKPTP
'If adaPO Then
'    MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'    Exit Sub
'Else
    Call ISIJMLPAYMENT
    If Check3.Value = 1 Then
        Frmpaidoff.Show
    End If
'End If
End Sub

Private Sub chkAppv_Click(Index As Integer)
'Select Case Index
'Case 0:
'    chkAppv(1).Value = 0
'Case 1:
'    chkAppv(0).Value = 0
'End Select
End Sub

Private Sub Chktenor_Click()
If Chktenor.Value = 1 Then
    txttenor.Enabled = True
    txttenor.BackColor = vbWhite
Else
    txttenor.Enabled = False
    txttenor.BackColor = &H4000&
    Chktenor.Value = 0
    txttenor.Value = 0
End If


End Sub

Private Sub CmbBaseOn_Click()
If CmbBaseOn.text = "PRINCIPLE" Then
CmbBaseOn.text = ""
End If
    Call cmbDiscount_Click
End Sub

Private Sub CmbBaseOn_LostFocus()
    'Call cmbDiscount_Click
End Sub

Private Sub cmbContacted_Click()
'DESCRIPTION CONTACTED
Dim I As Integer
cmbDescCon.clear

'If Left(vrcek, 2) = "BP" And Left(cmbContacted.Text, 3) = "POP" Then
'    cmbContacted.Text = ""
'End If

If Left(cmbContacted.text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.text = ""
    txtPayment.text = 0
    cmbDiscount.text = ""
    TdbPTP.text = ""
    TdbDatePTP.text = ""
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
     M_Objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_Objrs.EOF
        cmbDescCon.AddItem M_Objrs("Description")
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    C_Payment.Value = 0
   ' FrmPayment.Enabled = False
    Else
'    If Left(cmbContacted.Text, 2) = "NA" Then
'        cmbDescCon.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescCon.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        FrmPayment.Enabled = False
        
'    Else
         If Left(cmbContacted.text, 2) = "PT" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
            CmbBaseOn.text = "PRINCIPLE"
    Else
        If Left(cmbContacted.text, 2) = "BP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.text = ""
            txtPayment.text = 0
            cmbDiscount.text = ""
            TdbPTP.text = ""
            TdbDatePTP.text = ""
            C_Payment.Value = 0
           ' FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.text, 2) = "OP" Then
            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
          '  C_Payment.Value = 1
             'C_Payment.Value = False
            FrmPayment.Enabled = True
      Else
      
    If Left(cmbContacted.text, 2) = "PO" Or Left(cmbContacted.text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
        Set m_cust = New ADODB.Recordset
        m_cust.CursorLocation = adUseClient
        cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor, amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
        m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
           CmbBaseOn.text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(m_cust!dateptp), "", m_cust!dateptp))
            txttenor.Value = CStr(IIf(IsNull(m_cust!Tenor), "0", m_cust!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp))
            
      Set m_cust = Nothing
    End If
End If
End If
End If
End If
'End If

Set M_Objrs = Nothing
End Sub

Private Sub cmbContacted_KeyDown(KeyCode As Integer, Shift As Integer)

cmbContacted.text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_GotFocus()
'DESCRIPTION CONTACTED
Dim I As Integer
cmbDescCon.clear
If Left(cmbContacted.text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.text = ""
    txtPayment.text = 0
    cmbDiscount.text = ""
    TdbPTP.text = ""
    TdbDatePTP.text = ""
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
     M_Objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_Objrs.EOF
        cmbDescCon.AddItem M_Objrs("Description")
        M_Objrs.MoveNext
    Wend
    C_Payment.Value = 0
   ' FrmPayment.Enabled = False
    Set M_Objrs = Nothing
    Else
'    If Left(cmbContacted.Text, 2) = "NA" Then
'        cmbDescCon.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescCon.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        FrmPayment.Enabled = False
        
'    Else
         If Left(cmbContacted.text, 2) = "PT" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
            CmbBaseOn.text = "PRINCIPLE"
    Else
        If Left(cmbContacted.text, 2) = "BP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.text = ""
            txtPayment.text = 0
            cmbDiscount.text = ""
            TdbPTP.text = ""
            TdbDatePTP.text = ""
            C_Payment.Value = 0
'            FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.text, 2) = "OP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.text = ""
            txtPayment.text = 0
            cmbDiscount.text = ""
            TdbPTP.text = ""
            TdbDatePTP.text = ""
            C_Payment.Value = 0
           ' FrmPayment.Enabled = False
      Else
      
    If Left(cmbContacted.text, 2) = "PO" Or Left(cmbContacted.text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
            CmbBaseOn.text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
            
      Set m_cust = Nothing
    End If
End If
End If
End If
End If
'End If

Set M_Objrs = Nothing
End Sub

Private Sub cmbDescCon_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescCon.text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cmbDescUn_GotFocus()
Dim I As Integer
cmbDescUn.clear
If Left(cmbUncontacted.text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
          M_Objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_Objrs.EOF
            cmbDescUn.AddItem M_Objrs("Description")
            M_Objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_Objrs = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.text, 2) <> "MV" Then
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   M_Objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For I = 0 To 3
           cmbDescUn.AddItem M_Objrs("Description")
           M_Objrs.MoveNext
         Next I
   Set M_Objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_Objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_Objrs.EOF
           cmbDescUn.AddItem M_Objrs("Description")
           M_Objrs.MoveNext
       Wend
   Set M_Objrs = Nothing
   C_Payment.Value = 0
End If
End If
End Sub

Private Sub cmbDescUn_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescUn.text = ""
Exit Sub
End Sub

Private Sub cmbDiscount_Change()
Call ISIJMLPAYMENT
End Sub

Private Sub cmbDiscount_Click()
Call ISIJMLPAYMENT
'Check1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'If Left(cmbContacted.Text, 2) = "OP" Then
'    Check1.Enabled = False
'    Check3.Enabled = False
'End If
End Sub

Sub ISIJMLPAYMENT()
Dim M_Objrs As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If

M_Objrs.CursorLocation = adUseClient
M_Objrs.Open "Select * from tbldiscount where Description = '" + cmbDiscount.text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_Objrs.RecordCount <> 0 Then
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_Objrs!hari), 7, M_Objrs!hari)
Else
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
End If
If cmbDiscount.text = "0" Or cmbDiscount.text = "" Then
    If CmbBaseOn.text = "PRINCIPLE" Then
        txtPayment.Value = LblPrompA.Value
    Else
    
         txtPayment.Value = lblAmount.Value
         Exit Sub
         
'         If CmbBaseOn.Text = "TOTAL AMOUNT" Then
'            If lblAmount.Value = 0 Or lblAmount.ValueIsNull Or cmbDiscount = "" Then
'                txtPayment.Value = 0
'            Else
'                txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'                txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'            End If
'        End If
    End If
End If

        If CmbBaseOn.text = "TOTAL AMOUNT" Then
            If lblAmount.Value = 0 Or lblAmount.ValueIsNull Or cmbDiscount = "" Then
                txtPayment.Value = 0
            Else
                txtdiscount.text = CStr((cmbDiscount.text) / 100)
                txtPayment.Value = lblAmount.Value - (CCur(txtdiscount.text) * lblAmount.Value)
                End If

                
            End If
       ' End If

'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        If lblPromPA.Value = 0 Or lblPromPA.ValueIsNull Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblPromPA.Value - (CCur(txtDiscount.Text) * lblPromPA.Value)
'        End If
'    Else
'        If lblAmount.Value = 0 Or lblAmount.ValueIsNull Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'        End If
'    End If
'End If
'End If

End Sub

Private Sub cmbDiscount_LostFocus()
'Dim M_OBJRS As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If
'
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tbldiscount where Description = '" + cmbDiscount.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'If M_OBJRS.RecordCount <> 0 Then
'    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_OBJRS!hari), 7, M_OBJRS!hari)
'Else
'    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
'End If
'If cmbDiscount.Text = "0" Then
'Else
'
'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        If lblPromPA.Value = 0 Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblPromPA.Value - (CCur(txtDiscount.Text) * lblPromPA.Value)
'        End If
'    Else
'        If lblAmount.Value = 0 Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'        End If
'    End If
'End If
End Sub

Private Sub cmbNextAct_KeyDown(KeyCode As Integer, Shift As Integer)
cmbNextAct.text = ""
Exit Sub
End Sub

Private Sub CmbPhone_Click()
    CmbPhone.Locked = True
    If CmbPhone.text = "Add" Then
        Frm_Tambah_Telp.Show vbModal
    End If
End Sub

Private Sub CmbPhone_DropDown()
    CmbPhone.Locked = False
End Sub

Private Sub CmbPhone_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbUncontacted_Click()
'DESCRIPTION UNCONTACTED
Dim I As Integer
cmbDescUn.clear
If Left(cmbUncontacted.text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
          M_Objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_Objrs.EOF
            cmbDescUn.AddItem M_Objrs("Description")
            M_Objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_Objrs = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.text, 2) <> "MV" Then
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   M_Objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For I = 0 To 3
           cmbDescUn.AddItem M_Objrs("Description")
           M_Objrs.MoveNext
         Next I
   Set M_Objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_Objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_Objrs.EOF
           cmbDescUn.AddItem M_Objrs("Description")
           M_Objrs.MoveNext
       Wend
   Set M_Objrs = Nothing
   C_Payment.Value = 0
End If
End If
' Set M_OBJRS = New ADODB.Recordset
'   If kontak = False Then
'          M_OBJRS.Open "Select * from UncontactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'       While Not M_OBJRS.EOF
'           cmbDescUn.AddItem M_OBJRS("NMnoProdpresented")
'           M_OBJRS.MoveNext
'       Wend
'        Set M_OBJRS = Nothing
'   End If
'   C_Payment.Value = 0
'End If

End Sub

Private Sub headerDatePayment()
LstPayment.ColumnHeaders.ADD 1, , "", 0 * TXT
LstPayment.ColumnHeaders.ADD 2, , "ID", 1
LstPayment.ColumnHeaders.ADD 3, , "DATE", 1100
LstPayment.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstPayment.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstPayment.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT

LstReserve.ColumnHeaders.ADD 1, , "", 0 * TXT
LstReserve.ColumnHeaders.ADD 2, , "ID", 1
LstReserve.ColumnHeaders.ADD 3, , "DATE", 1100
LstReserve.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstReserve.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstReserve.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT

End Sub
Private Sub headerCustid_Double()
    LstDoubleId.ColumnHeaders.ADD 1, , "Id", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 2, , "Nama", 15 * TXT
    LstDoubleId.ColumnHeaders.ADD 3, , "DescColl", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 4, , "AmountWo", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 5, , "Principle", 20 * TXT
End Sub
Private Sub cmbUncontacted_KeyDown(KeyCode As Integer, Shift As Integer)
cmbUncontacted.text = ""
Exit Sub
End Sub
Private Sub Cmbwith_KeyDown(KeyCode As Integer, Shift As Integer)
Cmbwith.text = ""
Exit Sub
End Sub








Private Sub CmbStsKatHome1_Click()
    StsKategoriTelepon = Trim(CmbStsKatHome1.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHome1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub CmbStsKatHome2_Click()
    StsKategoriTelepon = Trim(CmbStsKatHome2.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHome2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub





Private Sub CmbStsKatHP1_Click()
    StsKategoriTelepon = Trim(CmbStsKatHP1.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHP1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub CmbStsKatHP2_Click()
    StsKategoriTelepon = Trim(CmbStsKatHP2.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHP2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbStsKatOffice1_Click()
    StsKategoriTelepon = Trim(CmbStsKatOffice1.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatOffice1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub CmbStsKatOffice2_Click()
    StsKategoriTelepon = Trim(CmbStsKatOffice2.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatOffice2_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub CmbViaPtp_Click()
    If UCase(Trim(CmbViaPtp.text)) = "ATM LAINNYA" Then
        FrmPilihanAtmLainnya.Show vbModal
    End If
     '@@09-04-2012
    CariTanggalTagih
End Sub

Private Sub CmbViaPtp_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub

Private Sub cmd_logcomplaint_Click()
    With Form_complaint
        .txt_custid.text = lblCustId.Caption
        .txt_custname.text = lblNama.Caption
        .txt_agent.text = lblaoc.Caption
        .Frame2.Enabled = False
        .cb_status.text = "OPEN"
        .lbl_ket = "N"
        .Show 1
    End With
End Sub

Private Sub CmdClaimAcc_Click()
    If UCase(lblaoc.Caption) <> "AKSESALL" Then
        MsgBox "Claim account hanya diperuntukkan bagi account yang di collect secara bersama-sama!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    Else
        'Pindahkan status account ke user claim
        FrmClaimAccount.TxtCustid.text = lblCustId.Caption
        FrmClaimAccount.TxtNama.text = lblNama.Caption
        FrmClaimAccount.Show vbModal
    End If
End Sub

Private Sub CmdDataMapping_Click()
    '@@ 30-03-2012 Data Mapping dinonaktifkan, udah jarang dipake
    'FrmDataMapping.Show vbModal
    
    Dim a As String
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim M_Objrs_Cek As ADODB.Recordset
    
    a = MsgBox("Apakah anda akan membuat account ini sebagai Kept account untuk anda?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbYes Then
        'cek data dulu
        cmdsql = "select * from tbl_keep_acc where date_part('year',tglkeep)="
        cmdsql = cmdsql + "date_part('year',now()) and date_part('month',tglkeep)="
        cmdsql = cmdsql + "date_part('month',now()) and agent='"
        cmdsql = cmdsql + lblaoc.Caption + "'"
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs.RecordCount >= 10 Then
           MsgBox "Account keep anda sudah lebih mencapai 10 account. Maksimal account keep 10!", vbOKOnly + vbInformation, "Informasi"
        Else
            
            'Cek apakah custid ini sudah termasuk keep account
            cmdsql = "select * from tbl_keep_acc where date_part('year',tglkeep)="
            cmdsql = cmdsql + "date_part('year',now()) and date_part('month',tglkeep)="
            cmdsql = cmdsql + "date_part('month',now()) and agent='"
            cmdsql = cmdsql + lblaoc.Caption + "' and custid='"
            cmdsql = cmdsql + lblCustId.Caption + "'"
            Set M_Objrs_Cek = New ADODB.Recordset
            M_Objrs_Cek.CursorLocation = adUseClient
            M_Objrs_Cek.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_Cek.RecordCount > 0 Then
                MsgBox "Account ini sudah di keep sebelumnya!", vbOKOnly + vbInformation, "Informasi"
                Set M_Objrs_Cek = Nothing
                Exit Sub
            End If
            
            Set M_Objrs_Cek = Nothing
            
            cmdsql = "insert into tbl_keep_acc (custid,agent,tglkeep,nama) values ('"
            cmdsql = cmdsql + lblCustId.Caption + "','"
            cmdsql = cmdsql + lblaoc.Caption + "','"
            cmdsql = cmdsql + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
            cmdsql = cmdsql + lblNama.Caption + "')"
            M_OBJCONN.execute cmdsql
            
            'Update juga di tabel mgm
            cmdsql = "update mgm set status_keep='1' where custid='"
            cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
            M_OBJCONN.execute cmdsql
            MsgBox "Keep account anda berhasil!", vbOKOnly + vbInformation, "Informasi"
        End If
        Set M_Objrs = Nothing
    End If
End Sub



Private Sub CmddetailPayment_Click()
    FrmDetailPayment.Show 1
End Sub

'@@ 05-10-2011, Penghapusan data di tabel lunas
'Private Sub CmdDeletePelunasan_Click()
'Dim m_msgbox As Variant
'If listview1(0).ListItems.Count = 0 Then
'    Exit Sub
'End If
'm_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
'If m_msgbox = vbOK Then
'    M_OBJCONN.Execute "Delete from tbllunas where id = " + listview1(0).SelectedItem.SubItems(4) + ""
'    listview1(0).ListItems.Remove listview1(0).SelectedItem.Index
'    MsgBox "Done"
'    Call isi_datapayment
'End If
'End Sub


Private Sub CmdHapusRemarks_Click()
    Dim cmdsql As String
    Dim a As String
    
    If listview1(1).ListItems.Count = 0 Then
        MsgBox "Tidak ada data remarks yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data: " & listview1(1).SelectedItem.SubItems(1) & " akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    cmdsql = "delete from mgm_hst where id='"
    cmdsql = cmdsql + Trim(listview1(1).SelectedItem.SubItems(7)) + "'"
    
    M_OBJCONN.execute cmdsql
    
    listview1(1).ListItems.Remove listview1(1).SelectedItem.Index
End Sub

Private Sub CmdKeep_Click()
 Dim cmdsql As String
 Dim M_Objrs As ADODB.Recordset
 Dim a As String
 
 cmdsql = "select * from mgm where custid='"
 cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
 Set M_Objrs = New ADODB.Recordset
 M_Objrs.CursorLocation = adUseClient
 M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
 If M_Objrs.RecordCount = 0 Then
    Set M_Objrs = Nothing
    Exit Sub
 End If
 
 If M_Objrs("status_htc") = "1" Then
    a = MsgBox("Apakah anda yakin akan menghilangkan status account ini tidak menjadi Hot Prospect?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        cmdsql = "update mgm set status_htc=null where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
        M_OBJCONN.execute cmdsql
        MsgBox "Status Hot Prospect untuk account ini telah dihilangkan!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    '@@ 03-04-2012, Tanyakan ke user, apakah ingin menghapus data ini sebagai keep account juga?
    a = MsgBox("Apakah anda juga akan menghapus Kept Account untuk CH ini?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        cmdsql = "delete from tbl_keep_acc where date_part('year',tglkeep)="
        cmdsql = cmdsql + "date_part('year',now()) and date_part('month',tglkeep)="
        cmdsql = cmdsql + "date_part('month',now()) and agent='"
        cmdsql = cmdsql + Trim(lblaoc.Caption) + "' and custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
        M_OBJCONN.execute cmdsql
        
        'Update status keep di mgm
        cmdsql = "update mgm set status_keep=null where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
        M_OBJCONN.execute cmdsql
        
        MsgBox "Kept Account untuk CH ini sudah dihapus!", vbOKOnly + vbInformation, "Informasi"
    End If
 End If
 
 If IsNull(M_Objrs("status_htc")) = True Then
    a = MsgBox("Apakah anda yakin akan  menjadikan account ini  menjadi Hot Prospect?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        cmdsql = "update mgm set status_htc='1' where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
        M_OBJCONN.execute cmdsql
        MsgBox "Status Hot Prospect telah ditandai dalam account ini!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    CmdDataMapping_Click
 End If
 
 
End Sub

Private Sub CmdOther_Click()
    FrmOther.Show vbModal
End Sub

Private Sub CmdRequest_Click()
'    '@@ 07-04-2011 Tambahan bikin Form Request
'    With Frm_Request
'        .TxtAgent.Text = lblaoc.Caption
'        .TxtCustid.Text = lblCustId.Caption
'        .TxtNamaCH.Text = lblNama.Caption
'
'        .TXtAmountWoPUM.Value = TDB_cur_bal.Value
'        .TxtPaymentDatePUM.Text = Format(lblPayDt.Value, "yyyy-mm-dd")
'        .Show vbModal
'    End With
    
    FrmListKeepAcc.Show vbModal
End Sub

Private Sub CmdRequestNumber_Click()
    With FrmReqTelepon
        .TxtCustid.text = lblCustId.Caption
        .Show vbModal
    End With
End Sub




Private Sub CmdSendPTP_Click()
    FrmSendPTP.Show vbModal
End Sub

Private Sub CmdViewRecording_Click()
    '@@31012013 diganti jadi view recording
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        MsgBox "Mohon maaf anda tidak memiliki akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    FrmRecording.TxtCustid.text = lblCustId.Caption
    FrmRecording.Show vbModal
End Sub

'@@ 05-10-2011, Tombol Unlock ditiadakan
'Private Sub CmdUnlock_Click()
'    '@@ 01/02/2011 UnLock Data Oleh agent
'    Dim a As String
'    Dim ID As String
'    Dim M_OBJRS As ADODB.Recordset
'    Dim m_objrs_cekid As ADODB.Recordset
'    Dim CMDSQL As String
'    Dim UpdateDtCloseSession As String
'    Dim m_objrs_ambilTL As ADODB.Recordset
'    Dim cmdsql_ambilTL As String
'
'    Dim pesan As String
'    Dim TglLock As String
'    Dim StartLock As String
'    Dim EndLock As String
'    Dim AccLock As String
'    Dim Status_lock As String
'
'    'Cek dulu apakah yang login agent?
'    If UCase(Trim(MDIForm1.Text2.Text)) <> "AGENT" Then
'        MsgBox "Unlock data ini hanya untuk AGENT!", vbOKOnly + vbExclamation, "Peringatan"
'        Exit Sub
'    End If
'
'    a = MsgBox("Anda yakin akan melakukan Unlock Data?", vbYesNo + vbQuestion, "Konfirmasi")
'    If a = vbNo Then
'        Exit Sub
'    End If
'
'    'Cek apakah ada data yang sedang di lock?
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    CMDSQL = "select * from usertbl where userid='"
'    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "'"
'    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_OBJRS("lockdarispv") = "" And M_OBJRS("lock_entry_lpd") = "" And M_OBJRS("lockmarkup") = "" Then
'        MsgBox "Tidak ada data yang akan di unlock!", vbOKOnly + vbInformation, "Informasi"
'        Set M_OBJRS = Nothing
'        Exit Sub
'    End If
'    Set M_OBJRS = Nothing
'
'    'Cari id data yang sedang di lock
'    CMDSQL = "select *,now() as tanggal_sekarang from tbltemplockacc_current where id in "
'    CMDSQL = CMDSQL + "(select max(idlock) as idlock from tblperformpersessionlock where agent='"
'    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "')"
'
'    Set m_objrs_cekid = New ADODB.Recordset
'    m_objrs_cekid.CursorLocation = adUseClient
'    m_objrs_cekid.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    ID = Trim(m_objrs_cekid("id"))
'    TglLock = Format(m_objrs_cekid("date_lock"), "yyyy-mm-dd hh:mm:ss")
'    StartLock = Format(m_objrs_cekid("start_lock"), "yyyy-mm-dd hh:mm:ss")
'    EndLock = Format(m_objrs_cekid("end_lock"), "yyyy-mm-dd hh:mm:ss")
'    AccLock = Trim(IIf(IsNull(m_objrs_cekid("account_lock")), "", m_objrs_cekid("account_lock")))
'    Status_lock = Trim(m_objrs_cekid("status_lock"))
'
'
'    'Catat ke dalam log
'    CMDSQL = "insert into log_unlock_agent (script_lock,date_lock,"
'    CMDSQL = CMDSQL + "start_lock,end_lock,account_lock,lock_by,f_locked,tgl_unlock,agent_unlock,status_lock,id) values ('"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("script_lock")), "", m_objrs_cekid("script_lock"))) + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("date_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("start_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("end_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("account_lock")), "", m_objrs_cekid("account_lock"))) + "','"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("lock_by")), "", m_objrs_cekid("lock_by"))) + "','"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("f_locked")), "", m_objrs_cekid("f_locked"))) + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("tanggal_sekarang"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "','"
'    CMDSQL = CMDSQL + Trim(m_objrs_cekid("status_lock")) + "','"
'    CMDSQL = CMDSQL + Trim(ID) + "')"
'
'    M_OBJCONN.Execute CMDSQL
'
'    'Bikin pesan ke TL,jika lock datanya sudah di unlock oleh agent
'    pesan = vbCrLf + "INFORMASI OLEH SISTEM : " + vbCrLf
'    pesan = pesan + "Agent: " + MDIForm1.Text1.Text + vbCrLf
'    pesan = pesan + "Melakukan Unlock data untuk accountnya sendiri." + vbCrLf
'    pesan = pesan + "Berikut informasi lock data yang di unlock:" + vbCrLf
'    pesan = pesan + "------------------------------------------------" + vbCrLf
'    pesan = pesan + "Tgl.Lock data :" + StartLock + vbCrLf
'    pesan = pesan + "Start.Lock data:" + EndLock + vbCrLf
'    pesan = pesan + "Account yang di lock:" + AccLock + vbCrLf
'    pesan = pesan + "Status yang di lock:" + Status_lock + vbCrLf
'    pesan = pesan + "------------------------------------------------" + vbCrLf
'    pesan = pesan + "Terima Kasih" + vbCrLf
'    pesan = pesan + "Message Created automatic by system"
'
'    MsgBox "Silahkan tunggu sebentar! Setelah menekan tombol OK ini, sistem akan melakukan unlock data. Harap Tunggu hingga muncul pesan Unlock data berhasil!", vbOKOnly + vbInformation, "Informasi"
'
'    'Pindahkan data ke tabel tblperformpersessionlock
'    DoEvents
'    UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(m_objrs_cekid("tanggal_sekarang"), "yyyy-mm-dd hh:mm:ss")) + "' from "
'    UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
'    UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
'    UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
'    UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
'    UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
'    UpdateDtCloseSession = UpdateDtCloseSession + Trim(ID) + "' and tblperformpersessionlock.agent='"
'    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.Text1.Text) + "'"
'    M_OBJCONN.Execute UpdateDtCloseSession
'
'    Set m_objrs_cekid = Nothing
'
'    cmdsqlserver = "update usertbl set dilockoleh='Release by:" + Trim(MDIForm1.Text2.Text) + "',"
'    cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
'    cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null where userid='"
'    cmdsqlserver = cmdsqlserver + Trim(MDIForm1.Text1.Text) + "'"
'    M_OBJCONN.Execute cmdsqlserver
'
'    'Berikan pesan ke TL-nya
'    cmdsql_ambilTL = "select * from usertbl where userid='"
'    cmdsql_ambilTL = cmdsql_ambilTL + Trim(MDIForm1.Text1.Text) + "'"
'    Set m_objrs_ambilTL = New ADODB.Recordset
'    m_objrs_ambilTL.CursorLocation = adUseClient
'    m_objrs_ambilTL.Open cmdsql_ambilTL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    CMDSQL = "insert into msgtbl  (recipient, datetime, sender, sentfrom, msg) VALUES ('"
'    CMDSQL = CMDSQL + Trim(m_objrs_ambilTL("team")) + "','"
'    CMDSQL = CMDSQL + CStr(Format(Now, "yyyymmdd")) + "','"
'    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "','"
'    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
'    CMDSQL = CMDSQL + Trim(pesan) + "')"
'    M_OBJCONN.Execute CMDSQL
'
'    Set m_objrs_ambilTL = Nothing
'
'    MsgBox "Data anda berhasil di Unlock!", vbOKOnly + vbInformation, "Informasi"
'    VIEW_MGMDATA.listview1.ListItems.CLEAR
'End Sub

Private Sub Command1_Click()
     If Command1.tag = 0 Then
        Tdbbalance.Visible = True
        
        '@@ 0408201 Dibuang
        'tdbprincipal.Visible = True
        
        Label11(14).Visible = True
        
        '@@ 04082011 dibuang
        'Label11(15).Visible = True
        
        Command1.tag = 1
        LblPrompA.Visible = True
        Label11(8).Visible = True
        Else
        Tdbbalance.Visible = False
        tdbprincipal.Visible = False
        Label11(14).Visible = False
        
        '@@ 04082011 dibuang
        'Label11(15).Visible = False
        
        Label11(8).Visible = False
        Command1.tag = 0
        LblPrompA.Visible = False
        End If
        
End Sub

Private Sub Command2_Click()
'Load FrmSendSMS
'FrmSendSMS.Show vbModal
'@@ 09031011, diubah formnya
FrmInboXSms.Show vbModal
End Sub

Private Sub Command3_Click()
    If MsgBox("Account ini akan diset set menjadi decease??", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        ' DELETE BEFORE
        M_OBJCONN.execute "DELETE FROM tblreq_decease WHERE custid='" & CStr(Trim(lblCustId.Caption)) & "'"
        M_OBJCONN.execute "INSERT INTO tblreq_decease(custid,agent) VALUES('" & CStr(Trim(lblCustId.Caption)) & "','" & MDIForm1.Text1.text & "')"
        MsgBox "Account telah diset menjadi Acc Decease, Tunggu approval dari TL", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub Form_Load()

On Error Resume Next

m_EditHWnd = FindEditChild(cboaccount.hwnd)
OldWindowProc = SetWindowLong(m_EditHWnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)

' ## Set Status Form Customer Aktif 12 Mei 2013 By Izuddin
bAktif_form_customer = True
' # 08 April 2013 Monitoring Activity By Izuddin
i_monitoring_activity = 0
MDIForm1.Timer2.Enabled = True

StsKategoriTelepon = ""
KelompokKategoriTlp = ""

If UCase(MDIForm1.Text2) = "AGENT" Then
    SSCommand1(4).Visible = False
    Command1.Visible = False
    'Jika agent c_ptp didisable 11 Juni 2012
    C_PTP.Enabled = False

ElseIf UCase(MDIForm1.Text2) = "SUPERVISOR" Or UCase(MDIForm1.Text2) = "ADMIN" Or UCase(MDIForm1.Text2) = "ADMINISTRATOR" Then
        SSCommand1(4).Visible = True
        Command1.Visible = False
        CmdHapusRemarks.Visible = True
        cmd_logcomplaint.Visible = True
End If

'@@19042012, Tombol Hangup Di nonaktifkan dulu
SSCommand1(1).Enabled = False


FrmCC_Colection.Left = 10
FrmCC_Colection.Top = 20

'cek list pelunasan
Dim I, iIndex As Integer
Dim sKata, cCombo As String


'------->>>  setting No Visit  <<<---------------

Text1.text = Format(Now, "yymmddhhmmss")
TDBDate1.Value = Now
'If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "SUPER" Then
If UCase(Left(MDIForm1.Text2.text, 5)) = "ADMIN" Then
    txtHomeNo1.Visible = True
    txtHomeNo1A.Visible = False
    txtHomeNo2.Visible = True
    txtHomeNo2A.Visible = False
    txtOfficeNo1.Visible = True
    txtOfficeNo1A.Visible = False
    txtOfficeNo2.Visible = True
    txtOfficeNo2A.Visible = False
    txtMobileNo1.Visible = True
    txtMobileNo1A.Visible = False
    txtMobileNo2.Visible = True
    txtMobileNo2A.Visible = False
    txtPhone.Visible = True
    txtPhoneA.Visible = False
    txtHomeAdd1.Visible = True
    txtHomeAdd1A.Visible = False
    txtHomeAdd2.Visible = True
    txtHomeAdd2A.Visible = False
    txtOfficeAdd1.Visible = True
    txtOfficeAdd1A.Visible = False
    txtOfficeAdd2.Visible = True
    txtOfficeAdd2A.Visible = False
    txtMobileAdd1.Visible = True
    txtMobileAdd1A.Visible = False
    txtMobileAdd2.Visible = True
    txtMobileAdd2A.Visible = False
    txtECno.Visible = True
    txtECnoA.Visible = False
    Tdbbalance.Visible = False
        '@@ 0408201 Dibuang
        'tdbprincipal.Visible = False
        
        Label11(14).Visible = False
        
        '@@ 04082011 Dibuang
        'Label11(15).Visible = False
        
        'aktifkan recsource @@ 160610
        label1(80).Visible = True
        lblRecsource.Visible = True
End If

If UCase(MDIForm1.Text2.text) = "AGENT" Then
        C_lunas.Enabled = False
        TdbLunas.Enabled = False
        'chkAppv(0).Enabled = False '@@25/01/2012 Buangin komponen tak terpakai 25012012
        'chkAppv(1).Enabled = False '@@25/01/2012 Buangin komponen tak terpakai 25012012
        TDBTot_payment.Enabled = False
        TxtFieldName.Enabled = False
        
        '@@ 05-10-2011 Tombol Hapus Tabel Lunas ditiadakan terlebih dahulu
        'CmdDeletePelunasan.Enabled = False
         
         ' Tampilkan PRincipal
        
        SSCommand2(3).Enabled = False
        SSCommand2(2).Enabled = False
        
        lblhapus.Enabled = False
        Label41.Enabled = False
        LblPrompA.Visible = True
        Label11(8).Visible = True
        Tdbbalance.Visible = False
        '@@ 0408201 Dibuang
        'tdbprincipal.Visible = False
        
        Label11(14).Visible = False
        
        '@@ 04082011 Dibuang
        'Label11(15).Visible = False
ElseIf UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        txtHomeAdd1.ReadOnly = False
        txtHomeAdd2.ReadOnly = False
        txtOfficeAdd1.ReadOnly = False
        txtOfficeAdd2.ReadOnly = False
        txtMobileAdd1.ReadOnly = False
        txtMobileAdd2.ReadOnly = False
        '@@ 06-01-2012 , Tombol Delete Reserved PTP untuk TL dibuka
        SSCommand2(3).Enabled = True
        SSCommand2(2).Enabled = True
        lblhapus.Enabled = False
        Label41.Enabled = False
        Command1.Visible = False
         ' Tampilkan PRincipal
        LblPrompA.Visible = True
        Label11(8).Visible = True
        Tdbbalance.Visible = False
        '@@ 0408201 Dibuang
        'tdbprincipal.Visible = False
        
        Label11(14).Visible = False
        
        '@@ 04082011 Dibuang
        'Label11(15).Visible = False
       
Else ' utk SPV tampilkan no telp
        txtHomeAdd1.ReadOnly = False
        txtHomeAdd2.ReadOnly = False
        txtOfficeAdd1.ReadOnly = False
        txtOfficeAdd2.ReadOnly = False
        txtMobileAdd1.ReadOnly = False
        txtMobileAdd2.ReadOnly = False
        SSCommand2(3).Enabled = True
        SSCommand2(2).Enabled = True
        lblhapus.Enabled = True
        Label41.Enabled = True
        
        txtHomeNo1.Visible = True
        txtHomeNo1A.Visible = False
        txtHomeNo2.Visible = True
        txtHomeNo2A.Visible = False
        
        txtOfficeNo1.Visible = True
        txtOfficeNo1A.Visible = False
        
        txtOfficeNo2.Visible = True
        txtOfficeNo2A.Visible = False
        
        txtMobileNo1.Visible = True
        txtMobileNo1A.Visible = False
        txtMobileNo2.Visible = True
        txtMobileNo2A.Visible = False
        
        txtECno.Visible = True
        txtECnoA.Visible = False
        
        txtHomeAdd1.Visible = True
        txtHomeAdd1A.Visible = False
        txtHomeAdd2.Visible = True
        txtHomeAdd2A.Visible = False
        
        txtOfficeAdd1.Visible = True
        txtOfficeAdd1A.Visible = False
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
        
        txtMobileAdd1.Visible = True
        txtMobileAdd1A.Visible = False
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
        ' Tampilkan PRincipal
        LblPrompA.Visible = True
        Label11(8).Visible = True
        'aktifkan recsource @@ 160610
        label1(80).Visible = True
        lblRecsource.Visible = True
        
End If
 
 '  FrmContacted.Enabled = False
   FrmUnContacted.Enabled = False
   'FrmPayment.Enabled = False
   
    Call headerDatePayment
    Call headerCustid_Double
    Call HEADER_HISTORY
    Call HEADER_HISTORY_PAID
    Call HEADER_RequestVisit
    'Call HEADER_SendSMS
    On Error Resume Next
    Call show_cust
    
    '@@ 05-06-2012, Jika Status Complain dan Paid OFF maka kategori telepon tidak dapat dipilih
    If StatusAccount = "CO-" Or StatusAccount = "PO-" Then
        CmbStsKatHome1.Enabled = False
        CmbStsKatHome2.Enabled = False
        CmbStsKatOffice1.Enabled = False
        CmbStsKatOffice2.Enabled = False
        CmbStsKatHP1.Enabled = False
        CmbStsKatHP2.Enabled = False
        CmdRequestNumber.Enabled = False
     End If
    
    Call VisitNo
'    Call isi_lastcall
    
    If UCase(MDIForm1.Text2.text) = "TEAMLEADER" Or UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Then
        Call aktifphone
    End If
    
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        Call aktifphoneAGENT
    End If
    
    '@@14031011
    Call CekSms
    
    '@@ 08032011 Cek Data Mapping
    Call CekDataMapping
        
  '  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
SSTab1.Tab = 0
cmbDateSch.Value = Now
cmbDateSch.Value = ""
'CONTACTED
CmbBaseOn.AddItem "PRINCIPLE"
CmbBaseOn.AddItem "TOTAL AMOUNT"


'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblvalid", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbovalid.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'    Set M_OBJRS = Nothing
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
M_Objrs.Open "Select * from tblptp where KdNoProdPresented not like 'PTP-PAID%' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_Objrs.EOF
        cboPTP.AddItem M_Objrs!KdNoProdPresented
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblskip", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cboskip.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'    Set M_OBJRS = Nothing

    
    
    
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "Select * from popspdesc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            cboPOPSP.AddItem M_Objrs!KdNoProdPresented
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
''M_OBJRS.Open "Select * from ContactedDesc where KdNoProdPresented not like 'ptp%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'If UCase(MDIForm1.Text2.Text) = "AGENT" Then
'M_OBJRS.Open "Select * from contacteddesc where KdNoProdPresented not like 'ptp%' and KdNoProdPresented <>'SP-SETTLE PAYMENT' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'Else
'M_OBJRS.Open "Select * from contacteddesc where KdNoProdPresented not like 'ptp%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
'    While Not M_OBJRS.EOF
'    '----tambahan 05 Maret 2007----'
'         scombo = M_OBJRS("KdNoProdPresented")
'            sKata = cmbContacted.Text
'            ' initialisasi index
'            If scombo = "BP-BROKEN PROMISE" Or scombo = "PTP-PROMISE TO PAY" Or scombo = "RP-REFUSE PAYMENT" Then
'                  iIndex = 1
'            ElseIf scombo = "POP-PROGRESS OF PAYMENT" Then
'                  iIndex = 2
'            ElseIf scombo = "SP-SETTLE PAYMENT" Then
'                  iIndex = 3
'            Else
'                  iIndex = 4
'            End If
'
'            ' saring tampilan
'            If iIndex = 1 Then
'               If iIndex = 4 Or sKata = "POP-PROGRESS OF PAYMENT" Or sKata = "SP-SETTLED PAYMENT" Then
'                  'lewat boo
'               Else
'                    If scombo = "BP-BROKEN PROMISE" And UCase(MDIForm1.Text2.Text) = "AGENT" Then
'                    Else
'                        cmbContacted.AddItem scombo
'                    End If
'               End If
'            ElseIf iIndex = 2 Then
'               If iIndex = 1 Or iIndex = 4 Or Left(sKata, 2) = "SP" Then
'                  'lewat boo
'               Else
'                  cmbContacted.AddItem scombo
'               End If
'            ElseIf iIndex = 3 Then
'                If UCase(MDIForm1.Text2.Text) = "AGENT" Then
'                Else
'                  cmbContacted.AddItem scombo
'                End If
'            Else
'                  If sKata = "BP-BROKEN PROMISE" Or sKata = "PTP-PROMISE TO PAY" Or sKata = "POP-PROGRESS OF PAYMENT" Or sKata = "SP-SETTLED PAYMENT" Then
'                     'lewat boo
'                  Else
'                     cmbContacted.AddItem scombo
'                  End If
'            End If
'            M_OBJRS.MoveNext
'    Wend
'
'
'Set M_OBJRS = Nothing

'If Left(cmbContacted.Text, 2) = "SP" Then
'    'C_Contacted.Enabled = False
'    'cmbContacted.Enabled = False
'    C_NotContacted.Enabled = False
'End If

'If Left(cmbContacted.Text, 3) = "POP" Then
'    'C_Contacted.Enabled = False
'    'cmbContacted.Enabled = False
'    C_NotContacted.Enabled = False
'End If

'UNCONTACTED
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
''If kontak = True Then
''    M_OBJRS.Open "Select * from UnContactedDesc where KdNoProdPresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
''Else
''    M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
''End If
'If kontak = True Then
'    M_OBJRS.Open "Select * from uncontacteddesc where kdnoprodpresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
'ElseIf Left(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8), 2) = "NA" Then
'    M_OBJRS.Open "Select * from uncontacteddesc  where kdnoprodpresented  IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
'Else
'    M_OBJRS.Open "Select * from uncontacteddesc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
'    While Not M_OBJRS.EOF
'        cmbUncontacted.AddItem M_OBJRS("KdNoProdPresented")
'        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing

'Set M_OBJRS = New ADODB.Recordset
'If kontak = True Then
'    C_NotContacted.Enabled = False
'Else
'    M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbUncontacted.AddItem M_OBJRS("KdNoProdPresented")
'        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
'        M_OBJRS.MoveNext
'    Wend
'End If
'Set M_OBJRS = Nothing




'DISCOUNT

'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.Open "Select * from tblDiscount", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbDiscount.AddItem M_OBJRS("Description")
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing

'NEXT ACTION
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from stsnextact", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not M_OBJRS.EOF
'    cmbNextAct.AddItem M_OBJRS("NmStsNextAct")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
    'untuk 108
    
'@@ 24 May 2012 Akses 108, untuk agent tertentu saja
Dim M_objrs_108 As ADODB.Recordset
cmdsql = "select sts_108 from usertbl where userid='"
cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "' and sts_108='1'"
Set M_objrs_108 = New ADODB.Recordset
M_objrs_108.CursorLocation = adUseClient
M_objrs_108.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs_108.RecordCount > 0 Then
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_Objrs.EOF
        CmbPhone.AddItem IIf(IsNull(M_Objrs("Nolayanan")), "", M_Objrs("Nolayanan"))
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End If
Set M_objrs_108 = Nothing

'@@25052012 Jika yang login Admin,Superviso
If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or _
   UCase(MDIForm1.Text2.text) = "ADMIN" Or _
   UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Then
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_Objrs.EOF
        CmbPhone.AddItem IIf(IsNull(M_Objrs("Nolayanan")), "", M_Objrs("Nolayanan"))
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End If

'sembunyiin principle kecuali SPV
If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
    LblPrompA.Visible = False
    Label11(8).Visible = False
Else
    LblPrompA.Visible = True
    Label11(8).Visible = True
End If

'@@ 23-11-10 ini tambahan buat sembunyikan/tampilkan tombol ost jika ada data
'Dim M_OBJRS_ost As New ADODB.Recordset
'Set M_OBJRS_ost = New ADODB.Recordset
'M_OBJRS_ost.CursorLocation = adUseClient
'M_OBJRS_ost.Open "SELECT * FROM opening_screen where name like '%" + Trim(FrmCC_Colection.lblNama.Caption) + "%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'If M_OBJRS_ost.RecordCount <> 0 Then
'    SSCommand1(7).Visible = True
'Else
'    SSCommand1(7).Visible = True
'End If
'Set M_OBJRS_ost = Nothing

'@@ 15-04-2011 Panggil CekCPA, jika ada data CPA maka kelap-kelip
Call CekCPA

'@@ 25-07-2011, OfferingDiscGuide tampil
'Call OfferingDiscGuide

'@@ 09092011 Form Offering
Call OfferingDiscGuideNew

    '@@11 Juni 2012 Jika Yang Login Agent maka form PTP disable
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        frmPTP.Enabled = False
    End If
End Sub

Sub isi_lastcall()
cbolastcall.clear
Dim M_Objrs As ADODB.Recordset
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

If MDIForm1.Text2.text = "AGENT" Then
    M_Objrs.Open "Select * from ContactedDesc where kdnoprodpresented <> 'SP-SETTLE PAYMENT' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Else
    M_Objrs.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If
While Not M_Objrs.EOF
    cbolastcall.AddItem Trim(M_Objrs("KdNoProdPresented"))
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_Objrs.EOF
    cbolastcall.AddItem Trim(M_Objrs("KdNoProdPresented"))
    M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
End Sub

Private Sub aktifphone()
'@@03-05-2012 DinonAktifkan
'AHomeAdd1(0).ReadOnly = False
'@@03-05-2012 Dinonaktifkan
'AHomeAdd2(1).ReadOnly = False

txtHomeAdd1.ReadOnly = False
txtHomeAdd1A.ReadOnly = False
txtHomeAdd2.ReadOnly = False
txtHomeAdd2A.ReadOnly = False

'@@03-05-2012 Dinonaktifkan
'AOfficeAdd(2).ReadOnly = False
'AOfficeAdd(3).ReadOnly = False

txtOfficeAdd1.ReadOnly = False
txtOfficeAdd1A.ReadOnly = False
txtOfficeAdd2.ReadOnly = False
txtOfficeAdd2A.ReadOnly = False
txtMobileAdd1.ReadOnly = False
txtMobileAdd1A.ReadOnly = False
txtMobileAdd2.ReadOnly = False
txtMobileAdd2A.ReadOnly = False

'txtECno.ReadOnly = False
'txtECnoA.ReadOnly = False
'@@11052012 EC dinonaktifkan
txtECno.ReadOnly = True
txtECnoA.ReadOnly = True
End Sub

Private Sub aktifphoneAGENT()
If txtHomeAdd1.Value = "" Then
    txtHomeAdd1.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd1(0).ReadOnly = False
End If
If txtHomeAdd1A.Value = "" Then
    txtHomeAdd1A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd1(0).ReadOnly = False
End If
If txtHomeAdd2.Value = "" Then
    txtHomeAdd2.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd2(1).ReadOnly = False
End If
If txtHomeAdd2A.Value = "" Then
    txtHomeAdd2A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd2(1).ReadOnly = False
End If
If txtOfficeAdd1.Value = "" Then
    txtOfficeAdd1.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(2).ReadOnly = False
End If
If txtOfficeAdd1A.Value = "" Then
    txtOfficeAdd1A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(2).ReadOnly = False
End If
If txtOfficeAdd2.Value = "" Then
    txtOfficeAdd2.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(3).ReadOnly = False
End If
If txtOfficeAdd2A.Value = "" Then
    txtOfficeAdd2A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(3).ReadOnly = False
End If
If txtMobileAdd1.Value = "" Then
    txtMobileAdd1.ReadOnly = False
End If
If txtMobileAdd1A.Value = "" Then
    txtMobileAdd1A.ReadOnly = False
End If
If txtMobileAdd2.Value = "" Then
    txtMobileAdd2.ReadOnly = False
End If
If txtMobileAdd2A.Value = "" Then
    txtMobileAdd2A.ReadOnly = False
End If
If txtECno.Value = "" Then
    txtECno.ReadOnly = True
End If
If txtECnoA.Value = "" Then
    txtECnoA.ReadOnly = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim n%
    For n = 1 To LstPayment.ListItems.Count
            If LstPayment.ListItems(n).SubItems(4) = "UNSCH" And regnego = True Then
                regnego = True
            End If
    Next n
    
    If regnego = False Or LstPayment.ListItems.Count = 0 Then
        kontak = False
        shedulePTP_Show = False
        regnego = False
        ' 'M_OBJCONN.Close
        M_OBJCONN.Close
        Set M_OBJCONN = Nothing
        M_OBJCONN.Open CMDSQLOPEN
        VIEW_MGMDATA.WindowState = 2
    Else
            MsgBox "Lakukan PTP yang benar,Jumlah PTP harus >= Deal Payment " & txtPayment.text & " , Atau data simpan dulu!!!"
            Cancel = 1
            Exit Sub
    End If
    ' Reset and disable monitoring
    i_monitoring_activity = 0
    MDIForm1.Timer2.Enabled = False
    ' ####
    ' Reset REMINDER ##############
    bAktif_form_customer = False
    bReminder_agent = False
    bAktif_Cust_Review = False
    ' #############################
    
    SetWindowLong m_EditHWnd, GWL_WNDPROC, OldWindowProc
End Sub





Private Sub Image1_Click(Index As Integer)
    Select Case Index
       Case 0
'          If Image1(0).Tag = 0 Then
'            Tdbbalance.Visible = True
'            tdbprincipal.Visible = True
'            Label11(14).Visible = True
'            Label11(15).Visible = True
'            Image1(0).Tag = 1
'            LblPrompA.Visible = True
'            Label11(8).Visible = True
'        Else
'            Tdbbalance.Visible = False
'            tdbprincipal.Visible = False
'            Label11(14).Visible = False
'            Label11(15).Visible = False
'            Label11(8).Visible = False
'            Image1(0).Tag = 0
'            LblPrompA.Visible = False
'        End If

    End Select
End Sub

Private Sub Label1_Click(Index As Integer)
  Dim ami As Integer
  
  Select Case Index
        Case 80
  'If label1(80).Tag = 0 Then
   If UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.text) = "ADMIN" Or UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Then
            Tdbbalance.Visible = True
            '@@ 0408201 Dibuang
            'tdbprincipal.Visible = True
            
            Label11(14).Visible = True
            
            '@@ 04082011 Dibuang
            'Label11(15).Visible = True
            
            label1(80).tag = 1
            LblPrompA.Visible = True
            Label11(8).Visible = True
            For ami = 1 To LstDoubleId.ListItems.Count
                LstDoubleId.ListItems(ami).SubItems(4) = ENCRIPY(True, LstDoubleId.ListItems(ami).SubItems(4))
            Next ami
        Else
            Tdbbalance.Visible = False
            
            '@@ 0408201 Dibuang
            'tdbprincipal.Visible = False
            
            Label11(14).Visible = False
            
            '@@ 04082011 Dibuang
            'Label11(15).Visible = False
            
            Label11(8).Visible = False
            label1(80).tag = 0
            LblPrompA.Visible = False
             For ami = 1 To LstDoubleId.ListItems.Count
                LstDoubleId.ListItems(ami).SubItems(4) = ENCRIPY(False, LstDoubleId.ListItems(ami).SubItems(4))
            Next ami
        End If
End Select

End Sub






Private Sub LblBlacklistAddHome1_Click()
    Dim cmdsql, a As String
    
    If txtHomeAdd1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtHomeAdd1.Value
            .LblTelp.Caption = "AddHome 1"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homenoadd1='1',f_valid_addhome1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addhome1='1', f_sts_valid_addhome1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHome2_Click()
    Dim cmdsql, a As String
    
    If txtHomeAdd2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtHomeAdd2.Value
            .LblTelp.Caption = "AddHome 2"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homenoadd2='1',f_valid_addhome2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addhome2='1', f_sts_valid_addhome2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHP1_Click()
      Dim cmdsql, a As String
    
    If txtMobileAdd1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtMobileAdd1.Value
            .LblTelp.Caption = "AddMobile 1"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobilenoadd1='1',f_valid_addmobile1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addmobile1='1', f_sts_valid_addmobile1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHP2_Click()
    
    If txtMobileAdd2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtMobileAdd2.Value
            .LblTelp.Caption = "AddMobile 2"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobilenoadd2='1',f_valid_addmobile2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addmobile2='1', f_sts_valid_addmobile2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
             MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddOffice1_Click()
    Dim cmdsql, a As String
    
    If txtOfficeAdd1.Value <> Empty Then
        
       a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtOfficeAdd1.Value
            .LblTelp.Caption = "AddOffice 1"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officenoadd1='1',f_valid_addoffice1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addoffice1='1', f_sts_valid_addoffice1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
             MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddOffice2_Click()
    Dim cmdsql, a As String
    
    If txtOfficeAdd2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtOfficeAdd2.Value
            .LblTelp.Caption = "AddOffice 2"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officenoadd2='1',f_valid_addoffice2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addoffice2='1', f_sts_valid_addoffice2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlackliSTEC_Click()
    Dim cmdsql, a As String
    
    If txtECno.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtECno.Value
            .LblTelp.Caption = "EC"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_ec_telp='1',f_valid_ec=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_ec='1', f_sts_valid_ec='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistHome2_Click()
    Dim cmdsql, a As String
    
    If txtHomeNo2.Value <> Empty Then
        
       a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtHomeNo2.Value
            .LblTelp.Caption = "Home 2"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'             If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homeno2='1',f_valid_home2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'             ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_home2='1', f_sts_valid_home2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
             'End If
             MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistHp1_Click()
    Dim cmdsql, a As String
    
    If txtMobileNo1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtMobileNo1.Value
            .LblTelp.Caption = "Mobile 1"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobileno='1',f_valid_mobile1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                 'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_mobile1='1', f_sts_valid_mobile1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistHp2_Click()
    Dim cmdsql, a As String
    
    If txtMobileNo2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtMobileNo2.Value
            .LblTelp.Caption = "Mobile 2"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobileno2='1',f_valid_mobile1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_mobile2='1', f_sts_valid_mobile2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistOffice1_Click()
    Dim cmdsql, a As String
    
    If txtOfficeNo1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtOfficeNo1.Value
            .LblTelp.Caption = "Office 1"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officeno='1',f_valid_office1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_office1='1', f_sts_valid_office1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
            'End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistOfficeno2_Click()
    Dim cmdsql, a As String
    
    If txtOfficeNo2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtOfficeNo2.Value
            .LblTelp.Caption = "Office 2"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'             If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officeno2='1',f_valid_office2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_office2='1', f_sts_valid_office2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlakcListHome1_Click()
    Dim cmdsql, a As String
    
    If txtHomeNo1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .txtNotelp.text = txtHomeNo1.Value
            .LblTelp = "Home 1"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homeno='1',f_valid_home1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_home1='1', f_sts_valid_home1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblMap_Click()
    TimerBlinkDetailMapping.Enabled = False
    FrmDetailMapping.Show vbModal
End Sub

Private Sub ListView1_Click(Index As Integer)
Dim KET As String
Select Case Index
Case 0

Case 1
If listview1(1).ListItems.Count = 0 Then
Exit Sub
Else
   KET = TXtDetails.text
      If Len(TXtDetails) = 0 Then
         TXtDetails.text = " - " + listview1(1).SelectedItem.SubItems(1)
      Else
         TXtDetails.text = KET + " - " + listview1(1).SelectedItem.SubItems(1)
      End If
End If
End Select
End Sub



Private Sub LstDoubleId_DblClick()
     If LstDoubleId.ListItems.Count = 0 Then
        Exit Sub
    End If
    FrmCC_Colection.Hide
    frmCC_Colection2.Show vbModal
End Sub

Private Sub LstPayment_DblClick()
If LstPayment.ListItems.Count = 0 Then
Exit Sub
Else
Call SSCommand2_Click(1)
End If
End Sub
Private Sub Lstscript_DblClick()
  If Lstscript.ListItems.Count > 0 Then
  StartMeUp (Lstscript.SelectedItem.SubItems(2))
  'MsgBox (LstScript.SelectedItem.SubItems(2))
   End If
End Sub
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'Private Sub LstSMS_DblClick()
'If LstSMS.ListItems.Count > 0 Then
'
'no_telp = LstSMS.SelectedItem.Text
'isi_Pesan = LstSMS.SelectedItem.SubItems(3)
'
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan)
'
'    Else
'    Exit Sub
' End If
'End Sub

'@@ 11-03-2011 Di remarks, udah tidak diapakai

'Private Sub LstSMS2_DblClick()
'If LstSMS2.ListItems.Count > 0 Then
'
'no_telp = LstSMS2.SelectedItem.Text
'isi_Pesan = LstSMS2.SelectedItem.SubItems(2)
'
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan)
'
'    Else
'    Exit Sub
' End If
'End Sub

Private Sub LstVisit_DblClick()
 If LstVisit.ListItems.Count > 0 Then
            
        
           With FRM_UpdateVisit
                .Text1.text = LstVisit.SelectedItem.SubItems(2)
                .Show vbModal
                

'                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)
'
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        'LstPayment.SelectedItem.SubItems(1) = ""
'                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
'                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
'
'
'                    On Error GoTo 0
'                    End If
'                End If
               End With
Else
Exit Sub
End If

End Sub

Private Sub Option1_Click()
'@@ 03-05-2012, Dinonaktifkan
'If Option1.Value = True Then
'TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AHome1.Value & txtHomeNo1.Value))
'   If txtHomeNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(AHome1.Value & txtHomeNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'    End If
'   Option2.Value = False
'   Option3.Value = False
'   Option4.Value = False
'   Option5.Value = False
'End If
End Sub

Private Sub Option2_Click()
'@@ 03-05-2012 Dinonaktifkan
'If Option2.Value = True Then
'TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AHome2.Value & txtHomeNo2.Value))
'   If txtHomeNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(AHome2.Value & txtHomeNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'    End If
'   Option1.Value = False
'   Option3.Value = False
'   Option4.Value = False
'   Option5.Value = False
'End If
End Sub

Private Sub Option3_Click()
    '@@ 03-05-2012 DinonAktifkan
'   If Option3.Value = True Then
'   TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AOffice2.Value & txtOfficeNo2.Value))
'   If txtOfficeNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(AOffice2.Value & txtOfficeNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option4.Value = False
'   Option1.Value = False
'   Option5.Value = False
'   End If
End Sub

Private Sub Option4_Click()
'@@DinonAktifkan 03-05-2012
'   If Option4.Value = True Then
'   TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AOffice1.Value & txtOfficeNo1.Value))
'   If txtOfficeNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(AOffice1.Value & txtOfficeNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option3.Value = False
'   Option1.Value = False
'   Option5.Value = False
'End If
End Sub

Private Sub Option5_Click()
 If Option5.Value = True Then
 TYPETELP = ""
   txtPhone.text = GetNumber(CStr(txtMobileNo2.Value))
    If txtMobileNo2.Value <> "" Then
        txtPhoneA.text = CStr(txtMobileNo2A.Value)
    Else
        txtPhoneA.text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option6.Value = False
   End If
End Sub

Private Sub Option6_Click()
 If Option6.Value = True Then
 TYPETELP = ""
   txtPhone.text = GetNumber(CStr(txtMobileNo1.Value))
   If txtMobileNo1.Value <> "" Then
        txtPhoneA.text = CStr(txtMobileNo1A.Value)
    Else
        txtPhoneA.text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option5.Value = False
   End If
End Sub

Private Sub Option7_Click(Index As Integer)
Select Case Index
Case 0
TxtAddress.text = AddrNow.text
Case 1
TxtAddress.text = lblAddr.text
Case 2
TxtAddress.text = lblOfficeAddr.text
End Select

End Sub

Private Sub Option8_Click(Index As Integer)
Select Case Index
Case 0
Frame8.Enabled = True
VisitYES
Case 1
VisitNo
Frame8.Enabled = False
End Select
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'LstSMS.Visible = True
'LstSMS2.Visible = False
End If
End Sub

Private Sub Option10_Click()
If Option10.Value = True Then
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'LstSMS.Visible = False
'LstSMS2.Visible = True
End If

End Sub

Private Sub SSCommand1_Click(Index As Integer)
Dim rsshut As New ADODB.Recordset
'On Error GoTo ke

Dim n As Integer
Select Case Index
  
  '@@ 05-10-2011 Skip Tracer ditiadakan
  'Case 7
  'frmdetailskip.Show 1
  
  Case 5
    'FRMSCRIPT.Show 1
    '@@ 09092011 Offering Discon digabung sama offering yang lama
    Call OfferingDiscGuide
  Case 0
    
    
        
 
'  If Len(CmbPhone.Text) < 2 Then
'    MsgBox "Pilihan No Telephone harus diisi"
'    CmbPhone.SetFocus
'    Exit Sub
'  End If
        
        '@@220610 --- Agar agent tidak dapat mengisi no.telepon di combo phone
'        If IsNumeric(CmbPhone.Text) = True Then
'            If CmbPhone.Text <> "108" Then
'                CmbPhone.Text = ""
'                MsgBox "Pilih no telepon!", vbOKOnly + vbCritical, "Peringatan"
'                Exit Sub
'            End If
'        End If
        
        StsKategoriTelepon = ""
        KelompokKategoriTlp = ""
        
        Select Case CmbPhone
            '@@02-05-2011 Tambahan Telp Additional
            Case "TelpAdditional"
                txtPhone.text = Trim(TxtAdditional.Value)
                telpno = txtPhone.text
                '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@02052012,Jika telepon additional pindahkan ke kotak additional yang baru
                'untuk memasukkan kategori telepon
                MsgBox "Sebelum anda melakukan call, harap pindahkan terlebih dahulu kategori teleponnya! Terima Kasih!", vbOKOnly + vbInformation, "Informasi"
                FrmReqTelepon.TxtCustid = Trim(lblCustId.Caption)
                FrmReqTelepon.txtNotelp.text = Trim(txtPhone.text)
                FrmReqTelepon.Show vbModal
                'Kosongkan telp_additional
                cmdsql = "update mgm set telp_additional=null where custid='"
                cmdsql = cmdsql + CStr(lblCustId.Caption) + "'"
                M_OBJCONN.execute cmdsql
            Case "Hp"
                txtPhone.text = Trim(txtMobileNo1.Value)
                telpno = txtPhone.text
                '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@11052012, Tambahan Kategori Telepon
                StsKategoriTelepon = "HP"
            Case "Hp2"
                txtPhone.text = txtMobileNo2.Value
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "HP"
            Case "HomePhone"
                '@@03-05-2012 DinonAktifkan
'                If AHome1.Value = "031" Or AHome1.Value = "" Then
'                    txtPhone.Text = Trim(txtHomeNo1.Value)
'                Else
'                    txtPhone.Text = Trim(AHome1.Value) & txtHomeNo1.Value
'                End If
                txtPhone.text = Trim(txtHomeNo1.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Home"
            Case "HomePhone2"
                '@@03-05-2012 Dinonaktifkan
'                If AHome1.Value = "031" Or AHome1.Value = "" Then
'                    txtPhone.Text = txtHomeNo2.Value
'                Else
'                    txtPhone.Text = Trim(AHome1.Value) & Trim(txtHomeNo2.Value)
'                End If
                txtPhone.text = Trim(txtHomeNo2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Home"
            Case "OfficePhone"
                '@@03-05-2012 DinonAktifkan
'                If AOffice1.Value = "031" Or AOffice1.Value = "" Then
'                    txtPhone.Text = txtOfficeNo1.Value
'                Else
'                    txtPhone.Text = AOffice1.Value & txtOfficeNo1.Value
'                End If
                txtPhone.text = Trim(txtOfficeNo1.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Office"
            Case "OfficePhone2"
                '@@03-05-2012 DinonAktifkan
'                If AOffice2.Value = "031" Or AOffice2.Value = "" Then
'                    txtPhone.Text = txtOfficeNo2.Value
'                Else
'                    txtPhone.Text = AOffice1.Value & txtOfficeNo2.Value
'                End If
                txtPhone.text = Trim(txtOfficeNo2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Office"
            Case "EconPhone"
                If txtECno.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                If Left(txtECno.text, 3) = "031" Then
                 txtPhone.text = Trim(Mid(txtECno.Value, 4, 16))
                 Else
                 txtPhone.text = Trim(txtECno.Value)
                End If
                txtPhone.text = txtECno.Value
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "EC"
            Case "AddHome1"
                If txtHomeAdd1A.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                '@@03-05-2012 Dinonaktifkan
'                If AHomeAdd1(0).Value = "031" Or AHomeAdd1(0).Value = "" Then
'                    txtPhone.Text = txtHomeAdd1.Value
'                Else
'                    txtPhone.Text = AHomeAdd1(0).Value & txtHomeAdd1.Value
'                End If
                txtPhone.text = Trim(txtHomeAdd1.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@ 02-05-2012,Tambahan Buat Nyatet Kategori Telepon
'                If CmbStsKatHome1.Text = "" Or CmbStsKatHome1.Text = "--Pilih Kategori Telepon--" Then
'                    MsgBox "Mohon maaf, tentukan terlebih dahulu kategori telepon di AddHome 1!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                End If
                StsKategoriTelepon = Trim(CmbStsKatHome1.text)
            Case "AddHome2"
                If txtHomeAdd2A.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                '@@03-05-2012 Dinonaktifkan
'                If AHomeAdd2(1).Value = "031" Or AHomeAdd2(1).Value = "" Then
'                    txtPhone.Text = txtHomeAdd2.Value
'                Else
'                    txtPhone.Text = AHomeAdd2(1).Value & txtHomeAdd2.Value
'                End If
                txtPhone.text = Trim(txtHomeAdd2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@ 02-05-2012,Tambahan Buat Nyatet Kategori Telepon
'                If CmbStsKatHome2.Text = "" Or CmbStsKatHome2.Text = "--Pilih Kategori Telepon--" Then
'                    MsgBox "Mohon maaf, tentukan terlebih dahulu kategori telepon di AddHome 2!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                End If
                StsKategoriTelepon = Trim(CmbStsKatHome2.text)
            Case "AddOffice1"
                If txtOfficeAdd1A.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                '@@03-05-2012 Dinonaktifkan
'                If AOfficeAdd(2).Value = "031" Or AOfficeAdd(2).Value = "" Then
'                    txtPhone.Text = txtOfficeAdd1.Value
'                Else
'                    txtPhone.Text = AOfficeAdd(2).Value & txtOfficeAdd1.Value
'                End If
                txtPhone.text = Trim(txtOfficeAdd1.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@ 02-05-2012,Tambahan Buat Nyatet Kategori Telepon
'                If CmbStsKatOffice1.Text = "" Or CmbStsKatOffice1.Text = "--Pilih Kategori Telepon--" Then
'                    MsgBox "Mohon maaf, tentukan terlebih dahulu kategori telepon di AddOffice 1!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                End If
                StsKategoriTelepon = Trim(CmbStsKatOffice1.text)
            Case "AddOffice2"
                If txtOfficeAdd2A.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                '@@03-05-2012 Dinonaktifkan
'                If AOfficeAdd(3).Value = "031" Or AOfficeAdd(3).Value = "" Then
'                    txtPhone.Text = Trim(txtOfficeAdd2.Value)
'                Else
'                    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'                End If
                txtPhone.text = Trim(txtOfficeAdd2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@ 02-05-2012,Tambahan Buat Nyatet Kategori Telepon
'                If CmbStsKatOffice2.Text = "" Or CmbStsKatOffice2.Text = "--Pilih Kategori Telepon--" Then
'                    MsgBox "Mohon maaf, tentukan terlebih dahulu kategori telepon di AddOffice 2!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                End If
                StsKategoriTelepon = Trim(CmbStsKatOffice2.text)
            Case "AddMobile1"
                If txtMobileAdd1A.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                txtPhone.text = Trim(txtMobileAdd1.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@ 02-05-2012,Tambahan Buat Nyatet Kategori Telepon
'                If CmbStsKatHP1.Text = "" Or CmbStsKatHP1.Text = "--Pilih Kategori Telepon--" Then
'                    MsgBox "Mohon maaf, tentukan terlebih dahulu kategori telepon di AddOffice 1!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                End If
                StsKategoriTelepon = Trim(CmbStsKatHP1.text)
            Case "AddMobile2"
                If txtMobileAdd2A.ReadOnly = False And UCase(MDIForm1.Text2.text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                txtPhone.text = Trim(txtMobileAdd2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@ 02-05-2012,Tambahan Buat Nyatet Kategori Telepon
'                If CmbStsKatHP2.Text = "" Or CmbStsKatHP2.Text = "--Pilih Kategori Telepon--" Then
'                    MsgBox "Mohon maaf, tentukan terlebih dahulu kategori telepon di AddOffice 1!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                End If
                StsKategoriTelepon = Trim(CmbStsKatHP2.text)
            Case Else
               
'               '@@ 17-04-2012, Cek dulu apakah ada telepon tambahan
'               If TxtNoTelpReq.Value = Empty Then
'                    Dim M_Objrs_Cek As ADODB.Recordset
'                    '@@09092011 Cek dulu apakah user telepon ada di tbllayanan telkom
'                     txtPhone.Text = Replace(CmbPhone.Text, " ", "")
'                    Cmdsql = "select * from tbllayanantelkom where nolayanan='"
'                    Cmdsql = Cmdsql + Trim(txtPhone.Text) + "'"
'                    Set M_Objrs_Cek = New ADODB.Recordset
'                    M_Objrs_Cek.CursorLocation = adUseClient
'                    M_Objrs_Cek.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                    TxtTelpKe.Text = CmbPhone.Text
'
'                    If M_Objrs_Cek.RecordCount = 0 Then
'                        MsgBox "Maaf, anda tidak dapat menelepon nomor yang tidak terdapat dalam database!", vbOKOnly + vbCritical, "Peringatan"
'                        Set M_Objrs_Cek = Nothing
'                        Exit Sub
'                    End If
'                Else
'                     txtPhone.Text = Trim(TxtNoTelpReq.Value)
'                     TxtTelpKe.Text = Trim(CmbPhone.Text)
'                     KelompokKategoriTlp = TxtKategori.Caption
'                     StsKategoriTelepon = TxtTelpKe.Text
'                End If
                
                '@@ 11 Juni 2012, Revisi Tambahan Telepon
                 txtPhone.text = Replace(CmbPhone.text, " ", "")
                 cmdsql = "select * from tbllayanantelkom where nolayanan='"
                 cmdsql = cmdsql + Trim(txtPhone.text) + "'"
                 Set M_Objrs_Cek = New ADODB.Recordset
                 M_Objrs_Cek.CursorLocation = adUseClient
                 M_Objrs_Cek.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                 If M_Objrs_Cek.RecordCount > 0 Then
                    TxtTelpKe.text = CmbPhone.text
                 Else
                    If TxtNoTelpReq.Value <> Empty Then
                        txtPhone.text = Trim(TxtNoTelpReq.Value)
                        TxtTelpKe.text = Trim(CmbPhone.text)
                        KelompokKategoriTlp = TxtKategori.Caption
                        StsKategoriTelepon = TxtTelpKe.text
                    Else
                       Set M_Objrs_Cek = Nothing
                       MsgBox "Maaf, anda tidak dapat menelepon nomor yang tidak terdapat dalam database!", vbOKOnly + vbCritical, "Peringatan"
                       Exit Sub
                    End If
                 End If
               Set M_Objrs_Cek = Nothing
        End Select
        
        '@@31-05-2012 Jika Status Account=PO dan CO maka tidak dapat di call
        If StatusAccount = "PO-" Or StatusAccount = "CO-" Then
            MsgBox "Mohon maaf! Status Account PAID OFF atau COMPLAIN tidak dapat di call!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        
        '@@ 02052012,, Tambahan Untuk SpeakWith
        Call PilihSpeakWith
        '@@ 03052012,,Tambahn Status Kategori
        Call CariKategoriTlp
        
    'Cek no telepon yang apakah masuk daftar blacklist. Jika masuk maka keluar sub!
    cmdsql = "select no_telp from tblblacklist where no_telp='"
    cmdsql = cmdsql + Replace(Trim(txtPhone.text), " ", "") + "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount <> 0 Then
            MsgBox "No.Telepon yang anda hubungi masuk dalam daftar blacklist!. Silahkan hubungi TL  anda!.", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
    Set M_Objrs = Nothing
    
    '@@ 07-05-2012, Cek Apakah termasuk Unvalid number?
    cmdsql = "select no_telp from tblunvalid_number where no_telp='"
    cmdsql = cmdsql + Replace(Trim(txtPhone.text), " ", "") + "' "
    '@@ 23-05-2012, Tambahkan yang blok hanya custid tertentu dengn nomor tertentu saja
    cmdsql = cmdsql + " and custid='"
    cmdsql = cmdsql + CStr(lblCustId.Caption) + "'"
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount <> 0 Then
            MsgBox "No.Telepon yang anda hubungi masuk dalam daftar Unvalid number!. Silahkan hubungi TL  anda!.", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
        
    ' ----------- CEK WIT OR WITA 05 FEB 2014 -----------
    If M_Objrs.state = 1 Then M_Objrs.Close
    M_Objrs.Open "SELECT now() as wkt_server"
    If M_Objrs.RecordCount > 0 Then
        waktu_server_skrg = M_Objrs!wkt_server
    End If
    
    If M_Objrs.state = 1 Then M_Objrs.Close
    M_Objrs.Open "SELECT * FROM tbl_timezone WHERE trim(kode)='" & Left(Replace(Trim(txtHomeNo1A.text), " ", ""), 4) & "'"
    If M_Objrs.RecordCount > 0 Then
        If Format(waktu_server_skrg, "hh:mm") >= Format(M_Objrs!time_limit, "hh:mm") Then
            MsgBox "Maaf anda tidak diperkenankan Telp pada Pukul atau melebihi " & M_Objrs!time_limit & " Pada area " & M_Objrs!group_time, vbCritical + vbOKOnly, "INFO"
            Exit Sub
        End If
    End If
    ' ---------------------------------------------------
    Set M_Objrs = Nothing
    

    
    ' 23-05-2013 untuk 5x Blok -------------------------
    sPhone_Agent = Trim(MDIForm1.Text1.text)
    sPhone_CustID = CStr(lblCustId.Caption)
    sPhone_TelpNo = Replace(Trim(txtPhone.text), " ", "")
    ' ---------------------------------------------------
    
    '@@ 18-04-2012, Cek setiap agent yang menelepon
    'ke nomor yang sama nomor teleponnya tidak bisa dihubungi lagi
    Dim M_Objrs_Cek_Panggilan As ADODB.Recordset
    
'    Cmdsql = "select * from tblphonemonitorhst where telpno='"
'    Cmdsql = Cmdsql + Trim(txtPhone.Text) + "' and userid='"
'    Cmdsql = Cmdsql + Trim(MDIForm1.Text1.Text) + "' and date(tgl)=date(now()) and flag_review='1' "
'    Set M_Objrs_Cek_Panggilan = New ADODB.Recordset
'    M_Objrs_Cek_Panggilan.CursorLocation = adUseClient
'    DoEvents
'    M_Objrs_Cek_Panggilan.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_Objrs_Cek_Panggilan.RecordCount >= 5 Then
'        MsgBox "Mohon maaf, anda sudah melakukan call ke nomor ini 5 kali. Anda hanya boleh melakukan call ke nomor yang sama, hanya 5 kali di hari yang sama. Silahkan call lagi besok atau hubungi TL/SPV anda!", vbOKOnly + vbInformation, "Informasi"
'        '@@18April2012, Ubah coding menjadi review
'        Cmdsql = "update mgm set agent='REVIEW' where custid='"
'        Cmdsql = Cmdsql + lblCustId.Caption + "'"
'        M_OBJCONN.Execute Cmdsql
'        MsgBox "Mohon maaf, untuk sementara custid: " & lblCustId.Caption & ", atas nama: " & lblNama.Caption + " dipindahkan ke coding REVIEW!", vbOKOnly + vbInformation, "Informasi"
'        Set M_Objrs_Cek_Panggilan = Nothing
'        Exit Sub
'    End If
'    Set M_Objrs_Cek_Panggilan = Nothing

    '@@19042012 Diganti searching ke icentra
'    CMDSQL = "select distinct durasi,acd_log_outgoing_session_id from outgoing_icentra where destination='"
'    CMDSQL = CMDSQL + CStr(Trim(txtPhone.Text)) + "' and custid='"
'    CMDSQL = CMDSQL + CStr(lblCustId.Caption) + "' and date(initiate)=date(now()) "
'    CMDSQL = CMDSQL + " and durasi >=40 "

    ' UPDATE 19 AGUSTUS 2014 BY IZUDDIN UNTUK ACC REVIEW
    If UCase(Trim(lblaoc.Caption)) = "AKSESALL" Or UCase(Trim(Left(lblaoc.Caption, 6))) = "REVIEW" Then
        lblagent_review = lbl_agentlama.Caption
    Else
        lblagent_review = lblaoc.Caption
    End If

     'Fitur telp 5x Blok Ditutup lagi 23 Mei 2013
     'Diaktifkan kembali 10 may 2013
    cmdsql = "SELECT * FROM user_phone_log WHERE custid='" & CStr(lblCustId.Caption) & "' AND date(call_log_time)=" & _
            "date(now()) AND no_telp='" & CStr(Trim(txtPhone.text)) & "' and agent='" & MDIForm1.Text1.text & "'"

    Set M_Objrs_Cek_Panggilan = New ADODB.Recordset
    M_Objrs_Cek_Panggilan.CursorLocation = adUseClient
    DoEvents
    M_Objrs_Cek_Panggilan.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ' Tambahan untuk tidak include layanan TELKOM
    If M_Objrs_Cek_Panggilan.RecordCount >= 5 And Trim(txtPhone.text) <> "031108" And Trim(txtPhone.text) <> "108" Then
        MsgBox "Mohon maaf, anda sudah melakukan call ke nomor ini 5 kali. Anda hanya boleh melakukan call ke nomor yang sama, hanya 5 kali di hari yang sama. Silahkan call lagi besok atau hubungi TL/SPV anda!", vbOKOnly + vbInformation, "Informasi"
        '@@18April2012, Ubah coding menjadi review
'        CMDSQL = "update mgm set agent='REVIEW' where custid='"
'        CMDSQL = CMDSQL + lblCustId.Caption + "'"
        '@@23042012, Pindah ke agent REVIEW sesuai dengan agentnya
        'SET AGENT ASLI!!
        cmdsql = "UPDATE mgm SET agent=agent_new,agent_asli='" & lblagent_review & "' "
        cmdsql = cmdsql + "from (select userid as agent_new from usertbl where userid like 'REVIEW%' "
        cmdsql = cmdsql + " and team in (select team from usertbl where userid='"
        ' cmdsql = cmdsql + MDIForm1.Text1.Text + "') ) as a "
        ' REVISI 28 AGUSTUS 2014
        cmdsql = cmdsql + lblagent_review + "') ) as a "
        cmdsql = cmdsql + " where mgm.custid='"
        cmdsql = cmdsql + lblCustId.Caption + "'"
        M_OBJCONN.execute cmdsql
        
        Set M_Objrs_Cek_Panggilan = Nothing
        
        '@@10052012 Inputkan Buat Bikin Log Custid Yang Masuk dalam Daftar Review
        cmdsql = "insert into tbl_log_acc_review (custid,agent,telp) values ('"
        cmdsql = cmdsql + CStr(lblCustId.Caption) + "','"
        cmdsql = cmdsql + CStr(lblagent_review) + "','"
        cmdsql = cmdsql + CStr(Trim(txtPhone.text)) + "')"
        M_OBJCONN.execute cmdsql
        MsgBox "Mohon maaf, untuk sementara custid: " & lblCustId.Caption & ", atas nama: " & lblNama.Caption + " dipindahkan ke coding REVIEW!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    Set M_Objrs_Cek_Panggilan = Nothing

    
   
    cmdsql = "Insert Into tblphonemonitorhst(UserId, CustId, NamaCh,StartDate, TelpNo, Recsource,status_telp,tgl) Values "
    cmdsql = cmdsql + " ('" + MDIForm1.Text1.text + "' , '" + FrmCC_Colection.lblCustId.Caption + "','"
    cmdsql = cmdsql + FrmCC_Colection.lblNama.Caption + "', '"
    cmdsql = cmdsql + Format(CStr(MDIForm1.TDBDate1.Value), "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
    cmdsql = cmdsql + "' , '" + Replace(txtPhone.text, " ", "") + "' ,'"
    cmdsql = cmdsql + FrmCC_Colection.lblRecsource.Caption + "','"
    cmdsql = cmdsql + IIf(IsNull(TxtKategori.Caption), "", TxtKategori.Caption) + "',now())"
    M_OBJCONN.execute cmdsql
    
    
    '@@19042012 Tombol Exit,Tombol Call di Nonaktifkan dulu
    SSCommand1(3).Enabled = False
    '@@19042012 Tombol Hangup Diaktifkan
    SSCommand1(1).Enabled = True
    '@@19042012 Tombol Call Dinonaktifkan
    SSCommand1(0).Enabled = False
    
    '@@25-05-2012 Tombol Save dinonaktifkan
    SSCommand1(2).Enabled = False
    
    '@@ Filter tanda baca ditelepon
    txtPhone.text = Replace(txtPhone.text, "/", "")
    txtPhone.text = Replace(txtPhone.text, "\", "")
    txtPhone.text = Replace(txtPhone.text, "'", "")
    txtPhone.text = Replace(txtPhone.text, ";", "")
    txtPhone.text = Replace(txtPhone.text, ":", "")
    txtPhone.text = Replace(txtPhone.text, "|", "")
    txtPhone.text = Replace(txtPhone.text, ".", "")
    txtPhone.text = Replace(txtPhone.text, ",", "")
    txtPhone.text = Replace(txtPhone.text, "?", "")
    txtPhone.text = Replace(txtPhone.text, "!", "")
    txtPhone.text = Replace(txtPhone.text, " ", "")
    
    
    'MDIForm1.ActionCTI ("DIAL|496821" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.Caption) & "|" & Trim(FrmCC_Colection.lblCustId.Caption))
    MDIForm1.ActionCTI ("DIAL|496821" & GetNumber(CStr(Replace(txtPhone.text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.Caption) & "|" & Trim(FrmCC_Colection.lblCustId.Caption)) & "-" & MDIForm1.Text1.text
    '@@ 25-07-2011 Dipindah, jadi di form load
    'Call OfferingDiscGuide
    
    MDIForm1.CmbNo.text = ""
    stscall = True
    TYPETELP = ""
   Case 2
        V_SAVE = CEK_DATA_VALID
        
        
        If V_SAVE = False Then
            Exit Sub
        Else
        End If
        If ADD_CUST Then
        Else
            Call CEK_UPDATE_PELANGGAN
            stscall = False
            Call isi_datapayment
        End If
        Call load_reminder
   Case 3
     If bRenderrecord = True Then
          '  VIEW_MGMDATA.renderdonk
      End If
      bRenderrecord = False
    kontak = False
        For n = 1 To LstPayment.ListItems.Count
            If LstPayment.ListItems(n).SubItems(4) = "UNSCH" And regnego = True Then
                regnego = True
            End If
        Next n
        If regnego = True And LstPayment.ListItems.Count <> 0 Then
            MsgBox "Lakukan PTP yang benar, Jumlah PTP harus >= Deal Payment " & txtPayment.text & " ,Atau data simpan dulu!!!"
            Exit Sub
        End If
     Strsql = "select * from tblshut where nshut=1"
     Set rsshut = New ADODB.Recordset
     rsshut.CursorLocation = adUseClient
     rsshut.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      If Not rsshut.EOF Then
         Strsql = "UPDATE  tblshut SET nshut=0"
        M_OBJCONN.execute (Strsql)
        End
        Exit Sub
      End If
      Set rsshut = Nothing
      
'      '@@ Awal 061110 cek lock account sesuai settingan timer
'        Dim m_objrsTemp As ADODB.Recordset
'        Dim m_objrsWaktuServer As ADODB.Recordset
'        Dim m_objrsCurrent As ADODB.Recordset
'
'
'        Dim cmdsqlserver As String
'        Dim WaktuServer As Date
'        Dim WaktuAkhirCurrent As Date
'
'        'ambil waktu server
'        cmdsqlserver = "select now() as WaktuServer "
'        Set m_objrsWaktuServer = New ADODB.Recordset
'        m_objrsWaktuServer.CursorLocation = adUseClient
'        m_objrsWaktuServer.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        WaktuServer = Format(m_objrsWaktuServer(0), "mm-dd-yyyy hh:mm")
'        Set m_objrsWaktuServer = Nothing
'
'        'Cek lock account yang sedang berjalan
'        cmdsqlserver = "select * from tbltemplockacc_current "
'        Set m_objrsCurrent = New ADODB.Recordset
'        m_objrsCurrent.CursorLocation = adUseClient
'        m_objrsCurrent.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If m_objrsCurrent.RecordCount <> 0 Then
'            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
'        Else
'            GoTo lockdata
'        End If
'
'        While Not m_objrsCurrent.EOF
'
'            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
'
'            If WaktuAkhirCurrent <= WaktuServer Then
'                'Cek dulu apakah ada user yang sedang mereset data
'                If Trim(m_objrsCurrent("f_locked")) = "2" Then
'                    GoTo KeluarLockAutoTL
'                End If
'
'                'update dulu status lock yang sedang berakhir, supaya agent lain ga ikut ngereset
'                cmdsqlserver = "update tbltemplockacc_current set f_locked='2' where id='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
'                cmdsqlserver = "update usertbl set dilockoleh='ClearByAutomatic',"
'                cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
'                cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null"
'                'Buat ambil kondisi agent yang sedang di lock
'                If Trim(m_objrsCurrent("account_lock")) = "ALL" Then
'                    cmdsqlserver = cmdsqlserver + " where usertype='1' "
'                ElseIf Left(Trim(m_objrsCurrent("account_lock")), 3) = "SPV" Then
'                    cmdsqlserver = cmdsqlserver + " where spvcode='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                Else
'                    cmdsqlserver = cmdsqlserver + " where userid='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                End If
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
'                cmdsqlserver = "update usertbl set f_pesanresetauto='1'"
'
'                'Buat mengupdate pesan kondisi agent yang di lock
'                If Trim(m_objrsCurrent("account_lock")) = "ALL" Then
'                    cmdsqlserver = cmdsqlserver + " where usertype='1'  "
'                ElseIf Left(Trim(m_objrsCurrent("account_lock")), 3) = "SPV" Then
'                    cmdsqlserver = cmdsqlserver + " where spvcode='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                Else
'                    cmdsqlserver = cmdsqlserver + " where userid='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                End If
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
'                cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current "
'                cmdsqlserver = cmdsqlserver + " where id='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Hapus data di tabel locktemp current
'                cmdsqlserver = "delete from tbltemplockacc_current where id='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
'                M_OBJCONN.Execute cmdsqlserver
'
'             End If
'KeluarLockAutoTL:
'                m_objrsCurrent.MoveNext
'            Wend
'            Set m_objrsCurrent = Nothing
'
'
'
'
'        '=======
'lockdata:
'        'Setelah cek waktu lock yang habis, sekarang cek lock yg masih dalam antrian
'        cmdsqlserver = "select * from tbltemplockacc where f_locked isnull order by start_lock asc "
'        Set m_objrsTemp = New ADODB.Recordset
'        m_objrsTemp.CursorLocation = adUseClient
'        m_objrsTemp.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'            'Cek ada ga data lock dalam antrian
'            If m_objrsTemp.RecordCount <> 0 Then
'                Dim WaktuAwal As Date
'                Dim WaktuAkhir As Date
'
'                While Not m_objrsTemp.EOF
'
'                    WaktuAwal = Format(m_objrsTemp("start_lock"), "mm-dd-yyyy hh:mm")
'                    WaktuAkhir = Format(m_objrsTemp("end_lock"), "mm-dd-yyyy hh:mm")
'
'                    If (WaktuAwal <= WaktuServer) And (WaktuAkhir > WaktuServer) Then
'                        'Cek apakah datanya sedang di lock sama agent lain?
'                        If Trim(m_objrsTemp("f_locked")) = "1" Then
'                            GoTo KeluarLockAutoTLLock
'                        End If
'
'                        'update status  f_lockednya jadi 1, supaya ga di log sama agent lain
'                        cmdsqlserver = "update tbltemplockacc set f_locked='1' where id='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
'                        M_OBJCONN.Execute cmdsqlserver
'
'                        'LAKUKAN LOCK DATA
'                        Dim i As Integer
'                        a = Split(m_objrsTemp("script_lock"), "|")
'
'                        For i = LBound(a) + 1 To UBound(a) - 1
'                            cmdsqlserver = Replace(a(i), "$", "'")
'                            M_OBJCONN.Execute cmdsqlserver
'                        Next i
'
'                        'Pindahin dulu data di tabel current ke tabel log, terus data di tabel current dihapus
''                        cmdsqlserver = "insert into tbltemplockacc_current "
''                        cmdsqlserver = cmdsqlserver + " select * from tbltemplockacc_log"
''                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
'
''                        cmdsqlserver = "delete from tbltemplockacc_current"
''                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
'
'                        'Pindahin data dari tabel temp lock ke tabel current log
'                        cmdsqlserver = "insert into tbltemplockacc_current "
'                        cmdsqlserver = cmdsqlserver + "select * from tbltemplockacc where id='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
'                        M_OBJCONN.Execute cmdsqlserver
'
'
'
'                       'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
'                        cmdsqlserver = "update usertbl set f_pesanlockauto='1',f_idsessstart='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "',f_idsessend='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "' "
'                        'Buat mengupdate pesan kondisi agent yang di lock
'                        If Trim(m_objrsTemp("account_lock")) = "ALL" Then
'                            cmdsqlserver = cmdsqlserver + " where usertype='1' "
'                        ElseIf Left(Trim(m_objrsTemp("account_lock")), 3) = "SPV" Then
'                            cmdsqlserver = cmdsqlserver + " where spvcode='"
'                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
'                        Else
'                            cmdsqlserver = cmdsqlserver + " where userid='"
'                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
'                        End If
'                        M_OBJCONN.Execute cmdsqlserver
'
'                        'Hapus data di templock
'                        cmdsqlserver = "delete from tbltemplockacc where id='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
'                        M_OBJCONN.Execute cmdsqlserver
'
'
'                    End If
'
'KeluarLockAutoTLLock:
'                    m_objrsTemp.MoveNext
'               Wend
'
'            End If
'
'        Set m_objrsTemp = Nothing

        '@@22072013 Tambahan cek aksesallacc
        Call CekAksessAllAcc

        Call PesanLockAuto
        
        '@@Buka lock account yang aksess ALL
        If Trim(UCase(lblaoc.Caption)) = "AKSESALL" Then
            cmdsql = "update mgm set monitor_akses=null,waktu_akses=null where custid='"
            cmdsql = cmdsql & lblCustId.Caption & "' and agent='AKSESALL'"
            M_OBJCONN.execute cmdsql
            
            '@@20022013, buat jaga2 nih khawatir tinsnya error, hapus juga deh berdasarkan agent
            cmdsql = "update mgm set monitor_akses=null,waktu_akses=null where monitor_akses like '%"
            cmdsql = cmdsql & MDIForm1.Text1.text & "%' and agent='AKSESALL'"
            M_OBJCONN.execute cmdsql
        End If
                
                
        '@@28012013 Cek nih apakah akunnya diblok
        Dim M_Objrs_Cek_Blok As ADODB.Recordset
        cmdsql = "select * from usertbl where userid='"
        cmdsql = cmdsql + Trim(MDIForm1.Text1.text) + "'"
        Set M_Objrs_Cek_Blok = New ADODB.Recordset
        M_Objrs_Cek_Blok.CursorLocation = adUseClient
        M_Objrs_Cek_Blok.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_Blok.RecordCount > 0 Then
            If Trim(M_Objrs_Cek_Blok("f_blok")) = "1" Then
                MsgBox "Mohon maaf, akun TINS anda di blok oleh SPV/Admin! Anda tidak dapat login ke aplikasi TINS. Konfirmasikan hal ini ke SPV/Admin!", vbOKOnly + vbCritical, "Informasi"
                End
            End If
        End If
        
        Set M_Objrs_Cek_Blok = Nothing
        
      '@@ Akhir 061110 cek lock account sesuai settingan timer
        Dim M_Objrs_Close As ADODB.Recordset
        cmdsql = "select sts_close from usertbl where userid='"
        cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "' and sts_close='1'"
        Set M_Objrs_Close = New ADODB.Recordset
        M_Objrs_Close.CursorLocation = adUseClient
        M_Objrs_Close.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Close.RecordCount > 0 Then
            MsgBox "Mohon maaf, ada perubahan system. Aplikasi TINS akan di tutup! Harap Login Ulang!", vbOKOnly + vbInformation, "Informasi"
            Set M_Objrs_Close = Nothing
            cmdsql = "update usertbl set sts_close=null where userid='"
            cmdsql = cmdsql + CStr(MDIForm1.Text1.text) + "' "
            M_OBJCONN.execute cmdsql
            End
        End If
        Set M_Objrs_Close = Nothing
        
        ' Matikan monitoring activity
        i_monitoring_activity = 0
        MDIForm1.Timer2.Enabled = False
        main_timer_activity = 0
        MDIForm1.Timer7.Enabled = True
        ' #####
        
        Unload Me
        Exit Sub
'KeluarLockAuto:
        'Unload Me
    Case 1
        DoEvents
        MDIForm1.ActionCTI ("HANGUP")
        SSCommand1(1).Enabled = False
        'WaitSecs (2)
        '@@ 18 April 2012, Catat ketika agent mengakhiri telepon
        cmdsql = "update tblphonemonitorhst set enddate=now() from "
        cmdsql = cmdsql + " (select id as idnew from "
        cmdsql = cmdsql + " tblphonemonitorhst where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and userid='"
        cmdsql = cmdsql + MDIForm1.Text1.text + "' order by id desc limit 1) as a "
        cmdsql = cmdsql + " where tblphonemonitorhst.id=idnew"
        DoEvents
        M_OBJCONN.execute cmdsql
        Call HitungDurasiCall
        DoEvents
        Call HitungDurasiDariIcentra
        '@@19042012 Tombol Exit,diaktifkan
        SSCommand1(3).Enabled = True
        '@@19042012 Tombol Hangup Dinonaktifkan
        SSCommand1(1).Enabled = False
        '@@19042012 Tombol Call Diaktifkan
        SSCommand1(0).Enabled = True
        '@@25-05-2012 Tombol Save Diaktifkan
        SSCommand1(2).Enabled = True
        
        ' Reset monitoring activity
        ' i_monitoring_activity = 0
        MDIForm1.Timer2.Enabled = True
        ' #####
        
           '@@08102012, Buat Hangup Xlite
        On Error Resume Next
        Dim iret As Long
        THandle = FindWindow(vbEmpty, "X-Lite")
        If THandle = 0 Then
            MsgBox "Maaf, X-Lite  tidak ditemukan!"
            Exit Sub
        End If
        iret = BringWindowToTop(THandle)
        Sendkeys "^h", 0.7
        WaitSecs 0.2
        Sendkeys "^h", 0.7
        
        
        txtremarks.SetFocus
    Case 4
        StatusCPA = "CPA Form 1"
        frmcpanew.Show 1
        
End Select
Exit Sub
'ke:
Strsql = "update usertbl set stsaplikasi=0  where userid ='" + MDIForm1.Text1.text + "'"
M_OBJCONN.execute (Strsql)
MsgBox err.Description
 Exit Sub
 
End Sub

Public Sub Show_NEGOPTP()
Dim showlist As New ADODB.Recordset
Dim ListItem As ListItem
Dim cmdsql As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblCustId.Caption + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
End If
'If ShowList.BOF And ShowList.EOF Then
'    'CMDSQL = "SELECT * FROM TBLNEGOPTP WHERE custid = '" + lblCustId.Caption + "'"
'    'AND CUSTID NOT IN (SELECT CUSTID FROM tbllunas)"
'    CMDSQL = "SELECT DISTINCT TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.ID,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.TYPE FROM TBLNEGOPTP,tbllunas WHERE "
'    CMDSQL = CMDSQL + "tbllunas.CUSTID<>TBLNEGOPTP.CUSTID AND TBLNEGOPTP.CUSTID='" + lblCustId.Caption + "' order by TBLNEGOPTP.promisedate desc"
'Else
'    CMDSQL = "SELECT distinct TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.ID,TBLNEGOPTP.TYPE "
'    CMDSQL = CMDSQL + "FROM VWLISTPTP,TBLNEGOPTP WHERE TBLNEGOPTP.CUSTID=VWLISTPTP.CUSTID AND "
'    CMDSQL = CMDSQL + "VWLISTPTP.PAYDATE<TBLNEGOPTP.PROMISEDATE AND TBLNEGOPTP.CUSTID='" + lblCustId.Caption + "' order by TBLNEGOPTP.promisedate desc"
'End If

cmdsql = "SELECT * FROM tblnegoptp where custid = '" + lblCustId.Caption + "' "
'@@ 08-02-2012 , Tambahan untuk filter tabel negoptp
'@@ 26-03-2012 Filter Bulan dan Tahun dinonaktifkan dulu
'CMDSQL = CMDSQL + " and date_part('month',promisedate)>=date_part('month',now()) and "
'CMDSQL = CMDSQL + " date_part('year',promisedate)>=date_part('year',now()) "
cmdsql = cmdsql + " order by promisedate desc"

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstPayment.ListItems.clear
Dim n As Currency
While Not showlist.EOF
    Set ListItem = LstPayment.ListItems.ADD(, , "")
        ListItem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
        ListItem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "dd/mm/yyyy")))
        ListItem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", Round((showlist!PromisePay), 0)))
        n = n + Val(ListItem.SubItems(3))
        If n <= TOTPTP Then
            ListItem.ListSubItems(1).ForeColor = vbRed
            ListItem.ListSubItems(2).ForeColor = vbRed
            ListItem.ListSubItems(3).ForeColor = vbRed
        End If
        
        ListItem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
        ListItem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "dd/mm/yyyy")))
     showlist.MoveNext
Wend

Set showlist = Nothing
End Sub
Public Sub show_cust()
Dim ListItem As ListItem
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_cust1 As ADODB.Recordset
Dim m_cust2 As ADODB.Recordset
Dim cmdsql As String
Dim CMDSQL2 As String
Dim sPending As String
Dim CEKREC As New ADODB.Recordset
'On Error GoTo HELL:
'CMDSQL = "SELECT mgm.*, mgm_DETAIL.* FROM mgm INNER JOIN "
'CMDSQL = CMDSQL + "mgm_DETAIL ON mgm.CUSTID = dbo.mgm_DETAIL.CUSTID"

cmdsql = "select * from mgm"
'CMDSQL2 = "select * from mgm_detail"





Set m_cust = New ADODB.Recordset
'Set m_cust2 = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
'm_cust2.CursorLocation = adUseClient
If shedulePTP_Show = True Then
    cmdsql = cmdsql + " where custid ='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
' Tambahan untuk reminder AGENT 27 Mei 2013 By Izuddin
ElseIf bReminder_agent = True Or bAktif_Cust_Review = True Then
    cmdsql = cmdsql + " where custid ='" & sReminder_CUST_ID & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
' +++++++++++++++++++++++++++++++++++++++++++++++++++++
Else
    cmdsql = cmdsql + " where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'CMDSQL2 = CMDSQL2 + " where custid ='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'"
    'm_cust2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'm_cust.Open "Select * from mgm where custid='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If

'tampilkan data tabel mgm
If Not m_cust.EOF Then
    
    On Error Resume Next
    
     
    '@@31052012 Buat Menyimpan Status Account
    StatusAccount = ""
    StatusAccount = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    
    '@@ 07-05-2012, Buat menandakan bahwa nomor tersebut UnValid Number
    If m_cust("f_unvalid_home1") = "1" Then
        txtHomeNo1A.BackColor = &HC0C0&
        txtHomeNo1.BackColor = &HC0C0&
        txtHomeNo1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home1")), "(Null)", m_cust("f_sts_unvalid_home1"))
        txtHomeNo1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home1")), "(Null)", m_cust("f_sts_unvalid_home1"))
    End If
    If m_cust("f_unvalid_home2") = "1" Then
        txtHomeNo2A.BackColor = &HC0C0&
        txtHomeNo2.BackColor = &HC0C0&
        txtHomeNo2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home2")), "(Null)", m_cust("f_sts_unvalid_home2"))
        txtHomeNo2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home2")), "(Null)", m_cust("f_sts_unvalid_home2"))
    End If
    If m_cust("f_unvalid_office1") = "1" Then
        txtOfficeNo1A.BackColor = &HC0C0&
        txtOfficeNo1.BackColor = &HC0C0&
        txtOfficeNo1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office1")), "(Null)", m_cust("f_sts_unvalid_office1"))
        txtOfficeNo1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office1")), "(Null)", m_cust("f_sts_unvalid_office1"))
    End If
    If m_cust("f_unvalid_office2") = "1" Then
        txtOfficeNo2A.BackColor = &HC0C0&
        txtOfficeNo2.BackColor = &HC0C0&
        txtOfficeNo2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office2")), "(Null)", m_cust("f_sts_unvalid_office2"))
        txtOfficeNo2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office2")), "(Null)", m_cust("f_sts_unvalid_office2"))
    End If
    If m_cust("f_unvalid_mobile1") = "1" Then
        txtMobileNo1A.BackColor = &HC0C0&
        txtMobileNo1.BackColor = &HC0C0&
        txtMobileNo1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile1")), "(Null)", m_cust("f_sts_unvalid_mobile1"))
        txtMobileNo1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile1")), "(Null)", m_cust("f_sts_unvalid_mobile1"))
    End If
    If m_cust("f_unvalid_mobile2") = "1" Then
        txtMobileNo2A.BackColor = &HC0C0&
        txtMobileNo2.BackColor = &HC0C0&
        txtMobileNo2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile2")), "(Null)", m_cust("f_sts_unvalid_mobile2"))
        txtMobileNo2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile2")), "(Null)", m_cust("f_sts_unvalid_mobile2"))
    End If
    If m_cust("f_unvalid_addhome1") = "1" Then
        txtHomeAdd1.BackColor = &HC0C0&
        txtHomeAdd1A.BackColor = &HC0C0&
        txtHomeAdd1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome1")), "(Null)", m_cust("f_sts_unvalid_addhome1"))
        txtHomeAdd1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome1")), "(Null)", m_cust("f_sts_unvalid_addhome1"))
    End If
    If m_cust("f_unvalid_addhome2") = "1" Then
        txtHomeAdd2.BackColor = &HC0C0&
        txtHomeAdd2A.BackColor = &HC0C0&
        txtHomeAdd2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome2")), "(Null)", m_cust("f_sts_unvalid_addhome2"))
        txtHomeAdd2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome2")), "(Null)", m_cust("f_sts_unvalid_addhome2"))
    End If
    If m_cust("f_unvalid_addoffice1") = "1" Then
        txtOfficeAdd1.BackColor = &HC0C0&
        txtOfficeAdd1A.BackColor = &HC0C0&
        txtOfficeAdd1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice1")), "(Null)", m_cust("f_sts_unvalid_addoffice1"))
        txtOfficeAdd1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice1")), "(Null)", m_cust("f_sts_unvalid_addoffice1"))
    End If
    If m_cust("f_unvalid_addoffice2") = "1" Then
        txtOfficeAdd2.BackColor = &HC0C0&
        txtOfficeAdd2A.BackColor = &HC0C0&
        txtOfficeAdd2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice2")), "(Null)", m_cust("f_sts_unvalid_addoffice2"))
        txtOfficeAdd2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice2")), "(Null)", m_cust("f_sts_unvalid_addoffice2"))
    End If
    If m_cust("f_unvalid_addmobile1") = "1" Then
        txtMobileAdd1.BackColor = &HC0C0&
        txtMobileAdd1A.BackColor = &HC0C0&
        txtMobileAdd1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile1")), "(Null)", m_cust("f_sts_unvalid_addmobile1"))
        txtMobileAdd1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile1")), "(Null)", m_cust("f_sts_unvalid_addmobile1"))
    End If
    If m_cust("f_unvalid_addmobile2") = "1" Then
        txtMobileAdd2.BackColor = &HC0C0&
        txtMobileAdd2A.BackColor = &HC0C0&
        txtMobileAdd2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile2")), "(Null)", m_cust("f_sts_unvalid_addmobile2"))
        txtMobileAdd2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile2")), "(Null)", m_cust("f_sts_unvalid_addmobile2"))
    End If
    If m_cust("f_unvalid_ec") = "1" Then
        txtECnoA.BackColor = &HC0C0&
        txtECno.BackColor = &HC0C0&
        txtECnoA.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_ec")), "(Null)", m_cust("f_sts_unvalid_ec"))
        txtECno.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_ec")), "(Null)", m_cust("f_sts_unvalid_ec"))
    End If
        
    '@@04-05-2012, Jika kategori call telah terisi, combo box dinonaktifkan
    If m_cust("homenoadd1") <> Empty And m_cust("stskathomeadd1") <> Empty Then
        CmbStsKatHome1.Enabled = False
    End If
    If m_cust("homenoadd2") <> Empty And m_cust("stskathomeadd2") <> Empty Then
        CmbStsKatHome2.Enabled = False
    End If
    If m_cust("officenoadd1") <> Empty And m_cust("stskatofficeadd1") <> Empty Then
        CmbStsKatOffice1.Enabled = False
    End If
    If m_cust("officenoadd2") <> Empty And m_cust("stskatofficeadd2") <> Empty Then
        CmbStsKatOffice2.Enabled = False
    End If
    If m_cust("mobilenoadd1") <> Empty And m_cust("stskathpadd1") <> Empty Then
        CmbStsKatHP1.Enabled = False
    End If
    If m_cust("mobilenoadd2") <> Empty And m_cust("stskathpadd2") <> Empty Then
        CmbStsKatHP2.Enabled = False
    End If
    
    '@@03-05-2012 buat nambahin tooltip dari keterangan nomor yang di black list
    Dim m_objrs_tooltip As ADODB.Recordset
    
    '@@220610 - Memberikan tanda merah pada no telepon yang di blacklist
    If m_cust("f_homeno") = 1 Then
        txtHomeNo1.ForeColor = vbRed
        txtHomeNo1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homeno") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("homeno"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeNo1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeNo1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    If m_cust("f_homeno2") = 1 Then
        txtHomeNo2.ForeColor = vbRed
        txtHomeNo2A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homeno2") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("homeno2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeNo2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeNo2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_officeno") = 1 Then
        txtOfficeNo1.ForeColor = vbRed
        txtOfficeNo1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officeno") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("officeno"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeNo1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeNo1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_officeno2") = 1 Then
        txtOfficeNo2.ForeColor = vbRed
        txtOfficeNo2A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officeno2") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("officeno2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeNo2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeNo2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobileno") = 1 Then
        txtMobileNo1.ForeColor = vbRed
        txtMobileNo1A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobileno") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("mobileno"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtMobileNo1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileNo1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobileno2") = 1 Then
        txtMobileNo2.ForeColor = vbRed
        txtMobileNo2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobileno2") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("mobileno2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtMobileNo2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileNo2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_homenoadd1") = 1 Then
        txtHomeAdd1.ForeColor = vbRed
        txtHomeAdd1A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homenoadd1") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("homenoadd1"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeAdd1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeAdd1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_homenoadd2") = 1 Then
        txtHomeAdd2.ForeColor = vbRed
        txtHomeAdd2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homenoadd2") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("homenoadd2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeAdd2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeAdd2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If

    If m_cust("f_officenoadd1") = 1 Then
         txtOfficeAdd1.ForeColor = vbRed
         txtOfficeAdd1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officenoadd1") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("officenoadd1"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeAdd1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeAdd1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_officenoadd2") = 1 Then
        txtOfficeAdd2.ForeColor = vbRed
        txtOfficeAdd2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officenoadd1") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("officenoadd2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeAdd2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeAdd2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobilenoadd1") = 1 Then
         txtMobileAdd1.ForeColor = vbRed
         txtMobileAdd1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobilenoadd1") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("mobilenoadd1"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtMobileAdd1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileAdd1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobilenoadd2") = 1 Then
        txtMobileAdd2.ForeColor = vbRed
        txtMobileAdd2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobilenoadd2") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("mobilenoadd2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtMobileAdd2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileAdd2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_ec_telp") = 1 Then
         txtECno.ForeColor = vbRed
         txtECnoA.ForeColor = vbRed
         '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("ec_telp") <> Empty Then
            cmdsql = "select * from tblblacklist where no_telp='"
            cmdsql = cmdsql + CStr(Trim(m_cust("ec_telp"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtECno.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtECnoA.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    
    '@@03-05-2012,,Buat Nandain Valid number -------------------------
    If m_cust("f_valid_home1") = 1 Then
        txtHomeNo1.ForeColor = vbBlue
        txtHomeNo1A.ForeColor = vbBlue
        
        txtHomeNo1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home1")), "", m_cust("f_sts_valid_home1"))
        txtHomeNo1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home1")), "", m_cust("f_sts_valid_home1"))
    End If
    If m_cust("f_valid_home2") = 1 Then
        txtHomeNo2.ForeColor = vbBlue
        txtHomeNo2A.ForeColor = vbBlue
        
        txtHomeNo2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home2")), "", m_cust("f_sts_valid_home2"))
        txtHomeNo2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home2")), "", m_cust("f_sts_valid_home2"))
    End If
    If m_cust("f_valid_office1") = 1 Then
        txtOfficeNo1.ForeColor = vbBlue
        txtOfficeNo1A.ForeColor = vbBlue
        
        txtOfficeNo1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office1")), "", m_cust("f_sts_valid_office1"))
        txtOfficeNo1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office1")), "", m_cust("f_sts_valid_office1"))
    End If
    If m_cust("f_valid_office2") = 1 Then
        txtOfficeNo2.ForeColor = vbBlue
        txtOfficeNo2A.ForeColor = vbBlue
        
        txtOfficeNo2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office2")), "", m_cust("f_sts_valid_office2"))
        txtOfficeNo2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office2")), "", m_cust("f_sts_valid_office2"))
    End If
    If m_cust("f_valid_mobile1") = 1 Then
        txtMobileNo1.ForeColor = vbBlue
        txtMobileNo1A.ForeColor = vbBlue
        
        txtMobileNo1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile1")), "", m_cust("f_sts_valid_mobile1"))
        txtMobileNo1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile1")), "", m_cust("f_sts_valid_mobile1"))
    End If
    If m_cust("f_valid_mobile2") = 1 Then
        txtMobileNo2.ForeColor = vbBlue
        txtMobileNo2A.ForeColor = vbBlue
        
        txtMobileNo2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile2")), "", m_cust("f_sts_valid_mobile2"))
        txtMobileNo2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile2")), "", m_cust("f_sts_valid_mobile2"))
    End If
    
    If m_cust("f_valid_addhome1") = 1 Then
        txtHomeAdd1.ForeColor = vbBlue
        txtHomeAdd1A.ForeColor = vbBlue
        
        txtHomeAdd1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome1")), "", m_cust("f_sts_valid_addhome1"))
        txtHomeAdd1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome1")), "", m_cust("f_sts_valid_addhome1"))
    End If
    If m_cust("f_valid_addhome2") = 1 Then
        txtHomeAdd2.ForeColor = vbBlue
        txtHomeAdd2A.ForeColor = vbBlue
        
        txtHomeAdd2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome2")), "", m_cust("f_sts_valid_addhome2"))
        txtHomeAdd2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome2")), "", m_cust("f_sts_valid_addhome2"))
    End If
    If m_cust("f_valid_addoffice1") = 1 Then
        txtOfficeAdd1.ForeColor = vbBlue
        txtOfficeAdd1A.ForeColor = vbBlue
        
        txtOfficeAdd1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice1")), "", m_cust("f_sts_valid_addoffice1"))
        txtOfficeAdd1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice1")), "", m_cust("f_sts_valid_addoffice1"))
    End If
    If m_cust("f_valid_addoffice2") = 1 Then
        txtOfficeAdd2.ForeColor = vbBlue
        txtOfficeAdd2A.ForeColor = vbBlue
        
        txtOfficeAdd2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice2")), "", m_cust("f_sts_valid_addoffice2"))
        txtOfficeAdd2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice2")), "", m_cust("f_sts_valid_addoffice2"))
    End If
    If m_cust("f_valid_addmobile1") = 1 Then
        txtMobileAdd1.ForeColor = vbBlue
        txtMobileAdd1A.ForeColor = vbBlue
        
        txtMobileAdd1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile1")), "", m_cust("f_sts_valid_addmobile1"))
        txtMobileAdd1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile1")), "", m_cust("f_sts_valid_addmobile1"))
    End If
    If m_cust("f_valid_addmobile2") = 1 Then
        txtMobileAdd2.ForeColor = vbBlue
        txtMobileAdd2A.ForeColor = vbBlue
        
        txtMobileAdd2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile2")), "", m_cust("f_sts_valid_addmobile2"))
        txtMobileAdd2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile2")), "", m_cust("f_sts_valid_addmobile2"))
    End If
    If m_cust("f_valid_ec") = 1 Then
        txtECnoA.ForeColor = vbBlue
        txtECno.ForeColor = vbBlue
        
        txtECnoA.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_ec")), "", m_cust("f_sts_valid_ec"))
        txtECno.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_ec")), "", m_cust("f_sts_valid_ec"))
    End If
    '@@03-05-2012,,AKHIR Buat Nandain Valid number -------------------------
    
    
    '@@ 08 Juni 2011 SEMUA TELEPON DIBUKA,STATUS APAPUN
'    '@@ 11-04-2011 , Sementara untuk custid yang diberikan
'    If m_cust("status_additional") = "1" Then
'        Frame15(5).Visible = True
'        Frame17.Visible = True
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
'
'    '@@ 02-05-2011, untuk memunculkan additional info dan EC disesuaikan dengan status
'    'Status ON-, VL-, PR- munculkan additional info
'    '@@ 26 May 2011, bp- dan ptp- digunakan untuk memunculkan additional dan ec
'    Dim CekStatus As String
'    CekStatus = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
'    If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'        Frame15(5).Visible = True
'        Frame17.Visible = True
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
'
'    'Jika status OS maka yang ditampilkan EC saja
'    If Trim(CekStatus) = "OS-" Then
'        Frame15(5).Visible = False
'        Frame17.Visible = False
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
'
'    'Jika status account masih kosong, maka tampilkan EC
'    '@@ 11-May-2011
'    If CekStatus = "" Then
'        Frame15(5).Visible = False
'        Frame17.Visible = False
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
    
    
    '@@02-05-2012, Tambahan untuk menampilkan kategori telepon di additional phone
     CmbStsKatHome1.text = IIf(IsNull(m_cust("stskathomeadd1")), "", m_cust("stskathomeadd1"))
     CmbStsKatHome2.text = IIf(IsNull(m_cust("stskathomeadd2")), "", m_cust("stskathomeadd2"))
     CmbStsKatOffice1.text = IIf(IsNull(m_cust("stskatofficeadd1")), "", m_cust("stskatofficeadd1"))
     CmbStsKatOffice2.text = IIf(IsNull(m_cust("stskatofficeadd2")), "", m_cust("stskatofficeadd2"))
     CmbStsKatHP1.text = IIf(IsNull(m_cust("stskathpadd1")), "", m_cust("stskathpadd1"))
     CmbStsKatHP2.text = IIf(IsNull(m_cust("stskathpadd2")), "", m_cust("stskathpadd2"))
    
    
    '@@ 17-04-2012, Tambahan untuk request number
    TxtKategori.Caption = IIf(IsNull(m_cust("status_telp")), "", m_cust("status_telp"))
    TxtNoTelpReq.text = IIf(IsNull(m_cust("req_nomor_telp")), "", Trim(m_cust("req_nomor_telp")))
    
    '@@ 09042012, Tambahan untuk Status Risk Account: POP1 dan PP1
    LblPop.Caption = IIf(IsNull(m_cust("status_pop1")), "", m_cust("status_pop1"))
    LblPP.Caption = IIf(IsNull(m_cust("status_pp1")), "", m_cust("status_pp1"))

    '01-02-2012, tambahkan status hot tobe collected
    If m_cust("status_htc") = "1" Then
        CmdKeep.BackColor = vbRed
        'CmdKeep.Caption = "Hot..."
    Else
        CmdKeep.BackColor = &H8000000F
        'CmdKeep.Caption = "Not Hot..."
    End If
    
    '@@ 29-03-2012 Tambahan status risk
    If IsNull(m_cust("status_risk")) = True Then
        LblStsRisk.ForeColor = &H80000012
    End If
    If IsNull(m_cust("status_risk")) = "1" Then
        LblStsRisk.ForeColor = &HFF&
    End If
    If IsNull(m_cust("status_risk")) = "2" Then
        LblStsRisk.ForeColor = &HFFFF&
    End If
    If IsNull(m_cust("status_risk")) = "3" Then
        LblStsRisk.ForeColor = &H80FF80
    End If
    
    '@@ 04082011 Tambahan Field
     On Error Resume Next
     TxtInstallment.Value = IIf(IsNull(m_cust("instalment")), "0", m_cust("instalment"))
     Txtperiod.Caption = IIf(IsNull(m_cust("period")), "", m_cust("period"))
     TxtCurpri.Value = IIf(IsNull(m_cust("curpri")), "", m_cust("curpri"))
     lbltype.Caption = IIf(IsNull(m_cust("acc_type")), "", m_cust("acc_type"))
     lblpurge.Caption = IIf(IsNull(m_cust("sts_purge")), "", m_cust("sts_purge"))
     
     '@@ 04082011 Jika type data card instalment dan period di hide
     If (UCase(lbltype.Caption) = "CARD") Then
        Label11(9).Visible = False
        TxtInstallment.Visible = False
        
        Label11(10).Visible = False
        Txtperiod.Visible = False
     End If
    
    '@@25/01/2012
    LblResultPTP.Caption = IIf(IsNull(m_cust("result_ptp")), "", m_cust("result_ptp"))
    
    '@@ 02031011
    LblMinPayment.Value = IIf(IsNull(m_cust("minpayment")), "0", m_cust("minpayment"))

    LblStatus.Caption = IIf(IsNull(m_cust("statusprior")), "", "Status : " & m_cust("statusprior"))
    lblCustId.Caption = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    LblMother.Caption = IIf(IsNull(m_cust("mother")), "", m_cust("mother"))
    'sql = "delete  from tblnegoptp where custid in (select custid from tbllunas where custid ='" + IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")) + "')"
    TxtCustid.text = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    TxtName.text = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblaoc.Caption = IIf(IsNull(m_cust("agent")), "", m_cust("Agent"))
    LblInterest.Caption = Format(IIf(IsNull(m_cust("INTEREST")), "0", m_cust("INTEREST")), "##,###")
    LblFees.Caption = Format(IIf(IsNull(m_cust("FEES")), "0", m_cust("FEES")), "##,###")
    lblregion.Caption = IIf(IsNull(m_cust("region")), "", m_cust("region"))
    
    '@@ 04082011 Komponennya dibuang
    'lblaging.Caption = IIf(IsNull(m_cust("Aging")), "            ", m_cust("Aging"))
    
    'lblwilling.Caption = IIf(IsNull(m_cust("Willing_Ness")), "              ", m_cust("Willing_Ness"))
    lblRecsource.Caption = IIf(IsNull(m_cust("RECSOURCE")), "", m_cust("RECSOURCE"))
    LBLEXP.Caption = IIf(IsNull(m_cust("date_into_clas")), "", "Expire date " & Format(DateAdd("d", 60, m_cust("date_into_clas")), "dd-mm-yyyy"))
    
    '@@ 04082011 Dibuang
    'LblRiskLevel.Caption = IIf(IsNull(m_cust("RiskLevel")), "", m_cust("RiskLevel"))
    
    'lblPriority.Caption = IIf(IsNull(m_cust("Priority")), "", m_cust("Priority"))
    lblNama.Caption = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblCardNo.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblID.Caption = IIf(IsNull(m_cust("ktpno")), "", m_cust("ktpno"))
    'lblDate.Value = IIf(IsNull(m_cust("BIRTHD")), "", Format(m_cust("BIRTHD"), "dd-mmm-yyyy"))
    LblDOB.Caption = IIf(IsNull(m_cust("DOB")), "", Left(m_cust("DOB"), 10))
    lblAddr.text = IIf(IsNull(m_cust("ADDRNOW")), "", m_cust("ADDRNOW"))
    TDB_cur_bal = IIf(IsNull(m_cust("CURBAL")), "", m_cust("CURBAL"))
    TXTRUMUS.text = IIf(IsNull(m_cust("typerumus")), "", m_cust("typerumus"))
    Combo1.text = IIf(IsNull(m_cust("stscallcust")), "", m_cust("stscallcust"))
    'tdbmaxad.Value = Format(IIf(IsNull(m_cust("maxad")), "0", m_cust("maxad")), "##,###")
    'tdbminad.Value = Format(IIf(IsNull(m_cust("minad")), "0", m_cust("minad")), "##,###")
    
    TxtInterest.Value = IIf(IsNull(m_cust("interest")), "", m_cust("interest"))
     
    '@@ Tambahan 2 field (map dan cycle)
    LblMap = IIf(IsNull(m_cust("map")), "0", m_cust("map"))
    LblCycle = IIf(IsNull(m_cust("cycle")), "0", m_cust("cycle"))

   Set CEKREC = New ADODB.Recordset
    CEKREC.CursorLocation = adUseClient
    CEKREC.Open "select * from opening_screen where custid='" + lblCustId.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    '@@ 12-10-2011, Blink OST dinonaktifkan
'    If CEKREC.RecordCount > 0 Then
'        'SSCommand1(7).BackColor = vbRed
'        TimerBlink.Enabled = True
'    Else
'        TimerBlink.Enabled = False
'    End If
    
     If InStr(1, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(3), "DE") > 0 Then
        txthasil.Visible = True
     Else
        txthasil.Visible = False
     End If
     
     Text6.text = IIf(IsNull(m_cust("disapp")), "0", m_cust("disapp"))
     
     '@@03-05-2012 DinonAktifkan
     'tdbhptrace.Value = IIf(IsNull(m_cust("hp1trace")), "", m_cust("hp1trace"))
     
     tdbtelptrace.Value = IIf(IsNull(m_cust("tlp1trace")), "", m_cust("tlp1trace"))
     txtremarkstrace.text = IIf(IsNull(m_cust("addrtrace")), "", m_cust("addrtrace"))
     
     bcekptp = False
    vrcek = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    
    '@@03062014 Catet Tanggal Paid Off
    TanggalPaidOff = IIf(IsNull(m_cust("tgl_paid_off")), "", m_cust("tgl_paid_off"))
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
     
    '@@ 04-03-2011 Ubah status jika TL/SPV/Admin yang buka dapat membuka semua status
    If UCase(Trim(MDIForm1.Text2.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.Text2.text)) = "TEAMLEADER" Then
       
        If vrcek <> "BP-" Or Mid(vrcek, 1, 3) = "PTP" Or Mid(vrcek, 1, 3) = "POP" Then
            Strsql = "Select * from contacteddesc WHERE status=1"
        ElseIf vrcek = "BP-" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','PO-','CO-') AND status=1"
        ElseIf Mid(vrcek, 1, 3) = "PTP" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','PO-','CO-') AND status=1"
        ElseIf Mid(vrcek, 1, 3) = "POP" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('POP') AND status=1"
        End If
        
    Else
    '@@ 04-03-2011 Nah ini jika yang login Agent
        If vrcek = "" Then
            Strsql = "Select * from contacteddesc WHERE status=1"
        Else
            If vrcek = "VL-" Then
                Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','ON-','PO-','CO-') and status=1"
            ElseIf vrcek = "OS-" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','ON-','SK-','PO-','CO-') AND status=1"
            ElseIf vrcek = "PR-" Then
                 Strsql = "Select * from contacteddesc WHERE  substring(KdNoProdPresented,1,3) in('PR-','ON-','PO-','CO-') AND status=1"
            ElseIf vrcek = "ON-" Then
                 Strsql = "Select * from contacteddesc WHERE  substring(KdNoProdPresented,1,3) in('ON-','PO-','CO-') AND status=1"
            ElseIf vrcek = "SK-" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','SK-','PO-','CO-') AND status=1"
            ElseIf vrcek = "SP-" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('SP-','PO-','CO-') AND status=1"
            ElseIf vrcek = "BP-" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','PO-','CO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "PTP" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','PO-','CO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "POP" Then
                 Strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('POP') AND status=1"
            '@@31052012Tambahan JIKA PAID OFF (PO-) DAN COMPLAIN (CO-)
            ElseIf Mid(vrcek, 1, 3) = "PO-" Then
                Strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('PO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "CO-" Then
                Strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('CO-') AND status=1"
            Else
                Strsql = " Select * from contacteddesc WHERE status=1 "
            End If
            
        End If
    End If
    'STRSQL = " Select * from contacteddesc WHERE status=1 "
    cboaccount.clear
    M_Objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_Objrs.EOF
        cboaccount.AddItem M_Objrs!KdNoProdPresented
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
'    '@@31-05-2012 Tambahan 2 Status Account, PAID OFF dan COMPLAIN
'    cboaccount.AddItem "PAID-OFF"
'    cboaccount.AddItem "COMPLAIN"
    
   If Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) <> "PTP" Then
    'cboaccount.Text = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    cboaccount.text = IIf(IsNull(m_cust("kethslkerja_new")), "", m_cust("kethslkerja_new"))
   ElseIf Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) = "PTP" Then
     cboPTP.text = IIf(IsNull(m_cust("kethslkerja_new")), "", m_cust("kethslkerja_new"))
     cboaccount = IIf(IsNull(m_cust("ptpdesc")), "", m_cust("ptpdesc"))
   End If
  
  
   
   If Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) = "PTP" Then
        C_PTP.Value = vbChecked
        '@@ 05-10-2011 Tambahan field PTP VIA
        CmbViaPtp.text = IIf(IsNull(m_cust("ptpvia")), "", m_cust("ptpvia"))
   End If
   
   If Trim(Mid(cboaccount, 1, 3)) = "POP" Or Trim(Mid(cboaccount, 1, 2)) = "BP" Then
       '@@ 05-10-2011 Tambahan field PTP VIA
        CmbViaPtp.text = IIf(IsNull(m_cust("ptpvia")), "", m_cust("ptpvia"))
   End If
   
   
   
 TglPTPNew = IIf(IsNull(m_cust("tglptpnew")), "", m_cust("tglptpnew"))
  If TglPTPNew <> "" Then
        'tdbptpnew.Value = Format(tglptpnew, "dd/mm/yyyy")
        tdbptpnew.Value = Format(m_cust("tglptpnew"), "mm/dd/yyyy")
  End If
  
If Left(vrcek, 3) = "PTP" Then
        SSCommand1(4).Visible = True
        Label43(2).Visible = True
Else
        SSCommand1(4).Visible = False
        Label43(2).Visible = False
End If

    If Left(vrcek, 2) = "BP" Then
  '  cboPOPSP.Enabled = False
'        FrmContacted.Enabled = False
'        C_Contacted.Enabled = False
'        cmbContacted.Enabled = False
'        cmbDescCon.Enabled = False
     End If
    
    lblOfficeAddr.text = IIf(IsNull(m_cust("ADDRPT")), "", m_cust("ADDRPT"))
    lblZIP.Caption = IIf(IsNull(m_cust("ZIPNOW")), "", m_cust("ZIPNOW"))
    '@@04082011 NoCard dihapus dulu
    'lblNoCard.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    
    lblNoPay.Caption = IIf(IsNull(m_cust("NoPay")), "", m_cust("NoPay"))
      
       
        
        
        
    'Else
    
     LblPrompA.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
     
        
   If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
        If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
            LblPrompA.Visible = False
            Label11(8).Visible = False
        Else
            LblPrompA.Visible = True
            Label11(8).Visible = True
       End If
    Else
          LblPrompA.Visible = False
          Label11(8).Visible = False
    End If
    
   ' End If
    
    '@@ 0408201 Dibuang
    'tdbprincipal.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
    
    lblOpenDate.Value = IIf(IsNull(m_cust("OpenDate")), "", m_cust("OpenDate"))
    lblLastBill.Value = IIf(IsNull(m_cust("LastBill")), "", m_cust("LastBill"))
    lblLcAtm.Value = IIf(IsNull(m_cust("LcATMP")), "", m_cust("LcATMP"))
    txttenor.Value = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    vrtenor = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    lblBrokenPromised.Caption = IIf(IsNull(m_cust("BrokenPromise")), "", m_cust("BrokenPromise"))
    lblBD.Value = IIf(IsNull(m_cust("B_D")), "", m_cust("B_D"))
    lblLimit.Value = IIf(IsNull(m_cust("Limit")), "", m_cust("Limit"))
    vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
    vrcekamont = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
    If listview1(0).ListItems.Count = 0 Then
    lblPayDt.Value = IIf(IsNull(m_cust("Pay_Dt")), "", m_cust("Pay_Dt"))
    End If
    
    
    If listview1(0).ListItems.Count = 0 Then
    lblLastPay.Value = IIf(IsNull(m_cust("LastPay")), "", m_cust("LastPay"))
    End If
    
    lblTtlPay.Value = IIf(IsNull(m_cust("TtlPay")), "", m_cust("TtlPay"))
    'If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
     '   lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
     '   If lblAmount.ValueIsNull Then
      '      lblAmount.Value = 0
      '  Else
       '     lblAmount.Value = lblAmount.Value + (lblAmount.Value * 18.26) / 100
       ' End If
        
    'Else
    
    
    lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    
    'End If
    
    If lblAmount.ValueIsNull Then
    
            tdbmaxad.Value = 0
        Else
            tdbmaxad.Value = lblAmount.Value - (lblAmount.Value * 24) / 100
        End If
    
    
     If lblAmount.ValueIsNull Then
            tdbminad.Value = tdbminad.Value - (lblAmount.Value * 35) / 100
        Else
            tdbminad.Value = lblAmount.Value - (lblAmount.Value * 31) / 100
        End If
        
    Tdbbalance.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    
    ' ----------- LATE FEE -------------
    TDBlate_fee.Value = IIf(IsNull(m_cust("late_fee")), "", Format(m_cust("late_fee"), "##.##0"))
    ' ----------------------------------
    
    ' ------------ CASE DECEASE -----------
    If lblClass.Caption = "835" Then
        Command3.Enabled = False
        Label11(19).Visible = True
    End If
    
    If IIf(IsNull(m_cust("f_decease")), "", m_cust("f_decease")) = 1 Then
        Command3.Enabled = False
        Label11(19).Visible = True
    End If
    ' -------------------------------------
        
    
    txtHomeNo1.Value = IIf(IsNull(m_cust("HOMENO")), "", m_cust("HOMENO"))
    '@@03-05-2012 DinonAktifkan
    'AHome1.Value = IIf(IsNull(m_cust("AHOMENO")), "", m_cust("AHOMENO"))
    
    
    
    If IsNull(m_cust("HOMENO")) = False And m_cust("HOMENO") <> "" Then
        'txtHomeNo1A.Value = Left(m_cust("HOMENO"), Len(m_cust("HOMENO")) - 3) & "XXX"
        txtHomeNo1A.Value = Left(m_cust("HOMENO"), 4) & "BBB" & Mid(m_cust("HOMENO"), 8, 15)
        CmbPhone.AddItem "HomePhone"
    End If
    
    '@@ 03-05-2012 DinonAktifkan
    'AHome2.Value = IIf(IsNull(m_cust("AHOMENO2")), "", m_cust("AHOMENO2"))
    
    txtHomeNo2.Value = IIf(IsNull(m_cust("HOMENO2")), "", m_cust("HOMENO2"))
    If IsNull(m_cust("HOMENO2")) = False And m_cust("HOMENO2") <> "" Then
        'txtHomeNo2A.Value = Left(m_cust("HOMENO2"), Len(m_cust("HOMENO2")) - 3) & "XXX"
        txtHomeNo2A.Value = Left(m_cust("HOMENO2"), 4) & "BBB" & Mid(m_cust("HOMENO2"), 8, 15)
        CmbPhone.AddItem "HomePhone2"
    End If
    
    '@@03-05-2012 DinonAktifkan
    'AOffice1.Value = IIf(IsNull(m_cust("AOFFICENO")), "", m_cust("AOFFICENO"))
    
    txtOfficeNo1.Value = IIf(IsNull(m_cust("OFFICENO")), "", m_cust("OFFICENO"))
    If IsNull(m_cust("OFFICENO")) = False And m_cust("OFFICENO") <> "" Then
        'txtOfficeNo1A.Value = Left(m_cust("OFFICENO"), Len(m_cust("OFFICENO")) - 3) & "XXX"
        txtOfficeNo1A.Value = Left(m_cust("OFFICENO"), 4) & "BBB" & Mid(m_cust("OFFICENO"), 8, 15)
        CmbPhone.AddItem "OfficePhone"
    End If
    
    '@@03-05-2012 DinonAktifkan
    'AOffice2.Value = IIf(IsNull(m_cust("AOFFICENO2")), "", m_cust("AOFFICENO2"))
    
    txtOfficeNo2.Value = IIf(IsNull(m_cust("OFFICENO2")), "", m_cust("OFFICENO2"))
    If IsNull(m_cust("OFFICENO2")) = False And m_cust("OFFICENO2") <> "" Then
        'txtOfficeNo2A.Value = Left(m_cust("OFFICENO2"), Len(m_cust("OFFICENO2")) - 3) & "XXX"
        txtOfficeNo2A.Value = Left(m_cust("OFFICENO2"), 4) & "BBB" & Mid(m_cust("OFFICENO2"), 8, 15)
        CmbPhone.AddItem "OfficePhone2"
    End If
    txtMobileNo1.Value = IIf(IsNull(m_cust("MOBILENO")), "", m_cust("MOBILENO"))
    If IsNull(m_cust("MOBILENO")) = False And m_cust("MOBILENO") <> "" Then
        'txtMobileNo1A.Value = Left(m_cust("MOBILENO"), Len(m_cust("MOBILENO")) - 3) & "XXX"
        txtMobileNo1A.Value = Left(m_cust("MOBILENO"), 4) & "BBB" & Mid(m_cust("MOBILENO"), 8, 15)
        CmbPhone.AddItem "Hp"
    End If
    txtMobileNo2.Value = IIf(IsNull(m_cust("MOBILENO2")), "", m_cust("MOBILENO2"))
    If IsNull(m_cust("MOBILENO2")) = False And m_cust("MOBILENO2") <> "" Then
        'txtMobileNo2A.Value = Left(m_cust("MOBILENO2"), Len(m_cust("MOBILENO2")) - 3) & "XXX"
        txtMobileNo2A.Value = Left(m_cust("MOBILENO2"), 4) & "BBB" & Mid(m_cust("MOBILENO2"), 8, 15)
        CmbPhone.AddItem "Hp2"
    End If
    
    '@@ 03-05-2012 Dinonaktifkan
    'AHomeAdd1(0).Value = IIf(IsNull(m_cust("AHOMENOADD1")), "", m_cust("AHOMENOADD1"))
    
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd2(1).Value = IIf(IsNull(m_cust("AHOMENOADD2")), "", m_cust("AHOMENOADD2"))
    
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(2).Value = IIf(IsNull(m_cust("AOFFICENOADD1")), "", m_cust("AOFFICENOADD1"))
    'AOfficeAdd(3).Value = IIf(IsNull(m_cust("AOFFICENOADD2")), "", m_cust("AOFFICENOADD2"))
   
    txtHomeAdd1.Value = IIf(IsNull(m_cust("HOMENOADD1")), "", m_cust("HOMENOADD1"))
    If IsNull(m_cust("HOMENOADD1")) = False And m_cust("HOMENOADD1") <> "" Then
        txtHomeAdd1A.Value = Left(m_cust("HOMENOADD1"), 4) & "BBB" & Mid(m_cust("HOMENOADD1"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddHome1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011 BP- dan Ptp- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddHome1"
'        End If
        '@@08-06-2011 Semua Telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@24-04-2012 Diaktifkan lagi
        CmbPhone.AddItem "AddHome1"
    Else
        txtHomeAdd1.Visible = True
        txtHomeAdd1A.Visible = False
    End If
    txtHomeAdd2.Value = IIf(IsNull(m_cust("HOMENOADD2")), "", m_cust("HOMENOADD2"))
    If IsNull(m_cust("HOMENOADD2")) = False And m_cust("HOMENOADD2") <> "" Then
        txtHomeAdd2A.Value = Left(m_cust("HOMENOADD2"), 4) & "BBB" & Mid(m_cust("HOMENOADD2"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddHome2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011, BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddHome2"
'        End If
        '@@08-06-2011 Telepon dibuka,status apa pun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@24-04-2012 Diaktifkan Lagi
        CmbPhone.AddItem "AddHome2"
    Else
        txtHomeAdd2A.Visible = False
        txtHomeAdd2.Visible = True
    End If
    txtOfficeAdd1.Value = IIf(IsNull(m_cust("OFFICENOADD1")), "", m_cust("OFFICENOADD1"))
    If IsNull(m_cust("OFFICENOADD1")) = False And m_cust("OFFICENOADD1") <> "" Then
        txtOfficeAdd1A.Value = Left(m_cust("OFFICENOADD1"), 4) & "BBB" & Mid(m_cust("OFFICENOADD1"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddOffice1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011, BP- dan PTP- ditampilkan juga
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddOffice1"
'        End If
        '@@08-06-2011 Telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddOffice1"
    Else
        txtOfficeAdd1A.Visible = False
        txtOfficeAdd1.Visible = True
    End If
    txtOfficeAdd2.Value = IIf(IsNull(m_cust("OFFICENOADD2")), "", m_cust("OFFICENOADD2"))
    If IsNull(m_cust("OFFICENOADD2")) = False And m_cust("OFFICENOADD2") <> "" Then
        
        anto = Trim(Left(m_cust("OFFICENOADD2"), 4) + " " + Mid(m_cust("OFFICENOADD2"), 8, 15))
        If Len(anto) = 0 Then
        txtOfficeAdd2A.Value = ""
        
        Else
        
        txtOfficeAdd2A.Value = Left(m_cust("OFFICENOADD2"), 4) & "BBB" & Mid(m_cust("OFFICENOADD2"), 8, 15)
        
        End If
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddOffice2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011 BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddOffice2"
'        End If
        '@@ 08-06-2011 Status telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddOffice2"
    Else
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
    End If
    txtMobileAdd1.Value = IIf(IsNull(m_cust("MOBILENOADD1")), "", m_cust("MOBILENOADD1"))
    If IsNull(m_cust("MOBILENOADD1")) = False And m_cust("MOBILENOADD1") <> "" Then
        txtMobileAdd1A.Value = Left(m_cust("MOBILENOADD1"), 4) & "BBB" & Mid(m_cust("MOBILENOADD1"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddMobile1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011 BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddMobile1"
'        End If
        '@@ 08-06-2011 Status Telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddMobile1"
    Else
        txtMobileAdd1.Visible = True
        txtMobileAdd1A.Visible = False
    End If
    txtMobileAdd2.Value = IIf(IsNull(m_cust("MOBILENOADD2")), "", m_cust("MOBILENOADD2"))
    If IsNull(m_cust("MOBILENOADD2")) = False And m_cust("MOBILENOADD2") <> "" Then
        txtMobileAdd2A.Value = Left(m_cust("MOBILENOADD2"), 4) & "BBB" & Mid(m_cust("MOBILENOADD2"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddMobile2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011, BP- dan PTP- ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddMobile2"
'        End If
        '@@ 08-06-2011, status telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddMobile2"
    Else
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
    End If
   
    AddrNow.text = IIf(IsNull(m_cust("TxtPtpAddr")), "", m_cust("TxtPtpAddr"))
    LblLunas.Caption = IIf(IsNull(m_cust!tgllunas), "", "TELAH LUNAS")
    TxtEC.text = IIf(IsNull(m_cust!ec_name), "", m_cust!ec_name)
    txtECno.Value = IIf(IsNull(m_cust!ec_telp), "", m_cust!ec_telp)
    If IsNull(m_cust("ec_telp")) = False And m_cust("ec_telp") <> "" Then
        txtECnoA.Value = Left(m_cust("ec_telp"), 4) & "BBB" & Mid(m_cust("ec_telp"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "EconPhone"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- dan kosong maka Additional&EC di tampilkan
'        '@@26 May 2011 BP- dan PTP juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "OS-" Or CekStatus = "" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "EconPhone"
'        End If
        '@@ 08-06-2011, Telepon dibuka status apapun
        CmbPhone.AddItem "EconPhone"
    Else
        txtECnoA.Visible = False
        txtECno.Visible = True
    End If
    
    '@@02-05-2011  Tambahan Additional
    TxtAdditional.Value = IIf(IsNull(m_cust("telp_additional")), "", m_cust("telp_additional"))
     If UCase(MDIForm1.Text2.text) = "AGENT" Then
            TxtAdditional.Enabled = False
        End If
    If TxtAdditional <> "" Then
        If UCase(MDIForm1.Text2.text) = "AGENT" Then
            TxtAdditional.Enabled = False
        End If
        '@@17-04-2012 Telepon di Non aktifkan
        '@@02052012 Diaktifkan Lagi
        CmbPhone.AddItem "TelpAdditional"
    End If
    
    '@@17-04-2012,Tambahan
    If TxtNoTelpReq.Value <> "" Then
        CmbPhone.AddItem TxtKategori.Caption
    End If
    
    txtECAdd.text = IIf(IsNull(m_cust!ECAddr), "", m_cust!ECAddr)
    cbolastcall.text = IIf(IsNull(m_cust!statuscall), "", Trim(m_cust!statuscall))
    cbolastcall.text = IIf(IsNull(m_cust!stscallwith), "", m_cust!stscallwith)
'    If cbolastcall.Text = "" Then
'        Call isi_lastcall
'    End If
' cari extension
    If InStr(1, txtOfficeNo1.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt1.Text = Right(txtOfficeNo1.Value, Len(txtOfficeNo1.Value) - InStr(1, txtOfficeNo1.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeNo2.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt2.Text = Right(txtOfficeNo2.Value, Len(txtOfficeNo2.Value) - InStr(1, txtOfficeNo2.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt3.Text = Right(txtOfficeAdd1.Value, Len(txtOfficeAdd1.Value) - InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt4.Text = Right(txtOfficeAdd2.Value, Len(txtOfficeAdd2.Value) - InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare))
    End If
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
        If Len(txtHomeAdd1.Value) > 2 Then
            txtHomeAdd1.ReadOnly = True
        End If
        If Len(txtHomeAdd2.Value) > 2 Then
            txtHomeAdd2.ReadOnly = True
        End If
        If Len(txtOfficeAdd1.Value) > 2 Then
            txtOfficeAdd1.ReadOnly = True
        End If
        If Len(txtOfficeAdd2.Value) > 2 Then
            txtOfficeAdd2.ReadOnly = True
        End If
        If Len(txtMobileAdd1.Value) > 2 Then
            txtMobileAdd1.ReadOnly = True
        End If
        If Len(txtMobileAdd2.Value) > 2 Then
            txtMobileAdd2.ReadOnly = True
        End If
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
    End If
   
    
    sPending = CStr(Trim(IIf(IsNull(m_cust!f_Pending), "", m_cust!f_Pending)))
     If sPending = "Pending" Then
         'chkAppv(0).Value = 0 '@@ 25/01/2012 Komponen Tak Terpakai
    End If
    
'    Select Case m_cust!RECSTATUS
'        Case "V"
'            C_VALID.Value = 1
'            cbovalid.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cbodescvalid.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'        Case "N"
'            C_NotContacted.Value = 1
'            cmbUncontacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cmbDescUn.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'        Case "C"
'            C_Contacted.Value = 1
'            kontak = True
'            If MDIForm1.Text2 = "Agent" Then
'                If Left(vrcek, 3) = "POP" Then
'                    C_SKIP.Enabled = False
'                    C_VALID.Enabled = False
'                    cboPOPSP.Enabled = False
'                    FrmPayment.Enabled = True
'                    C_Payment.Value = 1
'                End If
'            End If
'            cmbContacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'      Case "P"
'            C_PTP.Value = 1
'            cboPTP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'            If MDIForm1.Text2 = "Agent" Then
'                C_VALID.Enabled = False
'                C_Contacted.Enabled = False
'                FrMValid.Enabled = False
'                C_SKIP.Enabled = False
'                FrmSKIP.Enabled = False
'            End If
'         Case "S"
'            C_SKIP.Value = 1
'            cboskip.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cbodescskip.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'         Case "O"
'            'C_POPSP.Value = 1
'            cboPOPSP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))      cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'     End Select
     
    If MDIForm1.Text2 = "Agent" Then
'        If IIf(IsNull(m_cust!RECSTATUS), "", m_cust!RECSTATUS) <> "O" Then
'            frmpopsp.Enabled = False
'           cboPOPSP.Enabled = False
'        End If
    End If
        If IIf(IsNull(m_cust!f_cek_new), "", Left(m_cust!f_cek_new, 3)) = "PTP" Or Left(m_cust!f_cek_new, 3) = "POP" Or Left(m_cust!f_cek_new, 3) = "SP-" Or Left(m_cust!f_cek_new, 3) = "PRE" Then
            C_Payment.Value = 1
            TdbPTP.Value = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrtdbdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            TDBDate3.Value = IIf(IsNull(m_cust!dateptp), "", Format(m_cust!dateptp, "mm/dd/yyyy"))
            vrnewdate = IIf(IsNull(m_cust!dateptp), "", Format(m_cust!dateptp, "dd/mm/yyyy"))
            txtPayment.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            vrttlptp = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            Tdabamoint.Value = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            TxtPayment2.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp) 'tampilkan di detail payment
            cmbDiscount.text = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            vrdiskon = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            CmbBaseOn.text = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            vrbaseon = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            'TdbDatePTP.Value = IIf(IsNull(m_cust!TGLINCOMING), "", m_cust!TGLINCOMING)
            
            '@@25/01/2012 Tambahan, tambahkan data tanggal tagih
            TdbTglTagih.Value = IIf(IsNull(m_cust!tgl_tagih), "", Format(m_cust!tgl_tagih, "mm/dd/yyyy"))
        Else
        End If
End If
Call Custid_Double
'Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'", MDIForm1.Text2.Text)
Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'")
While Not m_cust1.EOF
    'Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("TGL"), 4) & "/" & Mid(m_cust1("TGL"), 5, 2) & "/" & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 7, 2)) & " " & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 9, 2)) & ":" & Right(m_cust1("TGL"), 2))
     Set ListItem = listview1(1).ListItems.ADD(, , Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss"))
        ListItem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
        ListItem.SubItems(2) = IIf(IsNull(m_cust1("user_log")), "", m_cust1("user_log"))
        ListItem.SubItems(3) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
        ListItem.SubItems(4) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
        ListItem.SubItems(5) = IIf(IsNull(m_cust1("statuscall")), "", m_cust1("statuscall"))
        ListItem.SubItems(6) = IIf(IsNull(m_cust1("ststelpwith")), "", m_cust1("ststelpwith"))
        ListItem.SubItems(7) = IIf(IsNull(m_cust1("id")), "", m_cust1("id"))
        'listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
m_cust1.MoveNext
Wend


Call isi_datapayment
Call Show_NEGOPTP
Call Show_Reserve
Call Show_Visit
Call Isi_listScript
Call Isi_SendSMS

Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient

'@@ 22-09-2011, penghitungan total payment di tabel lunas juga memperhatikan tgl data masuk
'total payment yang masuk adalah payment yang paydate-nya harus lebih besar dari data yang masuk
'CMDSQL = "Select custid, sum(payment) as jml from tbllunas where custid = '" + lblCustId.Caption + "' GROUP BY CUSTID"
cmdsql = "select sum(payment) as jml from "
cmdsql = cmdsql + "(SELECT b.custid as custid1, a.CUSTID,a.PayDate, "
cmdsql = cmdsql + " a.Payment,a.Agent,a.FieldName,a.Id from tbllunas a "
cmdsql = cmdsql + " inner join mgm b on "
cmdsql = cmdsql + " a.custid=b.custid  WHERE a.custid='"
cmdsql = cmdsql + lblCustId.Caption + "'  and date(a.Paydate)+1  > b.tglsource  order by a.PayDate asc ) as c"

M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_Objrs.EOF
        TxtAfterPay.Value = IIf(IsNull(M_Objrs("jml")), 0, M_Objrs("jml"))
        M_Objrs.MoveNext
Wend
 
 'hitung sisa hutang
 txtSisaHutang.Value = Val(TxtPayment2.Value) - Val(TxtAfterPay.Value)
 
 '---------->> hitung PRINCIPLE & AMOUNTWO  after pay  <<-----------------
 If TxtAfterPay.Value = 0 Then
    '@@04082011 Principle dibuang
    'txtPrinciple_A.Value = 0
    
    txtAmountwo_A.Value = 0
    Else
    If LblPrompA.ValueIsNull Or lblAmount.ValueIsNull Then
    Exit Sub
    End If
  '@@04082011 Principle dibuang
  'txtPrinciple_A.Value = Val(LblPrompA.Value) - Val(TxtAfterPay.Value)
  
  txtAmountwo_A.Value = Val(lblAmount.Value) - Val(TxtAfterPay.Value)
 End If
 
    If lblAmount.ValueIsNull Then
           '@@04082011 Dibuang
           'Woafter.Value = 0
       Else
           '@@04082011 Dibuang
           'Woafter.Value = lblAmount - TxtAfterPay.Value
    End If
  
    If listview1(0).ListItems.Count <> 0 Then
          '@@ 27-07-2011 , dimatiin dulu nih, cznya pay_dtnya jadi ke ambil dari payment disini
          'lblPayDt.Value = listview1(0).ListItems(listview1(0).ListItems.Count).Text
          'lblLastPay.Value = listview1(0).ListItems(listview1(0).ListItems.Count).SubItems(1)
          
'          TxtLPDPayment.Value = ListView1(0).ListItems(ListView1(0).ListItems.Count).Text
'          TxtLPAPayment.Value = ListView1(0).ListItems(ListView1(0).ListItems.Count).SubItems(1)
            
          '@@ 14042012, Karena list payment diubah berdasarkan desc, diubah
          TxtLPDPayment.Value = listview1(0).ListItems(1).text
          TxtLPAPayment.Value = listview1(0).ListItems(1).SubItems(1)
          LBLEXP.Caption = "Expire Date " + glexp
    End If
 
 
    Set m_cust = Nothing
    Set M_Objrs = Nothing

Exit Sub
'HELL:
   'MsgBox Err.Description
' Resume
 Set M_Objrs = Nothing
Set m_cust = Nothing
End Sub

Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
  'Static StartLoc
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
  tmpString1 = SourceString
 
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      
      FindReplace = tmpString1
End Function
Private Sub Isi_SendSMS()
'@@ 11-03-2011 di remarks, cznya udah tidak diapke
'Dim satu As String
'Dim dua As String
'Dim tiga As String
'Dim empat As String
'
'
'Dim RSsms_i As ADODB.Recordset
'Set RSsms_i = New ADODB.Recordset
'
'
'satu = FindReplace(TxtMobileno1.Text, "0", "+62")
'dua = FindReplace(TxtMobileno2.Text, "0", "+62")
'tiga = FindReplace(TxtMobileAdd1.Text, "0", "+62")
'empat = FindReplace(TxtMobileAdd2, "0", "+62")
'
'cmdsql_inbox = "Select receivingdatetime, sendernumber, textdecoded from inbox where (sendernumber='" + Trim$(satu) + "' or sendernumber='" + Trim$(dua) + "' or sendernumber='" + Trim$(tiga) + "' or sendernumber='" + Trim$(empat) + "') and processed='FALSE' "
'RSsms_i.Open cmdsql_inbox, M_OBJCONN1, adOpenDynamic, adLockOptimistic
'While Not RSsms_i.EOF
's = Format(RSsms_i!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
't = Trim(RSsms_i!sendernumber)
'u = Replace(RSsms_i!textdecoded, "'", " ")
'
''u1 = Replace(KATAUBAH, "- -", "-")
'v = FindReplace(t, "+62", "0")
'
'
'
'            CMDSQL = "INSERT INTO receive_sms (tgl_terima, notelp, pesan) VALUES ('" & s & "',"
'            CMDSQL = CMDSQL + " '" + v + "',"
'            CMDSQL = CMDSQL + " '" + u + "')"
'            M_OBJCONN.Execute CMDSQL
'
'            cmdsql_update = "update inbox set processed='TRUE'  where (sendernumber='" + Trim$(satu) + "' or sendernumber='" + Trim$(dua) + "' or sendernumber='" + Trim$(tiga) + "' or sendernumber='" + Trim$(empat) + "')"
'            M_OBJCONN1.Execute cmdsql_update
'
'
'RSsms_i.MoveNext
'Wend
'
''=======================================
'Dim RSsms As ADODB.Recordset
'Set RSsms = New ADODB.Recordset
'Dim lst As listitem
'RSsms.CursorLocation = adUseClient
'If Left(TxtMobileno1, 1) <> "0" And TxtMobileno1 <> "" Then
'satua = "031" & TxtMobileno1
'Else
'satua = TxtMobileno1
'End If
'
'If Left(TxtMobileno2, 1) <> "0" And TxtMobileno2 <> "" Then
'duaa = "031" & TxtMobileno2
'Else
'duaa = TxtMobileno2
'End If
'
'If Left(TxtMobileAdd1, 1) <> "0" And TxtMobileAdd1 <> "" Then
'tigaa = "031" & TxtMobileAdd1
'Else
'tigaa = TxtMobileAdd1
'End If
'
'If Left(TxtMobileAdd2, 1) <> "0" And TxtMobileAdd2 <> "" Then
'empata = "031" & TxtMobileAdd2
'Else
'empata = TxtMobileAdd2
'End If
'
'
'CMDSQL = "Select a.*, b.custid from receive_sms a, mgm b where (a.notelp='" + satua + "' or a.notelp='" + duaa + "' or a.notelp='" + tigaa + "' or a.notelp='" + empata + "') and b.custid='" + lblCustId + "'"
'RSsms.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms.EOF
'    Set lst = LstSMS.ListItems.ADD(, , IIf(IsNull(RSsms("notelp")), "", RSsms("notelp")))
'         lst.SubItems(1) = lblNama
'         lst.SubItems(2) = IIf(IsNull(RSsms("custid")), "", RSsms("custid"))
'         lst.SubItems(3) = IIf(IsNull(RSsms("pesan")), "", RSsms("pesan"))
'         lst.SubItems(4) = IIf(IsNull(RSsms("tgl_terima")), "", RSsms("tgl_terima"))
'
'RSsms.MoveNext
'Wend
'Set RSsms = Nothing
'Text3 = LstSMS.ListItems.Count
'
''--------------------------------
'If Text4.Text <> "0" Then
'If Int(Text3) > Int(Text2) Then
'
'Dim RSsms_cek As ADODB.Recordset
'Set RSsms_cek = New ADODB.Recordset
'
'RSsms_cek.CursorLocation = adUseClient
'cmdsql_cek = "select * from receive_sms order by tgl_terima desc limit 1"
'RSsms_cek.Open cmdsql_cek, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms_cek.EOF
'MsgBox "Anda mendapatkan satu SMS baru" & vbCrLf & "No Telepon : " & RSsms_cek("notelp") & vbCrLf & "Isi Pesan : " & Trim(RSsms_cek("pesan"))
'RSsms_cek.MoveNext
'Wend
'Set RSsms_cek = Nothing
'End If
'End If
'
'Text4.Text = "1"

End Sub
Private Sub Isi_SendSMS2()

Dim RSsms2 As ADODB.Recordset
'@@ 11-03-2011 Di remarks, udah tidak diapakai

'Set RSsms2 = New ADODB.Recordset
'Dim Lst2 As listitem
'RSsms2.CursorLocation = adUseClient
'CMDSQL = "Select * from sentitems where destinationnumber='" + TxtMobileno1 + "' or destinationnumber='" + TxtMobileno2 + "' or destinationnumber='" + TxtMobileAdd1 + "' or destinationnumber='" + TxtMobileAdd2 + "'"
'RSsms2.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic
'While Not RSsms2.EOF
'    Set Lst2 = LstSMS2.ListItems.ADD(, , IIf(IsNull(RSsms2("destinationnumber")), "", RSsms2("destinationnumber")))
'         Lst2.SubItems(1) = lblNama
'         Lst2.SubItems(2) = IIf(IsNull(RSsms2("textdecoded")), "", RSsms2("textdecoded"))
'         Lst2.SubItems(3) = IIf(IsNull(RSsms2("sendingdatetime")), "", RSsms2("sendingdatetime"))
'         Lst2.SubItems(4) = lblCustId
'         'Lst.SubItems(5) = IIf(IsNull(RSsms2("receivingdatetime")), "", RSsms2("receivingdatetime"))
''
'RSsms2.MoveNext
'Wend
'Set RSsms2 = Nothing
End Sub

Private Sub Isi_listScript()
'Mengisi Data di List LstScript
Set M_Objrs = New ADODB.Recordset
M_Objrs.CursorLocation = adUseClient
M_Objrs.Open "select * from tblinformationlokasi", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_Objrs.EOF
  Set ListItem = Lstscript.ListItems.ADD(, , M_Objrs.Bookmark)
      ListItem.SubItems(1) = M_Objrs("description")
      ListItem.SubItems(2) = M_Objrs("direktori")
  M_Objrs.MoveNext
Wend
Set M_Objrs = Nothing
End Sub

Private Sub isi_datapayment()
Dim m_cust2 As New ADODB.Recordset
Dim NilaiAfterPay As Currency
Dim M_DATA As New CLS_FRMCUST_CC
Set m_cust2 = M_DATA.QUERY_HIST_PAID(M_OBJCONN, "a.custid = '" + lblCustId.Caption + "' ")
listview1(0).ListItems.clear
While Not m_cust2.EOF
    Set ListItem = listview1(0).ListItems.ADD(, , IIf(IsNull(m_cust2("Paydate")), "", Format(m_cust2("Paydate"), "yyyy-mm-dd")))
        ListItem.SubItems(1) = IIf(IsNull(m_cust2("payment")), "0", Format(m_cust2("Payment"), "##,###"))
        ListItem.SubItems(2) = IIf(IsNull(m_cust2("AGENT")), "", m_cust2("AGENT"))
        ListItem.SubItems(3) = IIf(IsNull(m_cust2("FieldName")), "", m_cust2("FieldName"))
        ListItem.SubItems(4) = IIf(IsNull(m_cust2("Id")), "0", m_cust2("Id"))
        NilaiAfterPay = NilaiAfterPay + IIf(IsNull(m_cust2("payment")), "0", m_cust2("Payment"))
    m_cust2.MoveNext
Wend

Set m_cust2 = Nothing
TxtAfterPay.Value = NilaiAfterPay
txtSisaHutang.Value = Format(TxtPayment2.Value - TxtAfterPay.Value, "##,###")
End Sub
Private Sub Show_Visit()
Dim m_cust2 As New ADODB.Recordset
Dim m_Visit As New ClsVisit
Dim jml As String
Dim cmdsql As String
Set m_cust2 = New ADODB.Recordset
cmdsql = "SELECT requestdate,visitdate,detailsR,detailsV,visitke,VisitNo,id,F_CEK_new FROM tblvisit where custid='" + lblCustId.Caption + "'"
m_cust2.CursorLocation = adUseClient
m_cust2.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'Set m_cust2 = m_Visit.SELECT_RequestVisit(M_OBJCONN, lblCustId.Caption)
LstVisit.ListItems.clear
While Not m_cust2.EOF
    Set ListItem = LstVisit.ListItems.ADD(, , IIf(IsNull(m_cust2!RequestDate), "", m_cust2!RequestDate))
        ListItem.SubItems(1) = IIf(IsNull(m_cust2!VisitDate), "", m_cust2!VisitDate)
        ListItem.SubItems(2) = Trim(IIf(IsNull(m_cust2!VisitNo), "", m_cust2!VisitNo))
        ListItem.SubItems(3) = IIf(IsNull(m_cust2!DetailsR), "", m_cust2!DetailsR)
        ListItem.SubItems(4) = IIf(IsNull(m_cust2!DetailsV), "", m_cust2!DetailsV)
        ListItem.SubItems(5) = IIf(IsNull(m_cust2!VisitKe), "0", m_cust2!VisitKe)
        ListItem.SubItems(6) = IIf(IsNull(m_cust2!ID), "0", m_cust2!ID)
        ListItem.SubItems(7) = IIf(IsNull(m_cust2!f_cek_new), "0", m_cust2!f_cek_new)
        m_cust2.MoveNext
Wend
jml = m_cust2.RecordCount + 1
TDBNumber1.Value = jml
'Select Case Jml
'Case "0"
'Combo1.Text = "I"
'Case "1"
'Combo1.Text = "II"
'Case "2"
'Combo1.Text = "III"
'Case "3"
'Combo1.Text = "IV"
'Case "4"
'Combo1.Text = "V"
'Case "5"
'Combo1.Text = "VI"
'End Select
Set m_cust2 = Nothing

End Sub
Private Sub CEK_UPDATE_PELANGGAN()

    Dim M_DATA As New CLS_FRMCUST_CC_MGM
    Dim m_Visit As New ClsVisit
    Dim pStatusHstLstCall As String
    Dim StatusPTP As String

    Dim M_Objrs As ADODB.Recordset
    Dim cmdsql_waktu As String
    Dim waktu As String
    
    
    
    cmdsql_waktu = "select now() as waktu"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql_waktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    waktu = CDate(Format(M_Objrs("waktu"), "hh:nn:ss"))
    Set M_Objrs = Nothing


    Set M_update = New ADODB.Recordset
    M_update.CursorLocation = adUseServer
    M_update.Open "Select * from mgm where custid='" & lblCustId.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            
    '@@03062014 Buat nyatet Tanggal Paid Off
    If UCase(Trim(cboaccount.text)) = "PO-PAID OFF" Then
        'Cek apakah tanggal paid off masih kosong, jika ya update tanggal paid offnya
        If TanggalPaidOff = "" Or IsNull(TanggalPaidOff) = True Then
            M_update("tgl_paid_off") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & waktu
        End If
    End If
    ' ---------------------------------------
            
    '@@02-05-2012, Buat Simpan kategori telepon
    If txtHomeAdd1.Value <> Empty Then
        M_update("stskathomeadd1") = CmbStsKatHome1.text
    End If
    If txtHomeAdd2.Value <> Empty Then
        M_update("stskathomeadd2") = CmbStsKatHome2.text
    End If
    If txtOfficeAdd1.Value <> Empty Then
        M_update("stskatofficeadd1") = CmbStsKatOffice1.text
    End If
    If txtOfficeAdd2.Value <> Empty Then
        M_update("stskatofficeadd2") = CmbStsKatOffice2.text
    End If
    If txtMobileAdd1.Value <> Empty Then
        M_update("stskathpadd1") = CmbStsKatHP1.text
    End If
    If txtMobileAdd2.Value <> Empty Then
        M_update("stskathpadd2") = CmbStsKatHP2.text
    End If
            
    '@@ 19/08/2011 Untuk telpon additional hanya boleh admin/supervisor (sebelumnya agent bisa, tapi sekrg ngga)
    If UCase(Left(MDIForm1.Text2.text, 5)) = "ADMIN" Or _
       UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        M_update("telp_additional") = IIf(IsNull(TxtAdditional.Value), "", TxtAdditional.Value)
   End If
            
    '@@03-05-2012 Dinonaktifkan
    'M_update("AHOMENOADD1") = AHomeAdd1(0).Value
    
    '@@03-05-2012 Dinonaktifkan
    'M_update("AHOMENOADD2") = AHomeAdd2(1).Value
    'M_update("AOFFICENOADD1") = AOfficeAdd(2).Value
    'M_update("AOFFICENOADD2") = AOfficeAdd(3).Value
    
    M_update!maxad = tdbmaxad.Value
    M_update!minad = tdbminad.Value
    vrcekamont = Tdabamoint.Value
    '@@ 15 Juni 2011 Tambahkan SPV dan TeamLeader juga bisa save telepon
    If UCase(Left(MDIForm1.Text2.text, 5)) = "ADMIN" Or _
       UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.Text2.text) = "TEAMLEADER" Then
        M_update("HOMENOADD1") = txtHomeAdd1.Value
        M_update("HOMENOADD2") = txtHomeAdd2.Value
        M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        M_update("MOBILENOADD1") = txtMobileAdd1.Value
        M_update("MOBILENOADD2") = txtMobileAdd2.Value
        M_update!TxtPtpAddr = AddrNow.text
        M_update!ec_name = TxtEC.text
        M_update!ec_telp = txtECno.Value
    Else
        If txtHomeAdd1A.Value = "" And txtHomeAdd1A.Visible = True Then
            M_update("HOMENOADD1") = txtHomeAdd1A.Value
        ElseIf txtHomeAdd1.Value <> "" And txtHomeAdd1.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("HOMENOADD1") = txtHomeAdd1.Value
        End If
            
        If txtHomeAdd2A.Value = "" And txtHomeAdd2A.Visible = True Then
            M_update("HOMENOADD2") = txtHomeAdd2A.Value
        ElseIf txtHomeAdd2.Value <> "" And txtHomeAdd2.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("HOMENOADD2") = txtHomeAdd2.Value
        ElseIf txtHomeAdd2.Value = "" And txtHomeAdd2.Visible = True Then
            M_update("HOMENOADD2") = txtHomeAdd2.Value
        End If
                
        If txtOfficeAdd1A.Value = "" And txtOfficeAdd1A.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1A.Value
        ElseIf txtOfficeAdd1.Value <> "" And txtOfficeAdd1.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        ElseIf txtOfficeAdd1.Value = "" And txtOfficeAdd1.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        End If
                
        If txtOfficeAdd2A.Value = "" And txtOfficeAdd2A.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2A.Value
        ElseIf txtOfficeAdd2.Value <> "" And txtOfficeAdd2.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        ElseIf txtOfficeAdd2.Value = "" And txtOfficeAdd2.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        End If
            
        If txtMobileAdd1A.Value = "" And txtMobileAdd1A.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1A.Value
        ElseIf txtMobileAdd1.Value <> "" And txtMobileAdd1.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("MOBILENOADD1") = txtMobileAdd1.Value
        ElseIf txtMobileAdd1.Value = "" And txtMobileAdd1.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1.Value
        End If
            
        If txtMobileAdd2A.Value = "" And txtMobileAdd2A.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2A.Value
        ElseIf txtMobileAdd2.Value <> "" And txtMobileAdd2.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("MOBILENOADD2") = txtMobileAdd2.Value
        ElseIf txtMobileAdd2.Value = "" And txtMobileAdd2.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2.Value
        End If
            
        M_update!TxtPtpAddr = AddrNow.text
        M_update!ec_name = TxtEC.text
        M_update!ECAddr = txtECAdd.text
                 
        If txtECnoA.Value = "" And txtECnoA.Visible = True Then
            M_update("ec_telp") = txtECnoA.Value
        ElseIf txtECno.Value <> "" And txtECno.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update!ec_telp = txtECno.Value
        End If
    End If
        
    If UCase(MDIForm1.Text2.text) = "AGENT" Then
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
        If Len(txtHomeAdd1.Value) > 2 Then
            txtHomeAdd1.ReadOnly = True
        End If
        If Len(txtHomeAdd2.Value) > 2 Then
            txtHomeAdd2.ReadOnly = True
        End If
        If Len(txtOfficeAdd1.Value) > 2 Then
            txtOfficeAdd1.ReadOnly = True
        End If
        If Len(txtOfficeAdd2.Value) > 2 Then
            txtOfficeAdd2.ReadOnly = True
        End If
        If Len(txtMobileAdd1.Value) > 2 Then
            txtMobileAdd1.ReadOnly = True
        End If
        If Len(txtMobileAdd2.Value) > 2 Then
            txtMobileAdd2.ReadOnly = True
        End If
    End If
    
    '@@121110 Tambahan nih buat nyatet history perubahan status account
    If (IsNull(M_update!tglcall)) = True Then
        tglcalllalu = ""
    Else
        tglcalllalu = CStr(M_update("tglcall"))
    End If
        
    '@@ 05-10-2011, Jika status account=PTP or POP maka catat via mana dia bayarnya
    If Trim(Mid(cboaccount, 1, 3)) = "POP" Or Trim(Mid(cboaccount, 1, 2)) = "BP" Then
        M_update!ptpvia = IIf(IsNull(CmbViaPtp.text), "", CmbViaPtp.text)
    End If
        
        
    M_update("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & waktu
    'sebelum f_cek diubah statusnya
    StatusPTP = IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)

    Dim StatusAccCurrent As String  '@@ 121110 tambahan nih buat nyatet history f_cek_new
        
    If C_PTP.Value = vbChecked Then
        GoTo keptp
    End If
        
    If cboaccount.text <> "" Then
        pStatusLstCall = cboaccount.text
        M_update!f_cek_new = Left(cboaccount.text, 3)
        'txtResult.Text = pStatusLstCall '@@15/01/2012 KOmponen Tidak Terpakai
        '@@121110 tambahan buat nyatet history f_cek_new
        StatusAccCurrent = Left(cboaccount.text, 3)
    Else
keptp:
       
        Dim M_Objrs_PTPNew As New ADODB.Recordset
        Dim Cmdsql_PTPNew As String
        
        If C_PTP.Value Then
            M_update!ptpvia = IIf(IsNull(CmbViaPtp.text), "", CmbViaPtp.text)
            M_update!ptpdesc = cboaccount.text
            
            '//////////////////////// Awal Logika PTP 1 ////////////////////////////////////////////
            If vrcek = "BP-" And Len(TglPTPNew) > 0 And UCase(cboPTP.text) = "PTP-NEW" Then
                M_update!TglPTPNew = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                                        
                    If TDBDate1.ValueIsNull Then
                        M_update!dateptpnew = Null
                    Else
                        M_update!dateptpnew = Format(TDBDate3.Value, "yyyy-mm-dd")
                        '@@25/01/2012, tambahkan tanggal tagih
                        M_update!tgl_tagih = Format(TdbTglTagih.Value, "yyyy-mm-dd")
                    End If
                                        
                     '@@ 06-01-2012 amountnew yang digunakan untuk amountptp ptp-new
                     'sekarang diambil dari tblnegoptp id terakhir
'                    If Tdabamoint.ValueIsNull Then
'                        M_update!amountnew = 0
'                    Else
'                        M_update!amountnew = Tdabamoint.Value
'                    End If
                   
                    '@@ 16 APRIL 2012, bukan ID terakhir, tetapi inputdate terakhir
                    Cmdsql_PTPNew = "select * from tblnegoptp where custid='"
                    Cmdsql_PTPNew = Cmdsql_PTPNew + lblCustId.Caption + "' order by inputdate desc limit 1"
                    
                    
                    Set M_Objrs_PTPNew = New ADODB.Recordset
                    M_Objrs_PTPNew.CursorLocation = adUseClient
                    M_Objrs_PTPNew.Open Cmdsql_PTPNew, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    M_update!AmountNew = M_Objrs_PTPNew("promisepay")
                    Set M_Objrs_PTPNew = Nothing
            Else
                If cboPTP.text = "PTP-NEW" Then
                    If vrcek <> "PTP-NE" Then
                    
                        If UCase(cboPTP.text) = "PTP-NEW" And listview1(0).ListItems.Count = 0 Then
                            M_update!TglPTPNew = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                            If TDBDate1.ValueIsNull Then
                                M_update!dateptpnew = Null
                            Else
                                M_update!dateptpnew = Format(TDBDate3.Value, "yyyy-mm-dd")
                                '@@25/01/2012 , Tambahkan untuk tanggal tagih
                                M_update!tgl_tagih = Format(TdbTglTagih.Value, "yyyy-mm-dd")
                                
                            End If
                                        
                             '@@ 06-01-2012 amountnew yang digunakan untuk amountptp ptp-new
                            'sekarang diambil dari tblnegoptp id terakhir
'                            If Tdabamoint.ValueIsNull Then
'                                M_update!amountnew = 0
'                            Else
'                                M_update!amountnew = Tdabamoint.Value
'                            End If
                            
                            Cmdsql_PTPNew = "select * from tblnegoptp where custid='"
                            Cmdsql_PTPNew = Cmdsql_PTPNew + lblCustId.Caption + "' order by id desc limit 1"
                
                            Set M_Objrs_PTPNew = New ADODB.Recordset
                            M_Objrs_PTPNew.CursorLocation = adUseClient
                            M_Objrs_PTPNew.Open Cmdsql_PTPNew, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                            
                            If M_Objrs_PTPNew.RecordCount = 0 Then
                                M_update!AmountNew = 0
                            Else
                                M_update!AmountNew = M_Objrs_PTPNew("promisepay")
                            End If
                            
                            'M_update!amountnew = IIf(IsNull(M_Objrs_PTPNew("promisepay")), "0", M_Objrs_PTPNew("promisepay"))
                            Set M_Objrs_PTPNew = Nothing
                            
                        End If
                                                    
                    End If
                End If
            End If
            '//////////////////////// Akhir Logika PTP 1 ////////////////////////////////////////////
            
            '//////////////////////// Awal Logika PTP 2 ////////////////////////////////////////////
            If vrcek = "BP-" And Len(TglPTPNew) > 0 And Left(UCase(cboPTP.text), 3) = "PTP" Then
                M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
            Else
                If Left(cboPTP.text, 3) = "PTP" Then
                    If Left(vrcek, 6) <> Left(cboPTP.text, 6) Then
                        M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                    ElseIf vrnewdate <> TDBDate3.text Then
                        M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                    End If
                End If
            End If
            '//////////////////////// Akhir Logika PTP 2 ////////////////////////////////////////////
    
            pStatusLstCall = cboPTP.text
            'txtResult.Text = pStatusLstCall '@@15/01/2012 Komponen Tak Terpakai
            'txtResultDesc.Text = pStatusLstCalldesc '@@15/01/2012 Komponen Tak Terpakai
            M_update("RECSTATUS") = "P"
            M_update!f_cek_new = Left(cboPTP.text, 6)
                                
            '@@121110 tambahan buat nyatet history f_cek_new
            StatusAccCurrent = Left(cboPTP.text, 6)
            
        Else
        End If
    End If
        
    If C_Payment.Value Then
        If StatusPTP <> Empty Then
            If StatusPTP = M_update!f_cek_new Then
            Else
                M_update!TGLINCOMING = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
            End If
        End If
        M_update!ttlptp = txtPayment.Value
        'M_update!amountptp = Tdabamoint.Value
        '@@ 05-01-2012,tdabamoint sudah tidak dipakai, langsung pakai txtpayment
        M_update!amountptp = txtPayment.Value
        M_update!discpersen = cmbDiscount.text
        M_update!Tenor = txttenor.Value
        M_update!dateptp = Format(TDBDate3.Value, "yyyy/mm/dd")
        '@@25/01/2012, Update tanggal tagih
        If TdbTglTagih.ValueIsNull = False Then
         M_update!tgl_tagih = Format(TdbTglTagih.Value, "yyyy-mm-dd")
       End If
    Else
        M_update!ttlptp = 0
        M_update!discpersen = 0
    End If
               
    If Trim(UCase(IIf(IsNull(M_update("kethslkerja_new")), "", M_update("kethslkerja_new")))) = Trim(UCase(pStatusLstCall)) Then
        TGLSTATUS = IIf(IsNull(M_update("TGLSTATUS")), "", Format(M_update("TGLSTATUS"), "yyyy/mm/dd"))
    Else
        M_update("kethslkerja_new") = pStatusLstCall
        M_update("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
        TGLSTATUS = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")
    End If
        M_update!stscallwith = cbolastcall.text
        M_update("kethslkerja_new") = pStatusLstCall
        pStatusHstLstCall = IIf(IsNull(M_update("kethslkerja_new")), "", M_update("kethslkerja_new"))
        M_update("kethslkerjadesc_new") = cboaccount.text
        M_update("REMARKS") = Replace(txtremarks.text, "'", "`")
    If Not (cmbDateSch.ValueIsNull) Then
        M_update!NEXTACTDATE = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(cmbTimeSch.Value, "hh:nn")
    End If
        
    M_update("Statuscall") = Trim(cbolastcall.text)
    M_update("stscallcust") = Trim(Combo1.text)
        
    '@@ 12-11-10 ini nambahin history perubahan status f_cek_new
    'If statusptp <> "" Or IsNull(statusptp) = False Then
'            Dim HISTORYFCEK As String
'            'HISTORYFCEK = IIf(IsNull(M_update("f_cekhst")), "AWAL", M_update("f_cekhst")) + " > " + statusptp + " [" + CStr(tglcalllalu) + "] " + " > " + StatusAccCurrent + " [" + CStr(M_update("tglcall")) + "] "
'            HISTORYFCEK = IIf(IsNull(M_update("f_cekhst")), "AWAL", M_update("f_cekhst")) + " > " + statusptp + " | " + CStr(tglcalllalu) + " "
'            M_update("f_cekhst") = HISTORYFCEK
    'End If
    
    
    
    
    M_update.update
    
    '@@ 25-Januari-2012 Tulis Result PTPnya
    If C_PTP.Value Then
        FrmResultPTP.txtStatusAcc.text = Trim(cboPTP.text)
        FrmResultPTP.Show vbModal
    Else
        '@@ 28 Agustus 2013
        'Kalo yang statusnya POP tampilkan juga result ptp nya
        Dim M_Objrs_Cek_Status As ADODB.Recordset
        Dim cmdsql_cari As String
        If LstPayment.ListItems.Count > 0 Then
           cmdsql_cari = "select f_cek_new from mgm where custid='"
           cmdsql_cari = cmdsql_cari + CStr(lblCustId.Caption) + "'"
           Set M_Objrs_Cek_Status = New ADODB.Recordset
           M_Objrs_Cek_Status.CursorLocation = adUseClient
           M_Objrs_Cek_Status.Open cmdsql_cari, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
           
           If Trim(M_Objrs_Cek_Status("f_cek_new")) = "POP" Or _
              Trim(Left(M_Objrs_Cek_Status("f_cek_new"), 3)) = "PTP" Or _
              Trim(Left(M_Objrs_Cek_Status("f_cek_new"), 2)) = "BP" Then
                FrmResultPTP.txtStatusAcc = Trim(M_Objrs_Cek_Status("f_cek_new"))
                FrmResultPTP.Show vbModal
           End If
        End If
        Set M_Objrs_Cek_Status = Nothing
    End If
    
    If C_PTP.Value = vbChecked Then
        GoTo BRO
    End If
    
    
    '@@21 May 2012,Penulisan Remarks dipecah per 90 Karakter
    Dim BanyakBaris As Integer
    Dim AW As Integer
    Dim AwalRemarks As String
    Dim Pesan, Unik As String
    If cboaccount.text <> "" Then
        If txtremarks.text <> Empty Then
'            BanyakBaris = Ceiling(Val(Len(TxtRemarks.Text)) / 87)
'            Unik = Format(Now, "ddmmyyyyhhmmss")
'
'            'Bikin Baris KOsong....
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", "------------------------------------------------------------------------------------", CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.Text, 3), cbolastcall.Text, MDIForm1.Text1.Text, Unik, BanyakBaris + 1
'            For AW = 1 To BanyakBaris
'                'AwalRemarks = (87 * AW) - 87
'                AwalRemarks = (87 * ((BanyakBaris + 1) - AW)) - 87
'                pesan = "(" & BanyakBaris + 1 - AW & "/" & BanyakBaris & ") " & Mid(TxtRemarks.Text, AwalRemarks + 1, 87)
'                M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", IIf(IsNull(pesan), "", pesan), CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.Text, 3), cbolastcall.Text, MDIForm1.Text1.Text, Unik, BanyakBaris + 1 - AW
'            Next AW
            
            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.text, CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.text, Combo1.text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.text, 3), cbolastcall.text, MDIForm1.Text1.text, "", "0"
        End If
    End If
    
BRO:
    If C_PTP.Value = 1 Then
        If txtremarks.text <> Empty Then
'             BanyakBaris = Ceiling(Val(Len(TxtRemarks.Text)) / 87)
'            Unik = Format(Now, "ddmmyyyyhhmmss")
'
'            'Bikin Baris KOsong....
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", "------------------------------------------------------------------------------------", CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboPTP.Text, 5), cbolastcall.Text, MDIForm1.Text1.Text, Unik, BanyakBaris + 1
'            For AW = 1 To BanyakBaris
'                'AwalRemarks = (87 * AW) - 87
'                AwalRemarks = (87 * ((BanyakBaris + 1) - AW)) - 87
'                pesan = "(" & BanyakBaris + 1 - AW & "/" & BanyakBaris & ") " & Mid(TxtRemarks.Text, AwalRemarks + 1, 87)
'                M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", IIf(IsNull(pesan), "", pesan), CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboPTP.Text, 5), cbolastcall.Text, MDIForm1.Text1.Text, Unik, BanyakBaris + 1 - AW
'            Next AW
            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.text, CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.text, Combo1.text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboPTP.text, 5), cbolastcall.text, MDIForm1.Text1.text, "", "0"
        End If
    End If

    If Len(TDBTot_payment) > 2 Then
        M_DATA.ADD_tbllunas M_OBJCONN, lblCustId.Caption, Format(TdbLunas.Value, "yyyy/mm/dd"), CCur(TDBTot_payment.Value), VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), TxtFieldName.text, ""
    Else
        On Error Resume Next
    End If
    '------------>> simpan ke table Visit <<--------------------
    If Option8(0).Value Then
        m_Visit.ADD_RequestVisit M_OBJCONN, lblCustId.Caption, M_update!f_cek_new, Text1.text, Format(TDBDate1.Value, "yyyy-mm-dd"), TXtDetails.text, TDBNumber1.Value, TxtAddress.text, Trim(UCase(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)))
    Else
        On Error Resume Next
    End If

    MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
    
    kontak = False
    Set M_update = Nothing

    If shedulePTP_Show = True Then
    Else
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(7) = txtremarks.text
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8) = pStatusLstCall
        If cboaccount <> "" Then
            VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Left(cboaccount, 3)
        ElseIf cboPTP <> "" Then
            VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Left(cboPTP, 6)
        End If
    End If
    pStatusLstCall = ""
    pStatusHstLstCall = ""
    txtremarks.text = Empty


    Set M_DATA = Nothing
    Exit Sub
    Resume
End Sub

'@@ 11-03-2011 Di remarks, udah tidak diapakai
'Private Sub HEADER_SendSMS()
'LstSMS.ColumnHeaders.ADD 1, , "No Telp", 5 * TXT
'LstSMS.ColumnHeaders.ADD 2, , "Nama", 5 * TXT
'LstSMS.ColumnHeaders.ADD 3, , "Custid", 15 * TXT
'LstSMS.ColumnHeaders.ADD 4, , "Pesan", 5 * TXT
'LstSMS.ColumnHeaders.ADD 5, , "Tanggal Terima", 5 * TXT
'
'LstSMS2.ColumnHeaders.ADD 1, , "Sender", 5 * TXT
'LstSMS2.ColumnHeaders.ADD 2, , "Nama", 5 * TXT
'LstSMS2.ColumnHeaders.ADD 3, , "Pesan", 15 * TXT
'LstSMS2.ColumnHeaders.ADD 4, , "Jam", 5 * TXT
'LstSMS2.ColumnHeaders.ADD 5, , "Custid", 5 * TXT
'End Sub


Private Sub HEADER_HISTORY()
    listview1(1).ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    listview1(1).ColumnHeaders.ADD 2, , "History", 80 * TXT
    listview1(1).ColumnHeaders.ADD 3, , "User Log", 10 * TXT
    listview1(1).ColumnHeaders.ADD 4, , "Handle By", 10 * TXT
    listview1(1).ColumnHeaders.ADD 5, , "Sts Account", 10 * TXT
    listview1(1).ColumnHeaders.ADD 6, , "Sts Call", 10 * TXT
    listview1(1).ColumnHeaders.ADD 7, , "Sts Telp With", 25 * TXT
    listview1(1).ColumnHeaders.ADD 8, , "Id", 25 * TXT
End Sub
Private Sub HEADER_RequestVisit()
    LstVisit.ColumnHeaders.ADD 1, , "RequestDate", 10 * TXT
    LstVisit.ColumnHeaders.ADD 2, , "VisitDate", 10 * TXT
    LstVisit.ColumnHeaders.ADD 3, , "VisitNo", 10 * TXT
    LstVisit.ColumnHeaders.ADD 4, , "Details", 20 * TXT
    LstVisit.ColumnHeaders.ADD 5, , "Hasil Visit", 20 * TXT
    LstVisit.ColumnHeaders.ADD 6, , "VisitKe", 2 * TXT
    LstVisit.ColumnHeaders.ADD 7, , "ID", 1 * TXT
    LstVisit.ColumnHeaders.ADD 8, , "Status", 1 * TXT
    End Sub
Private Sub HEADER_HISTORY_PAID()
    listview1(0).ColumnHeaders.ADD 1, , "PayDate", 15 * TXT
    listview1(0).ColumnHeaders.ADD 2, , "Payment", 30 * TXT
    listview1(0).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    listview1(0).ColumnHeaders.ADD 4, , "FieldName", 30 * TXT
    listview1(0).ColumnHeaders.ADD 5, , "Id", 30 * TXT
End Sub
Private Function CEK_DATA_VALID() As Boolean
    Dim m_msgbox As Variant
    Dim cmdsql As String
    Dim M_Objrs_Cek_PTP  As ADODB.Recordset
    Dim m_objrs_reserve As ADODB.Recordset
    Dim TotalPtp As Double
    Dim Pesan As String
    
    If TDBTot_payment > 2 Then
        CEK_DATA_VALID = True
        Exit Function
    Else

        '@@02-05-2012, Tambahan Cek data nomor telepon, harus diisi kategorinya
'        If txtHomeAdd1.Value <> Empty Then
'            If CmbStsKatHome1.Text = "" Or CmbStsKatHome1.Text = "--Pilih Kategori Telepon--" Then
'                MsgBox "Additional Home 1 Tidak Kosong! Harap isi terlebih dahulu kategori teleponnya!", vbOKOnly + vbInformation, "Informasi"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'
'         If txtHomeAdd2.Value <> Empty Then
'            If CmbStsKatHome2.Text = "" Or CmbStsKatHome2.Text = "--Pilih Kategori Telepon--" Then
'                MsgBox "Additional Home 2 Tidak Kosong! Harap isi terlebih dahulu kategori teleponnya!", vbOKOnly + vbInformation, "Informasi"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'
'         If txtOfficeAdd1.Value <> Empty Then
'            If CmbStsKatOffice1.Text = "" Or CmbStsKatOffice1.Text = "--Pilih Kategori Telepon--" Then
'                MsgBox "Additional Office 1 Tidak Kosong! Harap isi terlebih dahulu kategori teleponnya!", vbOKOnly + vbInformation, "Informasi"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'
'         If txtOfficeAdd2.Value <> Empty Then
'            If CmbStsKatOffice2.Text = "" Or CmbStsKatOffice2.Text = "--Pilih Kategori Telepon--" Then
'                MsgBox "Additional Office 2 Tidak Kosong! Harap isi terlebih dahulu kategori teleponnya!", vbOKOnly + vbInformation, "Informasi"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'
'        If txtMobileAdd1.Value <> Empty Then
'            If CmbStsKatHP1.Text = "" Or CmbStsKatHP1.Text = "--Pilih Kategori Telepon--" Then
'                MsgBox "Additional Mobile 1 Tidak Kosong! Harap isi terlebih dahulu kategori teleponnya!", vbOKOnly + vbInformation, "Informasi"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'
'        If txtMobileAdd2.Value <> Empty Then
'            If CmbStsKatHP2.Text = "" Or CmbStsKatHP2.Text = "--Pilih Kategori Telepon--" Then
'                MsgBox "Additional Mobile 2 Tidak Kosong! Harap isi terlebih dahulu kategori teleponnya!", vbOKOnly + vbInformation, "Informasi"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'Dinonaktifkan
        
        '@@04-06-2012 Cek dulu apakah data ptp? kalo iya harus cek cpa
        If C_PTP.Value Then
            cmdsql = "select * from tblcpa where vcustid='"
            cmdsql = cmdsql + Trim(lblCustId.Caption) + "' order by nid desc limit 1 "
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs.RecordCount = 0 Then
             
                MsgBox "Untuk membuat status account PTP, harus dibuat terlebih dahulu CPA nya!", vbOKOnly + vbInformation, "Informasi"
                MsgBox "Data PTP gagal dibuat!", vbOKOnly + vbExclamation, "Peringatan"
                Set M_Objrs = Nothing
                CEK_DATA_VALID = False
                Exit Function
            End If
      
        End If
        
        '@@ 16 May 2012, Cek jika status PTP-POP atau PTP NEW tapi data di tblnegoptp tidak ada
        'Ubah otomastis ke BP
        Dim M_Objrs_NegoPTP As ADODB.Recordset
        Dim WA As String
        If cboPTP.text = "PTP-POP" Then
            'Cek Apakah data di tabelnegoptp ada?
            cmdsql = "select * from tblnegoptp where custid='"
            cmdsql = cmdsql + CStr(lblCustId.Caption) + "' order by promisedate desc limit 1 "
            Set M_Objrs_NegoPTP = New ADODB.Recordset
            M_Objrs_NegoPTP.CursorLocation = adUseClient
            M_Objrs_NegoPTP.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            'Ini Jika Tidak ditemukan data di tabel tblnegoptp, maka ubah status account menjadi BP-POP
            'Agar data bisa di dave
            If M_Objrs_NegoPTP.RecordCount = 0 Then
                WA = MsgBox("Benarkah account ini PTP? Jika benar, tolong sempurnakan datanya, List PTP Jatuh Tempo and masih kosong!. TEKAN YES jika anda ingin mengisi data PTP atau TEKAN NO jika data ini BUKAN PTP!", vbYesNo + vbQuestion, "Konfirmasi")
                If WA = vbYes Then
                    MsgBox "Sempurnakan terlebih dahulu Form PTP anda. Kemudian lakukan penyimpanan ulang remarks anda!", vbOKOnly + vbInformation, "Informasi"
                    CEK_DATA_VALID = False
                    Exit Function
                End If
                cmdsql = "update mgm set tglstatus= now() ,F_CEK='BP-',LASTSTATUS='BP-POP',"
                cmdsql = cmdsql + "KETHSLKERJA_NEW='BP-POP',F_CEK_NEW='BP-',"
                cmdsql = cmdsql + "KETHSLKERJADESC_NEW='BP-BROKEN PROMISE',"
                cmdsql = cmdsql + "KETHSLKERJA='BP-PTP POP BROKEN PROMISE',"
                cmdsql = cmdsql + "REMARKS = 'BP-POP BROKEN PROMISE @',"
                cmdsql = cmdsql + "RECSTATUS='C',OTO='Y' where f_cek_NEW like 'PTP-PO' and custid='"
                cmdsql = cmdsql + CStr(lblCustId.Caption) + "'"
                M_OBJCONN.execute cmdsql
                C_PTP.Value = vbUnchecked
                cboaccount.text = "BP-POP"
                C_Payment.Value = vbUnchecked
            End If
            Set M_Objrs_NegoPTP = Nothing
        End If
                
                
        If cboPTP.text = "PTP-NEW" Then
            'Cek Apakah data di tabelnegoptp ada?
            cmdsql = "select * from tblnegoptp where custid='"
            cmdsql = cmdsql + CStr(lblCustId.Caption) + "' order by promisedate desc limit 1 "
            Set M_Objrs_NegoPTP = New ADODB.Recordset
            M_Objrs_NegoPTP.CursorLocation = adUseClient
            M_Objrs_NegoPTP.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            'Ini Jika Tidak ditemukan data di tabel tblnegoptp, maka ubah status account menjadi BP-POP
            'Agar data bisa di dave
            If M_Objrs_NegoPTP.RecordCount = 0 Then
                WA = MsgBox("Benarkah account ini PTP? Jika benar, tolong sempurnakan datanya, List PTP Jatuh Tempo and masih kosong!. TEKAN YES jika anda ingin mengisi data PTP atau TEKAN NO jika data ini BUKAN PTP!", vbYesNo + vbQuestion, "Konfirmasi")
                If WA = vbYes Then
                    MsgBox "Sempurnakan terlebih dahulu Form PTP anda. Kemudian lakukan penyimpanan ulang remarks anda!", vbOKOnly + vbInformation, "Informasi"
                    CEK_DATA_VALID = False
                    Exit Function
                End If
                cmdsql = "update mgm set tglstatus= now() ,F_CEK='BP-',LASTSTATUS='BP-NEW',"
                cmdsql = cmdsql + "KETHSLKERJA_NEW='BP-NEW',F_CEK_NEW='BP-',"
                cmdsql = cmdsql + "KETHSLKERJADESC_NEW='BP-BROKEN PROMISE',"
                cmdsql = cmdsql + "KETHSLKERJA='BP-PTP NEW BROKEN PROMISE',"
                cmdsql = cmdsql + "REMARKS = 'BP-NEW BROKEN PROMISE @',"
                cmdsql = cmdsql + "RECSTATUS='C',OTO='Y' where f_cek_NEW like 'PTP-NE' and custid='"
                cmdsql = cmdsql + CStr(lblCustId.Caption) + "'"
                M_OBJCONN.execute cmdsql
                C_PTP.Value = vbUnchecked
                cboaccount.text = "BP-NEW"
                C_Payment.Value = vbUnchecked
            End If
            Set M_Objrs_NegoPTP = Nothing
        End If
                
        
        If Left(cmbContacted, 3) = "PTP" And LstPayment.ListItems.Count = 0 Then
            MsgBox "PTP harus buat Nego PTP di tabel yang hijau !!!", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        If Combo1.text = "" Then
            MsgBox "Status Call harus diisi!", vbInformation + vbOKOnly, "TINS"
            Combo1.SetFocus
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        If cboaccount.text = "" And C_PTP.Value = vbUnchecked Then
            MsgBox "Status Account harus diisi!", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        If cbolastcall.text = "" Then
            MsgBox "Status Telepon With harus diisi!", vbInformation + vbOKOnly, "TINS"
            cbolastcall.SetFocus
            CEK_DATA_VALID = False
            Exit Function
        End If
    
        If C_PTP.Value = vbChecked Then
              '@@ 11 Januari 2012 dinonaktifkan, tidak menggunakan tdabmoint
        '       If Val(vrcekamont) <> Tdabamoint.Value And bcekptp = False Then
        '            MsgBox "anda harus klik tambah di Call Activity untuk Negotiation", vbInformation + vbOKOnly, "TINS"
        '
        '            CEK_DATA_VALID = False
        '            Exit Function
        '        End If
        
            '@@ 05-10-2011, Jika melakukan PTP maka combo via ptp harus diisi
            If CmbViaPtp.text = "" Then
                MsgBox "Combo Via tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
            'Tambahan, Jika Status data PTP, hitung tanggal tagih
            If TDBDate3.ValueIsNull Then
                MsgBox "Anda belum menentukan tanggal effective pembayaran!", vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
'            If UCase(Trim(CmbViaPtp.Text)) = "HSBC" Then
'                TdbTglTagih.Value = Format(TDBDate3.Value - 1, "dd/mm/yyyy")
'            ElseIf UCase(Trim(CmbViaPtp.Text)) = "BERSAMA" Then
'                 TdbTglTagih.Value = Format(TDBDate3.Value - 1, "dd/mm/yyyy")
'            ElseIf UCase(Trim(CmbViaPtp.Text)) = "KANTOR POS" Then
'                 TdbTglTagih.Value = Format(TDBDate3.Value - 3, "dd/mm/yyyy")
'            ElseIf UCase(Trim(CmbViaPtp.Text)) = "PUM" Then
'                 TdbTglTagih.Value = Format(TDBDate3.Value - 1, "dd/mm/yyyy")
'            Else
'                 TdbTglTagih.Value = Format(TDBDate3.Value - 3, "dd/mm/yyyy")
'            End If
            
            Call CariTanggalTagih
            
        End If
    
        If C_Payment.Value = 1 Then
            CmbBaseOn.text = "TOTAL AMOUNT"
            If TDBDate3.ValueIsNull Then
                CEK_DATA_VALID = False
                MsgBox "Tanggal PTP Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                Exit Function
            End If
        End If
                   
        If C_PTP.Value = 1 Then
            If cboPTP.text = Empty Then
                CEK_DATA_VALID = False
                MsgBox "Description PTP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
                Exit Function
                SSTab1.Tab = 3
            End If
        End If

       
        If txtremarks.text = "" Then
            CEK_DATA_VALID = False
            MsgBox "Remarks Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            Exit Function
        End If
 
        If ADD_CUST = True Then
        Else
            If cboaccount.text <> "" Then
                Dim StatusRemarks As String
'                '@@ 16 Agustus 2011, pola remarks diubah
'                StatusRemarks = Combo1.Text & "-"
'                'StatusRemarks = StatusRemarks & cbolastcall.Text & "-"
'                '@@04-05-2012  Cbolastcall disingkat di statusspeakwith
'                StatusRemarks = StatusRemarks & StatusSpeakWith & "-"
'                StatusRemarks = StatusRemarks & "[" & cboaccount.Text & "] - "
'                StatusRemarks = StatusRemarks & TxtTelpKe.Text
'                '@@03-05-2012 Tambahan Status Telepon
'                StatusRemarks = StatusRemarks & "-" & IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp)
'                TxtRemarks.Text = StatusRemarks & " // " & TxtRemarks.Text
                 
                '@@10052012 Mengubah Pola Remarks
                StatusRemarks = IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp) & "/"
                StatusRemarks = StatusRemarks & IIf(Combo1.text = "Receive", "RCVD", "NRCV") & "/"
                StatusRemarks = StatusRemarks & StatusSpeakWith & "/"
                StatusRemarks = StatusRemarks & Mid(cboaccount.text, 1, 2) & ": "
                txtremarks.text = StatusRemarks & txtremarks.text
                
                
             ElseIf cboPTP.text <> "" Then
'                '@@ 16 Agustus 2011, pola remarks diubah
'                StatusRemarks = Combo1.Text & "-"
'                'StatusRemarks = StatusRemarks & cbolastcall.Text & "-"
'                '@@04-05-2012  Cbolastcall disingkat di statusspeakwith
'                StatusRemarks = StatusRemarks & StatusSpeakWith & "-"
'                StatusRemarks = StatusRemarks & " PTP Via:" & CmbViaPtp.Text + "-"
'                StatusRemarks = StatusRemarks & "[ " & cboPTP.Text & "-"
'                StatusRemarks = StatusRemarks & "AmountPTP:" & txtPayment.Text & "- "
'                StatusRemarks = StatusRemarks & "DatePTP:" & TDBDate3.Value & " ] -"
'                StatusRemarks = StatusRemarks & TxtTelpKe.Text
'                '@@03-05-2012 Tambahan Status Telepon
'                StatusRemarks = StatusRemarks & "-" & IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp)
'                TxtRemarks.Text = StatusRemarks & " // " & TxtRemarks.Text
                
                '@@10052012 Menubah Pola Remarks
                StatusRemarks = IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp) & "/"
                StatusRemarks = StatusRemarks & IIf(Combo1.text = "Receive", "RCVD", "NRCV") & "/"
                StatusRemarks = StatusRemarks & StatusSpeakWith & "/"
                StatusRemarks = StatusRemarks & cboPTP.text & "/"
                StatusRemarks = StatusRemarks & "PTP Via " & CmbViaPtp.text & "/"
                StatusRemarks = StatusRemarks & "Amount PTP " & txtPayment.text & "/"
                StatusRemarks = StatusRemarks & "Date PTP " & TDBDate3.Value & ": "
                txtremarks.text = StatusRemarks & txtremarks.text
                
            
            End If
            
            If stscall = True Then
                If C_PTP.Value = vbUnchecked And cboaccount.text = "" Then
                    CEK_DATA_VALID = False
                    MsgBox "Status Account Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                    SSTab1.Tab = 3
                    Exit Function
                End If
            End If
        End If
    End If

        
'        If cmbDiscount.Text = "" Then
'            MsgBox "Diskon harus diisi", vbInformation + vbOKOnly, "TINS"
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
      
    '@@23031012 Cek dulu apakah status data BP atau POP
    'JIka BP atau POP lewat saja pengecekan PTP
    If Mid(cboaccount.text, 1, 3) = "BP-" Or Mid(cboaccount.text, 1, 3) = "POP" Then
        GoTo Lanjut_1
    End If
      
    Pesan = "Informasi: " & vbCrLf
    Pesan = Pesan & "Anda hanya dapat membuat status PTP " & vbCrLf
    Pesan = Pesan & "jika CPA untuk account tersebut telah dibuat! " & vbCrLf
    Pesan = Pesan & "Mintalah kepada TL anda untuk membuat CPA!" & vbCrLf & vbCrLf
    Pesan = Pesan & "Jika anda mengalami kesulitan untuk menyimpan data remarks anda, kemungkinan adalah: " & vbCrLf
    Pesan = Pesan & "1. Ada data di list PTP Jatuh Tempo, tetapi Form PTP kosonng. Seperti Total Amount Deal dan Date Payment Effective." & vbCrLf
    Pesan = Pesan & "2. Ada data di Form PTP, tetapi data di list PTP Jatuh tempo kosong! " & vbCrLf
    Pesan = Pesan & "3. Jumlah data di list RESERVED PTP tidak sama dengan Tenor di Form PTP!" & vbCrLf
    Pesan = Pesan & "4. Ada data di list Reserved PTP, tetapi data di Form PTP masih kosong!" & vbCrLf
    Pesan = Pesan & "5. Date Payment Effective harus sama dengan tanggal di list PTP jatuh tempo!"
      
      
    '@@ 07-02-2012, cek data negoptp jika status data PTP
    If C_PTP.Value = 1 Then
                
        'Cek Nilai Payment
        If txtPayment.Value = "0" Or txtPayment.ValueIsNull = True Then
            MsgBox "Anda mencentang data PTP, Total Amount Deal tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
            MsgBox Pesan, vbOKOnly + vbInformation, "Informasi"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        'Cek Nilai Date Payment Effective
        If TDBDate3.ValueIsNull = True Then
            MsgBox "Anda mencentang data PTP, Date Payment Effective tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
            MsgBox Pesan, vbOKOnly + vbInformation, "Informasi"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        'Cek combo via
        If CmbViaPtp.text = "" Then
            MsgBox "Anda mencentang data PTP, Combo VIA tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
            MsgBox Pesan, vbOKOnly + vbInformation, "Informasi"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        '----------///////// Dinonaktifkan dulu, bermasalah pada saat penyimpanan Remarks ///////////////////
'
'        'Cek Data di tabel tblnegoptp, apakah sinkron/sama dengan data ptp di mgm
'        '@@ 26-03-2012 Filter Tanggal dinonaktifkan dulu
'        CMDSQL = "select * from tblnegoptp where custid='"
'        CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "' "
'        'CMDSQL = CMDSQL + " and date_part('month',promisedate)>="
'        'CMDSQL = CMDSQL + "date_part('month',now()) and date_part('year',promisedate)>="
'        'CMDSQL = CMDSQL + "date_part('year',now()) and promisepay>'0' "
'        CMDSQL = CMDSQL + " order by promisedate desc limit 1"
'        Set M_Objrs_Cek_PTP = New ADODB.Recordset
'        M_Objrs_Cek_PTP.CursorLocation = adUseClient
'        M_Objrs_Cek_PTP.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        'Jika data negoptp tidak ada, maka user harus mengklik tombol tambah PTP terlebih dahulu
'        If M_Objrs_Cek_PTP.RecordCount = 0 Then
'            MsgBox "Anda belum mengklik tombol ADD PTP!", vbOKOnly + vbInformation, "Informasi"
'            MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
'            Set M_Objrs_Cek_PTP = Nothing
'            CEK_DATA_VALID = False
'            Exit Function
'        Else
'            'Jika datanya ada cek apakah tanggalnya sama?
'            If Format(M_Objrs_Cek_PTP("promisedate"), "yyyy-mm-dd") <> Format(TDBDate3.Value, "yyyy-mm-dd") Then
'                MsgBox "Tanggal Date Payment Effective PTP berbeda dengan data yang ada di list PTP Jatuh Tempo! Date payment effective sama data di list PTP Jatuh Tempo harus sama!", vbOKOnly + vbInformation, "Informasi"
'                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
'                Set M_Objrs_Cek_PTP = Nothing
'                TxtRemarks.Text = ""
'                MsgBox "Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan"
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'        End If
'
'
'        'Cek data di tabel reserve
'        CMDSQL = "select * from tblreserve where custid='"
'        CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "' and stsmove='0'"
'        Set m_objrs_reserve = New ADODB.Recordset
'        m_objrs_reserve.CursorLocation = adUseClient
'        m_objrs_reserve.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        '@@ 16032012 Cek Reserve dinonaktifkan dulu
''        If txttenor.Value > 1 Then
''            'Jika jumlah tenor di listreserve tidak sama dengan jumlah tenor, keluar fungsi
''            If (txttenor.Value - 1) <> Val(m_objrs_reserve.RecordCount) Then
''                MsgBox "Count (jumlah) data di list reserve ptp tidak sama dengan jumlah tenor! Harap buat ulang PTP terlebih dahulu dengan mengklik tombol Add PTP!", vbOKOnly + vbInformation, "Informasi"
''                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
''                Set m_objrs_reserve = Nothing
''                CEK_DATA_VALID = False
''                Exit Function
''            End If
''        End If
''
''        If txttenor.Value = 0 Or txttenor.Value = 1 Then
''            If m_objrs_reserve.RecordCount > 0 Then
''                MsgBox "Count (jumlah) data di list reserve ptp tidak sama dengan jumlah tenor! Harap buat ulang PTP terlebih dahulu dengan mengklik tombol Add PTP!", vbOKOnly + vbInformation, "Informasi"
''                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
''                Set m_objrs_reserve = Nothing
''                CEK_DATA_VALID = False
''                Exit Function
''            End If
''        End If
'
'        Set M_Objrs_Cek_PTP = Nothing
'        Set m_objrs_reserve = Nothing
'----------///////// Dinonaktifkan dulu, bermasalah pada saat penyimpanan Remarks ///////////////////
    End If
    

'----------///////// Dinonaktifkan dulu, bermasalah pada saat penyimpanan Remarks ///////////////////
'    '@@ 08-02-2012 Jika Tanda PTP tidak dicentang tetapi ada data di tabel negoptp
'    'Maka form PTP harus diisi!
'    If C_PTP.Value = False Then
'         Dim WK As String
'
'        'Cek data di tabel negoptp
'        '@@ 26-03-2012 Filter Tanggal dinonaktifkan dulu
'        CMDSQL = "select * from tblnegoptp where custid='"
'        CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "' "
'        'CMDSQL = CMDSQL + "  and date_part('month',promisedate)>="
'        'CMDSQL = CMDSQL + "date_part('month',now()) and date_part('year',promisedate)>=date_part('year',now())"
'        Set M_Objrs_Cek_PTP = New ADODB.Recordset
'        M_Objrs_Cek_PTP.CursorLocation = adUseClient
'        M_Objrs_Cek_PTP.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        'Jika ada datanya
'        If M_Objrs_Cek_PTP.RecordCount > 0 Then
'
'            MsgBox "List PTP Jatuh Tempo tidak kosong! Tetapi Form PTP masih kosong. Anda dapat membuat PTP atau menghapus data di list PTP Jatuh Tempo, sebelum data disimpan!", vbOKOnly + vbInformation, "Informasi"
'            MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
'
'            '@@24031012, Kasih konfirmasi, supaya program bisa menghapus data
'            WK = MsgBox("Apakah anda ingin data di list PTP jatuh tempo dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
'            If WK = vbYes Then
'                '@@ 26-03-2012 Filter Tanggal dinonaktifkan terlebih dahulu
'                CMDSQL = "delete from tblnegoptp where custid='"
'                CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "' "
''                CMDSQL = CMDSQL + " and date_part('month',promisedate)>="
''                CMDSQL = CMDSQL + " date_part('month',now()) and date_part('year',promisedate)>=date_part('year',now())"
'                M_OBJCONN.Execute CMDSQL
'                TxtPayment.Value = 0
'                Chktenor.Value = vbUnchecked
'                txttenor.Value = 0
'                TDBDate3.Value = ""
'                CmbViaPtp.Text = ""
'                tdbptpnew.Value = ""
'                TdbTglTagih.Value = ""
'                LstPayment.ListItems.CLEAR
'                'Update data MGM nya
'                CMDSQL = "update mgm set ttlptp=null, tenor=null, dateptp=null,"
'                CMDSQL = CMDSQL + "ptpvia=null,tgl_tagih=null where custid='"
'                CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute CMDSQL
'                GoTo Cek_PTP_Reserved
'            Else
'                MsgBox "Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan"
'            End If
'
'            Set M_Objrs_Cek_PTP = Nothing
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
'----------///////// Dinonaktifkan dulu, bermasalah pada saat penyimpanan Remarks ///////////////////
Cek_PTP_Reserved:
        Set M_Objrs_Cek_PTP = Nothing
        
        'Cek data reserve
'        CMDSQL = "select * from tblreserve where custid='"
'        CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "' and stsmove='0'"
'        Set m_objrs_reserve = New ADODB.Recordset
'        m_objrs_reserve.CursorLocation = adUseClient
'        m_objrs_reserve.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
        '@@ 26-03-2012 Cek Reservednya dinonaktifkan dulu
        'Jika ada data reserve
'        If m_objrs_reserve.RecordCount > 0 Then
'            MsgBox "List Reserve PTP tidak kosong! Tetapi Form PTP masih kosong. Anda dapat membuat PTP atau menghapus data di list Reserve PTP, sebelum data disimpan!", vbOKOnly + vbInformation, "Informasi"
'            MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
'
'            '@@24031012, Kasih konfirmasi untuk menghapus reserved ptp
'            WK = MsgBox("Apakah anda ingin data reserved PTP dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
'            If WK = vbYes Then
'                CMDSQL = "delete from tblreserve where custid='"
'                CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "' and stsmove='0'"
'                M_OBJCONN.Execute CMDSQL
'                TxtPayment.Value = 0
'                Chktenor.Value = vbUnchecked
'                txttenor.Value = 0
'                TDBDate3.Value = ""
'                CmbViaPtp.Text = ""
'                tdbptpnew.Value = ""
'                TdbTglTagih.Value = ""
'                LstReserve.ListItems.CLEAR
'                Set m_objrs_reserve = Nothing
'                'Update data MGM nya
'                CMDSQL = "update mgm set ttlptp=null, tenor=null, dateptp=null,"
'                CMDSQL = CMDSQL + "ptpvia=null,tgl_tagih=null where custid='"
'                CMDSQL = CMDSQL + Trim(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute CMDSQL
'                GoTo Lanjut_1
'            Else
'                MsgBox "Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan"
'            End If
'
'            Set m_objrs_reserve = Nothing
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
'
'        Set m_objrs_reserve = Nothing
'   End If
Lanjut_1:
    
    
    If C_PTP.Value = 1 Then
        txtremarks.text = txtremarks.text
    End If

    If regnego = True Then
        Dim n%
        Dim jum As Currency
        For n = 1 To FrmCC_Colection.LstPayment.ListItems.Count
            jum = jum + FrmCC_Colection.LstPayment.ListItems(n).SubItems(3)
        Next n
        If jum < FrmCC_Colection.txtPayment.Value Then
            MsgBox "Jumlah PTP Belum sama dengan Jumlah Deal Payment"
            CEK_DATA_VALID = False
            txtremarks.text = ""
            Exit Function
        End If
    End If
    regnego = False
    CEK_DATA_VALID = True
    
End Function
Public Sub Custid_Double()
Dim ListItem As ListItem
Dim test As String
Dim cmdsql As String



Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
test = Format(LblDOB.Caption, "yyyy/mm/dd")

'@@ 26-11-10 Ubah logik double custid, harus cek ktpnya dulu
If Trim(lblID.Caption) <> "" Then
    cmdsql = "Select a.custid, a.name,a.agent, a.amountwo,"
    cmdsql = cmdsql + "a.principal,a.flaglead from mgm a where (a.name='"
    cmdsql = cmdsql + Trim(TxtName.text) + "' and dob='"
    cmdsql = cmdsql + test + "' or ktpno='"
    cmdsql = cmdsql + Trim(lblID.Caption) + "')  and a.custid <> '"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
Else
    cmdsql = "Select a.custid, a.name,a.agent, a.amountwo,"
    cmdsql = cmdsql + "a.principal,a.flaglead from mgm a where a.name='"
    cmdsql = cmdsql + Trim(TxtName.text) + "' and dob='"
    cmdsql = cmdsql + test + "'"
    cmdsql = cmdsql + " and a.custid <> '"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
End If


'm_cust.Open "Select a.custid, a.name,a.agent, a.amountwo,a.principal,a.flaglead from mgm a where (a.name='" + Trim(txtname.Text) + "' and dob='" + test + "' or ktpno='" & Trim(lblID.Caption) & "') and a.custid <> '" + Trim(lblCustId.Caption) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

While Not m_cust.EOF
    Set ListItem = LstDoubleId.ListItems.ADD(, , IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")))
        ListItem.SubItems(1) = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
        ListItem.SubItems(2) = IIf(IsNull(m_cust("AGENT")), "", m_cust("AGENT")) '
      '  If Format(IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")), "##,###") = 1 Then
         '    harga = IIf(IsNull(m_cust("AmountWo")), 0, m_cust("AmountWo"))
           '  harga = harga + (harga * 18.26) / 100
          '   listitem.SubItems(3) = Format(harga, "##,###")
        'Else
            ListItem.SubItems(3) = Format(IIf(IsNull(m_cust("AmountWo")), 0, m_cust("AmountWo")), "##,###")
        'End If
        
        
       ' If Format(IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")), "##,###") = 1 Then
        '     harga = IIf(IsNull(m_cust("principal")), 0, m_cust("principal"))
         '    harga = harga + (harga * 26.05) / 100
          '   listitem.SubItems(4) = Format(harga, "##,###")
        'Else
        
        
     If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
        If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
            ListItem.SubItems(4) = ""
        Else
           ListItem.SubItems(4) = ENCRIPY(False, CStr(Format(IIf(IsNull(m_cust("principal")), 0, m_cust("principal")), "##,###")))
        End If
    Else
            ListItem.SubItems(4) = ENCRIPY(False, CStr(Format(IIf(IsNull(m_cust("principal")), 0, m_cust("principal")), "##,###")))
    End If
      
        
   
      
     
       
       ' End If
        
    
    m_cust.MoveNext
Wend
Set m_cust = Nothing
End Sub

Private Sub SSCommand2_Click(Index As Integer)
Dim m_msgbox As Variant
Dim STATUS As String
Dim gaji As Currency
Dim gaji1 As String
Dim ListItem As ListItem
Dim M_DATA As New ClsNegoPTP
Dim JmlPay As Double
Dim I As Integer
Dim n As Integer
Dim Vrdate As String
Dim jatuhtempo As String
Dim M_Objrs_Cek_PTP As ADODB.Recordset
Dim m_objrs_cek_reserve As ADODB.Recordset

Select Case Index
    Case 0
         
        If TDBDate3.ValueIsNull Or Tdabamoint.ValueIsNull Or txttenor.ValueIsNull Then
            MsgBox "Pengisian Data Belum Lengkap (installment,tenor,dateptp)!"
            Exit Sub
        End If
        
        '@@ 26-03-2012, Dinonaktifkan dulu deh
'        If CDate(Format(TDBDate3.Value, "mm/dd/yyyy")) < CDate(Format(MDIForm1.TDBDate1.Value, "mm/dd/yyyy")) Then
'            MsgBox "Date 1st PTP tidak boleh lebih kecil dari tanggal hari ini!", vbOKOnly + vbInformation, "Informasi"
'            MsgBox "Data PTP gagal dibuat!", vbOKOnly + vbCritical, "Informasi"
'            Exit Sub
'        End If
                  
        '@@ 29 Desember 2011, Cek terlebih dahulu, apakah ada CPA atau tidak, jika tidak ada CPA maka
        'tidak bisa melakukan PTP
       cmdsql = "select * from tblcpa where vcustid='"
       cmdsql = cmdsql + Trim(lblCustId.Caption) + "' order by nid desc limit 1 "
       Set M_Objrs = New ADODB.Recordset
       M_Objrs.CursorLocation = adUseClient
       M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
       
       If M_Objrs.RecordCount = 0 Then
        'C_PTP.Value = vbUnchecked
        MsgBox "Untuk membuat status account PTP, harus dibuat terlebih dahulu CPA nya!", vbOKOnly + vbInformation, "Informasi"
        MsgBox "Data PTP gagal dibuat!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
       End If
       
             
       If txtPayment.Value < Val(M_Objrs("nttlpayment")) Then
        MsgBox "Total Amount Deal tidak boleh lebih kecil dari payment di CPA!", vbOKOnly + vbInformation, "Informasi"
        a = MsgBox("Payment di CPA adalah: Rp." + Format(M_Objrs("nttlpayment"), "##,###") + ". Anda ingin mengganti Total Amount Deal dengan nilai Payment di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
        If a = vbNo Then
            MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
            Exit Sub
        Else
            'Ambil Nilai Payment di CPA untuk di tempatkan di Total Amount Deal
            txtPayment.Value = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
            GoTo LanjutPtp
        End If
       End If
       
       If txtPayment.Value > Val(M_Objrs("nttlpayment")) Then
        MsgBox "Total Amount Deal tidak boleh lebih besar dari payment di CPA!", vbOKOnly + vbInformation, "Informasi"
        a = MsgBox("Payment di CPA adalah: Rp." + Format(M_Objrs("nttlpayment"), "##,###") + ". Anda ingin mengganti Total Amount Deal dengan nilai Payment di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
        If a = vbNo Then
            MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
            Exit Sub
        Else
            'Ambil Nilai Payment di CPA untuk di tempatkan di Total Amount Deal
            txtPayment.Value = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
            GoTo LanjutPtp
        End If
       End If
       
       
LanjutPtp:
        
         'Cek apakah Tenor, lebih kecil dari installment period di CPA
             If txttenor.Value < Val(M_Objrs("nperiod")) Then
                MsgBox "Tenor tidak boleh lebih kecil dari installment period di CPA!", vbOKOnly + vbInformation, "Informasi"
                a = MsgBox("Installment period di CPA adalah :" + Format(M_Objrs("nperiod"), "##,###") + ". Anda ingin mengganti Tenor dengan nilai Installment Period di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
                If a = vbNo Then
                    MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
                    Exit Sub
                Else
                    'Ambil Nilai Tenor dari Installment Period di CPA
                    txttenor.Value = IIf(IsNull(M_Objrs("nperiod")), "0", M_Objrs("nperiod"))
                    If txttenor > 1 Then
                        Chktenor.Value = vbChecked
                    End If
                    GoTo LanjutPtp2
                End If
            End If
            
            If txttenor.Value > Val(M_Objrs("nperiod")) Then
                MsgBox "Tenor tidak boleh lebih besar dari installment period di CPA!", vbOKOnly + vbInformation, "Informasi"
                a = MsgBox("Installment period di CPA adalah :" + Format(M_Objrs("nperiod"), "##,###") + ". Anda ingin mengganti Tenor dengan nilai Installment Period di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
                If a = vbNo Then
                    MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
                    Exit Sub
                Else
                    'Ambil Nilai Tenor dari Installment Period di CPA
                    txttenor.Value = IIf(IsNull(M_Objrs("nperiod")), "0", M_Objrs("nperiod"))
                    If txttenor > 1 Then
                        Chktenor.Value = vbChecked
                    End If
                    GoTo LanjutPtp2
                End If
            End If
            
            Set M_Objrs = Nothing

LanjutPtp2:
        
        '@@ 07-02-2012 Cek data dulu, apakah sebelumnya ada data di tblnegoptp? Buat Handle
        'Apakah ada data PTP sebelumnya, kalo ada data ptp sebelumnya dihapus
        '@@ 09-04-2012 filter tanggal dihapus dulu
        cmdsql = "select * from tblnegoptp where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'  "
        'Cmdsql = Cmdsql + " and date_part('month',promisedate)>=date_part('month',now())  "
        'Cmdsql = Cmdsql + " and date_part('year',promisedate)=date_part('year',now()) "
        '@@13-04-2012 Tambahkan Filter tanggal
        cmdsql = cmdsql + " and date(promisedate)='"
        cmdsql = cmdsql + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "' "
        cmdsql = cmdsql + " order by promisedate,id desc "
        Set M_Objrs_Cek_PTP = New ADODB.Recordset
        M_Objrs_Cek_PTP.CursorLocation = adUseClient
        M_Objrs_Cek_PTP.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_PTP.RecordCount > 0 Then
            Dim KonfirmasiPTP As String
            KonfirmasiPTP = MsgBox("Ada data PTP Sebelumnya dengan TANGGAL YANG SAMA, apakah anda akan menghapus data PTP lama dan menggantinya dengan yang baru?", vbYesNo + vbQuestion, "Konfirmasi")
            If KonfirmasiPTP = vbNo Then
                Set M_Objrs_Cek_PTP = Nothing
                MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbInformation, "Informasi"
                Exit Sub
            End If
            
            'Jika memilih Ya, maka cek reservenya
            Dim KonfirmasiReserve As String
            cmdsql = "select * from tblreserve where custid='"
            cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and stsmove='0'"
            Set m_objrs_cek_reserve = New ADODB.Recordset
            m_objrs_cek_reserve.CursorLocation = adUseClient
            m_objrs_cek_reserve.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If m_objrs_cek_reserve.RecordCount > 0 Then
                
                '@@ 14-04-2012, Cek dulu tenornya jika lebih dari 1 harus hapus data reservenya
                If txttenor.Value > 1 Then
                    KonfirmasiReserve = MsgBox("Tenor lebih dari 1.Apakah anda akan menghapus data reserve yang lama?", vbYesNo + vbQuestion, "Konfirmasi")
                
                    If KonfirmasiReserve = vbNo Then
                        MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Informasi"
                        Set m_objrs_cek_reserve = Nothing
                        Exit Sub
                    End If
                End If
                
                KonfirmasiReserve = vbYes
                
                If KonfirmasiReserve = vbYes Then
                    
                    If M_Objrs_Cek_PTP.RecordCount > 0 Then
                        'Hapus data PTPnya
                        While Not M_Objrs_Cek_PTP.EOF
                            cmdsql = "delete from tblnegoptp where id='"
                            cmdsql = cmdsql + CStr(M_Objrs_Cek_PTP("id")) + "'"
                            M_OBJCONN.execute cmdsql
                            M_Objrs_Cek_PTP.MoveNext
                        Wend
                    End If
                    
                    'Hapus data Reservenya
                    If m_objrs_cek_reserve.RecordCount > 0 Then
                        While Not m_objrs_cek_reserve.EOF
                            cmdsql = "delete from tblreserve where id='"
                            cmdsql = cmdsql + CStr(m_objrs_cek_reserve("id")) + "'"
                            M_OBJCONN.execute cmdsql
                            m_objrs_cek_reserve.MoveNext
                        Wend
                    End If
                    
                End If
                
            Else
                    'Jika tidak ada data reserve maka langsung hapus saja data ptp nya
                    If M_Objrs_Cek_PTP.RecordCount > 0 Then
                        While Not M_Objrs_Cek_PTP.EOF
                            cmdsql = "delete from tblnegoptp where id='"
                            cmdsql = cmdsql + CStr(M_Objrs_Cek_PTP("id")) + "'"
                            M_OBJCONN.execute cmdsql
                            M_Objrs_Cek_PTP.MoveNext
                        Wend
                    End If
            End If
            LstPayment.ListItems.clear
            LstReserve.ListItems.clear
            Set m_objrs_cek_reserve = Nothing
        Else
            'Ini jika PTP Jatuh Temponya kosong!
            'Konfirmasi lagi untuk penghapusan reserve data
            If txttenor.Value > 1 Then
                KonfirmasiReserve = MsgBox("Tenor lebih dari 1. Apakah anda akan membersihkan data reserve PTP?", vbYesNo + vbQuestion, "Konfirmasi")
                If KonfirmasiReserve = vbNo Then
                    MsgBox "Data PTP Gagal ditambahkan!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
                cmdsql = "delete from tblreserve where custid='"
                cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and stsmove='0'"
                M_OBJCONN.execute cmdsql
             End If
        End If
        
        Call CariTanggalTagih
        
        '@@ 22-12-2011 Menentukan nilai awal payment
        If Val(txttenor.Value) > 1 Then
            FrmDealPtp.Show vbModal
            Exit Sub
        End If
        
'        'Update amountptp dan amountnew ke database mgm
'        '@@ 22-09-2011
'        CMDSQL = "update mgm set amountnew='"
'        CMDSQL = CMDSQL + CStr(Tdabamoint.Value) + "', amountptp='"
'        CMDSQL = CMDSQL + CStr(Tdabamoint.Value) + "', tglptpnew=now() where custid='"
'        CMDSQL = CMDSQL + lblCustId.Caption + "'"
'        M_OBJCONN.Execute CMDSQL
        
        bcekptp = True
        '@@ 14 April 2012, Cek tanggal negoptp jika ada yang sama dengan yang diinputkan,
        'yang lama dihapus dan diganti dengan yang baru
        Dim M_Objrs_Cek_Tgl As ADODB.Recordset
           If Chktenor.Value = 0 Then
                
                jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
                
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp where custid='"
                cmdsql = cmdsql + lblCustId.Caption + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                 
                cmdsql = "INSERT INTO TblNegoPTP "
                cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                cmdsql = cmdsql + "VALUES "
                cmdsql = cmdsql + "('" + lblCustId + "', "
                cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "now(),"
                cmdsql = cmdsql + "'IPO')"
                M_OBJCONN.execute cmdsql
                
                '@@14042012, tblnegoptp_log di cek aja
                 '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp_log where custid='"
                cmdsql = cmdsql + lblCustId.Caption + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp_log where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                
                ' isi ke tbl log_ptp
                cmdsql = "INSERT INTO tblnegoptp_log "
                cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
                cmdsql = cmdsql + "VALUES "
                cmdsql = cmdsql + "('" + lblCustId + "', "
                cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "now(),"
                cmdsql = cmdsql + "'" + lblaoc.Caption + "','P')"
                M_OBJCONN.execute cmdsql
                
                Set ListItem = LstPayment.ListItems.ADD(, , "")
                ListItem.SubItems(1) = ""
                ListItem.SubItems(2) = Format(TDBDate3.Value, "dd/mm/yyyy")
                ListItem.SubItems(3) = CStr(Tdabamoint.Value)
                ListItem.SubItems(4) = "IPO"
                ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
            
            Else
            
                jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
                
                 '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp where custid='"
                cmdsql = cmdsql + lblCustId.Caption + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                
                cmdsql = "INSERT INTO TblNegoPTP "
                cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                cmdsql = cmdsql + "VALUES "
                cmdsql = cmdsql + "('" + lblCustId + "', "
                cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "now(),"
                cmdsql = cmdsql + "'IPO')"
                M_OBJCONN.execute cmdsql
                
                 '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp_log where custid='"
                cmdsql = cmdsql + lblCustId.Caption + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp_log where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                        
                ' isi ke tbl log_ptp
                cmdsql = "INSERT INTO tblnegoptp_log "
                cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
                cmdsql = cmdsql + "VALUES "
                cmdsql = cmdsql + "('" + lblCustId + "', "
                cmdsql = cmdsql + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                cmdsql = cmdsql + "now(),"
                cmdsql = cmdsql + "'" + lblaoc.Caption + "','P')"
                M_OBJCONN.execute cmdsql
                
                Set ListItem = LstPayment.ListItems.ADD(, , "")
                ListItem.SubItems(1) = ""
                ListItem.SubItems(2) = Format(TDBDate3.Value, "dd/mm/yyyy")
                ListItem.SubItems(3) = CStr(Tdabamoint.Value)
                ListItem.SubItems(4) = "IPO"
                ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
            
    

        n = 0
        For I = 1 To Val(txttenor - 1)
            n = n + 1
            JmlPay = (txtPayment - Tdabamoint) / (txttenor.Value - 1)
            'VRDATE = Format(DateAdd("m", n, TDBDate3.Value), "mm/dd/yyyy")
            Vrdate = DateAdd("m", n, Format(TDBDate3.Value, "yyyy-mm-dd"))
            
                '@@14-04-2012 Cek Data
                cmdsql = "select * from tblreserve where custid='"
                cmdsql = cmdsql + lblCustId.Caption + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(Vrdate, "yyyy-mm-dd")) + "' and stsmove='0'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblreserve where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            cmdsql = "INSERT INTO tblreserve "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblCustId + "', "
            cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "now(),"
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.execute cmdsql
            
            
            '@@14-04-2012 Cek Data
                cmdsql = "select * from tblnegoptp_log where custid='"
                cmdsql = cmdsql + lblCustId.Caption + "' and date(promisedate)='"
                cmdsql = cmdsql + CStr(Format(Vrdate, "yyyy-mm-dd")) + "' "
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        cmdsql = "delete from tblnegoptp_log where id='"
                        cmdsql = cmdsql + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.execute cmdsql
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            
            cmdsql = "INSERT INTO TblNegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblCustId + "', "
            cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "now(),"
            cmdsql = cmdsql + "'" + lblaoc.Caption + "','R')"
            M_OBJCONN.execute cmdsql

        Set ListItem = LstReserve.ListItems.ADD(, , "")
            ListItem.SubItems(1) = ""
                               'listitem.SubItems(2) = .TDBDate1.Value
            ListItem.SubItems(2) = Format(Vrdate, "dd/mm/yyyy")
            ListItem.SubItems(3) = JmlPay
            ListItem.SubItems(4) = "IPO"
            ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
    Next I
   End If
   
         '   regnego = True
          '  FrmNegoPTP.Show
            
'        With FrmNegoPTP
'                .Caption = "Tambah Data"
'                .Show vbModal
'                If .ok Then
'                 M_DATA.ADD_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), MDIForm1.TDBDate1.Value, jenis
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        Set listitem = LstPayment.ListItems.ADD(, , "")
'                            listitem.SubItems(1) = ""
'                            listitem.SubItems(2) = .TDBDate1.Value
'                            listitem.SubItems(3) = .TDBNumber1.Value
'                      On Error GoTo 0
'                    End If
'                End If
'                Unload FrmNegoPTP
'            End With
'        Exit Sub
     
    
    Case 1
        Dim M_Cek_Status As ADODB.Recordset
        Dim Cmdsql_Cek As String
        
        If LstPayment.ListItems.Count = 0 Then
            Exit Sub
        End If
        
        '@@ 11-04-2012 Cek status account terlebih dahulu, data bisa diedit jika status account PTP
        Cmdsql_Cek = "select f_cek_new from mgm where custid='"
        Cmdsql_Cek = Cmdsql_Cek + lblCustId.Caption + "'"
        Set M_Cek_Status = New ADODB.Recordset
        M_Cek_Status.CursorLocation = adUseClient
        M_Cek_Status.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If IsNull(M_Cek_Status("f_cek_new")) = True Then
            MsgBox "Data hanya dapat diedit jika status account=PTP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
        If Mid(M_Cek_Status("f_cek_new"), 1, 3) <> "PTP" Then
            MsgBox "Data hanya dapat diedit jika status account=PTP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
           With FrmNegoPTP
                .Caption = "Ubah Data"
                .SSCommand1(0).Caption = "Update"
                .TDBDate1.Value = Format(LstPayment.SelectedItem.SubItems(2), "dd/mm/yyyy")
                .TDBNumber1.Value = LstPayment.SelectedItem.SubItems(3)
                .Show vbModal
                If .ok Then
                    
                    '@@ Buat Update Tanggal Tagih
                    If C_PTP.Value = vbChecked Then
                                
                        '@@ 05-10-2011, Jika melakukan PTP maka combo via ptp harus diisi
                        If CmbViaPtp.text = "" Then
                            MsgBox "Combo Via tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
                            MsgBox "Data gagal diupdate!", vbOKOnly + vbInformation, "Informasi"
                            Unload Me
                            Exit Sub
                        End If
            
                        'Tambahan, Jika Status data PTP, hitung tanggal tagih
                        If TDBDate3.ValueIsNull Then
                            MsgBox "Anda belum menentukan tanggal effective pembayaran!", vbOKOnly + vbInformation, "Informasi"
                            MsgBox "Data gagal diupdate!", vbOKOnly + vbInformation, "Informasi"
                            Unload Me
                            Exit Sub
                        End If
            
                    End If
                    
                    
                    
                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.text, Format(.TDBDate1.Value, "yyyy-mm-dd"), CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)

                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        'LstPayment.SelectedItem.SubItems(1) = ""
                        LstPayment.SelectedItem.SubItems(2) = Format(.TDBDate1.Value, "mm/dd/yyyy")
                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
                        
                        Call CariTanggalTagih
                        
                        cmdsql = "update mgm set tgl_tagih='"
                        cmdsql = cmdsql + Format(TdbTglTagih.Value, "yyyy-mm-dd") + "' where custid='"
                        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
                        M_OBJCONN.execute cmdsql
                        
                    On Error GoTo 0
                    End If
                End If
                Unload FrmNegoPTP
            End With
        Exit Sub
    Case 2
         Frmdelete.Show vbModal
'        If LstPayment.ListItems.Count = 0 Then
'            Exit Sub
'        End If
'        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
'        If m_msgbox = 1 Then
'            M_DATA.DELETE_Nego_PTP M_OBJCONN, LstPayment.SelectedItem.SubItems(1)
'            If M_DATA.ADD_OK Then
'                LstPayment.ListItems.Remove LstPayment.SelectedItem.Index
'            End If
'        End If
'        Exit Sub
    
    Case 3
        MsgBox "Tidak dapat hapus reserved PTP!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
        frmdeletereserve.Show vbModal
End Select
add_error:
End Sub
Private Sub VisitYES()
Text1.BackColor = &HFF00&
TxtCustid.BackColor = &H80000005
TxtName.BackColor = &H80000005
TDBNumber1.BackColor = &H80000005
TXtDetails.BackColor = &H80000005
'LstVisit.BackColor = &HFF00&
TxtAddress.BackColor = &H80000005
TxtAddress.Enabled = True
TXtDetails.Enabled = True
Option7(0).Enabled = True
Option7(1).Enabled = True
Option7(2).Enabled = True
End Sub
Private Sub VisitNo()
Text1.BackColor = &H8000000F
TxtCustid.BackColor = &H8000000F
TxtName.BackColor = &H8000000F
TDBNumber1.BackColor = &H8000000F
TXtDetails.BackColor = &H8000000F
TxtAddress.BackColor = &H8000000F
'LstVisit.BackColor = &H8000000F
Option8(1).Value = True
Option7(0).Enabled = False
Option7(1).Enabled = False
Option7(2).Enabled = False

TxtAddress.Enabled = False
TXtDetails.Enabled = False
End Sub

Private Sub Tdabamoint_Change()
bcekptp = False
End Sub

Private Sub TDBDate3_Change()
   Dim cmdsql As String
   Dim M_Objrs As ADODB.Recordset
   Dim TglPtp As String
   
   If C_PTP.Value Then
        '@@ 09-04-2012
        Call CariTanggalTagih
        'Update tanggal negoptp
        cmdsql = "select * from tblnegoptp where custid='"
        cmdsql = cmdsql + lblCustId.Caption + "'"
        cmdsql = cmdsql + " order by promisedate desc limit 1"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs.RecordCount = 0 Then
             Set M_Objrs = Nothing
             Exit Sub
        End If
        
        If TDBDate3.Value = Empty Then
             TglPtp = "null"
        Else
             TglPtp = "'" + Format(TDBDate3.Value, "yyyy-mm-dd") + "'"
        End If
        
        On Error GoTo SALAH
        cmdsql = "update tblnegoptp set promisedate="
        cmdsql = cmdsql + TglPtp + " where id='"
        cmdsql = cmdsql + CStr(M_Objrs("id")) + "'"
        M_OBJCONN.execute cmdsql
        
        Call Show_NEGOPTP
   End If
   Exit Sub
SALAH:
   MsgBox "Ada error: " & err.Description
End Sub

Private Sub TdbPTP_Change()
TdbPTP.Value = TDBDate1.Value
End Sub

Private Sub Timer1_Timer()

End Sub

'Private Sub Timer_cek_inbox_Timer()
''@@ 11-03-2011 Di remarks, udah tidak diapakai
''    Text2 = LstSMS.ListItems.Count
''
''    LstSMS.ListItems.CLEAR
''    LstSMS2.ListItems.CLEAR
''    Isi_SendSMS
''    Isi_SendSMS2
'End Sub

Private Sub blink(Seconds As Single)
 Dim a As Long
 Seconds = Seconds + Timer
 While Seconds > Timer
  a = DoEvents
 Wend
End Sub

Private Sub TimerBlink_Timer()
'@@ 05-10-2011 tombol OST ditiadakan
   
'               If SSCommand1(7).BackColor = vbRed Then
'                 SSCommand1(7).BackColor = vbGreen
'                 KelapKelip = KelapKelip + 1
'               Else
'                 SSCommand1(7).BackColor = vbRed
'                 KelapKelip = KelapKelip + 1
'               End If
'
'           If KelapKelip = 7 Then
'            KelapKelip = 0
'            WaitSecs (3)
'            'TimerBlink.Enabled = False
'           End If
    
End Sub

Private Sub BlinkCPA_Timer()
    Dim kelapkelipCpa As Integer
    
    If SSCommand1(4).BackColor = vbBlack Then
        SSCommand1(4).BackColor = vbRed
        kelapkelipCpa = kelapkelipCpa + 1
    Else
        SSCommand1(4).BackColor = vbBlack
        kelapkelipCpa = kelapkelipCpa + 1
    End If
           
    If kelapkelipCpa = 7 Then
            kelapkelipCpa = 0
            WaitSecs (3)
            SSCommand1(4).BackColor = vbBlack
            TimerBlinkCPA.Enabled = False
    End If
End Sub

Private Sub TimerBlinkCPA_Timer()

End Sub

Private Sub TimerBlinkDetailMapping_Timer()
    'Dim kelapkelipDetail As Integer
    
    If Val(LblMap.Caption) > 0 Then
        If LblMap.BackColor = vbBlack Then
            LblMap.BackColor = vbRed
            kelapkelipDetail = kelapkelipDetail + 1
        Else
            LblMap.BackColor = vbBlack
            kelapkelipDetail = kelapkelipDetail + 1
        End If
               
'        If kelapkelipDetail = 7 Then
'                kelapkelipDetail = 0
'                WaitSecs (3)
'                LblMap.BackColor = vbBlack
'                TimerBlinkDetailMapping.Enabled = False
'        End If
    Else
        TimerBlinkDetailMapping.Enabled = False
    End If
End Sub

Private Sub TimerBlinkSms_Timer()
    If LabelSms.ForeColor = vbBlack Then
        LabelSms.ForeColor = vbRed
        Command2.BackColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        LabelSms.ForeColor = vbBlack
        Command2.BackColor = vbYellow
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
            KelapKelip = 0
            WaitSecs (3)
            'TimerBlink.Enabled = False
    End If
End Sub

Private Sub TimerCekMapping_Timer()
     If CmdDataMapping.BackColor = vbGreen Then
        CmdDataMapping.BackColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        CmdDataMapping.BackColor = vbYellow
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
            KelapKelip = 0
            WaitSecs (3)
            'TimerBlink.Enabled = False
    End If
End Sub



Private Sub TimerOfferingDiscon_Timer()
    OfferingDiscGuide
    TimerOfferingDiscon.Enabled = False
End Sub

'Private Sub TimerCekSms_Timer()
'
'    On Error Resume Next
'    Dim M_OBJRS As New ADODB.Recordset
'    Dim cmdsql34 As String
'    Dim TELPo As String
'    Dim codea As String
'    Dim m_objrscek As ADODB.Recordset
'
'    If Left(MDIForm1.Text1, 1) = "D" Or Text1 = "JOKO" Or Text1 = "SPV1" Or Left(MDIForm1.Text1, 1) = "T" Then
'        Select Case Text1.Text
'            Case "TL1"
'                codea = "ACC1"
'            Case "TL2"
'                codea = "ACC2"
'            Case "TL3"
'                codea = "ACC3"
'            Case "TL4"
'                codea = "ACC4"
'            Case "TL5"
'                codea = "ACC5"
'            Case "TL6"
'                codea = "ACC6"
'            Case "TL7"
'                codea = "ACC7"
'            Case "TL8"
'                codea = "ACC8"
'            Case "TL9"
'                codea = "ACC9"
'            Case "TL10"
'                codea = "ACC10"
'            Case Else
'                codea = MDIForm1.Text1.Text
'        End Select
'
'        TELPo = "Select count(*) as banyak from inbox where sendernumber in ('a',"
'
'        Set M_OBJRS = New ADODB.Recordset
'        M_OBJRS.CursorLocation = adUseClient
'        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + codea + "'"
'        M_OBJRS.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If M_OBJRS.RecordCount = 0 Then
'            Timer6.Interval = 60000
'            Exit Sub
'        End If
'
'        While Not M_OBJRS.EOF
'
'            If Len(M_OBJRS("mobileno")) <> 0 Then
'                satu = FindReplace(M_OBJRS("mobileno"), "0", "+62")
'                TELPo = TELPo + "'" + satu + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            If Len(M_OBJRS("mobileno2")) <> 0 Then
'                dua = FindReplace(M_OBJRS("mobileno2"), "0", "+62")
'                TELPo = TELPo + "'" + dua + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            If Len(M_OBJRS("mobilenoadd1")) <> 0 Then
'                tiga = FindReplace(M_OBJRS("mobilenoadd1"), "0", "+62")
'                TELPo = TELPo + "'" + tiga + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            If Len(M_OBJRS("mobilenoadd2")) <> 0 Then
'                empat = FindReplace(M_OBJRS("mobilenoadd2"), "0", "+62")
'                TELPo = TELPo + "'" + empat + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            M_OBJRS.MoveNext
'        Wend
'        Set M_OBJRS = Nothing
'
'
'        TELPo = Left(TELPo, Len(TELPo) - 1)
'        Dim TELPo1
'
'
'        TELPo1 = TELPo + ") and processed='f'"
'        'TELPo2 = TELPo + ") and processed='t'"
'
'        Set m_objrscek = New ADODB.Recordset
'        m_objrscek.CursorLocation = adUseClient
'        m_objrscek.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
''        While Not M_OBJRS.EOF
''            'LblJmlSmsBaru.Caption = M_OBJRS("banyak")
''            LabelSms.Caption = "ADA SMS BARU!" '& LblJmlSmsBaru.Caption & " SMS"
''            M_OBJRS.MoveNext
''        Wend
'
''        'JIKA ADA SMS BARU MASUK
''        If Trim(LabelSms.Caption) = "SMS BARU 0 SMS" Then
''            'MsgBox "Tidak ada sms baru!"
''            TimerBlink.Enabled = False
''            LabelSms.ForeColor = vbBlack
''        Else
''            If Trim(LabelSms.Caption) <> "" Then
''                TimerBlink.Enabled = True
''                MsgBox "Ada SMS BARU MASUK! Silahkan cek!", vbOKOnly + vbInformation, "Informasi"
''            End If
''        End If
'         If m_objrscek(0) > 0 Then
'            TimerBlinkSms.Enabled = True
'            LabelSms.Caption = "Ada SMS Baru!"
'         Else
'            LabelSms.Caption = "Tidak ada SMS baru!"
'            LabelSms.ForeColor = vbBlack
'            Command2.BackColor = vbGreen
'            TimerBlinkSms.Enabled = False
'         End If
'
'        Set m_objrscek = Nothing
'End If
'        Timer6.Interval = 60000
'End Sub



Private Sub txtECno_Click()
TYPETELP = "Emergency Contact"
txtPhone.text = txtECno.Value
txtPhoneA.text = txtECnoA.Value
CmbPhone.text = "EconPhone"
End Sub


Private Sub txtECnoA_Change()
'txtECno.Text = txtECnoA.Text
End Sub

Private Sub txtECnoA_Click()
TYPETELP = "Emergency Contact"
txtPhone.text = txtECno.Value
txtPhoneA.text = txtECnoA.Value
CmbPhone.text = "EconPhone"
End Sub

Private Sub txtFaxAdd1_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox "Anda tidak boleh mengisi di fax, kecuali SPV!"
End Sub

Private Sub txtFaxAdd2_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox "Anda tidak boleh mengisi di fax, kecuali SPV!"
End Sub
Private Sub txtECnoA_DblClick()
txthasil.text = txtECno.text
End Sub

Private Sub txthasil_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtHomeAdd1_Click()
TYPETELP = "HOME1"
    '@@03-05-2012 DinonAktifkan
'    If Trim(AHomeAdd1(0).Value) = "031" Or AHomeAdd1(0).Value = "" Then
'        txtPhone.Text = txtHomeAdd1.Value
'        txtPhoneA.Text = txtHomeAdd1.Value
'    Else
'        txtPhone.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
'        txtPhoneA.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
'    End If
    CmbPhone.text = "AddHome1"
End Sub
Private Sub txtHomeAdd1A_Click()
TYPETELP = "HOME1"
    '@@03-05-2012 Dinonaktifkan
'    If Trim(AHomeAdd1(0).Value) = "031" Or AHomeAdd1(0).Value = "" Then
'        txtPhone.Text = txtHomeAdd1.Value
'        txtPhoneA.Text = txtHomeAdd1A.Value
'
'    Else
'        txtPhone.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
'        txtPhoneA.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1A.Value)
'    End If
    CmbPhone.text = "AddHome1"
End Sub

Private Sub txtHomeAdd1A_DblClick()
txthasil.text = txtHomeAdd1.text

End Sub

Private Sub txtHomeAdd2_Click()
TYPETELP = "HOME2"
'@@03-05-2012 Dinonaktikan
'If Trim(AHomeAdd2(1).Value) = "031" Or AHomeAdd2(1).Value = "" Then
'    txtPhone.Text = txtHomeAdd2.Value
'    txtPhoneA.Text = txtHomeAdd2.Value
'Else
'    txtPhone.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
'    txtPhoneA.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
'End If
CmbPhone.text = "AddHome2"
End Sub
Private Sub txtHomeAdd2A_Change()
'txtHomeAdd2.Text = txtHomeAdd2A.Text
End Sub
Private Sub txtHomeAdd2A_Click()
TYPETELP = "HOME2"
'@@03-05-2012 Dinonaktifkan
'If Trim(AHomeAdd2(1).Value) = "031" Or AHomeAdd2(1).Value = "" Then
'    txtPhone.Text = txtHomeAdd2.Value
'    txtPhoneA.Text = txtHomeAdd2A.Value
'Else
'    txtPhone.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
'    txtPhoneA.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2A.Value)
'End If
CmbPhone.text = "AddHome2"
End Sub

Private Sub txtHomeAdd2A_DblClick()
txthasil.text = txtHomeAdd2.text
End Sub

Private Sub txtHomeNo1_Click()
    If Len(txtHomeNo1.text) > 3 Then
    CmbPhone.text = "HomePhone"
    Else
    CmbPhone.text = ""
    End If
End Sub

Private Sub txtHomeNo1A_Click()
If Len(txtHomeNo1A.text) > 3 Then
    CmbPhone.text = "HomePhone"
    Else
    CmbPhone.text = ""
    End If
End Sub
Private Sub txtHomeNo1A_DblClick()
txthasil.text = txtHomeNo1.text
End Sub

Private Sub txtHomeNo2_Click()
    If Len(txtHomeNo2.text) > 3 Then
    CmbPhone.text = "HomePhone2"
    Else
    CmbPhone.text = ""
    End If
End Sub

Private Sub txtHomeNo2A_Click()
  If Len(txtHomeNo2A.text) > 3 Then
    CmbPhone.text = "HomePhone2"
    Else
    CmbPhone.text = ""
    End If
End Sub
Private Sub txtHomeNo2A_DblClick()
txthasil.text = txtHomeNo2.text
End Sub

Private Sub txtMobileAdd1A_Click()
TYPETELP = "MOBILE1"
    txtPhone.text = txtMobileAdd1.Value
    txtPhoneA.text = txtMobileAdd1A.Value
    
    CmbPhone.text = "AddMobile1"
End Sub

Private Sub txtMobileAdd1A_DblClick()
txthasil.text = txtMobileAdd1.text
End Sub

Private Sub txtMobileAdd2A_Change()
'    txtMobileAdd2.Text = txtMobileAdd2A.Text
End Sub
Private Sub txtMobileAdd2A_Click()
TYPETELP = "MOBILE2"
    txtPhone.text = txtMobileAdd2.Value
    txtPhoneA.text = txtMobileAdd2A.Value
    If Len(txtMobileAdd2A.text) > 3 Then
    CmbPhone.text = "AddMobile2"
    Else
    CmbPhone.text = ""
    End If
End Sub

Private Sub txtMobileAdd2A_DblClick()
txthasil.text = txtMobileAdd2.text
End Sub

Private Sub txtMobileNo1_Click()
If Len(txtMobileNo1.text) > 3 Then
CmbPhone.text = "Hp"
Else
CmbPhone.text = ""
End If
End Sub

Private Sub txtMobileNo1A_Click()
If Len(txtMobileNo1A.text) > 3 Then
CmbPhone.text = "Hp"
Else
CmbPhone.text = ""
End If
End Sub

Private Sub txtMobileNo1A_DblClick()
txthasil.text = txtMobileNo1.text
End Sub

Private Sub txtMobileNo2_Click()
If Len(txtMobileNo2.text) > 3 Then
CmbPhone.text = "Hp2"
Else
CmbPhone.text = ""
End If
End Sub
Private Sub txtMobileNo2A_Click()
If Len(txtMobileNo2A.text) > 3 Then
CmbPhone.text = "Hp2"
Else
CmbPhone.text = ""
End If
End Sub
Private Sub txtMobileNo2A_DblClick()
    txthasil.text = txtMobileNo2.text
End Sub

Private Sub txtOfficeAdd1_Click()
TYPETELP = "OFFICE1"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(2).Value) = "031" Or AOfficeAdd(2).Value = "" Then
'    txtPhone.Text = txtOfficeAdd1.Value
'    txtPhoneA.Text = txtOfficeAdd1.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'End If
CmbPhone.text = "AddOffice1"
End Sub

Private Sub txtOfficeAdd1A_Change()
'    txtOfficeAdd1.Text = txtOfficeAdd1A.Text
End Sub

Private Sub txtOfficeAdd1A_Click()
TYPETELP = "OFFICE1"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(2).Value) = "031" Or AOfficeAdd(2).Value = "" Then
'    txtPhone.Text = txtOfficeAdd1.Value
'    txtPhoneA.Text = txtOfficeAdd1A.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1A.Value)
'End If
CmbPhone.text = "AddOffice1"
End Sub
Private Sub txtOfficeAdd1A_DblClick()
    txthasil.text = txtOfficeAdd1.text
End Sub

Private Sub txtOfficeAdd2_Click()
TYPETELP = "OFFICE2"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(3).Value) = "031" Or AOfficeAdd(3).Value = "" Then
'    txtPhone.Text = txtOfficeAdd2.Value
'    txtPhoneA.Text = txtOfficeAdd2.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'End If
CmbPhone.text = "AddOffice2"
End Sub

Private Sub txtMobileAdd1_Click()
TYPETELP = "MOBILE1"
    txtPhone.text = txtMobileAdd1.Value
    txtPhoneA.text = txtMobileAdd1.Value
If Len(txtMobileAdd1.text) > 3 Then
    CmbPhone.text = "AddMobile1"
    Else
    CmbPhone.text = ""
End If
End Sub

Private Sub txtMobileAdd2_Click()
TYPETELP = "MOBILE2"
    txtPhone.text = txtMobileAdd2.Value
    txtPhoneA.text = txtMobileAdd2.Value

If Len(txtMobileAdd2.text) > 3 Then
    CmbPhone.text = "AddMobile2"
    Else
    CmbPhone.text = ""
End If
    
End Sub
Public Sub UpdateAppv()
'If chkAppv(0).Value Then
'    x = MsgBox("Pindahkan data ke Agent DA ?", vbYesNo + vbExclamation, "Info !")
'    If x = vbYes Then
'        CMDSQL = "update mgm set F_pending='Pending',Agent='DA',PO_Agent='" & lblaoc.Caption & "' where custid='" + lblCustId.Caption + "'"
'        M_OBJCONN.Execute CMDSQL
'        spend = True
'        MsgBox "Data berhasil dipindah ke agent DA", vbInformation
'        VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
'        MDIForm1.LstGrade.ListItems.CLEAR
'    End If
'Else
'    If chkAppv(1).Value Then
'        Dim spo As ADODB.Recordset
'        Set spo = New ADODB.Recordset
'        spo.CursorLocation = adUseClient
'        spo.Open "select PO_Agent from mgm where custid='" + lblCustId.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        If spo!PO_AGENT <> "" And IsNull(spo!PO_AGENT) = False Then
'            CMDSQL = "update mgm set F_pending='',AGENT=PO_Agent where custid='" + lblCustId.Caption + "'"
'            M_OBJCONN.Execute CMDSQL
'            CMDSQL = "update mgm set PO_Agent='' where custid='" + lblCustId.Caption + "'"
'            M_OBJCONN.Execute CMDSQL
'            MsgBox "Data berhasil dikembalikan", vbInformation
'            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
'            MDIForm1.LstGrade.ListItems.CLEAR
'        Else
'            MsgBox "Silahkan Pilih Status !," & vbCrLf & "untuk menyimpan hilangkan ceklist NO !", vbInformation
'            Exit Sub
'        End If
'    End If
'End If
End Sub

Private Sub txtOfficeAdd2A_Change()
'    txtOfficeAdd2.Text = txtOfficeAdd2A.Text
End Sub

Private Sub txtOfficeAdd2A_Click()
TYPETELP = "OFFICE2"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(3).Value) = "031" Or AOfficeAdd(3).Value = "" Then
'    txtPhone.Text = txtOfficeAdd2.Value
'    txtPhoneA.Text = txtOfficeAdd2A.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2A.Value)
'End If

CmbPhone.text = "AddOffice2"
End Sub

Private Sub txtOfficeAdd2A_DblClick()
txthasil.text = txtOfficeAdd2.text
End Sub

Private Sub txtOfficeNo1_Click()
If Len(txtOfficeNo1.text) > 3 Then
CmbPhone.text = "OfficePhone"
Else
CmbPhone.text = ""
End If
End Sub
Private Sub txtOfficeNo1A_DblClick()
 txthasil.text = txtOfficeNo1.text
End Sub

Private Sub txtOfficeNo1A_Click()
If Len(txtOfficeNo1A.text) > 3 Then
CmbPhone.text = "OfficePhone"
Else
CmbPhone.text = ""
End If

End Sub
Private Sub txtOfficeNo2_Click()
If Len(txtOfficeNo2.text) > 3 Then
CmbPhone.text = "OfficePhone2"
Else
CmbPhone.text = ""
End If

End Sub
Private Sub txtOfficeNo2A_Click()
If Len(txtOfficeNo2A.text) > 3 Then
CmbPhone.text = "OfficePhone2"
Else
CmbPhone.text = ""
End If

End Sub
Public Sub Show_Reserve()
Dim showlist As New ADODB.Recordset
Dim ListItem As ListItem
Dim cmdsql As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblCustId.Caption + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
End If
'If ShowList.BOF And ShowList.EOF Then
'    'CMDSQL = "SELECT * FROM TBLNEGOPTP WHERE custid = '" + lblCustId.Caption + "'"
'    'AND CUSTID NOT IN (SELECT CUSTID FROM tbllunas)"
'    CMDSQL = "SELECT DISTINCT TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.ID,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.TYPE FROM TBLNEGOPTP,tbllunas WHERE "
'    CMDSQL = CMDSQL + "tbllunas.CUSTID<>TBLNEGOPTP.CUSTID AND TBLNEGOPTP.CUSTID='" + lblCustId.Caption + "' order by TBLNEGOPTP.promisedate desc"
'Else
'    CMDSQL = "SELECT distinct TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.ID,TBLNEGOPTP.TYPE "
'    CMDSQL = CMDSQL + "FROM VWLISTPTP,TBLNEGOPTP WHERE TBLNEGOPTP.CUSTID=VWLISTPTP.CUSTID AND "
'    CMDSQL = CMDSQL + "VWLISTPTP.PAYDATE<TBLNEGOPTP.PROMISEDATE AND TBLNEGOPTP.CUSTID='" + lblCustId.Caption + "' order by TBLNEGOPTP.promisedate desc"
'End If
If MDIForm1.Text2.text = "SUPERVISOR" Then
    cmdsql = "SELECT * FROM tblreserve where custid = '" + lblCustId.Caption + "' order by promisedate"
Else
    cmdsql = "SELECT * FROM tblreserve where custid = '" + lblCustId.Caption + "' and stsmove=0 order by promisedate"
End If

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstReserve.ListItems.clear
Dim n As Currency
While Not showlist.EOF
    Set ListItem = LstReserve.ListItems.ADD(, , "")
        ListItem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
        ListItem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "dd/mm/yyyy")))
        ListItem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", (Round(showlist!PromisePay, 0))))
        n = n + Val(ListItem.SubItems(3))
        If n <= TOTPTP Then
            ListItem.ListSubItems(1).ForeColor = vbRed
            ListItem.ListSubItems(2).ForeColor = vbRed
            ListItem.ListSubItems(3).ForeColor = vbRed
        End If
        
        ListItem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
        ListItem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "dd/mm/yyyy")))
     showlist.MoveNext
Wend

Set showlist = Nothing
End Sub

Private Sub txtOfficeNo2A_DblClick()
txthasil.text = txtOfficeNo2.text
End Sub

Public Sub PesanLockAuto()
    Dim m_objrsPesanReset As ADODB.Recordset
    Dim m_objrsPesanLock As ADODB.Recordset
    Dim M_ObjWktServer As ADODB.Recordset
    Dim WaktuServer As Date
    Dim cmdsql As String
    
    'Ambil Waktu Server Sekarang
    Set M_ObjWktServer = New ADODB.Recordset
    M_ObjWktServer.CursorLocation = adUseClient
    M_ObjWktServer.Open "Select now() as WktSrv ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    WaktuServer = Format(M_ObjWktServer(0), "yyyy-mm-dd hh:mm")
    Set M_ObjWktServer = Nothing
    
    'Cek pesan reset
    cmdsql = "select f_pesanresetauto,f_idsessend from usertbl where userid='"
    cmdsql = cmdsql + Trim(MDIForm1.Text1.text) + "'"
    Set m_objrsPesanReset = New ADODB.Recordset
    m_objrsPesanReset.CursorLocation = adUseClient
    m_objrsPesanReset.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If m_objrsPesanReset.RecordCount <> 0 Then
        If m_objrsPesanReset("f_pesanresetauto") = "1" Then
            MsgBox "Reset Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
           
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.clear
            '@@20-11-10 akhiri session dengan mencatat hasil akhir perubahan status data yang dikerjain agent
                If m_objrsPesanReset("f_idsessend") <> "" Or IsNull(m_objrsPesanReset("f_idsessend")) = False Or m_objrsPesanReset("f_idsessend") <> Empty Then
                    Dim UpdateDtCloseSession As String
'                    UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "' from "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
'                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(m_objrsPesanReset("f_idsessend")) + "' and tblperformpersessionlock.agent='"
'                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.Text1.Text) + "'"
'                    M_OBJCONN.Execute UpdateDtCloseSession
                    'bikin null lagi nilai f_idsessend
                    UpdateDtCloseSession = "update usertbl set f_idsessend=null where userid='"
                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.Text1.text) + "'"
                    M_OBJCONN.execute UpdateDtCloseSession
                End If
            '@@20-11-10 akhiri session dengan mencatat hasil akhir perubahan status data yang dikerjain agent
             
            cmdsql = "update usertbl set f_pesanresetauto=null where userid='"
            cmdsql = cmdsql + Trim(MDIForm1.Text1.text) + "'"
            M_OBJCONN.execute cmdsql
        End If
    End If
    
    Set m_objrsPesanReset = Nothing
    
    'Cek pesan Lock
    cmdsql = "select f_pesanlockauto from usertbl where userid='"
    cmdsql = cmdsql + Trim(MDIForm1.Text1.text) + "'"
    Set m_objrsPesanLock = New ADODB.Recordset
    m_objrsPesanLock.CursorLocation = adUseClient
    m_objrsPesanLock.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If m_objrsPesanLock.RecordCount <> 0 Then
        If m_objrsPesanLock("f_pesanlockauto") = "1" Then
            MsgBox "Lock Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
            cmdsql = "update usertbl set f_pesanlockauto=null where userid='"
            cmdsql = cmdsql + Trim(MDIForm1.Text1.text) + "'"
            M_OBJCONN.execute cmdsql
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.clear
        End If
     End If
    
    Set m_objrsPesanLock = Nothing
End Sub

'@@ 14031011
Private Sub CekSms()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    '@@ 14/02/2010,, Cek smsnya melalui field blink di usertbl aja, jadinya lebih ringan
    If UCase(Trim(MDIForm1.Text2.text)) = "AGENT" Then
        cmdsql = "select status_sms from usertbl where userid='"
        cmdsql = cmdsql + Trim(MDIForm1.Text1.text) + "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs("status_sms") <> "" Then
            TimerBlinkSms.Enabled = True
            LabelSms.Caption = "Ada SMS Baru!"
        Else
            LabelSms.Caption = "Tidak ada SMS baru!"
            LabelSms.ForeColor = vbBlack
            Command2.BackColor = vbGreen
            TimerBlinkSms.Enabled = False
        End If
        
        Set M_Objrs = Nothing
    End If
End Sub



'@@ 08-03-2011 Cek data mapping
Private Sub CekDataMapping()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    
    cmdsql = "select * from mgm_mapping_pil where custidcard='"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "' or ktpno='"
    '@@ 25-07-2011 , Tambahan cari juga berdasarkan Nomor KTP
    cmdsql = cmdsql + Trim(lblID.Caption) + "'"

    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
    
    
    If M_Objrs.RecordCount = 0 Then
        CmdDataMapping.BackColor = vbGreen
        TimerCekMapping.Enabled = False
    Else
        TimerCekMapping.Enabled = True
    End If
        
    Set M_Objrs = Nothing
End Sub

'@@ 15-04-2011, Cek CPA , jika ada data cpa  maka kelap-kelip
Private Sub CekCPA()
    Dim Strsql As String
    Dim M_Objrs As ADODB.Recordset
    
    Strsql = "select * from tblcpa where vcustid='" + Trim(lblCustId.Caption) + "'"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        TimerBlinkCPA.Enabled = True
    Else
         TimerBlinkCPA.Enabled = False
    End If
    Set M_Objrs = Nothing
End Sub

'@@ 06-May 2011 Tambahan Offering Discon Guide
Private Sub OfferingDiscGuide()
    '@@06 May 2011 Tambahan Offering
        Dim K As Integer
        Dim W As String
        Dim l As Integer
        Dim diskon As Integer
        
        Dim M_Objrs As ADODB.Recordset
        Dim m_objrs_waktu As ADODB.Recordset
        Dim cmdsql As String
              
        
        'Cek dulu ada pembayaran apa ngga di tabel lunas
        cmdsql = "select * from tbllunas where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        'Ambil waktu sekarang
        cmdsql = "select now() as waktu "
        Set m_objrs_waktu = New ADODB.Recordset
        m_objrs_waktu.CursorLocation = adUseClient
        m_objrs_waktu.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '@@ 08-06-2011, Jika lblpaydt=kosong on error goto salah
        On Error GoTo SALAH
        l = DateDiff("M", Format(lblPayDt.Value, "mm/dd/yyyy"), Format(CDate(m_objrs_waktu("waktu")), "mm/dd/yyyy"))
        
        '@@ 09-05-2011 Jika tidak ada nopay atau lpd > 4 bulan dari current date maka
        'tampilkan offering
        
        
        If M_Objrs.RecordCount = 0 Or _
            l > 4 Then
            On Error GoTo SALAH
            K = DateDiff("M", lblOpenDate.Value, lblBD.Value)
            If K < 12 Then
                W = "Penawaran Diskon Maximal 60%"
                diskon = 60
            ElseIf K >= 12 And K <= 17 Then
                W = "Penawaran Diskon Maximal 50%"
                diskon = 50
            ElseIf K >= 18 And K <= 36 Then
                W = "Penawaran Diskon Maximal 40%"
                diskon = 40
            ElseIf K > 37 Then
                W = "Cicilan panjang " & " dan diskon 30%"
                diskon = 30
            End If
        
            'MsgBox "Pemandu Offering: " & w, vbOKOnly + vbInformation, "Offering Disc.Guide..."
            'With FrmOfferingGuide
            ' Di hilangkan Cek email 09-04-2013 by gustav
'            With FRMSCRIPT
'                On Error Resume Next
'                .LblTextGuide.Caption = "Pemandu Offering: " & W
'                .TdbBalance.Value = lblAmount.Value
'                .TdbMaxDisc.Value = diskon
'                .Show vbModal
'            End With
        End If
        
        Set M_Objrs = Nothing
        Set m_objrs_waktu = Nothing
        Exit Sub
SALAH:
    Set M_Objrs = Nothing
    Set m_objrs_waktu = Nothing
    MsgBox "Ada error: " & err.Description
End Sub


'@@ 09092011, Skrip Ofering yang awalnya di FormOfferingGuide, Sekarang Dipindah ke FormScript
Private Sub OfferingDiscGuideNew()
    '@@06 May 2011 Tambahan Offering
        Dim K As Integer
        Dim W As String
        Dim l As Integer
        Dim diskon As Integer
        
        Dim M_Objrs As ADODB.Recordset
        Dim m_objrs_waktu As ADODB.Recordset
        Dim cmdsql As String
              
        
        'Cek dulu ada pembayaran apa ngga di tabel lunas
        cmdsql = "select * from tbllunas where custid='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        'Ambil waktu sekarang
        cmdsql = "select now() as waktu "
        Set m_objrs_waktu = New ADODB.Recordset
        m_objrs_waktu.CursorLocation = adUseClient
        m_objrs_waktu.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '@@ 08-06-2011, Jika lblpaydt=kosong on error goto salah
        On Error GoTo SALAH
        l = DateDiff("M", Format(lblPayDt.Value, "mm/dd/yyyy"), Format(CDate(m_objrs_waktu("waktu")), "mm/dd/yyyy"))
        
        '@@ 09-05-2011 Jika tidak ada nopay atau lpd > 4 bulan dari current date maka
        'tampilkan offering
        
        
        If M_Objrs.RecordCount = 0 Or _
            l > 4 Then
            On Error GoTo SALAH
            K = DateDiff("M", Format(lblOpenDate.Value, "mm/dd/yyyy"), Format(lblBD.Value, "mm/dd/yyyy"))
            If K < 12 Then
                W = "Penawaran Diskon Maximal 60%"
                diskon = 60
            ElseIf K >= 12 And K <= 17 Then
                W = "Penawaran Diskon Maximal 50%"
                diskon = 50
            ElseIf K >= 18 And K <= 36 Then
                W = "Penawaran Diskon Maximal 40%"
                diskon = 40
            ElseIf K > 37 Then
                W = "Cicilan panjang " & " dan diskon 30%"
                diskon = 30
            End If
        
            'MsgBox "Pemandu Offering: " & w, vbOKOnly + vbInformation, "Offering Disc.Guide..."
            With FRMSCRIPT
                .LblTextGuide.Caption = "Pemandu Offering: " & W
                .Tdbbalance.Value = lblAmount.Value
                .TdbMaxDisc.Value = diskon
                '.Show vbModal
            End With
        End If
        
        Set M_Objrs = Nothing
        Set m_objrs_waktu = Nothing
        Exit Sub
SALAH:
    Set M_Objrs = Nothing
    Set m_objrs_waktu = Nothing
End Sub

'@@22-09-2011 Hitung InstallmentPtp
Private Sub HitungInstallmentPtp()
    Dim installment As Double
    
    If txttenor.Value = 0 Then
        installment = txtPayment.Value / 1
    Else
        installment = txtPayment.Value / txttenor.Value
    End If
    Tdabamoint.Value = installment
End Sub

Private Sub txtPayment_Change()
    HitungInstallmentPtp
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtremarks.text)) >= 80 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttenor_Change()
    HitungInstallmentPtp
End Sub

Private Sub CariTanggalTagih()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim TglPaymentEffective As String
    
    If IsNull(TDBDate3.Value) = True Then
        MsgBox "Payment effective tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    TglPaymentEffective = Format(TDBDate3.Value, "yyyy-mm-dd")
    
    cmdsql = "Select  date('" + TglPaymentEffective + "')-"
    If UCase(Trim(CmbViaPtp.text)) = "HSBC" Then
        cmdsql = cmdsql + "1"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "BERSAMA" Then
        cmdsql = cmdsql + "1"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "KANTOR POS" Then
        cmdsql = cmdsql + "3"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "PUM" Then
        cmdsql = cmdsql + "1"
    Else
        cmdsql = cmdsql + "3"
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    On Error GoTo SALAH
    TdbTglTagih.Value = Format(M_Objrs(0), "mm/dd/yyyy")
    
    Set M_Objrs = Nothing
    Exit Sub
SALAH:
    MsgBox "Ada Error: " & err.Description
End Sub

'@@ 17-04-2012, Ini buat hitung durasi call
Private Sub HitungDurasiCall()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim JAM, Menit, Detik As Long
     
    cmdsql = "select id,enddate-tgl as durasi from tblphonemonitorhst where custid='"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and userid='"
    cmdsql = cmdsql + MDIForm1.Text1.text + "' order by id desc limit 1"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    DoEvents
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    JAM = Val(Mid(M_Objrs("durasi"), 1, 2)) * 3600
    Menit = Val(Mid(M_Objrs("durasi"), 4, 2)) * 60
    Detik = Val(Mid(M_Objrs("durasi"), 7, 2)) + JAM + Menit
    
    If Detik >= 40 Then
        cmdsql = "update tblphonemonitorhst set durasi='"
        cmdsql = cmdsql + CStr(Detik) + "', flag_review='1' where id='"
        cmdsql = cmdsql + CStr(M_Objrs("id")) + "'"
    Else
        cmdsql = "update tblphonemonitorhst set durasi='"
        cmdsql = cmdsql + CStr(Detik) + "' where id='"
        cmdsql = cmdsql + CStr(M_Objrs("id")) + "'"
    End If
    DoEvents
    M_OBJCONN.execute cmdsql
    Set M_Objrs = Nothing
End Sub

'@@ 19042012,, Buat Hitung Durasi Call dari Icentra
Private Sub HitungDurasiDariIcentra()
    Dim connIcentra As ADODB.Connection
    Dim StrKoneksi As String
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Initiate, Start, Finish As String
    Dim JAM, Menit, Detik As Long
    
    
    Set connIcentra = New ADODB.Connection
'    If Trim(MDIForm1.TxtIPIcentra.Text) = "192.168.10.4" Then
'       '-- Lokal --
'       'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
'       '-- Database --
'       StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
'    ElseIf Trim(MDIForm1.TxtIPIcentra.Text) = "192.168.10.5" Then
'       '-- Lokal --
'       'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_5; UID=admin; PWD=admin321"
'       '-- Database --
'       StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
'    Else
'        '@@ 02052012, Jika IP Kosong,, coba dicari dulu di database
'        Dim M_Objrs_IP_Icentra As ADODB.Recordset
'
'        CMDSQL = "select * from tbl_ip_icentra where ip='"
'        CMDSQL = CMDSQL + CStr(MDIForm1.WskCTI.LocalIP) + "'"
'        Set M_Objrs_IP_Icentra = New ADODB.Recordset
'        M_Objrs_IP_Icentra.CursorLocation = adUseClient
'        M_Objrs_IP_Icentra.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If M_Objrs_IP_Icentra.RecordCount = 0 Then
'            MDIForm1.TxtIPIcentra.Text = ""
'            Set M_Objrs_IP_Icentra = Nothing
'            '@@ Jika IP tidak ditemukan langsung exit, Tapi Cek dulu manual dengan
'            'menelusuri server 4 dan 5
'            'Call CariIPIcentra
'            '@@ 24 May 2012, Cari Berdasarkan Waktu Login aja
'            Call CariIPIcentraByWaktuLogin
'            Exit Sub
'        Else
'            MDIForm1.TxtIPIcentra.Text = IIf(IsNull(M_Objrs_IP_Icentra("ip_icentra")), "", Trim(M_Objrs_IP_Icentra("ip_icentra")))
'            StrKoneksi = "Driver={PostgreSQL ANSI}; Server=" & MDIForm1.TxtIPIcentra.Text & "; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
'            Set M_Objrs_IP_Icentra = Nothing
'        End If
'    End If
    '------------ LOKAL ICENTRA --------------------
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
    '------------ ICENTRA BANDUNG ---------------------
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.11.1; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    '------------ ICENTRA SURABAYA ----------------------
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.11.1; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo SALAH
    connIcentra.Open StrKoneksi
    
    cmdsql = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
    cmdsql = cmdsql + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and date(initiate)=date(now()) "
    cmdsql = cmdsql + " and start is not null and finish is not null  "
    cmdsql = cmdsql + " order by acd_log_outgoing_session_id desc limit 1 "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        'Pindahin data dari icentra ke database card
        Initiate = IIf(IsNull(M_Objrs("initiate")), "null", "'" & Format(M_Objrs("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
        Start = IIf(IsNull(M_Objrs("start")), "null", "'" & Format(M_Objrs("start"), "yyyy-mm-dd hh:mm:ss") + "'")
        Finish = IIf(IsNull(M_Objrs("finish")), "null", "'" & Format(M_Objrs("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
        
        'Hitung Konevrsi Selisih ke detik
        JAM = Val(Mid(M_Objrs("durasi"), 1, 2)) * 3600
        Menit = Val(Mid(M_Objrs("durasi"), 4, 2)) * 60
        Detik = Val(Mid(M_Objrs("durasi"), 7, 2)) + JAM + Menit
        
        cmdsql = "insert into outgoing_icentra (destination,"
        cmdsql = cmdsql + "initiate,start,finish,recording_filename,"
        cmdsql = cmdsql + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("destination")), "", CStr(M_Objrs("destination"))) + "',"
        cmdsql = cmdsql + Initiate + "," + Start + "," + Finish + ",'"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("recording_filename")), "", CStr(M_Objrs("recording_filename"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("campaign")), "", CStr(M_Objrs("campaign"))) + "','"
        cmdsql = cmdsql + CStr(Detik) + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(M_Objrs("acd_log_outgoing_session_id")) + "')"
        M_OBJCONN.execute cmdsql
    End If
    
    Set M_Objrs = Nothing
    Set connIcentra = Nothing
    Exit Sub
SALAH:
    Exit Sub
    'MsgBox "Anda tidak terhubung ke Icentra!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

'@@ 02052012, Tambahkan Pilihan Speak With
Private Sub PilihSpeakWith()
    cbolastcall.clear
    If UCase(Trim(TxtTelpKe.text)) = "OTHER CH OFFICE" Or _
       StsKategoriTelepon = "OTHER CH OFFICE" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman kantor"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "OTHER CH HOME" Or _
       StsKategoriTelepon = "OTHER CH HOME" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Pembantu/Supir"
        cbolastcall.AddItem "Kontrakan"
        cbolastcall.AddItem "Other"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "FAMILY" Or _
       StsKategoriTelepon = "FAMILY" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "NEIGHBOUR" Or _
       StsKategoriTelepon = "NEIGHBOUR" Then
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "RELATED PERSON" Or _
       StsKategoriTelepon = "RELATED PERSON" Then
        cbolastcall.AddItem "Lawyer"
        cbolastcall.AddItem "Teman"
        cbolastcall.AddItem "Other"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman kantor"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    
        
    If UCase(Trim(TxtTelpKe.text)) = "OTHER CH MOBILE" Or _
        StsKategoriTelepon = "OTHER CH MOBILE" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "SPOUSE"
        cbolastcall.AddItem "OTHER"
    End If
    
    If UCase(Trim(TxtTelpKe.text)) = "HOMEPHONE" Or _
       UCase(Trim(TxtTelpKe.text)) = "HOMEPHONE2" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Pembantu/Supir"
        cbolastcall.AddItem "Kontrakan"
        cbolastcall.AddItem "Other"
    End If
    
    If UCase(Trim(TxtTelpKe.text)) = "OFFICEPHONE" Or _
       UCase(Trim(TxtTelpKe.text)) = "OFFICEPHONE2" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman Kantor"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "ECONPHONE" Or _
       UCase(Trim(TxtTelpKe.text)) = "ECONPHONE" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "EC"
        cbolastcall.AddItem "LAWYER"
        cbolastcall.AddItem "Teman"
        cbolastcall.AddItem "OTHER"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman Kantor"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    
    If UCase(Trim(TxtTelpKe.text)) = "HP" Or _
       UCase(Trim(TxtTelpKe.text)) = "HP2" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Other"
    End If
    
    
    If UCase(Trim(TxtTelpKe.text)) = "OTHER EC" Or _
       StsKategoriTelepon = "OTHER EC" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "EC"
        cbolastcall.AddItem "LAWYER"
        cbolastcall.AddItem "Teman"
        cbolastcall.AddItem "OTHER"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman Kantor"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    
    cbolastcall.AddItem "UnReceive"
    
End Sub

Private Sub CariKategoriTlp()
    If StsKategoriTelepon = "OTHER CH OFFICE" Then
        KelompokKategoriTlp = "OCO"
    ElseIf StsKategoriTelepon = "OTHER CH HOME" Then
        KelompokKategoriTlp = "OCH"
    ElseIf StsKategoriTelepon = "FAMILY" Then
        KelompokKategoriTlp = "FAM"
    ElseIf StsKategoriTelepon = "NEIGHBOUR" Then
        KelompokKategoriTlp = "NEB"
    ElseIf StsKategoriTelepon = "RELATED PERSON" Then
        KelompokKategoriTlp = "RLP"
    ElseIf StsKategoriTelepon = "OTHER EC" Then
        KelompokKategoriTlp = "OEC"
    ElseIf StsKategoriTelepon = "OTHER CH MOBILE" Then
        KelompokKategoriTlp = "OCM"
    ElseIf StsKategoriTelepon = "HP" Then
        KelompokKategoriTlp = "HP"
    ElseIf StsKategoriTelepon = "Home" Then
        KelompokKategoriTlp = "HOME"
    ElseIf StsKategoriTelepon = "Office" Then
        KelompokKategoriTlp = "OFF"
    ElseIf StsKategoriTelepon = "EC" Then
        KelompokKategoriTlp = "EC"
    End If
End Sub

'@@ 16 May 2012, Khusus HSBC JAKARTA
Private Sub CariIPIcentra()
    Dim connIcentra As ADODB.Connection
    Dim StrKoneksi As String
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Initiate, Start, Finish As String
    Dim JAM, Menit, Detik As Long
    
    '@@ Cek Ke server 4 dulu ---------------------------------------------------------------------------
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo SALAH
    connIcentra.Open StrKoneksi
    
    cmdsql = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
    cmdsql = cmdsql + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and date(initiate)=date(now()) "
    cmdsql = cmdsql + " and start is not null and finish is not null  "
    cmdsql = cmdsql + " order by acd_log_outgoing_session_id desc limit 1 "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        'Pindahin data dari icentra ke database card
        Initiate = IIf(IsNull(M_Objrs("initiate")), "null", "'" & Format(M_Objrs("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
        Start = IIf(IsNull(M_Objrs("start")), "null", "'" & Format(M_Objrs("start"), "yyyy-mm-dd hh:mm:ss") + "'")
        Finish = IIf(IsNull(M_Objrs("finish")), "null", "'" & Format(M_Objrs("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
        
        'Hitung Konevrsi Selisih ke detik
        JAM = Val(Mid(M_Objrs("durasi"), 1, 2)) * 3600
        Menit = Val(Mid(M_Objrs("durasi"), 4, 2)) * 60
        Detik = Val(Mid(M_Objrs("durasi"), 7, 2)) + JAM + Menit
        
        cmdsql = "insert into outgoing_icentra (destination,"
        cmdsql = cmdsql + "initiate,start,finish,recording_filename,"
        cmdsql = cmdsql + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("destination")), "", CStr(M_Objrs("destination"))) + "',"
        cmdsql = cmdsql + Initiate + "," + Start + "," + Finish + ",'"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("recording_filename")), "", CStr(M_Objrs("recording_filename"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("campaign")), "", CStr(M_Objrs("campaign"))) + "','"
        cmdsql = cmdsql + CStr(Detik) + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(M_Objrs("acd_log_outgoing_session_id")) + "')"
        M_OBJCONN.execute cmdsql
        
        MDIForm1.TxtIPIcentra.text = "192.168.10.4"
        
        Set M_Objrs = Nothing
        Set connIcentra = Nothing
        Exit Sub
    End If
    Set M_Objrs = Nothing
    Set connIcentra = Nothing
    
    '-------------------------------------------------------------------------------------
    
    '---- Cek Server 5 -------------------------------------------------------------------
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo SALAH
    connIcentra.Open StrKoneksi
    
    cmdsql = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
    cmdsql = cmdsql + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
    cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and date(initiate)=date(now()) "
    cmdsql = cmdsql + " and start is not null and finish is not null  "
    cmdsql = cmdsql + " order by acd_log_outgoing_session_id desc limit 1 "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        'Pindahin data dari icentra ke database card
        Initiate = IIf(IsNull(M_Objrs("initiate")), "null", "'" & Format(M_Objrs("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
        Start = IIf(IsNull(M_Objrs("start")), "null", "'" & Format(M_Objrs("start"), "yyyy-mm-dd hh:mm:ss") + "'")
        Finish = IIf(IsNull(M_Objrs("finish")), "null", "'" & Format(M_Objrs("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
        
        'Hitung Konevrsi Selisih ke detik
        JAM = Val(Mid(M_Objrs("durasi"), 1, 2)) * 3600
        Menit = Val(Mid(M_Objrs("durasi"), 4, 2)) * 60
        Detik = Val(Mid(M_Objrs("durasi"), 7, 2)) + JAM + Menit
        
        cmdsql = "insert into outgoing_icentra (destination,"
        cmdsql = cmdsql + "initiate,start,finish,recording_filename,"
        cmdsql = cmdsql + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("destination")), "", CStr(M_Objrs("destination"))) + "',"
        cmdsql = cmdsql + Initiate + "," + Start + "," + Finish + ",'"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("recording_filename")), "", CStr(M_Objrs("recording_filename"))) + "','"
        cmdsql = cmdsql + IIf(IsNull(M_Objrs("campaign")), "", CStr(M_Objrs("campaign"))) + "','"
        cmdsql = cmdsql + CStr(Detik) + "','"
        cmdsql = cmdsql + MDIForm1.Text1.text + "','"
        cmdsql = cmdsql + CStr(M_Objrs("acd_log_outgoing_session_id")) + "')"
        M_OBJCONN.execute cmdsql
        
        MDIForm1.TxtIPIcentra.text = "192.168.10.5"
    End If
    Set M_Objrs = Nothing
    Set connIcentra = Nothing
    Exit Sub
SALAH:
    Exit Sub
    'MsgBox "Maaf anda tidak terhubung ke Icentra!", vbOKOnly + vbInformation, "Informasi"
End Sub

'@@ 21 May 2012, Tambahan Buat bikin beberapa baris  dari remarks
Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function

'@@ 24 May 2012, Mencari IP Centra Berdasarkan Waktu Login
Private Sub CariIPIcentraByWaktuLogin()
    Dim KoneksiIcentra As ADODB.Connection
    Dim StrKoneksiIcentra As String
    Dim M_Objrs_Icentra As ADODB.Recordset
    Dim M_Objrs_Telp As ADODB.Recordset
    Dim Initiate, Start, Finish As String
    Dim JAM, Menit, Detik As Long
    
    Set KoneksiIcentra = New ADODB.Connection
    
    'Cek di Server4 Dulu
    StrKoneksiIcentra = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo SALAH
    KoneksiIcentra.Open StrKoneksiIcentra
    cmdsql = "select * from acd_log_agent_session,acd_agent where "
    cmdsql = cmdsql + " acd_log_agent_session.acd_agent_id=acd_agent.acd_agent_id "
    cmdsql = cmdsql + " and acd_agent.name='"
    cmdsql = cmdsql + Trim(Replace(MDIForm1.Text1.text, "TL", "TLCARD")) + "' "
    cmdsql = cmdsql + " and date(login_time)=date(now()) limit 1 "
    Set M_Objrs_Icentra = New ADODB.Recordset
    M_Objrs_Icentra.CursorLocation = adUseClient
    DoEvents
    M_Objrs_Icentra.Open cmdsql, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        
    If M_Objrs_Icentra.RecordCount > 0 Then
        MDIForm1.TxtIPIcentra.text = "192.168.10.4"
        
        'Cari No Telepon yang terakhir
        cmdsql = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
        cmdsql = cmdsql + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and date(initiate)=date(now()) "
        cmdsql = cmdsql + " and start is not null and finish is not null  "
        cmdsql = cmdsql + " order by acd_log_outgoing_session_id desc limit 1 "
        Set M_Objrs_Telp = New ADODB.Recordset
        M_Objrs_Telp.CursorLocation = adUseClient
        DoEvents
        M_Objrs_Telp.Open cmdsql, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Telp.RecordCount > 0 Then
            'Pindahin data dari icentra ke database card
            Initiate = IIf(IsNull(M_Objrs_Telp("initiate")), "null", "'" & Format(M_Objrs_Telp("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
            Start = IIf(IsNull(M_Objrs_Telp("start")), "null", "'" & Format(M_Objrs_Telp("start"), "yyyy-mm-dd hh:mm:ss") + "'")
            Finish = IIf(IsNull(M_Objrs_Telp("finish")), "null", "'" & Format(M_Objrs_Telp("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
            
            'Hitung Konevrsi Selisih ke detik
            JAM = Val(Mid(M_Objrs_Telp("durasi"), 1, 2)) * 3600
            Menit = Val(Mid(M_Objrs_Telp("durasi"), 4, 2)) * 60
            Detik = Val(Mid(M_Objrs_Telp("durasi"), 7, 2)) + JAM + Menit
            
            cmdsql = "insert into outgoing_icentra (destination,"
            cmdsql = cmdsql + "initiate,start,finish,recording_filename,"
            cmdsql = cmdsql + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs_Telp("destination")), "", CStr(M_Objrs_Telp("destination"))) + "',"
            cmdsql = cmdsql + Initiate + "," + Start + "," + Finish + ",'"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs_Telp("recording_filename")), "", CStr(M_Objrs_Telp("recording_filename"))) + "','"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs_Telp("campaign")), "", CStr(M_Objrs_Telp("campaign"))) + "','"
            cmdsql = cmdsql + CStr(Detik) + "','"
            cmdsql = cmdsql + MDIForm1.Text1.text + "','"
            cmdsql = cmdsql + CStr(M_Objrs_Telp("acd_log_outgoing_session_id")) + "')"
            M_OBJCONN.execute cmdsql
            
            Set M_Objrs_Telp = Nothing
            Set M_Objrs_Icentra = Nothing
            Set KoneksiIcentra = Nothing
            Exit Sub
        End If
    End If
    Set M_Objrs_Icentra = Nothing
    Set KoneksiIcentra = Nothing
    
    '/////////////////////////////----------- Server 5 ----------------------------------------
    Set KoneksiIcentra = New ADODB.Connection
    StrKoneksiIcentra = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo SALAH
    KoneksiIcentra.Open StrKoneksiIcentra
    cmdsql = "select * from acd_log_agent_session,acd_agent where "
    cmdsql = cmdsql + " acd_log_agent_session.acd_agent_id=acd_agent.acd_agent_id "
    cmdsql = cmdsql + " and acd_agent.name='"
    cmdsql = cmdsql + Trim(Replace(MDIForm1.Text1.text, "TL", "TLCARD")) + "' "
    cmdsql = cmdsql + " and date(login_time)=date(now()) limit 1 "
    Set M_Objrs_Icentra = New ADODB.Recordset
    M_Objrs_Icentra.CursorLocation = adUseClient
    DoEvents
    M_Objrs_Icentra.Open cmdsql, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        
    If M_Objrs_Icentra.RecordCount > 0 Then
        MDIForm1.TxtIPIcentra.text = "192.168.10.5"
        
        'Cari No Telepon yang terakhir
        cmdsql = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
        cmdsql = cmdsql + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
        cmdsql = cmdsql + Trim(lblCustId.Caption) + "' and date(initiate)=date(now()) "
        cmdsql = cmdsql + " and start is not null and finish is not null  "
        cmdsql = cmdsql + " order by acd_log_outgoing_session_id desc limit 1 "
        Set M_Objrs_Telp = New ADODB.Recordset
        M_Objrs_Telp.CursorLocation = adUseClient
        DoEvents
        M_Objrs_Telp.Open cmdsql, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Telp.RecordCount > 0 Then
            'Pindahin data dari icentra ke database card
            Initiate = IIf(IsNull(M_Objrs_Telp("initiate")), "null", "'" & Format(M_Objrs_Telp("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
            Start = IIf(IsNull(M_Objrs_Telp("start")), "null", "'" & Format(M_Objrs_Telp("start"), "yyyy-mm-dd hh:mm:ss") + "'")
            Finish = IIf(IsNull(M_Objrs_Telp("finish")), "null", "'" & Format(M_Objrs_Telp("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
            
            'Hitung Konevrsi Selisih ke detik
            JAM = Val(Mid(M_Objrs_Telp("durasi"), 1, 2)) * 3600
            Menit = Val(Mid(M_Objrs_Telp("durasi"), 4, 2)) * 60
            Detik = Val(Mid(M_Objrs_Telp("durasi"), 7, 2)) + JAM + Menit
            
            cmdsql = "insert into outgoing_icentra (destination,"
            cmdsql = cmdsql + "initiate,start,finish,recording_filename,"
            cmdsql = cmdsql + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs_Telp("destination")), "", CStr(M_Objrs_Telp("destination"))) + "',"
            cmdsql = cmdsql + Initiate + "," + Start + "," + Finish + ",'"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs_Telp("recording_filename")), "", CStr(M_Objrs_Telp("recording_filename"))) + "','"
            cmdsql = cmdsql + IIf(IsNull(M_Objrs_Telp("campaign")), "", CStr(M_Objrs_Telp("campaign"))) + "','"
            cmdsql = cmdsql + CStr(Detik) + "','"
            cmdsql = cmdsql + MDIForm1.Text1.text + "','"
            cmdsql = cmdsql + CStr(M_Objrs_Telp("acd_log_outgoing_session_id")) + "')"
            M_OBJCONN.execute cmdsql
            
            Set M_Objrs_Telp = Nothing
            Set M_Objrs_Icentra = Nothing
            Set KoneksiIcentra = Nothing
            Exit Sub
        End If
    End If
    Set M_Objrs_Icentra = Nothing
    Set KoneksiIcentra = Nothing
    Exit Sub
SALAH:
    Exit Sub
End Sub

Private Sub CekAksessAllAcc()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    On Error GoTo SALAH
    If UCase(MDIForm1.Text2.text) = "ADMINISTRATOR" Or _
       UCase(MDIForm1.Text2.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.Text2.text) = "ADMIN" Then
        Exit Sub
    End If
    
    DoEvents
    
    ' # Unset account monitor_akses
'    Cmdsql = "update mgm set monitor_akses=null"
'    Cmdsql = Cmdsql + ",waktu_akses=null where custid='" & Trim(lblcustid.Caption) & "'"
'    M_OBJCONN.Execute Cmdsql
    
    cmdsql = "select * from tbl_cust_aksesall WHERE kd_profile in " & _
            "(SELECT a.kd_profile FROM tbl_profile_aksesall a, usertbl b WHERE a.kd_profile=b.profile_akses_all " & _
            " AND b.userid='"
    cmdsql = cmdsql + MDIForm1.Text1.text + "' AND a.waktu_awal < now() and "
    cmdsql = cmdsql + " a.waktu_akhir > now() )"
    
    'cek di tabel distribusi
'    Cmdsql = "select * from tbl_distribusi_account where agent='"
'    Cmdsql = Cmdsql + MDIForm1.Text1.Text + "' and waktu_awal < now() and "
'    Cmdsql = Cmdsql + " waktu_akhir > now() "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        'cek akses allnya
        If AksesAllAcc <> "1" Then
            'update di f_pesanresetauto nya
            cmdsql = "update usertbl set f_akses_all_acc='1',f_pesanresetauto='1' where "
            cmdsql = cmdsql + " userid='"
            cmdsql = cmdsql + MDIForm1.Text1.text + "'"
            M_OBJCONN.execute cmdsql
            AksesAllAcc = "1"
        End If
    Else
        'Hapus datanya dari tbl_distribusi_account
        ' CLOSE - UPDATE 22 MEI 2013 BY IZUDDIN
'        Cmdsql = "delete from tbl_distribusi_account where waktu_akhir < now() and agent='"
'        Cmdsql = Cmdsql + MDIForm1.Text1.Text + "'"
'        M_OBJCONN.Execute Cmdsql
'
'        'Update kembalikan agent semula
'        Cmdsql = "update mgm set agent=agent_asli,agent_asli=null WHERE monitor_akses is null" & _
'                " AND agent='AKSESALL'"
'        M_OBJCONN.Execute Cmdsql
'
        'update statusnya
        ' CLOSE - UPDATE 22 MEI 2013 BY IZUDDIN
'        Cmdsql = "update usertbl set f_akses_all_acc=null where "
'        Cmdsql = Cmdsql + " userid='"
'        Cmdsql = Cmdsql + MDIForm1.Text1.Text + "'"
'        M_OBJCONN.Execute Cmdsql
'        AksesAllAcc = ""
        cmdsql = "DELETE FROM tbl_cust_aksesall WHERE kd_profile in " & _
                "(SELECT a.kd_profile FROM tbl_profile_aksesall a, usertbl b WHERE a.kd_profile=b.profile_akses_all " & _
                " AND b.userid='"
        cmdsql = cmdsql + MDIForm1.Text1.text + "' AND a.waktu_awal < now() and "
        cmdsql = cmdsql + " a.waktu_akhir > now() )"
        M_OBJCONN.execute cmdsql
        AksesAllAcc = ""
    End If

    Set M_Objrs = Nothing
    Exit Sub
SALAH:
    MsgBox "Mohon maaf ada error! " & err.Description, vbOKOnly + vbExclamation, "Pesan error"
    
End Sub
