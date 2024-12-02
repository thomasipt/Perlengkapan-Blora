VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form LAPBI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BARANG INVENTARIS"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      Caption         =   "KIR Inventaris Lainnya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4837
      TabIndex        =   16
      Top             =   4665
      Width           =   2190
   End
   Begin VB.Frame Frame6 
      Caption         =   "Tahun"
      Height          =   1365
      Left            =   3195
      TabIndex        =   35
      Top             =   5010
      Width           =   5505
      Begin VB.CommandButton Command8 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   19
         Top             =   675
         Width           =   870
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1335
         TabIndex        =   39
         Text            =   "Combo5"
         Top             =   315
         Width           =   4020
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   18
         Text            =   "Text7"
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Lokasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   180
         TabIndex        =   37
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   36
         Top             =   975
         Width           =   1380
      End
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "KIB A - Tanah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2212
      TabIndex        =   0
      Top             =   285
      Width           =   1560
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "KIB B - Peralatan dan Mesin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1635
      TabIndex        =   4
      Top             =   2475
      Width           =   2715
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "KIB C - Gedung dan Bangunan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7410
      TabIndex        =   7
      Top             =   285
      Width           =   2925
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      Caption         =   "KIR Ruangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8145
      TabIndex        =   12
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   240
      TabIndex        =   28
      Top             =   600
      Width           =   5505
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   675
         Width           =   3300
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   990
         Width           =   3300
      End
      Begin VB.CommandButton cmdCTK1 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4545
         TabIndex        =   3
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   195
         TabIndex        =   30
         Top             =   975
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   195
         TabIndex        =   29
         Top             =   660
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tahun"
      Height          =   1365
      Left            =   240
      TabIndex        =   26
      Top             =   2790
      Width           =   5505
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   990
         Width           =   1905
      End
      Begin VB.CommandButton Command5 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4530
         TabIndex        =   6
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Lokasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   40
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   27
         Top             =   975
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   6135
      TabIndex        =   22
      Top             =   615
      Width           =   5505
      Begin VB.CommandButton Command6 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4530
         TabIndex        =   11
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   990
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1065
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   675
         Width           =   3255
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1065
         TabIndex        =   8
         Text            =   "Combo3"
         Top             =   315
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   25
         Top             =   660
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   24
         Top             =   975
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Lokasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   180
         TabIndex        =   23
         Top             =   315
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tahun"
      Height          =   1365
      Left            =   6135
      TabIndex        =   17
      Top             =   2805
      Width           =   5505
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1065
         TabIndex        =   14
         Text            =   "Text6"
         Top             =   990
         Width           =   1815
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1065
         TabIndex        =   13
         Text            =   "Combo4"
         Top             =   315
         Width           =   4380
      End
      Begin VB.CommandButton Command7 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   15
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   180
         TabIndex        =   21
         Top             =   975
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Lokasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   20
         Top             =   315
         Width           =   1380
      End
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   5722
      Top             =   1972
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Height          =   1920
      Left            =   165
      TabIndex        =   31
      Top             =   135
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Height          =   1920
      Left            =   6045
      TabIndex        =   32
      Top             =   2310
      Width           =   5655
   End
   Begin VB.CommandButton Command3 
      Height          =   1920
      Left            =   6045
      TabIndex        =   33
      Top             =   150
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Height          =   1920
      Left            =   165
      TabIndex        =   34
      Top             =   2310
      Width           =   5655
   End
   Begin VB.CommandButton Command9 
      Height          =   1920
      Left            =   3105
      TabIndex        =   38
      Top             =   4515
      Width           =   5655
   End
End
Attribute VB_Name = "LAPBI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim a, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String
Private TTL

Private Sub Label8_Click()
End Sub

Private Sub cmdCTK1_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\KIBATANAH2.rpt"
CRPT.SelectionFormula = "{BRGINVENTARIS.JENISLOKASI} = '" + Trim(Combo1) + "'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option1.Value = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Command5_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\KIBBPERALATAN2.rpt"
'crpt.SelectionFormula = "{BRGINVENTARIS.JENISLOKASI} = '" + Trim(Combo2) + "'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option2.Value = False
End Sub

Private Sub Command6_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\KIBCGEDUNG.rpt"
CRPT.SelectionFormula = "{BRGINVENTARIS.KIB} = 'KIB C'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option1.Value = False
End Sub

Private Sub Command7_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\KIR.rpt"
CRPT.SelectionFormula = "{BRGINVENTARIS.KIB} = 'KIR'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option1.Value = False
End Sub

Private Sub Command8_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\KIRLAIN.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option1.Value = False
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""
Combo5 = ""

Call Ilang
Call Mati
'Call IsiCombo

Text1 = TAHUNc
Text2 = SEMESTERc
Text3 = TAHUNc
Text4 = SEMESTERc
Text5 = TAHUNc
Text6 = TAHUNc
Text7 = TAHUNc

End Sub

Private Sub IsiCombo()
SCombo = "SELECT BRGINVENTARIS.JENISLOKASI From BRGINVENTARIS GROUP BY BRGINVENTARIS.JENISLOKASI"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdConcurRowVer)

If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo1.AddItem RCombo("JENISLOKASI")
        Combo2.AddItem RCombo("JENISLOKASI")
        Combo3.AddItem RCombo("JENISLOKASI")
        Combo4.AddItem RCombo("JENISLOKASI")
        Combo5.AddItem RCombo("JENISLOKASI")
    RCombo.MoveNext
    Loop
Else
    MsgBox "DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If
RCombo.Close
Set RCombo = Nothing
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 0
End Sub

Private Sub Mati()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
End Sub

Private Sub Ilang()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Combo1.SetFocus
End Sub

Private Sub Option2_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option3_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = True
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option4_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = True
Frame6.Visible = False
End Sub

Private Sub Option5_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub
