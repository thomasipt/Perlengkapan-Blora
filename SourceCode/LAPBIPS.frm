VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form LAPBIPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN MUTASI"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Frame Frame5 
      Caption         =   "Tahun"
      Height          =   1365
      Left            =   6135
      TabIndex        =   21
      Top             =   2805
      Width           =   5505
      Begin VB.CommandButton Command7 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   15
         Top             =   675
         Width           =   870
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1335
         TabIndex        =   13
         Text            =   "Combo4"
         Top             =   315
         Width           =   4020
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   14
         Text            =   "Text6"
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
         Index           =   9
         Left            =   180
         TabIndex        =   35
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
         Index           =   8
         Left            =   180
         TabIndex        =   34
         Top             =   975
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   6135
      TabIndex        =   20
      Top             =   615
      Width           =   5505
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1335
         TabIndex        =   9
         Text            =   "Combo3"
         Top             =   315
         Width           =   4020
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   675
         Width           =   1500
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   990
         Width           =   1500
      End
      Begin VB.CommandButton Command6 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4485
         TabIndex        =   12
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
         Index           =   7
         Left            =   180
         TabIndex        =   33
         Top             =   315
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
         TabIndex        =   32
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
         Index           =   5
         Left            =   180
         TabIndex        =   31
         Top             =   660
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tahun"
      Height          =   1365
      Left            =   240
      TabIndex        =   19
      Top             =   2790
      Width           =   5505
      Begin VB.CommandButton Command5 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   8
         Top             =   675
         Width           =   870
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1335
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   315
         Width           =   4020
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   7
         Text            =   "Text3"
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
         Index           =   4
         Left            =   180
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   975
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   5505
      Begin VB.CommandButton cmdCTK1 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   4
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1350
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   675
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   315
         Width           =   4020
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
         TabIndex        =   28
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
         Index           =   1
         Left            =   195
         TabIndex        =   27
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
         Index           =   0
         Left            =   195
         TabIndex        =   26
         Top             =   315
         Width           =   1380
      End
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      Caption         =   "Mutasi Barang Pakai Habis per Tahun"
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
      Left            =   6937
      TabIndex        =   17
      Top             =   2460
      Width           =   3870
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "Mutasi Barang Pakai Habis per Semester"
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
      Left            =   6937
      TabIndex        =   16
      Top             =   285
      Width           =   3870
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "Barang Inventaris per Tahun"
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
      Left            =   1410
      TabIndex        =   5
      Top             =   2460
      Width           =   3165
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "Barang Inventaris per Semester"
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
      Left            =   1410
      TabIndex        =   0
      Top             =   285
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Height          =   1920
      Left            =   165
      TabIndex        =   22
      Top             =   135
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Height          =   1920
      Left            =   6045
      TabIndex        =   25
      Top             =   2310
      Width           =   5655
   End
   Begin VB.CommandButton Command3 
      Height          =   1920
      Left            =   6045
      TabIndex        =   24
      Top             =   150
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Height          =   1920
      Left            =   165
      TabIndex        =   23
      Top             =   2310
      Width           =   5655
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   36
      Top             =   4290
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6932
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6932
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6932
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "LAPBIPS"
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
CRPT.ReportFileName = "c:\Windows\RPRL\LapBRperS.rpt"
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

Private Sub Command5_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\LapBRperT.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option2.Value = False
End Sub

Private Sub Command6_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\LapBRpkHABISperS.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option2.Value = False
End Sub

Private Sub Command7_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\LapBRpkHABISperT.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = True
CRPT.WindowMinButton = True
CRPT.Action = 1
Option2.Value = False
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""

Call Ilang
Call Mati
Call IsiCombo

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

End Sub

Private Sub IsiCombo()
SCombo = "Select NAMALOKASI from LOKASI order by KODELOKASI"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdConcurRowVer)

If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo1.AddItem RCombo("NAMALOKASI")
        Combo2.AddItem RCombo("NAMALOKASI")
        Combo3.AddItem RCombo("NAMALOKASI")
        Combo4.AddItem RCombo("NAMALOKASI")
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
End Sub

Private Sub Mati()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
End Sub

Private Sub Ilang()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Option1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Combo1.SetFocus

Text2 = SEMESTERc
Text1 = TAHUNc

End Sub

Private Sub Option2_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Combo2.SetFocus

Text3 = SEMESTERc

End Sub

Private Sub Option3_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = True
Frame5.Visible = False
Combo3.SetFocus

Text5 = SEMESTERc
Text4 = TAHUNc

End Sub

Private Sub Option4_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = True
Combo4.SetFocus

Text6 = SEMESTERc

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
