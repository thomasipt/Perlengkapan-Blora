VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form LAPPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN PENERIMAAN & PENGELUARAN BARANG HABIS PAKAI"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cetak Penerimaan - Pengeluaran per Semester"
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
      Left            =   750
      TabIndex        =   15
      Top             =   285
      Width           =   4425
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cetak Penerimaan - Pengeluaran per Tahun"
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
      Left            =   750
      TabIndex        =   14
      Top             =   2415
      Width           =   4425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   210
      TabIndex        =   6
      Top             =   600
      Width           =   5505
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   315
         Width           =   4020
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1350
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   675
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   990
         Width           =   2895
      End
      Begin VB.CommandButton cmdCTK1 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   7
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
         Index           =   0
         Left            =   195
         TabIndex        =   13
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
         Index           =   1
         Left            =   195
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   660
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   225
      TabIndex        =   0
      Top             =   2715
      Width           =   5505
      Begin VB.CommandButton Command6 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4485
         TabIndex        =   3
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   990
         Width           =   1500
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1335
         TabIndex        =   1
         Text            =   "Combo3"
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
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   960
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
         TabIndex        =   4
         Top             =   315
         Width           =   1380
      End
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   2520
      Top             =   4410
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
      Left            =   135
      TabIndex        =   16
      Top             =   135
      Width           =   5655
   End
   Begin VB.CommandButton Command3 
      Height          =   1920
      Left            =   135
      TabIndex        =   17
      Top             =   2250
      Width           =   5655
   End
End
Attribute VB_Name = "LAPPP"
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

Private Sub cmdCTK1_Click()
CRPT.ReportFileName = "c:\Windows\RPRL\CTKperS.rpt"
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
Combo3 = ""

Call Ilang
Call Mati
Call IsiCombo
End Sub

Private Sub IsiCombo()
SCombo = "Select NAMALOKASI from LOKASI order by KODELOKASI"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdConcurRowVer)

If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo1.AddItem RCombo("NAMALOKASI")
        Combo3.AddItem RCombo("NAMALOKASI")
    RCombo.MoveNext
    Loop
Else
    MsgBox "DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If
RCombo.Close
Set RCombo = Nothing
Combo1.ListIndex = 0
Combo3.ListIndex = 0
End Sub

Private Sub Mati()
Option1.Value = False
Option3.Value = False
End Sub

Private Sub Ilang()
Frame1.Visible = False
Frame4.Visible = False
End Sub

Private Sub Option1_Click()
Frame1.Visible = True
Frame4.Visible = False
Combo1.SetFocus
End Sub

Private Sub Option3_Click()
Frame1.Visible = False
Frame4.Visible = True
Combo3.SetFocus
End Sub
