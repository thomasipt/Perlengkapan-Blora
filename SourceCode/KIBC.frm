VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form KIBC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Inventaris Barang KIB - C ( Gedung & Bangunan )"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7110
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   5962
      Width           =   3585
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>"
      Height          =   285
      Left            =   10755
      TabIndex        =   2
      Top             =   5977
      Width           =   300
   End
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>"
      Height          =   285
      Left            =   3855
      TabIndex        =   1
      Top             =   5977
      Width           =   300
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   5962
      Width           =   3585
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   5460
      Left            =   75
      TabIndex        =   4
      Top             =   127
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9631
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   12640511
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      Height          =   750
      Left            =   6975
      Shape           =   4  'Rounded Rectangle
      Top             =   5647
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Ruang"
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
      Left            =   7110
      TabIndex        =   6
      Top             =   5692
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      Height          =   750
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   5647
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pencarian Kode Lokasi"
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
      Left            =   210
      TabIndex        =   5
      Top             =   5692
      Width           =   4005
   End
End
Attribute VB_Name = "KIBC"
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

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""
Combo2 = ""
Call SiapkanGrid
End Sub

Private Sub SiapkanGrid()
With GRID
     .Cols = 19
     .Row = 0
     .Col = 0: .ColWidth(0) = 1500: .Text = "Kode Barang": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 2000: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 1500: .Text = "Jenis Barang": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 1000: .Text = "Register": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1500: .Text = "Kondisi Bangn": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 1500: .Text = "Bertingkat": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 1500: .Text = "Beton": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 1000: .Text = "Luas Lnt": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 3000: .Text = "Letak/Alamat": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 1000: .Text = "Tanggal Bgn": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 1000: .Text = "No. Gedung": .CellAlignment = 4
     .Col = 11: .ColWidth(11) = 1000: .Text = "Luas (M2)": .CellAlignment = 4
     .Col = 12: .ColWidth(12) = 1000: .Text = "Stat Tanah": .CellAlignment = 4
     .Col = 13: .ColWidth(13) = 1000: .Text = "No. Tanah": .CellAlignment = 4
     .Col = 14: .ColWidth(14) = 3000: .Text = "Asal Usul": .CellAlignment = 4
     .Col = 15: .ColWidth(15) = 1000: .Text = "Harga Pasar": .CellAlignment = 4
     .Col = 16: .ColWidth(16) = 1000: .Text = "Nilai Lain": .CellAlignment = 4
     .Col = 17: .ColWidth(17) = 1500: .Text = "Kondisi": .CellAlignment = 4
     .Col = 18: .ColWidth(18) = 3000: .Text = "Keterangan": .CellAlignment = 4
End With
End Sub





