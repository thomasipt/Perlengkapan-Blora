VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form JENISBARANG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL JENIS BARANG"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpt 
      Left            =   6690
      Top             =   3060
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1785
      Left            =   6120
      TabIndex        =   8
      Top             =   4305
      Width           =   1335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         TabIndex        =   10
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Refresh Tabel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         TabIndex        =   9
         Top             =   975
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   6195
      TabIndex        =   4
      Top             =   30
      Width           =   1335
      Begin VB.CommandButton cmdDEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   7
         Top             =   1470
         Width           =   1080
      End
      Begin VB.CommandButton cmdEDIT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1080
      End
      Begin VB.CommandButton cmdADD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   1080
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6127
      Width           =   5970
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   5985
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   10557
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   6495
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4419
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4419
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4419
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TABEL JENIS BARANG KOSONG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   90
      TabIndex        =   2
      Top             =   2617
      Width           =   5970
   End
End
Attribute VB_Name = "JENISBARANG"
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

Private Sub cmdADD_Click()
INISIAL = "ENTRY"
TIPE = "1"
Unload Me
JENISBARANGADD.Show 1
End Sub

Private Sub cmdDEL_Click()
If Text1 = "" Then Exit Sub
Call DEL
End Sub

Private Sub DEL()
TANYA = MsgBox("DELETE " + Text1, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SCari = "Delete from JENISBRG where JNSBRG = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Close
Set SCari = Nothing

GRID.Clear
GRID.Refresh

MsgBox "UPDATE DATA", vbCritical, "KONFIRMASI"

Unload Me
JENISBARANG.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Then Exit Sub
INISIAL = "EDIT"
TIPE = "2"
Unload Me
JENISBARANGADD.Show 1
End Sub

Private Sub Command1_Click()
Unload Me
JENISBARANG.Show 1
End Sub

Private Sub Command2_Click()
crpt.ReportFileName = "c:\windows\RPRL\TABELJENISBARANG.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Call IsiGrid

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

End Sub

Private Sub SiapkanGrid()
With GRID
     .Cols = 2
     .Row = 0
     .Col = 0: .ColWidth(0) = 3500: .Text = "Jenis Barang": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 2000: .Text = "KIB": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
Call SiapkanGrid
SCari = "Select * From JENISBRG order by KIB"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
RCari.MoveFirst
    Brs = 1
    BB = 0
    Do Until RCari.EOF
    If BB = 0 Then
        BB = 1
    ElseIf BB = 1 Then
        BB = 0
    End If
        With GRID
            If BB = 0 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("JNSBRG"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari("KIB"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("JNSBRG")
                .Col = 1: .Text = RCari("KIB")
            End If
        End With
            RCari.MoveNext
            Brs = Brs + 1
    Loop
Else
    GRID.Visible = False
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub GRID_Click()
GRID.Col = 0
GRIDKLIK = ""
GRIDKLIK2 = ""

Clipboard.SetText (GRID.Text)
GRIDKLIK = GRID.Text
Text1 = GRIDKLIK
GRIDKLIK2 = (GRID.TextMatrix(GRID.Row, 1))
End Sub

