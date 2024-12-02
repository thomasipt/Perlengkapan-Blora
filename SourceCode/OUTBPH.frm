VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form OUTBPH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OUTBPH"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1785
      Left            =   9851
      TabIndex        =   8
      Top             =   4350
      Width           =   1335
      Begin VB.CommandButton Command2 
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
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   86
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6180
      Width           =   9645
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   9851
      TabIndex        =   0
      Top             =   105
      Width           =   1335
      Begin VB.CommandButton cmdDEL 
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
         TabIndex        =   3
         Top             =   1470
         Width           =   1080
      End
      Begin VB.CommandButton cmdEDIT 
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
         TabIndex        =   2
         Top             =   840
         Width           =   1080
      End
      Begin VB.CommandButton cmdADD 
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
         TabIndex        =   1
         Top             =   210
         Width           =   1080
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6090
      Left            =   90
      TabIndex        =   5
      Top             =   45
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   10742
      _Version        =   393216
      Rows            =   4
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   14737632
      BackColorBkg    =   12640511
      MergeCells      =   2
      AllowUserResizing=   3
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   6495
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
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
   Begin Crystal.CrystalReport crpt 
      Left            =   10350
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR KOSONG"
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
      Left            =   75
      TabIndex        =   7
      Top             =   2640
      Width           =   9645
   End
End
Attribute VB_Name = "OUTBPH"
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
OUTBPHADD.Show 1
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

SCari = "Delete from BRGPAKAIHABIS where KODEBRG = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)


Unload Me
OUTBPH.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Then Exit Sub
INISIAL = "EDIT"
TIPE = "2"
Unload Me
OUTBPHADD.Show 1
End Sub

Private Sub Command1_Click()
Unload Me
OUTBPH.Show 1
End Sub

Private Sub Command2_Click()
crpt.ReportFileName = "c:\windows\RPRL\TABELLOKASI.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub GRID_Click()
GRID.Col = 5
GRIDKLIK = ""
Clipboard.SetText (GRID.Text)
GRIDKLIK = GRID.Text
Text1 = GRIDKLIK

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Call IsiGrid
End Sub

Private Sub SiapkanGrid()
With GRID
     .Cols = 11
     .Rows = 3
     .Row = 0
     .Col = 0: .ColWidth(0) = 1500: .Text = "No Bukti": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 2000: .Text = "Kode Lokasi": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 2000: .Text = "Jenis Lokasi": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 1500: .Text = "Ruang": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1500: .Text = "Tanggal": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 2000: .Text = "Kode Barang": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 2000: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 1500: .Text = "Satuan": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 2500: .Text = "Keterangan": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 1500: .Text = "QTY": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 1500: .Text = "QTY": .CellAlignment = 4
     
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeCol(10) = True
    .MergeRow(0) = True
    .MergeRow(1) = True
    
     .Row = 1
     .Col = 0: .ColWidth(0) = 1500: .Text = "No Bukti": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 2000: .Text = "Kode Lokasi": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 2000: .Text = "Jenis Lokasi": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 1500: .Text = "Ruang": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1500: .Text = "Tanggal": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 2000: .Text = "Kode Barang": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 2000: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 1500: .Text = "Satuan": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 2500: .Text = "Keterangan": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 1500: .Text = "Masuk": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 1500: .Text = "Keluar": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
Call SiapkanGrid
SCari = "Select * From BRGPAKAIHABIS where JUMLAH <> 0 order by NOBUKTI DESC"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

If RCari.RowCount <> 0 Then
RCari.MoveFirst
    Brs = 2
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
                .Col = 0: .Text = RCari("NOBUKTI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari("JENISLOKASI"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari("RUANG"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari("TANGGAL"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari("KODEBRG"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari("SATUAN"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari("KETERANGAN"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari("QTYMASUK"): .CellBackColor = &HFFFFC0: .CellAlignment = 4
                .Col = 10: .Text = RCari("QTYKELUAR"): .CellBackColor = &HFFFFC0: .CellAlignment = 4
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("NOBUKTI")
                .Col = 1: .Text = RCari("KODELOKASI")
                .Col = 2: .Text = RCari("JENISLOKASI")
                .Col = 3: .Text = RCari("RUANG")
                .Col = 4: .Text = RCari("TANGGAL")
                .Col = 5: .Text = RCari("KODEBRG")
                .Col = 6: .Text = RCari("NAMABRG")
                .Col = 7: .Text = RCari("SATUAN")
                .Col = 8: .Text = RCari("KETERANGAN")
                .Col = 9: .Text = RCari("QTYMASUK"): .CellAlignment = 4
                .Col = 10: .Text = RCari("QTYKELUAR"): .CellAlignment = 4
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

