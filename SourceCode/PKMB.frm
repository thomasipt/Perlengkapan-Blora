VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form PKMB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR PERUBAHAN DAN MUTASI BARANG"
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
      Left            =   9963
      TabIndex        =   6
      Top             =   4590
      Width           =   1110
   End
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
      Left            =   9978
      TabIndex        =   5
      Top             =   1530
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
      Left            =   9978
      TabIndex        =   4
      Top             =   900
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
      Left            =   9978
      TabIndex        =   0
      Top             =   270
      Width           =   1080
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
      Left            =   9963
      TabIndex        =   1
      Top             =   5475
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   86
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6180
      Width           =   9645
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6090
      Left            =   86
      TabIndex        =   3
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
   Begin Crystal.CrystalReport crpt 
      Left            =   10170
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "PKMB"
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
PKMBADD.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Then Exit Sub
INISIAL = "EDIT"
TIPE = "2"
Unload Me
PKMBADD.Show 1
End Sub

Private Sub Command1_Click()
Unload Me
PKMB.Show 1
End Sub

Private Sub Command2_Click()
CRPT.ReportFileName = "c:\windows\RPRL\MUTASIBARANG.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Call IsiGrid

cmdEDIT.Visible = False
cmdDEL.Visible = False

End Sub

Private Sub SiapkanGrid()
With GRID
     .Cols = 17
     .Rows = 3
     .Row = 0
     
     .Col = 0: .ColWidth(0) = 1000: .Text = "No. Bukti": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 2000: .Text = "Tanggal": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 3000: .Text = "Kode Lokasi": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 5000: .Text = "Jenis Lokasi": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1000: .Text = "Jenis Barang": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 3000: .Text = "Kode Barang": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 3000: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 3000: .Text = "Ruang": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 2000: .Text = "Keterangan": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 3000: .Text = "KONDISI": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 3000: .Text = "KONDISI": .CellAlignment = 4
     .Col = 11: .ColWidth(11) = 3000: .Text = "KODE LOKASI": .CellAlignment = 4
     .Col = 12: .ColWidth(12) = 3000: .Text = "KODE LOKASI": .CellAlignment = 4
     .Col = 13: .ColWidth(13) = 5000: .Text = "JENIS LOKASI": .CellAlignment = 4
     .Col = 14: .ColWidth(14) = 5000: .Text = "JENIS LOKASI": .CellAlignment = 4
     .Col = 15: .ColWidth(15) = 3000: .Text = "RUANG": .CellAlignment = 4
     .Col = 16: .ColWidth(16) = 3000: .Text = "RUANG": .CellAlignment = 4
     
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
     .Col = 0: .ColWidth(0) = 1000: .Text = "No. Bukti": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 1000: .Text = "Tanggal": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 3000: .Text = "Kode Lokasi": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 5000: .Text = "Jenis Lokasi": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1000: .Text = "Jenis Barang": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 3000: .Text = "Kode Barang": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 3000: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 3000: .Text = "Ruang": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 2000: .Text = "Keterangan": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 3000: .Text = "SEBELUM": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 3000: .Text = "SESUDAH": .CellAlignment = 4
     .Col = 11: .ColWidth(11) = 3000: .Text = "SEBELUM": .CellAlignment = 4
     .Col = 12: .ColWidth(12) = 3000: .Text = "SESUDAH": .CellAlignment = 4
     .Col = 13: .ColWidth(13) = 5000: .Text = "SEBELUM": .CellAlignment = 4
     .Col = 14: .ColWidth(14) = 5000: .Text = "SESUDAH": .CellAlignment = 4
     .Col = 15: .ColWidth(15) = 3000: .Text = "SEBELUM": .CellAlignment = 4
     .Col = 16: .ColWidth(16) = 3000: .Text = "SESUDAH": .CellAlignment = 4

End With
End Sub

Private Sub IsiGrid()
Call SiapkanGrid
SCari = "Select * From MUTASI order by NOBUKTI ASC"
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
                .Col = 1: .Text = RCari("TANGGAL"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari("JENISLOKASI"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari("KODEBARANG"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari("NAMABARANG"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari("RUANG"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari("KETERANGAN"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari("KONDISISEBELUM"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari("KONDISISESUDAH"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari("KODELOKASISEBELUM"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari("KODELOKASISESUDAH"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari("JENISLOKASISEBELUM"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari("JENISLOKASISESUDAH"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari("RUANGSEBELUM"): .CellBackColor = &HFFFFC0
                .Col = 16: .Text = RCari("RUANGSESUDAH"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("NOBUKTI")
                .Col = 1: .Text = RCari("TANGGAL")
                .Col = 2: .Text = RCari("KODELOKASI")
                .Col = 3: .Text = RCari("JENISLOKASI")
                .Col = 4: .Text = RCari("JENISBRG")
                .Col = 5: .Text = RCari("KODEBARANG")
                .Col = 6: .Text = RCari("NAMABARANG")
                .Col = 7: .Text = RCari("RUANG")
                .Col = 8: .Text = RCari("KETERANGAN")
                .Col = 9: .Text = RCari("KONDISISEBELUM")
                .Col = 10: .Text = RCari("KONDISISESUDAH")
                .Col = 11: .Text = RCari("KODELOKASISEBELUM")
                .Col = 12: .Text = RCari("KODELOKASISESUDAH")
                .Col = 13: .Text = RCari("JENISLOKASISEBELUM")
                .Col = 14: .Text = RCari("JENISLOKASISESUDAH")
                .Col = 15: .Text = RCari("RUANGSEBELUM")
                .Col = 16: .Text = RCari("RUANGSESUDAH")
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
Clipboard.SetText (GRID.Text)
GRIDKLIK = GRID.Text
Text1 = Trim(GRIDKLIK)
End Sub
