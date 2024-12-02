VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form MAINMENU 
   BackColor       =   &H00000000&
   Caption         =   "MENU UTAMA"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   861
   StartUpPosition =   2  'CenterScreen
   Begin PERLENGKAPAN.net_Resize net_Resize1 
      Left            =   45
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      ResizeFont      =   -1  'True
      KeepRatio       =   -1  'True
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   7650
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   741
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7514
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7514
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7514
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   60
      TabIndex        =   4
      Top             =   1530
      Width           =   12840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   60
      TabIndex        =   3
      Top             =   1170
      Width           =   12840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   12840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   67
      TabIndex        =   1
      Top             =   90
      Width           =   12840
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   1530
      Picture         =   "MAINMENU.frx":0000
      Stretch         =   -1  'True
      Top             =   -225
      Width           =   9915
   End
   Begin VB.Menu SS 
      Caption         =   "SISTEM"
      Index           =   100
      Begin VB.Menu SSS 
         Caption         =   "Konfigurasi"
         Index           =   101
      End
      Begin VB.Menu SSS 
         Caption         =   "Password"
         Index           =   102
      End
   End
   Begin VB.Menu MS 
      Caption         =   "MASTER DATA"
      Index           =   200
      Begin VB.Menu MSS 
         Caption         =   "Jenis Barang"
         Index           =   201
      End
      Begin VB.Menu MSS 
         Caption         =   "Kode Lokasi Diknas"
         Index           =   202
      End
      Begin VB.Menu MSS 
         Caption         =   "Kelompok Barang KIR"
         Index           =   203
      End
      Begin VB.Menu MSS 
         Caption         =   "-"
         Index           =   204
         Visible         =   0   'False
      End
      Begin VB.Menu MSS 
         Caption         =   "Daftar Cabang Dinas Diknas"
         Index           =   205
         Visible         =   0   'False
      End
      Begin VB.Menu MSS 
         Caption         =   "Daftar Sekolah"
         Index           =   206
         Visible         =   0   'False
      End
      Begin VB.Menu MSS 
         Caption         =   "Daftar Ruang | Subdin"
         Index           =   207
         Visible         =   0   'False
      End
   End
   Begin VB.Menu AA 
      Caption         =   "AKTIFITAS"
      Index           =   300
      Begin VB.Menu AAA 
         Caption         =   "Entry Inventaris KIB A Tanah"
         Index           =   301
      End
      Begin VB.Menu AAA 
         Caption         =   "Entry Inventaris KIB B Peralatan dan Mesin"
         Index           =   302
      End
      Begin VB.Menu AAA 
         Caption         =   "Entry Inventaris KIB C Gedung dan Bangunan"
         Index           =   303
      End
      Begin VB.Menu AAA 
         Caption         =   "Entry Inventaris KIR Ruang"
         Index           =   304
      End
      Begin VB.Menu AAA 
         Caption         =   "-"
         Index           =   305
      End
      Begin VB.Menu AAA 
         Caption         =   "Entry Penerimaan / Pengeluaran Barang Pakai Habis"
         Index           =   306
      End
      Begin VB.Menu AAA 
         Caption         =   "Entry Pengeluaran Barang Pakai Habis"
         Index           =   307
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MT 
      Caption         =   "MUTASI"
      Index           =   500
      Begin VB.Menu MTT 
         Caption         =   "Perubahan Kondisi dan Mutasi Barang"
         Index           =   501
      End
      Begin VB.Menu MTT 
         Caption         =   "-"
         Index           =   502
         Visible         =   0   'False
      End
      Begin VB.Menu MTT 
         Caption         =   "Daftar Mutasi"
         Index           =   503
         Visible         =   0   'False
      End
   End
   Begin VB.Menu LP 
      Caption         =   "LAPORAN"
      Index           =   600
      Begin VB.Menu LPP 
         Caption         =   "Laporan Barang"
         Index           =   601
      End
      Begin VB.Menu LPP 
         Caption         =   "Laporan Pengeluaran dan Penerimaan"
         Index           =   602
         Visible         =   0   'False
      End
      Begin VB.Menu LPP 
         Caption         =   "Kompilasi Aset Tahunan"
         Index           =   603
         Visible         =   0   'False
      End
      Begin VB.Menu LPP 
         Caption         =   "-"
         Index           =   604
      End
      Begin VB.Menu LPP 
         Caption         =   "Laporan Barang Inventaris"
         Index           =   610
      End
      Begin VB.Menu LPP 
         Caption         =   "Laporan Barang Habis Pakai"
         Index           =   611
      End
      Begin VB.Menu TLL 
         Caption         =   "Backup"
         Index           =   701
         Visible         =   0   'False
      End
      Begin VB.Menu TLL 
         Caption         =   "Restore"
         Index           =   702
         Visible         =   0   'False
      End
      Begin VB.Menu TLL 
         Caption         =   "-"
         Index           =   703
         Visible         =   0   'False
      End
      Begin VB.Menu TLL 
         Caption         =   "Export Data"
         Index           =   704
         Visible         =   0   'False
      End
      Begin VB.Menu TLL 
         Caption         =   "Import Data"
         Index           =   705
         Visible         =   0   'False
      End
      Begin VB.Menu TLL 
         Caption         =   "-"
         Index           =   706
         Visible         =   0   'False
      End
      Begin VB.Menu TLL 
         Caption         =   "PROSES SEMESTER"
         Index           =   707
         Visible         =   0   'False
      End
   End
   Begin VB.Menu E 
      Caption         =   "EXIT"
      Index           =   900
   End
End
Attribute VB_Name = "MAINMENU"
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
Private Lolos

Private Sub AAA_Click(Index As Integer)
Select Case Index
    Case 301
        EKIBA.Show
    Case 302
        EKIBB.Show
    Case 303
        EKIBC.Show
    Case 304
        EKIR.Show
    Case 306
        INBPH.Show 1
    'Case 307
    '    OUTBPH.Show
End Select
End Sub

Private Sub DKK_Click(Index As Integer)
Select Case Index
    Case 401
        KIBA.Show
    Case 402
        KIBB.Show
    Case 403
        KIBC.Show
    Case 404
        KIBK.Show
    Case 406
        KIBBHP.Show
End Select
End Sub

Private Sub Command1_Click()
IPT = 0
SCari = "Select * from BRGINVENTARIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    RCari.MoveFirst
    Do Until RCari.EOF
    KODE_BRG = Format(RCari("KODEBRG"), ">")
        RCari.Edit
            RCari("KODEBRG") = KODE_BRG + "." + IPT
        RCari.Update
        IPT = IPT + 1
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing

MsgBox "SELESAI BUNG.......!!!!   MERDEKA..............!!!"
End Sub

Private Sub E_Click(Index As Integer)
TANYA = MsgBox("KELUAR DARI SISTEM PERLENGKAPAN", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Unload Me
    LOGON.Show
Else
    End
End If
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

Call CEKKONFIG

With StatusBar1.Panels
    .Item(1).Style = sbrText
    .Item(1).Text = "USERCODE : " & Operator
    .Item(1).AutoSize = sbrSpring
    .Item(2).Style = sbrText
    .Item(2).AutoSize = sbrSpring
    .Item(2).Text = "TANGGAL : " & Date
    .Item(3).Style = sbrText
    .Item(3).AutoSize = sbrSpring
    .Item(3).Text = "Copyright® EDP IPT"
End With

If KODELOKASIc = "" Then
    MsgBox "KONFIGURASI BELUM DIJALANKAN", vbCritical, "KONFIRMASI"
End If

End Sub

Private Sub CEKKONFIG()
SCari2 = "Select * from KONFIG"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    KODELOKASIc = RCari2("KODELOKASI")
    JENISLOKASIc = RCari2("JENISLOKASI")
    NAMALOKASIc = RCari2("NAMALOKASI")
    ALAMATc = RCari2("ALAMAT")
    PENGURUSc = RCari2("PENGURUS")
    KEPALA_SKPDc = RCari2("KEPALA_SKPD")
    SEMESTERc = RCari2("SEMESTER")
    TAHUNc = RCari2("TAHUN")
    
    Label1 = NAMALOKASIc + "  " + JENISLOKASIc
    Label2 = ALAMATc
    Label3 = "SEMESTER " + Trim(SEMESTERc)
    Label4 = "TAHUN " + Trim(TAHUNc)
Else
    Label1 = ""
    Label2 = ""
    Label3 = ""
    Label4 = ""
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub LPP_Click(Index As Integer)
Select Case Index
    Case 601
        LAPALL.Show 1
        'LAPBIPS.Show 1
    'Case 602
    '    LAPPP.Show 1
    Case 603
        LAPKAT.Show 1
    Case 610
        LAPRKPT.Show 1
        'LAPBI.Show 1
    Case 611
        LAPRKPT_BPH.Show 1
End Select
End Sub

Private Sub MSS_Click(Index As Integer)
Select Case Index
    Case 201
        JENISBARANG.Show
    Case 202
        KODELOKASI.Show
    Case 203
        KELOMPOK.Show
    Case 205
        Call DAFTARCABANGDINASDIKNAS
    Case 206
        Call DAFTARSEKOLAH
    Case 207
        Call DAFTARRUANG
End Select
End Sub

Private Sub DAFTARSEKOLAH()
CRPT.ReportFileName = "c:\windows\RPRL\DAFTARSEKOLAH.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub DAFTARCABANGDINASDIKNAS()
CRPT.ReportFileName = "c:\windows\RPRL\DAFTARCABANG.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub DAFTARRUANG()
CRPT.ReportFileName = "c:\windows\RPRL\DAFTARRUANG.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub MTT_Click(Index As Integer)
    PKMB.Show 1
End Sub

Private Sub SSS_Click(Index As Integer)
Select Case Index
    Case 101
        Unload Me
        KONFIGURASI.Show
    Case 102
        PASWORD.Show 1
End Select
End Sub
