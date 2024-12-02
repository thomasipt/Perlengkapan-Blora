VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form LAPRKPT_BPH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BARANG PAKAI HABIS"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1575
      Width           =   4545
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MUTASI"
      Height          =   600
      Left            =   128
      TabIndex        =   8
      Top             =   2025
      Width           =   5955
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CETAK REKAPITULASI"
      Height          =   600
      Left            =   128
      TabIndex        =   7
      Top             =   585
      Width           =   5955
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1560
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4485
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1560
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   1080
      Top             =   6030
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
   Begin VB.Label Label4 
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   180
      TabIndex        =   10
      Top             =   1590
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "S/D  Tahun Akhir"
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
      Left            =   6780
      TabIndex        =   6
      Top             =   525
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Semester Awal"
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
      Left            =   165
      TabIndex        =   5
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Tahun Awal"
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
      Left            =   3315
      TabIndex        =   4
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8220
      TabIndex        =   3
      Top             =   180
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8220
      TabIndex        =   2
      Top             =   540
      Width           =   930
   End
End
Attribute VB_Name = "LAPRKPT_BPH"
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

Private Sub Command1_Click()

Call CariRekap2
TANYA = MsgBox("CETAK LAPORAN", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
Else
    crpt.ReportFileName = "c:\Windows\RPRL\LapMutasiBRpkHABIS.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If

End Sub

Private Sub Command6_Click()

Call CariRekap

TANYA = MsgBox("CETAK LAPORAN", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
Else
    crpt.ReportFileName = "c:\Windows\RPRL\LapBRpkHABISperS.rpt"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = True
    crpt.WindowMinButton = True
    crpt.Action = 1
End If



'MsgBox "SELESAI BUNG.......!!!!   MERDEKA..............!!!"

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Call IsiCombo

Label2 = SEMESTERc
Label3 = TAHUNc

End Sub

Private Sub IsiCombo()
SCombo = "SELECT BRGPAKAIHABIS_HIS.SEMESTER, BRGPAKAIHABIS_HIS.TAHUN From BRGPAKAIHABIS_HIS GROUP BY BRGPAKAIHABIS_HIS.SEMESTER, BRGPAKAIHABIS_HIS.TAHUN"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdConcurRowVer)
If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo2.AddItem RCombo("TAHUN")
    RCombo.MoveNext
    Loop
    Combo2.ListIndex = 0
Else
    Combo2.AddItem TAHUNc
    Combo2.ListIndex = 0
End If
RCombo.Close
Set RCombo = Nothing

Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.ListIndex = 0

SCombo2 = "Select NAMABRG from BRGPAKAIHABIS Order by NAMABRG Asc"
Set RCombo2 = RDCO.OpenResultset(SCombo2, rdOpenDynamic, rdConcurRowVer)
If RCombo2.RowCount <> 0 Then
    RCombo2.MoveFirst
    Do Until RCombo2.EOF
        Combo3.AddItem RCombo2("NAMABRG")
    RCombo2.MoveNext
    Loop
    Combo3.ListIndex = 0
End If
RCombo2.Close
Set RCombo2 = Nothing

End Sub

Private Sub CariRekap()
SCari = "Delete from REKAPITULASI_BPH"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Close
Set SCari = Nothing

SCari2 = "SELECT * from BRGPAKAIHABIS"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Do Until RCari2.EOF
    TOKET = RCari2("KODEBRG")
    
        SCari3 = "SELECT * from BRGPAKAIHABIS_HIS_2 where KODEBRG = '" + Trim(TOKET) + "' and SEMESTER ='" + Trim(Combo1) + "' and TAHUN ='" + Trim(Combo2) + "'"
        Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenDynamic, rdConcurRowVer)
        If RCari3.RowCount <> 0 Then
            NoMin = RCari3("SumOfQTYMASUK")
            NoMax = RCari3("SumOfQTYKELUAR")
        Else
            NoMin = 0
            NoMax = 0
        End If

    SSave = "SELECT * from REKAPITULASI_BPH"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.AddNew
        RSave("KODE_BARANG") = RCari2("KODEBRG")
        RSave("NAMA_BARANG") = RCari2("NAMABRG")
        RSave("SATUAN") = RCari2("SATUAN")
        RSave("HARGASAT") = RCari2("HARGASAT")
        RSave("KETERANGAN") = RCari2("KETERANGAN")
        
        RSave("JUMLAH_AKHIR") = RCari2("JUMLAH")
        RSave("M_MASUK") = NoMin
        RSave("M_KELUAR") = NoMax
        
        RSave("SEMESTER") = Trim(Combo1)
        RSave("TAHUN") = Trim(Combo2)
    RSave.Update
    RSave.Close
    Set RSave = Nothing

        RCari3.Close
        Set RCari3 = Nothing

    RCari2.MoveNext
    Loop
End If
RCari2.Close
Set RCari2 = Nothing

End Sub

Private Sub CariRekap2()
SCari = "Delete from LapMutasiBRpkHABIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Close
Set SCari = Nothing

SCari2 = "SELECT * from BRGPAKAIHABIS_HIS where NAMABRG = '" + Combo3 + "' Order By NO_URUT Asc"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Do Until RCari2.EOF
    
    SSave = "SELECT * from LapMutasiBRpkHABIS"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    RSave.AddNew
        RSave("NO") = RCari2("NO_URUT")
        RSave("KODE") = RCari2("KODEBRG")
        RSave("NAMA") = RCari2("NAMABRG")
        RSave("MASUK") = RCari2("QTYMASUK")
        RSave("KELUAR") = RCari2("QTYKELUAR")
        RSave("JUMLAH") = RCari2("JUMLAH")
        RSave("KETERANGAN") = RCari2("KETERANGAN")
        RSave("SEMESTER") = RCari2("SEMESTER")
        RSave("TAHUN") = RCari2("TAHUN")
        RSave("LOKASI") = RCari2("KODELOKASI")
        RSave("JENIS") = RCari2("KODELOKASI")
        RSave("RUANG") = RCari2("RUANG")
    RSave.Update
    RSave.Close
    Set RSave = Nothing

    RCari2.MoveNext
    Loop
End If
RCari2.Close
Set RCari2 = Nothing

End Sub
