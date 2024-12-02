VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form LAPKAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN KOMPILASI ASET TAHUNAN"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Semester"
      Height          =   1365
      Left            =   203
      TabIndex        =   0
      Top             =   157
      Width           =   5505
      Begin VB.CommandButton cmdCTK1 
         Caption         =   "CETAK"
         Height          =   600
         Left            =   4500
         TabIndex        =   3
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   990
         Width           =   2895
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
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   315
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   1500
      Left            =   128
      TabIndex        =   6
      Top             =   112
      Width           =   5655
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   525
      Top             =   2415
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
End
Attribute VB_Name = "LAPKAT"
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

Call CariRekap

TANYA = MsgBox("CETAK LAPORAN", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
Else
    CRPT.ReportFileName = "c:\Windows\RPRL\ASET.rpt"
    CRPT.WindowState = crptMaximized
    CRPT.WindowMaxButton = True
    CRPT.WindowMinButton = True
    CRPT.Action = 1
End If

End Sub

Private Sub CariRekap()
SCari = "Delete from ASET"
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

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""

Call IsiCombo
Text1 = TAHUNc

End Sub

Private Sub IsiCombo()
SCombo = "Select NAMALOKASI from LOKASI order by KODELOKASI"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdConcurRowVer)

If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo1.AddItem RCombo("NAMALOKASI")
    RCombo.MoveNext
    Loop
Else
    MsgBox "DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If
RCombo.Close
Set RCombo = Nothing
Combo1.ListIndex = 0
End Sub
