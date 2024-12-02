VERSION 5.00
Begin VB.Form EOD001 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROSES AKHIR SEMESTER"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdADD 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   135
      TabIndex        =   5
      Top             =   1305
      Width           =   7065
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   3
      Text            =   "4"
      Top             =   720
      Width           =   1860
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   2
      Text            =   "3"
      Top             =   315
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1845
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   315
      Width           =   1860
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1845
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "TAHUN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   7
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEMESTER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   6
      Top             =   315
      Width           =   1665
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3645
      TabIndex        =   4
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "EOD001"
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
OYEN = ""

TANYA = MsgBox("JALANKAN PROSES PERGANTIAN SEMESTER", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
Else
    OYEN = "Select * from LapBRperS where KIB = 'KIB A'"
    Call CariMutasi
        OYEN = "Select * from LapBRperS where KIB = 'KIB B'"
        Call CariMutasi2
    OYEN = "Select * from LapBRperS where KIB = 'KIB C'"
    Call CariMutasi3
        OYEN = "Select * from LapBRperS where KIB = 'KIR'"
        Call CariMutasi4
    MsgBox "PROSES PERGANTIAN TELAH SELESAI", vbCritical, "WARNING"
    
    'Call SimpanBPH
    
    
    Call STS_KONFIG
End If

Unload Me
LOGON.Show
End Sub

Private Sub Simpan_BPH()
SCari = "Select * from BRGPAKAIHABIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    SCari2 = "Select * from BRGPAKAIHABIS_AWAL"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
        RCari2("NOBUKTI") = RCari("NOBUKTI")
        RCari2("KODELOKASI") = RCari("KODELOKASI")
        RCari2("JENISLOKASI") = RCari("JENISLOKASI")
        RCari2("RUANG") = RCari("RUANG")
        RCari2("TANGGAL") = RCari("TANGGAL")
        RCari2("KODEBRG") = RCari("KODEBRG")
        RCari2("NAMABRG") = RCari("NAMABRG")
        RCari2("QTYMASUK") = RCari("QTYMASUK")
        RCari2("QTYKELUAR") = RCari("QTYKELUAR")
        RCari2("JUMLAH") = RCari("JUMLAH")
        RCari2("SATUAN") = Trim(Text5)
        RCari2("HARGASAT") = CCur(Text6)
        RCari2("KETERANGAN") = Trim(Text7)
        RCari2("SEMESTER") = SEMESTERc
        RCari2("TAHUN") = TAHUNc
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
End If
RCari.Close
Set RCari = Nothing
End Sub

'KIB A
Private Sub CariMutasi()
SCari = OYEN
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        SCari2 = "Select * from ZZZ_AWAL"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AWAL") = RCari("SumOfALuas")
            RCari2("JUMLAH_AWAL") = RCari("CountOfREGISTER")
            RCari2("HARGA_AWAL") = RCari("SumOfAHARGA")
            RCari2("SEMESTER") = SEMESTERc
            RCari2("TAHUN") = TAHUNc
        RCari2.Update
        RCari2.Close
        Set RCari2 = Nothing
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

'KIB B
Private Sub CariMutasi2()
SCari = OYEN
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        SCari2 = "Select * from ZZZ_AWAL"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AWAL") = 0
            RCari2("JUMLAH_AWAL") = RCari("CountOfREGISTER")
            RCari2("HARGA_AWAL") = RCari("SumOfBHARGA")
            RCari2("SEMESTER") = SEMESTERc
            RCari2("TAHUN") = TAHUNc
        RCari2.Update
        RCari2.Close
        Set RCari2 = Nothing
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

'KIB C
Private Sub CariMutasi3()
SCari = OYEN
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        SCari2 = "Select * from ZZZ_AWAL"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AWAL") = RCari("SumOfCLuas")
            RCari2("JUMLAH_AWAL") = RCari("CountOfREGISTER")
            RCari2("HARGA_AWAL") = RCari("SumOfCHARGAPASAR")
            RCari2("SEMESTER") = SEMESTERc
            RCari2("TAHUN") = TAHUNc
        RCari2.Update
        RCari2.Close
        Set RCari2 = Nothing
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

'KIR
Private Sub CariMutasi4()
SCari = OYEN
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        SCari2 = "Select * from ZZZ_AWAL"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AWAL") = 0
            RCari2("JUMLAH_AWAL") = RCari("SumOfRJUMLAHBRG")
            RCari2("HARGA_AWAL") = RCari("SumOfRNILAIPASAR")
            RCari2("SEMESTER") = SEMESTERc
            RCari2("TAHUN") = TAHUNc
        RCari2.Update
        RCari2.Close
        Set RCari2 = Nothing
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub STS_KONFIG()
SCari3 = "Select * from KONFIG"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurRowVer)
RCari3.Edit
    RCari3("SEMESTER") = Text3
    RCari3("TAHUN") = Text4
RCari3.Update
RCari3.Close
Set RCari3 = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Text1 = SEMESTERc
Text2 = TAHUNc

If Text1 = "1" Then
    Text3 = "2"
    Text4 = Text2
Else
    Text3 = "1"
    Text4 = CCur(Text2) + 1
End If

End Sub
