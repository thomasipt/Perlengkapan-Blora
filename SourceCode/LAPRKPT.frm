VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form LAPRKPT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BARANG INVENTARIS"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Rekapitulasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1950
      Left            =   82
      TabIndex        =   0
      Top             =   90
      Width           =   6270
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   1560
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   1560
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CETAK"
         Height          =   600
         Left            =   135
         TabIndex        =   1
         Top             =   1215
         Width           =   5955
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
         Left            =   150
         TabIndex        =   9
         Top             =   315
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
         Left            =   150
         TabIndex        =   8
         Top             =   675
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
         Left            =   3300
         TabIndex        =   7
         Top             =   675
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "S/D  Semester Akhir"
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
         Left            =   3300
         TabIndex        =   6
         Top             =   315
         Width           =   1785
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
         Left            =   5145
         TabIndex        =   5
         Top             =   330
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
         Left            =   5145
         TabIndex        =   4
         Top             =   690
         Width           =   930
      End
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   675
      Top             =   6390
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
Attribute VB_Name = "LAPRKPT"
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

Private Sub Command6_Click()
If Combo1 > Label2 And Combo2 >= Label3 Then
    MsgBox "SEMESTER AWAL SALAH", vbCritical, "WARNING"
    Exit Sub
End If

'SCari = "Select * From ZZZ_AWAL where SEMESTER = '" + Trim(Combo1) + "' and TAHUN = '" + Trim(Combo2) + "'"
'Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
'If RCari.RowCount = 0 Then
'    MsgBox "REKAP AWAL TIDAK ADA", vbCritical, "KONFIRMASI"
'    Exit Sub
'End If
'RCari.Close
'Set RCari = Nothing

Call KumpulinRekapitulasiSampeCape_Dech

TANYA = MsgBox("TAMPILKAN LAPORAN", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
Else
    Crpt.ReportFileName = "c:\Windows\RPRL\REKAPITULASI.rpt"
    Crpt.WindowState = crptMaximized
    Crpt.WindowMaxButton = True
    Crpt.WindowMinButton = True
    Crpt.Action = 1
End If

Call CariRekap

End Sub

Private Sub KumpulinRekapitulasiSampeCape_Dech()
SCari7 = "Select * from REKAPITULASI"
Set RCari7 = RDCO.OpenResultset(SCari7, rdOpenDynamic, rdConcurRowVer)

    SCari5 = "Select * from ZZZ order by KIB Asc"
    Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenDynamic, rdConcurRowVer)
    RCari5.MoveFirst
    Do Until RCari5.EOF
        
        NOVI = RCari5("NAMA_BARANG")
        
        SCari6 = "Select * from ZZZ_AWAL where NAMA_BARANG = '" + Trim(NOVI) + "' and SEMESTER='" + Trim(Combo1) + "' and TAHUN = '" + Trim(Combo2) + "'"
        Set RCari6 = RDCO.OpenResultset(SCari6, rdOpenDynamic, rdConcurRowVer)
        If RCari6.RowCount <> 0 Then
        
            RCari7.AddNew
            RCari7("KIB") = RCari5("KIB")
            RCari7("KODE_BARANG") = RCari5("KODE_BARANG")
            RCari7("NAMA_BARANG") = RCari5("NAMA_BARANG")
            
            RCari7("LUAS_AWAL") = RCari6("LUAS_AWAL")
            RCari7("JUMLAH_AWAL") = RCari6("JUMLAH_AWAL")
            RCari7("HARGA_AWAL") = RCari6("HARGA_AWAL")
            RCari7("SEMESTER_AWAL") = RCari6("SEMESTER")
            RCari7("TAHUN_AWAL") = RCari6("TAHUN")
            
            RCari7("LUAS_AKHIR") = RCari5("LUAS_AKHIR")
            RCari7("JUMLAH_AKHIR") = RCari5("JUMLAH_AKHIR")
            RCari7("HARGA_AKHIR") = RCari5("HARGA_AKHIR")
            RCari7("SEMESTER_AKHIR") = RCari5("SEMESTER")
            RCari7("TAHUN_AKHIR") = RCari5("TAHUN")
            RCari7.Update
        
        ElseIf RCari6.RowCount = 0 Then
        
            RCari7.AddNew
            RCari7("KIB") = RCari5("KIB")
            RCari7("KODE_BARANG") = RCari5("KODE_BARANG")
            RCari7("NAMA_BARANG") = RCari5("NAMA_BARANG")
            
            RCari7("LUAS_AWAL") = 0
            RCari7("JUMLAH_AWAL") = 0
            RCari7("HARGA_AWAL") = 0
            RCari7("SEMESTER_AWAL") = RCari5("SEMESTER")
            RCari7("TAHUN_AWAL") = RCari5("TAHUN")
            
            RCari7("LUAS_AKHIR") = RCari5("LUAS_AKHIR")
            RCari7("JUMLAH_AKHIR") = RCari5("JUMLAH_AKHIR")
            RCari7("HARGA_AKHIR") = RCari5("HARGA_AKHIR")
            RCari7("SEMESTER_AKHIR") = RCari5("SEMESTER")
            RCari7("TAHUN_AKHIR") = RCari5("TAHUN")
            RCari7.Update
            
        End If
            RCari6.Close
            Set RCari6 = Nothing
            
    RCari5.MoveNext
    Loop
    RCari5.Close
    Set RCari5 = Nothing

RCari7.Close
Set RCari7 = Nothing

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Call IsiCombo

Call CariRekap

Label2 = SEMESTERc
Label3 = TAHUNc

End Sub

Private Sub CariRekap()
SCari = "Delete from REKAPITULASI"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Close
Set SCari = Nothing

SCari2 = "Delete from ZZZ"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
RCari2.Close
Set SCari2 = Nothing

OYEN = "Select * from LapBRperS where KIB = 'KIB A'"
Call CariMutasi
    OYEN = "Select * from LapBRperS where KIB = 'KIB B'"
    Call CariMutasi2
OYEN = "Select * from LapBRperS where KIB = 'KIB C'"
Call CariMutasi3
    OYEN = "Select * from LapBRperS where KIB = 'KIR'"
    Call CariMutasi4
    


End Sub

Private Sub CariMutasi()
SCari = OYEN
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Do Until RCari.EOF
        SCari2 = "Select * from ZZZ"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AKHIR") = RCari("SumOfALuas")
            RCari2("JUMLAH_AKHIR") = RCari("CountOfREGISTER")
            RCari2("HARGA_AKHIR") = RCari("SumOfAHARGA")
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
        SCari2 = "Select * from ZZZ"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AKHIR") = 0
            RCari2("JUMLAH_AKHIR") = RCari("CountOfREGISTER")
            RCari2("HARGA_AKHIR") = RCari("SumOfBHARGA")
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
        SCari2 = "Select * from ZZZ"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AKHIR") = RCari("SumOfCLuas")
            RCari2("JUMLAH_AKHIR") = RCari("CountOfREGISTER")
            RCari2("HARGA_AKHIR") = RCari("SumOfCHARGAPASAR")
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
        SCari2 = "Select * from ZZZ"
        Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
        RCari2.AddNew
            RCari2("KIB") = Trim(RCari("KIB"))
            RCari2("KODE_BARANG") = Trim(RCari("KODEBRG"))
            RCari2("NAMA_BARANG") = RCari("NAMABRG")
            RCari2("LUAS_AKHIR") = 0
            RCari2("JUMLAH_AKHIR") = RCari("SumOfRJUMLAHBRG")
            RCari2("HARGA_AKHIR") = RCari("SumOfRNILAIPASAR")
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

Private Sub IsiCombo()
SCombo = "SELECT ZZZ_AWAL.TAHUN From ZZZ_AWAL GROUP BY ZZZ_AWAL.TAHUN"
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
    
End Sub
