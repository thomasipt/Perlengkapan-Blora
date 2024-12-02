VERSION 5.00
Begin VB.Form ZRESTORE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESTORE DATABASE"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Update Data Cabang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2513
      TabIndex        =   23
      Top             =   6397
      Width           =   1485
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4770
      TabIndex        =   19
      Top             =   6397
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "INFORMASI DATABASE CABANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5370
      Left            =   255
      TabIndex        =   2
      Top             =   142
      Width           =   6000
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "Text7"
         Top             =   4860
         Width           =   1170
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "Text6"
         Top             =   4440
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   3930
         Width           =   4005
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   3510
         Width           =   4005
      End
      Begin VB.TextBox Text3 
         Height          =   1785
         Left            =   1755
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "ZRESTORE.frx":0000
         Top             =   1620
         Width           =   4005
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   1200
         Width           =   4005
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   780
         Width           =   4005
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text8"
         Top             =   360
         Width           =   4005
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Label10"
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
         Height          =   315
         Left            =   3015
         TabIndex        =   22
         Top             =   4440
         Width           =   2910
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Label9"
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
         Height          =   315
         Left            =   3015
         TabIndex        =   21
         Top             =   4860
         Width           =   2910
      End
      Begin VB.Label Label8 
         Caption         =   "Pengurus"
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
         Left            =   180
         TabIndex        =   18
         Top             =   3510
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Alamat"
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
         Left            =   180
         TabIndex        =   17
         Top             =   2355
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Nama Lokasi"
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
         Left            =   180
         TabIndex        =   16
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label Label5 
         Caption         =   "Jenis Lokasi"
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
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Kepala SKPD"
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
         Left            =   180
         TabIndex        =   14
         Top             =   3930
         Width           =   1380
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   13
         Top             =   4440
         Width           =   1380
      End
      Begin VB.Label Label2 
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
         Left            =   180
         TabIndex        =   12
         Top             =   4860
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Lokasi"
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
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cek Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   255
      TabIndex        =   0
      Top             =   5542
      Width           =   6000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore Data Cabang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   225
      TabIndex        =   1
      Top             =   6397
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Height          =   1545
      Left            =   -1440
      ScaleHeight     =   1485
      ScaleWidth      =   14865
      TabIndex        =   20
      Top             =   6165
      Width           =   14925
   End
End
Attribute VB_Name = "ZRESTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RST, RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private RSQL, SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String


Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private Lolos

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo ErrorHandler
Set CR = New ADODB.Connection
    CR.CursorLocation = adUseClient
    CR.Open "Provider=Microsoft.Jet.Oledb.4.0;Data Source=D:\DATABASE\DATABASE_CABANG.MDB;Jet OLEDB:Database Password="
    
Call Data_Cabang
Call Data_His

ErrorHandler:
Select Case Err.Number
    Case -2147467259
    MsgBox "FILE TIDAK DITEMUKAN", vbCritical, "WARNING"
End Select

End Sub

Private Sub Data_Cabang()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from KONFIG"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic
    Text8 = Format(TLogin("KODELOKASI"), ">")
    Text1 = Format(TLogin("JENISLOKASI"), ">")
    Text2 = Format(TLogin("NAMALOKASI"), ">")
    Text3 = Format(TLogin("ALAMAT"), ">")
    Text4 = Format(TLogin("PENGURUS"), ">")
    Text5 = Format(TLogin("KEPALA_SKPD"), ">")
    Text6 = Format(TLogin("SEMESTER"), ">")
    Text7 = TLogin("TAHUN")
TLogin.Close
Set TLogin = Nothing

End Sub

Private Sub Data_His()
SCari = "Select * from KONFIG_HIS where KODELOKASI = '" + Trim(Text8) + "' and SEMESTER = '" + Trim(Text6) + "' and TAHUN = '" + Trim(Text7) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    MsgBox "DATA CABANG SUDAH ADA", vbCritical, "WARNING"
    Command2.Enabled = False
    Label10 = "RESTORE TERAKHIR TANGGAL"
    Label9 = RCari("TGL")
    
    Command2.Enabled = False
    Command3.Enabled = True
Else
    Command2.Enabled = True
    Command3.Enabled = False
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub Command2_Click()
Call DATA_BRGINVENTARIS
Call BRGPAKAIHABIS
Call BRGPAKAIHABIS_HIS

'Call DATA_LOKASI

Call DATA_KONFIG_HIS

End
End Sub

Private Sub DATA_BRGINVENTARIS()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from BRGINVENTARIS order by NO_URUT Asc"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From BRGINVENTARIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("KODEBRG") = TLogin("KODEBRG")
        RCari2("REGISTER") = TLogin("REGISTER")
        RCari2("NAMABRG") = TLogin("NAMABRG")
        RCari2("KELOMPOK") = TLogin("KELOMPOK")
        RCari2("SUBKELOMPOK") = TLogin("SUBKELOMPOK")
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("RUANG") = TLogin("RUANG")
        RCari2("JENISBRG") = TLogin("JENISBRG")
        RCari2("KIB") = TLogin("KIB")
        RCari2("TAHUN") = TLogin("TAHUN")
        RCari2("SEMESTER") = TLogin("SEMESTER")
        RCari2("KONDISI") = TLogin("KONDISI")
        RCari2("KODELOKASISEBELUM") = TLogin("KODELOKASISEBELUM")
        RCari2("JENISLOKASISEBELUM") = TLogin("JENISLOKASISEBELUM")
        RCari2("RUANGSEBELUM") = TLogin("RUANGSEBELUM")
        RCari2("KODELOKASISESUDAH") = TLogin("KODELOKASISESUDAH")
        RCari2("JENISLOKASISESUDAH") = TLogin("JENISLOKASISESUDAH")
        RCari2("RUANGSESUDAH") = TLogin("RUANGSESUDAH")
        RCari2("ALUAS") = TLogin("ALUAS")
        RCari2("ALETAK") = TLogin("ALETAK")
        RCari2("AHAK") = TLogin("AHAK")
        RCari2("ATGLSERT") = TLogin("ATGLSERT")
        RCari2("ANOMORSERT") = TLogin("ANOMORSERT")
        RCari2("AGUNA") = TLogin("AGUNA")
        RCari2("AASAL") = TLogin("AASAL")
        RCari2("AHARGA") = TLogin("AHARGA")
        RCari2("AKETERANGAN") = TLogin("AKETERANGAN")
        RCari2("BMERK") = TLogin("BMERK")
        RCari2("BTAHUN") = TLogin("BTAHUN")
        RCari2("BUKURAN") = TLogin("BUKURAN")
        RCari2("BBAHAN") = TLogin("BBAHAN")
        RCari2("BPABRIK") = TLogin("BPABRIK")
        RCari2("BRANGKA") = TLogin("BRANGKA")
        RCari2("BMESIN") = TLogin("BMESIN")
        RCari2("BPOLISI") = TLogin("BPOLISI")
        RCari2("BBPKB") = TLogin("BBPKB")
        RCari2("BASAL") = TLogin("BASAL")
        RCari2("BHARGA") = TLogin("BHARGA")
        RCari2("BBAIK") = TLogin("BBAIK")
        RCari2("BKURANG") = TLogin("BKURANG")
        RCari2("BRUSAK") = TLogin("BRUSAK")
        RCari2("CKONDISIBGN") = TLogin("CKONDISIBGN")
        RCari2("CBERTINGKAT") = TLogin("CBERTINGKAT")
        RCari2("CBETON") = TLogin("CBETON")
        RCari2("CLUASLNT") = TLogin("CLUASLNT")
        RCari2("CALAMAT") = TLogin("CALAMAT")
        RCari2("CNOMORDOC") = TLogin("CNOMORDOC")
        RCari2("CTANGGALDOC") = TLogin("CTANGGALDOC")
        RCari2("CLUAS") = TLogin("CLUAS")
        RCari2("CSTATUSTANAH") = TLogin("CSTATUSTANAH")
        RCari2("CKODETANAH") = TLogin("CKODETANAH")
        RCari2("CASAL") = TLogin("CASAL")
        RCari2("CHARGAPASAR") = TLogin("CHARGAPASAR")
        RCari2("CNILAILAIN") = TLogin("CNILAILAIN")
        RCari2("CBAIK") = TLogin("CBAIK")
        RCari2("CKURANG") = TLogin("CKURANG")
        RCari2("CRUSAK") = TLogin("CRUSAK")
        RCari2("CKETERANGAN") = TLogin("CKETERANGAN")
        RCari2("RMERK") = TLogin("RMERK")
        RCari2("RNOSERI") = TLogin("RNOSERI")
        RCari2("RUKURAN") = TLogin("RUKURAN")
        RCari2("RBAHAN") = TLogin("RBAHAN")
        RCari2("RJUMLAHBRG") = TLogin("RJUMLAHBRG")
        RCari2("RHARGABELI") = TLogin("RHARGABELI")
        RCari2("RNILAIPASAR") = TLogin("RNILAIPASAR")
        RCari2("RMUTASI") = TLogin("RMUTASI")
        RCari2("RTAHUN") = TLogin("RTAHUN")
        RCari2("RBAIK") = TLogin("RBAIK")
        RCari2("RKURANG") = TLogin("RKURANG")
        RCari2("RRUSAK") = TLogin("RRUSAK")
        RCari2("PHOTO") = TLogin("PHOTO")
        RCari2("S_AWAL") = TLogin("S_AWAL")
        RCari2("DEBET") = TLogin("DEBET")
        RCari2("CREDIT") = TLogin("CREDIT")
        RCari2("SALDO") = TLogin("SALDO")
        RCari2("TYPE") = TLogin("TYPE")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.Edit
        TLogin("RESTORE") = Date
    TLogin.Update
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE INVENTARIS BARANG SELESAI", vbCritical, "WARNING"

End Sub

Private Sub BRGPAKAIHABIS()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from BRGPAKAIHABIS order by NOBUKTI Asc"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From BRGPAKAIHABIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("NOBUKTI") = TLogin("NOBUKTI")
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("RUANG") = TLogin("RUANG")
        RCari2("TANGGAL") = TLogin("TANGGAL")
        RCari2("KODEBRG") = TLogin("KODEBRG")
        RCari2("NAMABRG") = TLogin("NAMABRG")
        RCari2("QTYMASUK") = TLogin("QTYMASUK")
        RCari2("QTYKELUAR") = TLogin("QTYKELUAR")
        RCari2("JUMLAH") = TLogin("JUMLAH")
        RCari2("SATUAN") = TLogin("SATUAN")
        RCari2("HARGASAT") = TLogin("HARGASAT")
        RCari2("KETERANGAN") = TLogin("KETERANGAN")
        RCari2("SEMESTER") = TLogin("SEMESTER")
        RCari2("TAHUN") = TLogin("TAHUN")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.Edit
        TLogin("RESTORE") = Date
    TLogin.Update
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE INVENTARIS BARANG SELESAI", vbCritical, "WARNING"

End Sub

Private Sub BRGPAKAIHABIS_HIS()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from BRGPAKAIHABIS order by NOBUKTI Asc"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From BRGPAKAIHABIS_HIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("NOBUKTI") = TLogin("NOBUKTI")
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("RUANG") = TLogin("RUANG")
        RCari2("TANGGAL") = TLogin("TANGGAL")
        RCari2("KODEBRG") = TLogin("KODEBRG")
        RCari2("NAMABRG") = TLogin("NAMABRG")
        RCari2("QTYMASUK") = TLogin("QTYMASUK")
        RCari2("QTYKELUAR") = TLogin("QTYKELUAR")
        RCari2("JUMLAH") = TLogin("JUMLAH")
        RCari2("SATUAN") = TLogin("SATUAN")
        RCari2("HARGASAT") = TLogin("HARGASAT")
        RCari2("KETERANGAN") = TLogin("KETERANGAN")
        RCari2("SEMESTER") = TLogin("SEMESTER")
        RCari2("TAHUN") = TLogin("TAHUN")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.Edit
        TLogin("RESTORE") = Date
    TLogin.Update
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE INVENTARIS BARANG SELESAI", vbCritical, "WARNING"

End Sub

Private Sub DATA_LOKASI()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from LOKASI Where KODELOKASI <> '" + Trim(K_KODELOKASI) + "'"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From LOKASI"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("NAMALOKASI") = TLogin("NAMALOKASI")
        RCari2("RUANG") = TLogin("RUANG")

    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE CABANG DINAS DIKNAS SELESAI", vbCritical, "WARNING"

End Sub

Private Sub DATA_KONFIG_HIS()
SCari2 = "Select * From KONFIG_HIS"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
RCari2.AddNew
On Error Resume Next
    RCari2("KODELOKASI") = Trim(Text8)
    RCari2("JENISLOKASI") = Trim(Text1)
    RCari2("NAMALOKASI") = Trim(Text2)
    RCari2("ALAMAT") = Trim(Text3)
    RCari2("PENGURUS") = Trim(Text4)
    RCari2("KEPALA_SKPD") = Trim(Text5)
    RCari2("SEMESTER") = Trim(Text6)
    RCari2("TAHUN") = Trim(Text7)
    RCari2("TGL") = Date
    
    If Command2.Enabled = True Then
        RCari2("KETERANGAN") = "RESTORE"
    Else
        RCari2("KETERANGAN") = "UPDATE"
    End If
    
RCari2.Update
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Label10 = ""
Label9 = ""

Command2.Enabled = False
Command3.Enabled = False

End Sub

'UPDATE.....UPDATE.....UPDATE.....UPDATE.....UPDATE.....
Private Sub Command3_Click()
Call DATA_BRGINVENTARIS2
Call BRGPAKAIHABIS2
Call BRGPAKAIHABIS_HIS2

'Call DATA_LOKASI

Call DATA_KONFIG_HIS
End Sub

Private Sub DATA_BRGINVENTARIS2()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from BRGINVENTARIS where RESTORE not like '" + Trim(Label9) + "' order by NO_URUT Asc"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From BRGINVENTARIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("KODEBRG") = TLogin("KODEBRG")
        RCari2("REGISTER") = TLogin("REGISTER")
        RCari2("NAMABRG") = TLogin("NAMABRG")
        RCari2("KELOMPOK") = TLogin("KELOMPOK")
        RCari2("SUBKELOMPOK") = TLogin("SUBKELOMPOK")
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("RUANG") = TLogin("RUANG")
        RCari2("JENISBRG") = TLogin("JENISBRG")
        RCari2("KIB") = TLogin("KIB")
        RCari2("TAHUN") = TLogin("TAHUN")
        RCari2("SEMESTER") = TLogin("SEMESTER")
        RCari2("KONDISI") = TLogin("KONDISI")
        RCari2("KODELOKASISEBELUM") = TLogin("KODELOKASISEBELUM")
        RCari2("JENISLOKASISEBELUM") = TLogin("JENISLOKASISEBELUM")
        RCari2("RUANGSEBELUM") = TLogin("RUANGSEBELUM")
        RCari2("KODELOKASISESUDAH") = TLogin("KODELOKASISESUDAH")
        RCari2("JENISLOKASISESUDAH") = TLogin("JENISLOKASISESUDAH")
        RCari2("RUANGSESUDAH") = TLogin("RUANGSESUDAH")
        RCari2("ALUAS") = TLogin("ALUAS")
        RCari2("ALETAK") = TLogin("ALETAK")
        RCari2("AHAK") = TLogin("AHAK")
        RCari2("ATGLSERT") = TLogin("ATGLSERT")
        RCari2("ANOMORSERT") = TLogin("ANOMORSERT")
        RCari2("AGUNA") = TLogin("AGUNA")
        RCari2("AASAL") = TLogin("AASAL")
        RCari2("AHARGA") = TLogin("AHARGA")
        RCari2("AKETERANGAN") = TLogin("AKETERANGAN")
        RCari2("BMERK") = TLogin("BMERK")
        RCari2("BTAHUN") = TLogin("BTAHUN")
        RCari2("BUKURAN") = TLogin("BUKURAN")
        RCari2("BBAHAN") = TLogin("BBAHAN")
        RCari2("BPABRIK") = TLogin("BPABRIK")
        RCari2("BRANGKA") = TLogin("BRANGKA")
        RCari2("BMESIN") = TLogin("BMESIN")
        RCari2("BPOLISI") = TLogin("BPOLISI")
        RCari2("BBPKB") = TLogin("BBPKB")
        RCari2("BASAL") = TLogin("BASAL")
        RCari2("BHARGA") = TLogin("BHARGA")
        RCari2("BBAIK") = TLogin("BBAIK")
        RCari2("BKURANG") = TLogin("BKURANG")
        RCari2("BRUSAK") = TLogin("BRUSAK")
        RCari2("CKONDISIBGN") = TLogin("CKONDISIBGN")
        RCari2("CBERTINGKAT") = TLogin("CBERTINGKAT")
        RCari2("CBETON") = TLogin("CBETON")
        RCari2("CLUASLNT") = TLogin("CLUASLNT")
        RCari2("CALAMAT") = TLogin("CALAMAT")
        RCari2("CNOMORDOC") = TLogin("CNOMORDOC")
        RCari2("CTANGGALDOC") = TLogin("CTANGGALDOC")
        RCari2("CLUAS") = TLogin("CLUAS")
        RCari2("CSTATUSTANAH") = TLogin("CSTATUSTANAH")
        RCari2("CKODETANAH") = TLogin("CKODETANAH")
        RCari2("CASAL") = TLogin("CASAL")
        RCari2("CHARGAPASAR") = TLogin("CHARGAPASAR")
        RCari2("CNILAILAIN") = TLogin("CNILAILAIN")
        RCari2("CBAIK") = TLogin("CBAIK")
        RCari2("CKURANG") = TLogin("CKURANG")
        RCari2("CRUSAK") = TLogin("CRUSAK")
        RCari2("CKETERANGAN") = TLogin("CKETERANGAN")
        RCari2("RMERK") = TLogin("RMERK")
        RCari2("RNOSERI") = TLogin("RNOSERI")
        RCari2("RUKURAN") = TLogin("RUKURAN")
        RCari2("RBAHAN") = TLogin("RBAHAN")
        RCari2("RJUMLAHBRG") = TLogin("RJUMLAHBRG")
        RCari2("RHARGABELI") = TLogin("RHARGABELI")
        RCari2("RNILAIPASAR") = TLogin("RNILAIPASAR")
        RCari2("RMUTASI") = TLogin("RMUTASI")
        RCari2("RTAHUN") = TLogin("RTAHUN")
        RCari2("RBAIK") = TLogin("RBAIK")
        RCari2("RKURANG") = TLogin("RKURANG")
        RCari2("RRUSAK") = TLogin("RRUSAK")
        RCari2("PHOTO") = TLogin("PHOTO")
        RCari2("S_AWAL") = TLogin("S_AWAL")
        RCari2("DEBET") = TLogin("DEBET")
        RCari2("CREDIT") = TLogin("CREDIT")
        RCari2("SALDO") = TLogin("SALDO")
        RCari2("TYPE") = TLogin("TYPE")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.Edit
        TLogin("RESTORE") = Date
    TLogin.Update
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE INVENTARIS BARANG SELESAI", vbCritical, "WARNING"

End Sub

Private Sub BRGPAKAIHABIS2()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from BRGPAKAIHABIS where RESTORE not like '" + Trim(Label9) + "' order by NOBUKTI Asc"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From BRGPAKAIHABIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("NOBUKTI") = TLogin("NOBUKTI")
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("RUANG") = TLogin("RUANG")
        RCari2("TANGGAL") = TLogin("TANGGAL")
        RCari2("KODEBRG") = TLogin("KODEBRG")
        RCari2("NAMABRG") = TLogin("NAMABRG")
        RCari2("QTYMASUK") = TLogin("QTYMASUK")
        RCari2("QTYKELUAR") = TLogin("QTYKELUAR")
        RCari2("JUMLAH") = TLogin("JUMLAH")
        RCari2("SATUAN") = TLogin("SATUAN")
        RCari2("HARGASAT") = TLogin("HARGASAT")
        RCari2("KETERANGAN") = TLogin("KETERANGAN")
        RCari2("SEMESTER") = TLogin("SEMESTER")
        RCari2("TAHUN") = TLogin("TAHUN")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.Edit
        TLogin("RESTORE") = Date
    TLogin.Update
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE INVENTARIS BARANG SELESAI", vbCritical, "WARNING"

End Sub

Private Sub BRGPAKAIHABIS_HIS2()
Set TLogin = New ADODB.Recordset
NATAL = "Select * from BRGPAKAIHABIS where RESTORE not like '" + Trim(Label9) + "' order by NOBUKTI Asc"
TLogin.Open NATAL, CR, adOpenDynamic, adLockPessimistic

Do

    SCari2 = "Select * From BRGPAKAIHABIS_HIS"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    RCari2.AddNew
    On Error Resume Next
        RCari2("NOBUKTI") = TLogin("NOBUKTI")
        RCari2("KODELOKASI") = TLogin("KODELOKASI")
        RCari2("JENISLOKASI") = TLogin("JENISLOKASI")
        RCari2("RUANG") = TLogin("RUANG")
        RCari2("TANGGAL") = TLogin("TANGGAL")
        RCari2("KODEBRG") = TLogin("KODEBRG")
        RCari2("NAMABRG") = TLogin("NAMABRG")
        RCari2("QTYMASUK") = TLogin("QTYMASUK")
        RCari2("QTYKELUAR") = TLogin("QTYKELUAR")
        RCari2("JUMLAH") = TLogin("JUMLAH")
        RCari2("SATUAN") = TLogin("SATUAN")
        RCari2("HARGASAT") = TLogin("HARGASAT")
        RCari2("KETERANGAN") = TLogin("KETERANGAN")
        RCari2("SEMESTER") = TLogin("SEMESTER")
        RCari2("TAHUN") = TLogin("TAHUN")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing
    
    TLogin.Edit
        TLogin("RESTORE") = Date
    TLogin.Update
    
    TLogin.MoveNext
Loop Until TLogin.EOF

TLogin.Close
Set TLogin = Nothing
    
    MsgBox "UPDATE INVENTARIS BARANG SELESAI", vbCritical, "WARNING"

End Sub
