VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form OUTBPHADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OUTBPHADD"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "Save"
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
      Left            =   240
      TabIndex        =   22
      Top             =   4695
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   1410
      Left            =   -315
      TabIndex        =   10
      Top             =   -135
      Width           =   11220
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2010
         TabIndex        =   12
         Text            =   "Combo4"
         Top             =   900
         Width           =   2955
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2010
         TabIndex        =   11
         Text            =   "Combo3"
         Top             =   225
         Width           =   4305
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label19"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2010
         TabIndex        =   15
         Top             =   585
         Width           =   4305
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ruang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   585
         TabIndex        =   14
         Top             =   900
         Width           =   1380
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Kode Lokasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   585
         TabIndex        =   13
         Top             =   225
         Width           =   1380
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1965
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1395
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1965
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   2520
      Width           =   4005
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1965
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2895
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1965
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   3315
      Width           =   1020
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1965
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "OUTBPHADD.frx":0000
      Top             =   3690
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1965
      TabIndex        =   4
      Text            =   "Text7"
      Top             =   4065
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1965
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2145
      Width           =   4005
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
      Left            =   2550
      TabIndex        =   1
      Top             =   4695
      Width           =   1080
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
      Height          =   480
      Left            =   4875
      TabIndex        =   0
      Top             =   4650
      Width           =   1080
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   1965
      TabIndex        =   3
      Top             =   1770
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62521345
      CurrentDate     =   39585
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   5115
      Top             =   1440
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   16
      Top             =   5475
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3572
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3572
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3572
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
   Begin VB.PictureBox Picture1 
      Height          =   960
      Left            =   -75
      ScaleHeight     =   900
      ScaleWidth      =   14865
      TabIndex        =   23
      Top             =   4455
      Width           =   14925
   End
   Begin VB.CommandButton cmdTMBH 
      Caption         =   "Tambah"
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
      Left            =   240
      TabIndex        =   24
      Top             =   4695
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1140
      Left            =   6390
      TabIndex        =   17
      Top             =   3195
      Width           =   6225
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   19
         Text            =   "Text8"
         Top             =   630
         Width           =   2280
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3420
         TabIndex        =   18
         Text            =   "Text9"
         Top             =   630
         Width           =   2280
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Jumlah Sedia"
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
         Left            =   90
         TabIndex        =   21
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Total"
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
         Left            =   3645
         TabIndex        =   20
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "No.Bukti"
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
      Left            =   210
      TabIndex        =   32
      Top             =   1410
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal"
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
      Left            =   210
      TabIndex        =   31
      Top             =   1785
      Width           =   1380
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Barang"
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
      Left            =   210
      TabIndex        =   30
      Top             =   2160
      Width           =   1380
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
      Left            =   210
      TabIndex        =   29
      Top             =   2535
      Width           =   1380
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah"
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
      Left            =   210
      TabIndex        =   28
      Top             =   2910
      Width           =   2145
   End
   Begin VB.Label Label6 
      Caption         =   "Satuan"
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
      Left            =   210
      TabIndex        =   27
      Top             =   3330
      Width           =   1380
   End
   Begin VB.Label Label7 
      Caption         =   "Harga Satuan"
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
      Left            =   210
      TabIndex        =   26
      Top             =   3705
      Width           =   1380
   End
   Begin VB.Label Label9 
      Caption         =   "Keterangan"
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
      Left            =   210
      TabIndex        =   25
      Top             =   4080
      Width           =   1380
   End
End
Attribute VB_Name = "OUTBPHADD"
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

Private Sub cmdCANCEL_Click()
Unload Me
INBPH.Show 1
End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Text2 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Simpan
ClearTextBoxes Me
Text1.SetFocus
Unload Me
INBPH.Show 1
End Sub

Private Sub Simpan()
SCari = "Select * from BRGPAKAIHABIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.AddNew
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Label19
    RCari("RUANG") = Combo4
    
    RCari("NOBUKTI") = Trim(Text1)
    RCari("TANGGAL") = DTPicker2
    RCari("KODEBRG") = Trim(Text2)
    RCari("NAMABRG") = Trim(Text3)
    RCari("QTYMASUK") = 0
    RCari("QTYKELUAR") = Text4
    RCari("JUMLAH") = 0
    RCari("SATUAN") = Trim(Text5)
    RCari("HARGASAT") = Text6
    RCari("KETERANGAN") = Trim(Text7)
   
RCari.Update
RCari.Close
Set RCari = Nothing

End Sub

Private Sub cmdTMBH_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Text2 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Simpan
ClearTextBoxes Me
Text1.SetFocus
Unload Me
OUTBPH.Show 1
End Sub

Private Sub Edit()
TANYA = MsgBox("EDIT KODE BARANG " + Text2 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SEdit = "Delete from BRGPAKAIHABIS where KODEBRG = '" + Trim(Text2) + "'"
Set REdit = RDCO.OpenResultset(SEdit, rdOpenDynamic, rdConcurRowVer)

Call Simpan

ClearTextBoxes Me
Text1.SetFocus
Unload Me
EKIR.Show 1

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
Combo4.Clear
SKelompok = "Select * from LOKASI where KODELOKASI = '" + Combo3 + "'"
Set RKelompok = RDCO.OpenResultset(SKelompok, rdOpenDynamic, rdConcurRowVer)
If RKelompok.RowCount <> 0 Then
    Label19 = Trim(RKelompok("JENISLOKASI")) + " / " + Trim(RKelompok("NAMALOKASI"))
    RKelompok.MoveFirst
    Do While Not RKelompok.EOF
        Combo4.AddItem Trim(RKelompok("RUANG"))
    RKelompok.MoveNext
    Loop
End If
RKelompok.Close
Set RKelompok = Nothing
Combo4.ListIndex = 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Text15 = 0
Text17 = 0
Combo3 = ""
Combo4 = ""
DTPicker2 = Date
Label19 = ""

Call IsiCombo

Me.Caption = INISIAL + " PENGELUARAN BARANG PAKAI HABIS"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text1 = GRIDKLIK
        Call Cari
    End If

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Frame1.Visible = False
cmdTMBH.Visible = False

End Sub

Private Sub IsiCombo()
SCombo3 = "Select KODELOKASI from V_LOKASI order by KODELOKASI"
Set RCombo3 = RDCO.OpenResultset(SCombo3, rdOpenDynamic, rdConcurRowVer)
If RCombo3.RowCount <> 0 Then
    RCombo3.MoveFirst
    Do Until RCombo3.EOF
        Combo3.AddItem RCombo3("KODELOKASI")
    RCombo3.MoveNext
    Loop
Else
    MsgBox "DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If
RCombo3.Close
Set RCombo3 = Nothing
Combo3.ListIndex = 0
End Sub

Private Sub Cari()
SCari = "Select * from BRGINVENTARIS where KODEBRG = '" + Trim(Text2) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text2 = RCari("KODEBRG")
    Text1 = RCari("REGISTER")
    Text3 = RCari("NAMABRG")
    Combo2 = RCari("KELOMPOK")
    Label18 = RCari("SUBKELOMPOK")
    Combo3 = RCari("KODELOKASI")
    Label19 = RCari("JENISLOKASI")
    Combo4 = RCari("RUANG")
    Text5 = RCari("TAHUN")
    Text14 = RCari("SEMESTER")
    Combo5 = RCari("KONDISI")
    
    Text6 = RCari("RMERK")
    Text7 = RCari("RNOSERI")
    Text8 = RCari("RUKURAN")
    Text12 = RCari("RBAHAN")
    Text9 = RCari("RJUMLAHBRG")
    Text10 = Format(RCari("RHARGABELI"), "##,###.00")
    Text11 = Format(RCari("RNILAIPASAR"), "##,###.00")
    Text4 = RCari("RMUTASI")
    
    Text13 = RCari("PHOTO")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub GAMBAR()
On Error GoTo ErrorHandler
    Image1.Picture = LoadPicture(Text13)
Exit Sub

ErrorHandler:
    
    Select Case Err.Number
    Case 76
        MsgBox "DATA GAMBAR TIDAK ADA", vbCritical, "WARNING"
    End Select
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then Exit Sub
Text1 = Digit(10, Val(Text1))
STiket = "Select NOBUKTI from BRGPAKAIHABIS where NOBUKTI = '" + Trim(Text1) + "'"
Set RTiket = RDCO.OpenResultset(STiket, rdOpenDynamic, rdConcurRowVer)
If RTiket.RowCount <> 0 Then
    MsgBox " NOMOR BUKTI TELAH DIGUNAKAN", vbOKOnly, "KONFIRMASI"
    Text1.SetFocus
End If
RTiket.Close
Set RTiket = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
    Exit Sub
Else
    Text2 = Format(Text2, ">")
End If

'SCari = "Select KODEBRG from BRGPAKAIHABIS where KODEBRG='" + Trim(Text2) + "'"
'Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
'
'If RCari.RowCount <> 0 Then
'    TANYA = MsgBox("KODE BARANG TELAH DIGUNAKAN, LAKUKAN PENAMBAHAN SALDO", vbQuestion + vbOKCancel, "KONFIRMASI")
'        If TANYA = vbCancel Then
'            Text3.SetFocus
'        Else
'            Call TBHJML
'        End If
'Else
'    Frame1.Visible = False
'    cmdTMBH.Visible = False
'    Text3 = ""
'    Text4 = ""
'    Text5 = ""
'    Text6 = ""
'    Text7 = ""
'    Label5.Caption = "Jumlah"
'    Label5.FontBold = False
    
'End If
'RCari.Close
'Set RCari = Nothing
End Sub

Private Sub TBHJML()
Label5.Caption = "JUMLAH MASUK"
Label5.FontBold = True
Frame1.Visible = True
cmdTMBH.Visible = True

    SCari2 = "Select * from BRGPAKAIHABIS where KODEBRG = '" + Trim(Text2) + "'"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari2.RowCount <> 0 Then
        Text3 = RCari2("NAMABRG")
        Text8 = RCari2("JUMLAH")
        Text5 = RCari2("SATUAN")
        Text6 = Format(RCari2("HARGASAT"), "##,###.00")
        Text7 = RCari2("KETERANGAN")
    End If
    RCari2.Close
    Set RCari2 = Nothing

Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False

Text4.SetFocus

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    If Frame1.Visible = True Then
        Text9 = CCur(Text8) + CCur(Text4)
    Else
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_gotFocus()
Text6 = ""
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Text6 = 0
Text6 = Format(CCur(Text6), "##,###.00")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub


