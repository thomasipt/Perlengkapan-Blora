VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form INBPHADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INBPHADD"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1942
      TabIndex        =   4
      Top             =   2093
      Width           =   4005
   End
   Begin VB.CommandButton cmdCANCEL 
      BackColor       =   &H00C0C0C0&
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
      Left            =   4845
      TabIndex        =   12
      Top             =   4695
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
      Left            =   2520
      TabIndex        =   11
      Top             =   4695
      Width           =   1080
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Top             =   1725
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20840449
      CurrentDate     =   39585
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1942
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   4065
      Width           =   4005
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1942
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "INBPHADD.frx":0000
      Top             =   3690
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1942
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   3315
      Width           =   1020
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1942
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   2850
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1942
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2475
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1935
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1350
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   1410
      Left            =   -338
      TabIndex        =   22
      Top             =   -135
      Width           =   11220
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2010
         TabIndex        =   0
         Top             =   225
         Width           =   4305
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2010
         TabIndex        =   1
         Top             =   900
         Width           =   2955
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   26
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   25
         Top             =   900
         Width           =   1380
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
         Left            =   570
         TabIndex        =   23
         Top             =   585
         Width           =   5745
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   24
      Top             =   5490
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
      Left            =   6660
      TabIndex        =   27
      Top             =   3195
      Width           =   5820
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
         TabIndex        =   30
         Text            =   "Text9"
         Top             =   630
         Width           =   2280
      End
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
         TabIndex        =   28
         Text            =   "Text8"
         Top             =   630
         Width           =   2280
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
         TabIndex        =   31
         Top             =   315
         Width           =   1695
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
         TabIndex        =   29
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSAVE 
      BackColor       =   &H00C0C0C0&
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
      Left            =   210
      TabIndex        =   10
      Top             =   4695
      Width           =   1080
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
      Left            =   210
      TabIndex        =   32
      Top             =   4695
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      Height          =   960
      Left            =   -105
      ScaleHeight     =   900
      ScaleWidth      =   14865
      TabIndex        =   13
      Top             =   4455
      Width           =   14925
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
      Left            =   180
      TabIndex        =   21
      Top             =   4080
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
      Left            =   180
      TabIndex        =   20
      Top             =   3705
      Width           =   1380
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
      Left            =   180
      TabIndex        =   19
      Top             =   3330
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
      Left            =   180
      TabIndex        =   18
      Top             =   2872
      Width           =   2145
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
      TabIndex        =   17
      Top             =   2490
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
      Left            =   180
      TabIndex        =   16
      Top             =   2115
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
      Left            =   180
      TabIndex        =   15
      Top             =   1740
      Width           =   1380
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
      Left            =   180
      TabIndex        =   14
      Top             =   1365
      Width           =   1380
   End
End
Attribute VB_Name = "INBPHADD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim a, Isi As String
Dim NOBUKTI_KRK As String

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
Private TTL, SALDO

Private Sub cmdCANCEL_Click()
Unload Me
INBPH.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Combo1 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Combo1 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Simpan2
'Call Simpan_His
ClearTextBoxes Me
Unload Me
INBPH.Show 1
End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Combo1 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Combo1 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

If TIPE = 1 Then
    If TOKET = 0 Then
        Call Simpan
    ElseIf TOKET = 1 Then
        Call SimpanTambah
    End If
    Call Simpan_His
ElseIf TIPE = 3 Then
    Call Simpan3
    Call Simpan_His
End If

Unload Me
INBPH.Show 1
End Sub

Private Sub SimpanTambah()
SCari = "Select * from BRGPAKAIHABIS where KODELOKASI = '" + Combo3 + "' and RUANG = '" + Combo4 + "' and KODEBRG ='" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Edit

    RCari("QTYMASUK") = CCur(RCari("QTYMASUK")) + CCur(Text4)
    RCari("JUMLAH") = CCur(RCari("JUMLAH")) + CCur(Text4)
    
    RCari("RESTORE") = Date
    
RCari.Update
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Simpan()
SCari = "Select * from BRGPAKAIHABIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.AddNew
    RCari("NOBUKTI") = Trim(Text1)
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Label19
    RCari("RUANG") = Combo4
    RCari("TANGGAL") = DTPicker2
    RCari("KODEBRG") = Trim(Combo1)
    RCari("NAMABRG") = Trim(Text3)
    RCari("QTYMASUK") = CCur(Text4)
    RCari("QTYKELUAR") = 0
    RCari("JUMLAH") = CCur(Text4)
    RCari("SATUAN") = Trim(Text5)
    RCari("HARGASAT") = CCur(Text6)
    RCari("KETERANGAN") = Trim(Text7)
    RCari("SEMESTER") = SEMESTERc
    RCari("TAHUN") = TAHUNc
    
    RCari("RESTORE") = Date
    
RCari.Update
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Simpan_His()
SCari = "Select * from BRGPAKAIHABIS_HIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.AddNew
    RCari("NOBUKTI") = Trim(Text1)
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Label19
    RCari("RUANG") = Combo4
    RCari("TANGGAL") = DTPicker2
    RCari("KODEBRG") = Trim(Combo1)
    RCari("NAMABRG") = Trim(Text3)
    If TIPE = 3 Then
        RCari("QTYMASUK") = 0
        RCari("QTYKELUAR") = CCur(Text4)
    Else
        RCari("QTYMASUK") = CCur(Text4)
        RCari("QTYKELUAR") = 0
    End If
    If TOKET = 1 Then
        RCari("JUMLAH") = CCur(Text4) + CCur(Text8)
    Else
        If TIPE = 3 Then
            RCari("JUMLAH") = CCur(Text8) - CCur(Text4)
        Else
            RCari("JUMLAH") = CCur(Text4)
        End If
    End If
    RCari("SATUAN") = Trim(Text5)
    RCari("HARGASAT") = CCur(Text6)
    If TIPE = 2 Then
        RCari("KETERANGAN") = "KOREKSI." + NOBUKTI_KRK
    Else
        RCari("KETERANGAN") = Trim(Text7)
    End If
    RCari("SEMESTER") = SEMESTERc
    RCari("TAHUN") = TAHUNc
    
    RCari("RESTORE") = Date
    
RCari.Update
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Simpan2()
SCari = "Select * from BRGPAKAIHABIS where NOBUKTI='" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.Edit
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Label19
    RCari("RUANG") = Combo4
    RCari("TANGGAL") = DTPicker2
    RCari("KODEBRG") = Trim(Combo1)
    RCari("NAMABRG") = Trim(Text3)
    RCari("QTYMASUK") = CCur(Text4)
    RCari("QTYKELUAR") = 0
    RCari("JUMLAH") = CCur(Text4)
    RCari("SATUAN") = Trim(Text5)
    RCari("HARGASAT") = CCur(Text6)
    RCari("KETERANGAN") = Trim(Text7)
    RCari("SEMESTER") = SEMESTERc
    RCari("TAHUN") = TAHUNc
    
    RCari("RESTORE") = Date
    
RCari.Update
RCari.Close
Set RCari = Nothing

SCari2 = "Select * from BRGPAKAIHABIS_HIS where KODEBRG='" + Trim(NOVI) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
Do Until RCari2.EOF
    RCari2.Edit
        RCari2("KODELOKASI") = Trim(Combo3)
        RCari2("JENISLOKASI") = Label19
        RCari2("RUANG") = Combo4
        RCari2("KODEBRG") = Trim(Combo1)
        RCari2("NAMABRG") = Trim(Text3)
        RCari2("SATUAN") = Trim(Text5)
        RCari2("HARGASAT") = CCur(Text6)
        
        RCari("RESTORE") = Date
            
    RCari2.Update
    RCari2.MoveNext
Loop
RCari2.Close
Set RCari2 = Nothing

End Sub

Private Sub Simpan3()
SCari = "Select * from BRGPAKAIHABIS where KODEBRG ='" + Trim(GRIDKLIK) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    SALDO = RCari("JUMLAH")
RCari.Edit
    RCari("QTYKELUAR") = CCur(RCari("QTYKELUAR")) + CCur(Text4)
    RCari("JUMLAH") = CCur(SALDO) - CCur(Text4)
    
    RCari("RESTORE") = Date
        
RCari.Update
RCari.Close
Set RCari = Nothing
End Sub

Private Sub cmdTMBH_Click()
If Text1 = "" Or Combo1 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Combo1 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Simpan
ClearTextBoxes Me
Text1.SetFocus
Unload Me
INBPH.Show 1
End Sub

Private Sub Edit()
TANYA = MsgBox("EDIT KODE BARANG " + Combo1 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SEdit = "Delete from BRGPAKAIHABIS where KODEBRG = '" + Trim(Combo1) + "'"
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
    Combo4.ListIndex = 0
End If
RKelompok.Close
Set RKelompok = Nothing
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo4_LostFocus()
Dim No As Double

If TIPE = 2 Then Exit Sub

SKelompok = "Select * from BRGPAKAIHABIS where KODELOKASI = '" + Combo3 + "' and RUANG = '" + Combo4 + "'"
Set RKelompok = RDCO.OpenResultset(SKelompok, rdOpenDynamic, rdConcurRowVer)
If RKelompok.RowCount <> 0 Then
    RKelompok.MoveFirst
    Do While Not RKelompok.EOF
        Combo1.AddItem Trim(RKelompok("KODEBRG"))
    RKelompok.MoveNext
    Loop
    Combo1.ListIndex = 0
Else
    Combo1 = "<Add New>"
End If
RKelompok.Close
Set RKelompok = Nothing

STiket = "Select count(*) as No from BRGPAKAIHABIS_HIS"
Set RTiket = RDCO.OpenResultset(STiket, rdOpenDynamic, rdConcurRowVer)
If RTiket.RowCount <> 0 Then
    No = Val(RTiket("No")) + 1
    Text1 = No
End If
RTiket.Close
Set RTiket = Nothing

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

Me.Caption = INISIAL + " PENERIMAAN BARANG PAKAI HABIS"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text3 = GRIDKLIK
        Call Cari
        Text1.Enabled = False
        DTPicker2.Visible = False
        Text4.Enabled = False
        Text7.Visible = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Combo1.Enabled = False
    ElseIf TIPE = 3 Then
        Me.Caption = INISIAL + " PENGELUARAN BARANG PAKAI HABIS"
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
        Text3 = GRIDKLIK
        Call Cari2
        Call AutoNumber
        Text1.Enabled = False
        Combo1.Enabled = False
        Text3.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Frame2.BackColor = &HFFFFC0
        Label17.BackColor = &HFFFFC0
        Label20.BackColor = &HFFFFC0
    End If

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Frame1.Visible = False
cmdTMBH.Visible = False

DTPicker2 = Date
TOKET = 0

End Sub

Private Sub AutoNumber()
STiket = "Select count(*) as No from BRGPAKAIHABIS_HIS"
Set RTiket = RDCO.OpenResultset(STiket, rdOpenDynamic, rdConcurRowVer)
If RTiket.RowCount <> 0 Then
    No = Val(RTiket("No")) + 1
    Text1 = No
End If
RTiket.Close
Set RTiket = Nothing

Text7 = "KELUAR." + Trim(Text1) + "." + Trim(DTPicker2)
Text7.Enabled = False

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

Combo1 = ""

End Sub

Private Sub Cari()
SCari = "Select * from BRGPAKAIHABIS where KODEBRG = '" + Trim(Text3) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text1 = RCari("NOBUKTI")
    NOBUKTI_KRK = RCari("NOBUKTI")
    Combo3 = Format(RCari("KODELOKASI"), ">")
    Label19 = Format(RCari("JENISLOKASI"), ">")
    Combo4 = Format(RCari("RUANG"), ">")
    DTPicker2 = RCari("TANGGAL")
    Combo1 = Format(RCari("KODEBRG"), ">")
    Text3 = Format(RCari("NAMABRG"), ">")
    Text4 = RCari("QTYMASUK")
    Text5 = Format(RCari("SATUAN"), ">")
    Text6 = Format(RCari("HARGASAT"), "##,###.00")
    Text7 = Format(RCari("KETERANGAN"), ">")
    
    NOVI = RCari("KODEBRG")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Cari2()
SCari = "Select * from BRGPAKAIHABIS where KODEBRG = '" + Trim(Text3) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Combo3 = Format(RCari("KODELOKASI"), ">")
    Label19 = Format(RCari("JENISLOKASI"), ">")
    Combo4 = Format(RCari("RUANG"), ">")
    DTPicker2 = RCari("TANGGAL")
    Combo1 = Format(RCari("KODEBRG"), ">")
    Text3 = Format(RCari("NAMABRG"), ">")
    Text5 = Format(RCari("SATUAN"), ">")
    Text6 = Format(RCari("HARGASAT"), "##,###.00")
    Text7 = Format(RCari("KETERANGAN"), ">")
    Text8 = RCari("JUMLAH")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then Exit Sub
STiket = "Select NOBUKTI from BRGPAKAIHABIS_HIS where NOBUKTI = '" + Trim(Text1) + "'"
Set RTiket = RDCO.OpenResultset(STiket, rdOpenDynamic, rdConcurRowVer)
If RTiket.RowCount <> 0 Then
    MsgBox " NOMOR BUKTI TELAH DIGUNAKAN", vbOKOnly, "KONFIRMASI"
    Text1.SetFocus
    Text1 = ""
End If
RTiket.Close
Set RTiket = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
TOKET = 0

If Combo1.Text = "" Then
    Exit Sub
Else
    Combo1 = Format(Combo1, ">")
End If

SCari = "Select KODEBRG from BRGPAKAIHABIS where KODEBRG ='" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

If RCari.RowCount <> 0 Then
    TANYA = MsgBox("KODE BARANG TELAH DIGUNAKAN, LAKUKAN PENAMBAHAN SALDO", vbQuestion + vbOKCancel, "KONFIRMASI")
        If TANYA = vbCancel Then
            Combo1.SetFocus
            toket1 = 0
        Else
            Call TBHJML
        End If
Else
    Frame1.Visible = False
    cmdTMBH.Visible = False
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Label5.Caption = "Jumlah"
    Label5.FontBold = False
    
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub TBHJML()
Label5.Caption = "JUMLAH MASUK"
Label5.FontBold = True
Frame1.Visible = True
cmdTMBH.Visible = True

    SCari2 = "Select * from BRGPAKAIHABIS where KODEBRG = '" + Trim(Combo1) + "'"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
    If RCari2.RowCount <> 0 Then
        Text3 = RCari2("NAMABRG")
        Text8 = RCari2("JUMLAH")
        Text5 = RCari2("SATUAN")
        Text6 = Format(RCari2("HARGASAT"), "##,###.00")
    End If
    RCari2.Close
    Set RCari2 = Nothing

Text5.Enabled = False
Text6.Enabled = False

Text7 = "MASUK." + Trim(Text1) + "." + Trim(DTPicker2)
Text7.Enabled = False

Text4.SetFocus

TOKET = 1

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
        If Text4 = "" Then
            Text4 = 0
        End If
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


