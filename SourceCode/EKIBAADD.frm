VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form EKIBAADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EKIBAADD"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3915
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "T14"
      Top             =   2835
      Width           =   960
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Text            =   "Text12"
      Top             =   6000
      Width           =   4005
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   7350
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5556
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5556
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   5556
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "FOTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   5865
      TabIndex        =   21
      Top             =   2385
      Width           =   3585
      Begin VB.CommandButton cmdBROWSE 
         Caption         =   "Load Picture"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   16
         Top             =   3570
         Width           =   3465
      End
      Begin VB.Image Image1 
         Height          =   3435
         Left            =   60
         Stretch         =   -1  'True
         Top             =   45
         Width           =   3465
      End
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
      Left            =   4230
      TabIndex        =   18
      Top             =   6615
      Width           =   1080
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   4530
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59441153
      CurrentDate     =   39584
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Text            =   "Text11"
      Top             =   5715
      Width           =   1800
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Text            =   "Text10"
      Top             =   5430
      Width           =   4005
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Text            =   "Text9"
      Top             =   5145
      Width           =   4005
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   4860
      Width           =   2430
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   4245
      Width           =   4005
   End
   Begin VB.TextBox Text6 
      Height          =   840
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "EKIBAADD.frx":0000
      Top             =   3390
      Width           =   4005
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "T5"
      Top             =   2835
      Width           =   960
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2460
      Width           =   1695
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
      Left            =   270
      TabIndex        =   17
      Top             =   6615
      Width           =   1080
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
      Height          =   525
      Left            =   8157
      TabIndex        =   19
      Top             =   6615
      Width           =   1080
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   5895
      Top             =   945
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   -315
      ScaleHeight     =   735
      ScaleWidth      =   14865
      TabIndex        =   20
      Top             =   6480
      Width           =   14925
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   6210
      TabIndex        =   22
      Text            =   "Text13"
      Top             =   3240
      Width           =   2850
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Height          =   2535
      Left            =   -135
      TabIndex        =   34
      Top             =   -270
      Width           =   10140
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1875
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   675
         Width           =   1905
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1875
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   990
         Width           =   3270
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1875
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1905
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1875
         TabIndex        =   3
         Text            =   "Combo3"
         Top             =   1380
         Width           =   2955
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2100
         Width           =   2955
      End
      Begin VB.Label Label100 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   9045
         TabIndex        =   42
         Top             =   315
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nama"
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
         Left            =   225
         TabIndex        =   40
         Top             =   990
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Kode"
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
         Left            =   225
         TabIndex        =   39
         Top             =   675
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nomor Register"
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
         Left            =   225
         TabIndex        =   38
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0FF&
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
         Left            =   225
         TabIndex        =   37
         Top             =   1380
         Width           =   2055
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
         Left            =   1875
         TabIndex        =   36
         Top             =   1770
         Width           =   7725
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Kondisi"
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
         Left            =   225
         TabIndex        =   35
         Top             =   2100
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "JENIS LOKASI / NAMA LOKASI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5790
         TabIndex        =   41
         Top             =   1417
         Width           =   3810
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Luas                                                 meter"
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
      Left            =   195
      TabIndex        =   33
      Top             =   2445
      Width           =   4845
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tahun                                      Semester"
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
      Left            =   195
      TabIndex        =   32
      Top             =   2820
      Width           =   4845
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Letak / Alamat"
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
      Left            =   195
      TabIndex        =   31
      Top             =   3570
      Width           =   1380
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hak"
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
      Left            =   195
      TabIndex        =   30
      Top             =   4230
      Width           =   1380
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "No. Sertifikat"
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
      Left            =   195
      TabIndex        =   29
      Top             =   4845
      Width           =   1380
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Penggunaan"
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
      Left            =   195
      TabIndex        =   28
      Top             =   5130
      Width           =   1380
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Asal Usul"
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
      Left            =   195
      TabIndex        =   27
      Top             =   5415
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Harga"
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
      Left            =   195
      TabIndex        =   26
      Top             =   5700
      Width           =   1380
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
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
      Height          =   315
      Left            =   195
      TabIndex        =   25
      Top             =   5985
      Width           =   1380
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tgl Sertifikat"
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
      Left            =   195
      TabIndex        =   24
      Top             =   4530
      Width           =   1380
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   90
      Top             =   2385
      Width           =   5685
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   3075
      Left            =   90
      Top             =   3285
      Width           =   5685
   End
End
Attribute VB_Name = "EKIBAADD"
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

Private Sub cmdBROWSE_Click()
cdOpen.ShowOpen
If Not vbCancel Then
   Text13 = cdOpen.FileName
End If
Image1.Picture = LoadPicture(Text13)
Image1.Stretch = True
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
EKIBA.Show 1
End Sub

Private Sub cmdEDIT_Click()
Call Edit
End Sub

Private Sub Edit()
TANYA = MsgBox("EDIT KODE BARANG " + Text2 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SEdit = "Delete from BRGINVENTARIS where REGISTER = '" + Trim(GRIDKLIK2) + "'"
Set REdit = RDCO.OpenResultset(SEdit, rdOpenDynamic, rdConcurRowVer)

Call Simpan

ClearTextBoxes Me
Text1.SetFocus
Unload Me
EKIBA.Show 1

End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Or Text13 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Text1 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Simpan
ClearTextBoxes Me
Text1.SetFocus
Unload Me
EKIBA.Show 1
End Sub

Private Sub Simpan()
SCari = "Select * from BRGINVENTARIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

RCari.AddNew
   
    RCari("REGISTER") = Trim(Text1)
    RCari("KODEBRG") = Trim(Text2)
    RCari("NAMABRG") = Trim(Text3)
    
    RCari("KELOMPOK") = "BARANG TDK BERGERAK"
    RCari("SUBKELOMPOK") = Label18
    
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Label19
    
    RCari("RUANG") = "-"
    RCari("JENISBRG") = "TANAH"
    RCari("KIB") = "KIB A"
    RCari("TAHUN") = Trim(Text5)
    RCari("SEMESTER") = Trim(Text14)
    
    RCari("KONDISI") = Combo5
    
    RCari("ALUAS") = Text4
    
    RCari("ALETAK") = Trim(Text6)
    RCari("AHAK") = Trim(Text7)
    RCari("ANOMORSERT") = Trim(Text8)
    RCari("AGUNA") = Trim(Text9)
    RCari("AASAL") = Trim(Text10)
    RCari("AHARGA") = CCur(Text11)
    RCari("AKETERANGAN") = Trim(Text12)
    RCari("PHOTO") = Trim(Text13)
    RCari("ATGLSERT") = DTPicker1
    
    RCari("KODELOKASISEBELUM") = Combo3
    RCari("JENISLOKASISEBELUM") = Trim(Label19)
    RCari("RUANGSEBELUM") = "-"
    
    RCari("KODELOKASISESUDAH") = "-"
    RCari("JENISLOKASISESUDAH") = "-"
    RCari("RUANGSESUDAH") = "-"
    
    RCari("S_AWAL") = CCur(Text4)
    RCari("DEBET") = 0
    RCari("CREDIT") = 0
    RCari("SALDO") = CCur(Text4)
    
    RCari("RESTORE") = Date
    
RCari.Update
RCari.Close
Set RCari = Nothing

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo2_LostFocus()
If Combo2 = "" Then Exit Sub
SKelompok = "Select KELOMPOK,SUBKELOMPOK from KELOMPOKBRG where KODEKELOMPOK = '" + Combo2 + "'"
Set RKelompok = RDCO.OpenResultset(SKelompok, rdOpenDynamic, rdConcurRowVer)
If RKelompok.RowCount <> 0 Then
    Label18 = Trim(RKelompok("SUBKELOMPOK"))
Else
    Combo2.ListIndex = 0
End If
RKelompok.Close
Set RKelompok = Nothing
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
'Combo4.Clear
SKelompok = "Select * from LOKASI where KODELOKASI = '" + Combo3 + "'"
Set RKelompok = RDCO.OpenResultset(SKelompok, rdOpenDynamic, rdConcurRowVer)
If RKelompok.RowCount <> 0 Then
    Label19 = Trim(RKelompok("NAMALOKASI"))
    RKelompok.MoveFirst
'    Do While Not RKelompok.EOF
'        Combo4.AddItem Trim(RKelompok("RUANG"))
'    RKelompok.MoveNext
'    Loop
'    Combo4.ListIndex = 0
Else
    Combo3.ListIndex = 0
End If
RKelompok.Close
Set RKelompok = Nothing

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo5_LostFocus()
Combo5.AddItem "BAIK"
Combo5.AddItem "KURANG"
Combo5.AddItem "RUSAK"
Combo5.ListIndex = 0
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Text11 = 0

Label19 = ""
Label18 = ""

Text5 = TAHUNc
Text14 = SEMESTERc
DTPicker1 = Date

ZZ = 0

Me.Caption = INISIAL + " INVENTARIS BARANG KIB A - Tanah"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text2 = GRIDKLIK
        Text1 = GRIDKLIK2
        Call Cari
        Call GAMBAR
    End If
    
Call IsiCombo

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With



End Sub

Private Sub IsiCombo()
'SCombo2 = "Select KODEKELOMPOK from KELOMPOKBRG order by KODEKELOMPOK"
'Set RCombo2 = RDCO.OpenResultset(SCombo2, rdOpenDynamic, rdConcurRowVer)
'If RCombo2.RowCount <> 0 Then
'    RCombo2.MoveFirst
'    Do Until RCombo2.EOF
'        Combo2.AddItem RCombo2("KODEKELOMPOK")
'    RCombo2.MoveNext
'    Loop
'    RCombo2.Close
'    Set RCombo2 = Nothing
'    Combo2.ListIndex = 0
'Else
'    MsgBox "MASUKAN DATA KELOMPOK BARANG DAHULU", vbCritical, "KONFIRMASI"
'End If


SCombo3 = "Select KODELOKASI from V_LOKASI order by KODELOKASI"
Set RCombo3 = RDCO.OpenResultset(SCombo3, rdOpenDynamic, rdConcurRowVer)
If RCombo3.RowCount <> 0 Then
    RCombo3.MoveFirst
    Do Until RCombo3.EOF
        Combo3.AddItem RCombo3("KODELOKASI")
    RCombo3.MoveNext
    Loop
    RCombo3.Close
    Set RCombo3 = Nothing
    Combo3.ListIndex = 0
Else
    MsgBox "DATA KOSONG", vbCritical, "KONFIRMASI"
End If


Combo5.AddItem "BAIK"
Combo5.AddItem "KURANG"
Combo5.AddItem "RUSAK"
Combo5.ListIndex = 0

End Sub

Private Sub Cari()
SCari = "Select * from BRGINVENTARIS where REGISTER = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text2 = RCari("KODEBRG")
    Text1 = RCari("REGISTER")
    Text3 = RCari("NAMABRG")
    
    Combo3.AddItem RCari("KODELOKASI")
    Combo3.ListIndex = 0
    
    Label19 = RCari("JENISLOKASI")
    Text5 = RCari("TAHUN")
    Text14 = RCari("SEMESTER")
    Combo5.AddItem RCari("KONDISI")
    
    Text4 = RCari("ALUAS")
    Text5 = RCari("TAHUN")
    Text6 = RCari("ALETAK")
    Text7 = RCari("AHAK")
    Text8 = RCari("ANOMORSERT")
    Text9 = RCari("AGUNA")
    Text10 = RCari("AASAL")
    Text11 = Format(RCari("AHARGA"), "##,###.00")
    Text12 = RCari("AKETERANGAN")
    Text13 = RCari("PHOTO")
    DTPicker1 = RCari("ATGLSERT")
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
Text1 = Format(Text1, ">")
Call CekBarang
End Sub

Private Sub CekBarang()
SCari = "Select * from BRGINVENTARIS where REGISTER = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    MsgBox "DATA SUDAH TERDAFTAR", vbCritical, "KONFIRMASI"
    Label100.Visible = True
Else
    Label100.Visible = False
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
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
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text8_LostFocus()
Text8 = Format(Text8, ">")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_LostFocus()
Text9 = Format(Text9, ">")
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, ">")
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text11_gotFocus()
Text11 = ""
End Sub

Private Sub Text11_LostFocus()
If Text11 = "" Then Text11 = 0
Text11 = Format(CCur(Text11), "##,###.00")
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text12_LostFocus()
Text12 = Format(Text12, ">")
End Sub

