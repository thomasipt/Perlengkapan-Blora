VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form EKIBCADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EKIBCADD"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Text            =   "Text18"
      Top             =   5985
      Width           =   4230
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   4005
      TabIndex        =   18
      Text            =   "Text17"
      Top             =   5640
      Width           =   1665
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
      Left            =   5835
      TabIndex        =   24
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
         TabIndex        =   20
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
      Left            =   4223
      TabIndex        =   22
      Top             =   6615
      Width           =   1080
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
      Left            =   293
      TabIndex        =   21
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
      Left            =   8168
      TabIndex        =   23
      Top             =   6615
      Width           =   1080
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3900
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "T14"
      Top             =   2475
      Width           =   960
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "T5"
      Top             =   2475
      Width           =   960
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4005
      TabIndex        =   10
      Text            =   "Text12"
      Top             =   3330
      Width           =   1665
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Text            =   "Text11"
      Top             =   4815
      Width           =   1665
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Text            =   "Text10"
      Top             =   4530
      Width           =   1665
   End
   Begin VB.TextBox Text9 
      Height          =   780
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "EKIBCADD.frx":0000
      Top             =   3660
      Width           =   4230
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   3330
      Width           =   1665
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4005
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   3030
      Width           =   1665
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   3030
      Width           =   1665
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4005
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   4815
      Width           =   1665
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Text            =   "Text15"
      Top             =   5640
      Width           =   1665
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Text            =   "Text16"
      Top             =   5220
      Width           =   4230
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   7005
      Top             =   405
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   25
      Top             =   7395
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
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   -2700
      ScaleHeight     =   735
      ScaleWidth      =   14865
      TabIndex        =   26
      Top             =   6480
      Width           =   14925
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      Height          =   2580
      Left            =   -150
      TabIndex        =   27
      Top             =   -270
      Width           =   10140
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1965
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   675
         Width           =   1905
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1965
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   990
         Width           =   3270
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1965
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1905
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1965
         TabIndex        =   3
         Text            =   "Combo3"
         Top             =   1380
         Width           =   2955
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1965
         TabIndex        =   4
         Text            =   "Combo5"
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
         Left            =   9090
         TabIndex        =   50
         Top             =   315
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Nama Barang"
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
         Left            =   270
         TabIndex        =   33
         Top             =   990
         Width           =   1830
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Kode Barang"
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
         Left            =   270
         TabIndex        =   32
         Top             =   675
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
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
         Left            =   270
         TabIndex        =   31
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label Label17 
         BackColor       =   &H0000FF00&
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
         Left            =   270
         TabIndex        =   30
         Top             =   1380
         Width           =   1830
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
         Left            =   1965
         TabIndex        =   29
         Top             =   1770
         Width           =   7590
      End
      Begin VB.Label Label24 
         BackColor       =   &H0000FF00&
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
         Left            =   270
         TabIndex        =   28
         Top             =   2100
         Width           =   1830
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
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
         TabIndex        =   34
         Top             =   1410
         Width           =   3810
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4005
      TabIndex        =   13
      Top             =   4515
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59441153
      CurrentDate     =   39584
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   6690
      TabIndex        =   35
      Text            =   "Text13"
      Top             =   3105
      Width           =   1845
   End
   Begin VB.Label Label25 
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
      Left            =   180
      TabIndex        =   49
      Top             =   5970
      Width           =   1380
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lainnya"
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
      Left            =   3285
      TabIndex        =   48
      Top             =   5625
      Width           =   1380
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Harga Pasar"
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
      TabIndex        =   47
      Top             =   5625
      Width           =   1380
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
      Left            =   180
      TabIndex        =   46
      Top             =   2460
      Width           =   4845
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bertingkat"
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
      TabIndex        =   45
      Top             =   3015
      Width           =   930
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Beton"
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
      Left            =   3330
      TabIndex        =   44
      Top             =   3015
      Width           =   930
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Luas Lantai"
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
      TabIndex        =   43
      Top             =   3315
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   42
      Top             =   3645
      Width           =   1380
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nomor Doc"
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
      TabIndex        =   41
      Top             =   4515
      Width           =   1065
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
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
      Height          =   315
      Left            =   3285
      TabIndex        =   40
      Top             =   4515
      Width           =   1065
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Luas"
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
      Left            =   3330
      TabIndex        =   39
      Top             =   3315
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Status Tanah"
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
      TabIndex        =   38
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kode"
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
      Left            =   3285
      TabIndex        =   37
      Top             =   4860
      Width           =   1290
   End
   Begin VB.Label Label16 
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
      Left            =   180
      TabIndex        =   36
      Top             =   5205
      Width           =   1380
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   3435
      Left            =   75
      Top             =   2925
      Width           =   5685
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Left            =   75
      Top             =   2385
      Width           =   5685
   End
End
Attribute VB_Name = "EKIBCADD"
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
EKIBC.Show 1
End Sub

Private Sub cmdEDIT_Click()
Call Edit
End Sub

Private Sub Edit()
TANYA = MsgBox("EDIT KODE BARANG " + Text2 + " & NAMA BARANG " + Text3, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SEdit = "Delete from BRGINVENTARIS where KODEBRG = '" + Trim(GRIDKLIK) + "'"
Set REdit = RDCO.OpenResultset(SEdit, rdOpenDynamic, rdConcurRowVer)

Call Simpan

ClearTextBoxes Me
Text1.SetFocus
Unload Me
EKIBC.Show 1

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
EKIBC.Show 1
End Sub

Private Sub Simpan()
SCari = "Select * from BRGINVENTARIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

RCari.AddNew
    RCari("KODEBRG") = Trim(Text2)
    RCari("REGISTER") = Trim(Text1)
    RCari("NAMABRG") = Trim(Text3)
    RCari("KELOMPOK") = "BARANG TDK BERGERAK"
    RCari("SUBKELOMPOK") = Label18
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Label19
    RCari("RUANG") = "-"
    RCari("JENISBRG") = "GEDUNG DAN BANGUNAN"
    RCari("KIB") = "KIB C"
    RCari("TAHUN") = Trim(Text5)
    RCari("SEMESTER") = Trim(Text14)
    RCari("KONDISI") = Combo5
    
    RCari("CKONDISIBGN") = Trim(Combo5)
    RCari("CBERTINGKAT") = Trim(Text6)
    RCari("CBETON") = Trim(Text7)
    RCari("CLUASLNT") = CCur(Text8)
    RCari("CALAMAT") = Trim(Text9)
    RCari("CNOMORDOC") = Trim(Text10)
    RCari("CTANGGALDOC") = DTPicker1
    RCari("CLUAS") = CCur(Text12)
    RCari("CSTATUSTANAH") = Trim(Text11)
    RCari("CKODETANAH") = Trim(Text4)
    RCari("CASAL") = Trim(Text16)
    RCari("CHARGAPASAR") = CCur(Text15)
    RCari("CNILAILAIN") = Trim(Text17)
    RCari("CKETERANGAN") = Trim(Text18)
    
    RCari("PHOTO") = Trim(Text13)
    
    RCari("KODELOKASISEBELUM") = Combo3
    RCari("JENISLOKASISEBELUM") = Trim(Label19)
    RCari("RUANGSEBELUM") = "-"
    
    RCari("KODELOKASISESUDAH") = "-"
    RCari("JENISLOKASISESUDAH") = "-"
    RCari("RUANGSESUDAH") = "-"
    
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
SKelompok = "Select KODEKELOMPOK,SUBKELOMPOK from KELOMPOKBRG where KODEKELOMPOK = '" + Trim(Combo2) + "'"
Set RKelompok = RDCO.OpenResultset(SKelompok, rdOpenDynamic, rdConcurRowVer)
If RKelompok.RowCount <> 0 Then
    Label18 = Trim(RKelompok("SUBKELOMPOK"))
End If
RKelompok.Close
Set RKelompok = Nothing
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
SKelompok = "Select * from LOKASI where KODELOKASI = '" + Combo3 + "'"
Set RKelompok = RDCO.OpenResultset(SKelompok, rdOpenDynamic, rdConcurRowVer)
If RKelompok.RowCount <> 0 Then
    Label19 = Trim(RKelompok("NAMALOKASI"))
    RKelompok.MoveFirst
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

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Text15 = 0
Text17 = 0
Combo2 = ""
Combo3 = ""
Combo4 = ""

Label19 = ""
Label18 = ""

Call IsiCombo

Me.Caption = INISIAL + " INVENTARIS BARANG KIB C - Gedung dan Bangunan"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text1 = GRIDKLIK
        Call Cari
        Call GAMBAR
    End If

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Text5 = TAHUNc
Text14 = SEMESTERc
DTPicker1 = Date

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
    
    Combo5 = RCari("CKONDISIBGN")
    Text6 = RCari("CBERTINGKAT")
    Text7 = RCari("CBETON")
    Text8 = RCari("CLUASLNT")
    Text9 = RCari("CALAMAT")
    Text10 = RCari("CNOMORDOC")
    DTPicker1 = RCari("CTANGGALDOC")
    Text12 = RCari("CLUAS")
    Text11 = RCari("CSTATUSTANAH")
    Text4 = RCari("CKODETANAH")
    Text16 = RCari("CASAL")
    Text15 = RCari("CHARGAPASAR")
    Text17 = RCari("CNILAILAIN")
    Text18 = RCari("CKETERANGAN")
    
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

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, ">")
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text11_LostFocus()
Text11 = Format(Text11, ">")
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text15_gotFocus()
Text15 = ""
End Sub

Private Sub Text15_LostFocus()
If Text15 = "" Then Text15 = 0
Text15 = Format(CCur(Text15), "##,###.00")
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text16_LostFocus()
Text16 = Format(Text16, ">")
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text17_gotFocus()
Text17 = ""
End Sub

Private Sub Text17_LostFocus()
If Text17 = "" Then Text17 = 0
Text17 = Format(CCur(Text17), "##,###.00")
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text18_LostFocus()
Text18 = Format(Text18, ">")
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
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_LostFocus()
Text9 = Format(Text9, ">")
End Sub

