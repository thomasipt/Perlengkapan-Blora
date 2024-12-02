VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EKIRADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EKIRADD"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   -2790
      TabIndex        =   49
      Text            =   "Text3"
      Top             =   2520
      Width           =   2265
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4635
      MaxLength       =   4
      TabIndex        =   17
      Text            =   "22"
      Top             =   7020
      Width           =   675
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3570
      MaxLength       =   4
      TabIndex        =   16
      Text            =   "21"
      Top             =   7020
      Width           =   675
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2430
      MaxLength       =   4
      TabIndex        =   15
      Text            =   "20"
      Top             =   7020
      Width           =   675
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   2970
      TabIndex        =   14
      Text            =   "15"
      Top             =   6225
      Width           =   2700
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2970
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   5895
      Width           =   2700
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2970
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   3525
      Width           =   2700
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2970
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   3840
      Width           =   2700
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2970
      TabIndex        =   8
      Text            =   "Text8"
      Top             =   4185
      Width           =   2700
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2970
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   4875
      Width           =   2700
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2970
      TabIndex        =   11
      Text            =   "Text10"
      Top             =   5205
      Width           =   2700
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   2970
      TabIndex        =   12
      Text            =   "Text11"
      Top             =   5535
      Width           =   2700
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2970
      TabIndex        =   9
      Text            =   "Text12"
      Top             =   4545
      Width           =   2700
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   22
      Text            =   "T5"
      Top             =   2970
      Width           =   960
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3900
      MaxLength       =   2
      TabIndex        =   23
      Text            =   "T14"
      Top             =   2970
      Width           =   960
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
      Left            =   8175
      TabIndex        =   21
      Top             =   7560
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
      Left            =   300
      TabIndex        =   19
      Top             =   7575
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
      Left            =   4230
      TabIndex        =   20
      Top             =   7575
      Width           =   1080
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   7920
      Top             =   1125
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   26
      Top             =   8295
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
      Left            =   -2340
      ScaleHeight     =   735
      ScaleWidth      =   12420
      TabIndex        =   27
      Top             =   7425
      Width           =   12480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   3030
      Left            =   -150
      TabIndex        =   28
      Top             =   -270
      Width           =   10140
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   2460
         TabIndex        =   2
         Text            =   "Text16"
         Top             =   990
         Width           =   4380
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2460
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2595
         Width           =   4980
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2460
         TabIndex        =   3
         Text            =   "Combo3"
         Top             =   1380
         Width           =   4980
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2460
         TabIndex        =   4
         Text            =   "Combo4"
         Top             =   2190
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2460
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   4380
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2460
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   675
         Width           =   4380
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2460
         TabIndex        =   51
         Top             =   1800
         Width           =   7140
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
         TabIndex        =   50
         Top             =   315
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Kelompok"
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
         Left            =   315
         TabIndex        =   48
         Top             =   2595
         Width           =   1380
      End
      Begin VB.Label Label13 
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
         Height          =   285
         Left            =   315
         TabIndex        =   43
         Top             =   1395
         Width           =   2595
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
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
         Left            =   315
         TabIndex        =   42
         Top             =   990
         Width           =   2595
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
         Left            =   315
         TabIndex        =   31
         Top             =   2190
         Width           =   2595
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   315
         TabIndex        =   30
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
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
         Left            =   315
         TabIndex        =   29
         Top             =   675
         Width           =   2595
      End
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
      Height          =   3930
      Left            =   5835
      TabIndex        =   25
      Top             =   3420
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
         Left            =   45
         TabIndex        =   18
         Top             =   3525
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
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   6690
      TabIndex        =   24
      Text            =   "Text13"
      Top             =   3600
      Width           =   1845
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "KEADAAN BARANG"
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
      Left            =   75
      TabIndex        =   47
      Top             =   6660
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   75
      Top             =   6615
      Width           =   2190
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Rusak"
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
      Left            =   4590
      TabIndex        =   46
      Top             =   6795
      Width           =   750
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Kurang"
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
      Left            =   3525
      TabIndex        =   45
      Top             =   6795
      Width           =   750
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Baik"
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
      Left            =   2385
      TabIndex        =   44
      Top             =   6795
      Width           =   750
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tahun Pembuatan / Pembelian"
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
      Left            =   225
      TabIndex        =   41
      Top             =   5520
      Width           =   2910
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mutasi"
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
      Left            =   225
      TabIndex        =   40
      Top             =   5880
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nilai Pasar"
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
      Left            =   225
      TabIndex        =   39
      Top             =   6210
      Width           =   1155
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bahan"
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
      Left            =   225
      TabIndex        =   38
      Top             =   4530
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Harga Beli"
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
      Left            =   225
      TabIndex        =   37
      Top             =   5190
      Width           =   1065
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Jumlah Barang / Register"
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
      Left            =   225
      TabIndex        =   36
      Top             =   4860
      Width           =   2910
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ukuran"
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
      Left            =   225
      TabIndex        =   35
      Top             =   4170
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "No. Seri"
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
      Left            =   225
      TabIndex        =   34
      Top             =   3825
      Width           =   930
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Merk"
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
      Left            =   225
      TabIndex        =   33
      Top             =   3510
      Width           =   930
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
      TabIndex        =   32
      Top             =   2955
      Width           =   4845
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Left            =   90
      Top             =   2880
      Width           =   5685
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   3165
      Left            =   75
      Top             =   3420
      Width           =   5685
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   75
      Top             =   6660
      Width           =   5685
   End
End
Attribute VB_Name = "EKIRADD"
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
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
EKIR.Show 1
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
EKIR.Show 1

End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Or Text13 = "" Or Text20 = "" Or Text21 = "" Or Text22 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

If CCur(Text9) <> CCur(Text20) + CCur(Text21) + CCur(Text22) Then
    MsgBox "JUMLAH BARANG REGISTER TIDAK SAMA DENGAN TOTAL KEADAAN", vbCritical, "KONFIRMASI"
    Text20.SetFocus
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
EKIR.Show 1
End Sub

Private Sub Simpan()
SCari = "Select * from BRGINVENTARIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

RCari.AddNew
    RCari("KODEBRG") = Trim(Text2)
    RCari("REGISTER") = Trim(Text1)
    RCari("NAMABRG") = Trim(Text16)
    RCari("SUBKELOMPOK") = Trim(Combo2)
    RCari("KODELOKASI") = Trim(Combo3)
    RCari("JENISLOKASI") = Trim(Label21)
    
    RCari("KELOMPOK") = Trim(Combo1)

    RCari("RUANG") = Combo4
    RCari("JENISBRG") = "RUANG DAN SUBDIN"
    RCari("KIB") = "KIR"
    RCari("TAHUN") = Trim(Text5)
    RCari("SEMESTER") = Trim(Text14)
    
    RCari("KONDISI") = "BAIK." + Trim(Text20) + " KURANG." + Trim(Text21) + " RUSAK." + Trim(Text22)
    
    RCari("RMERK") = Trim(Text6)
    RCari("RNOSERI") = Trim(Text7)
    RCari("RUKURAN") = Trim(Text8)
    RCari("RBAHAN") = Trim(Text12)
    RCari("RJUMLAHBRG") = CCur(Text9)
    RCari("RHARGABELI") = CCur(Text10)
    RCari("RNILAIPASAR") = CCur(Text15)
    RCari("RMUTASI") = Trim(Text4)
    RCari("RTAHUN") = Text11
    
    RCari("PHOTO") = Trim(Text13)
    
    RCari("KODELOKASISEBELUM") = Combo3
    RCari("JENISLOKASISEBELUM") = Trim(Label21)
    RCari("RUANGSEBELUM") = Combo4
    
    RCari("KODELOKASISESUDAH") = "-"
    RCari("JENISLOKASISESUDAH") = "-"
    RCari("RUANGSESUDAH") = "-"
    
    RCari("RBAIK") = CCur(Text20)
    RCari("RKURANG") = CCur(Text21)
    RCari("RRUSAK") = CCur(Text22)
    
    RCari("RESTORE") = Date

RCari.Update
RCari.Close
Set RCari = Nothing

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
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
    Label21 = Trim(RKelompok("NAMALOKASI"))
    RKelompok.MoveFirst
    Do While Not RKelompok.EOF
        Combo4.AddItem Trim(RKelompok("RUANG"))
    RKelompok.MoveNext
    Loop
    Combo4.ListIndex = 0
Else
    MsgBox "LOKASI TIDAK TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo3.ListIndex = 0
    Combo3.SetFocus
    Exit Sub
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
Text17 = 0

Call IsiCombo

Me.Caption = INISIAL + " INVENTARIS BARANG KIR - Ruang dan SubDin"

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Text5 = TAHUNc
Text14 = SEMESTERc

Text20 = 0
Text21 = 0
Text22 = 0

    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
        'Call AutoKode
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text3 = GRIDKLIK2
        Call Cari
        Call GAMBAR
    End If

End Sub

'Private Sub AutoKode()
'Dim No As Double
'SCari = "Select KODE_LOKASI from TEMP"
'Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
'If RCari.RowCount <> 0 Then
'    No = Right(RCari("KODE_LOKASI"), 2) + 1
'    Text1 = Left(RCari("KODE_LOKASI"), (Len(RCari("KODE_LOKASI")) - 3))
'    Text11 = Digit(2, No)
'End If
'RCari.Close
'Set RCari = Nothing
'End Sub

Private Sub AutoKode()
Dim No As Double
SCari = "Select KODE_BARANG from TEMP"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    No = Right(RCari("KODE_BARANG"), 2) + 1
    Text1 = Trim(Left(RCari("KODE_BARANG"), (Len(RCari("KODE_BARANG")) - 2))) + Trim(No)
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiCombo()

SCombo3 = "Select KODELOKASI from V_LOKASI order by KODELOKASI Asc"
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

SCombo1 = "Select URAIAN from KELOMPOKBRG order by URAIAN Asc"
Set RCombo1 = RDCO.OpenResultset(SCombo1, rdOpenDynamic, rdConcurRowVer)
If RCombo1.RowCount <> 0 Then
    RCombo1.MoveFirst
    Do Until RCombo1.EOF
        Combo1.AddItem RCombo1("URAIAN")
    RCombo1.MoveNext
    Loop
Else
    MsgBox "MASUKAN DATA KELOMPOK", vbCritical, "KONFIRMASI"
    Exit Sub
End If
RCombo1.Close
Set RCombo1 = Nothing
Combo1.ListIndex = 0

Combo4 = ""

End Sub

Private Sub Cari()
SCari = "Select * from BRGINVENTARIS where NO_URUT like '" + Text3 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Text2 = RCari("KODEBRG")
    Text1 = RCari("REGISTER")
    Text16 = Trim(RCari("NAMABRG"))
    Combo3 = Trim(RCari("KODELOKASI"))
    Combo1 = Trim(RCari("KELOMPOK"))
    Combo4 = Trim(RCari("RUANG"))
    Text5 = RCari("TAHUN")
    Text14 = RCari("SEMESTER")
    
    Text6 = RCari("RMERK")
    Text7 = RCari("RNOSERI")
    Text8 = RCari("RUKURAN")
    Text12 = RCari("RBAHAN")
    Text9 = CCur(RCari("RJUMLAHBRG"))
    Text10 = CCur(RCari("RHARGABELI"))
    Text15 = CCur(RCari("RNILAIPASAR"))
    Text4 = Trim(RCari("RMUTASI"))
    Text11 = (RCari("RTAHUN"))
    
    Text13 = (RCari("PHOTO"))
    
    Label21 = RCari("JENISLOKASISEBELUM")
    Text20 = RCari("RBAIK")
    Text21 = RCari("RKURANG")
    Text22 = RCari("RRUSAK")
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
'Call CekData
'Text2 = Text1
End Sub

Private Sub CekBarang()
If TIPE = 2 Then Exit Sub

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

Private Sub CekData()
SCari = "Select * From BRGINVENTARIS where KODEBRG = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    MsgBox "KODE SUDAH ADA", vbCritical, "KONFIRMASI"
    Text1 = ""
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_gotFocus()
Text10 = ""
End Sub

Private Sub Text10_LostFocus()
If Text10 = "" Then Text10 = 0
Text10 = Format(CCur(Text10), "##,###.00")
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
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text12_LostFocus()
Text12 = Format(Text12, ">")
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
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

Private Sub Text15_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text16_LostFocus()
Text16 = Format(Text16, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub Text20_gotFocus()
Text20 = ""
End Sub

Private Sub Text20_LostFocus()
If Text20 = "" Then Text20 = 0
End Sub

Private Sub Text21_gotFocus()
Text21 = ""
End Sub

Private Sub Text21_LostFocus()
If Text21 = "" Then Text21 = 0
End Sub

Private Sub Text22_gotFocus()
Text22 = ""
End Sub

Private Sub Text22_LostFocus()
If Text22 = "" Then Text22 = 0
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
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
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
