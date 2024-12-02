VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form PKMBADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PKMBADD"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text100 
      Height          =   915
      Left            =   10350
      TabIndex        =   41
      Text            =   "Text100"
      Top             =   2025
      Width           =   1005
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2745
      TabIndex        =   40
      Text            =   "Text20"
      Top             =   3615
      Width           =   690
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   4725
      TabIndex        =   39
      Text            =   "Text21"
      Top             =   3615
      Width           =   690
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   6540
      TabIndex        =   38
      Text            =   "Text22"
      Top             =   3615
      Width           =   690
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Text4"
      Top             =   3195
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Text            =   "Combo3"
      Top             =   1935
      Width           =   2595
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1710
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2655
      Width           =   6135
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   1230
      Width           =   2595
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mutasi Barang Ke"
      Height          =   1380
      Left            =   240
      TabIndex        =   20
      Top             =   5640
      Width           =   5790
      Begin VB.ComboBox Combo11 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         TabIndex        =   7
         Text            =   "Combo11"
         Top             =   645
         Width           =   3375
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Text            =   "Combo10"
         Top             =   315
         Width           =   3375
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         Text            =   "Combo9"
         Top             =   975
         Width           =   3375
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
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
         Left            =   210
         TabIndex        =   23
         Top             =   645
         Width           =   1485
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
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
         Left            =   210
         TabIndex        =   22
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ruang"
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
         Left            =   210
         TabIndex        =   21
         Top             =   975
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mutasi Barang Dari"
      Height          =   1335
      Left            =   240
      TabIndex        =   17
      Top             =   4260
      Width           =   5790
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   315
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text7"
         Top             =   975
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text6"
         Top             =   645
         Width           =   3375
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
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
         Left            =   210
         TabIndex        =   32
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ruang"
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
         Left            =   210
         TabIndex        =   19
         Top             =   975
         Width           =   1485
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
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
         Left            =   210
         TabIndex        =   18
         Top             =   645
         Width           =   1485
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   450
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59244545
      CurrentDate     =   39585
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1710
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   840
      Width           =   2955
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1665
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   105
      Width           =   2535
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
      Left            =   3412
      TabIndex        =   10
      Top             =   7335
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
      Height          =   480
      Left            =   6626
      TabIndex        =   11
      Top             =   7335
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
      Left            =   225
      TabIndex        =   9
      Top             =   7335
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   -210
      ScaleHeight     =   795
      ScaleWidth      =   14865
      TabIndex        =   24
      Top             =   7170
      Width           =   14925
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   26
      Top             =   8235
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4683
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4683
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4683
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
   Begin VB.PictureBox Picture2 
      Height          =   2925
      Left            =   -315
      ScaleHeight     =   2865
      ScaleWidth      =   14865
      TabIndex        =   25
      Top             =   4185
      Width           =   14925
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   2160
      TabIndex        =   34
      Text            =   "Combo5"
      Top             =   3615
      Width           =   3375
   End
   Begin VB.Label Label8 
      Caption         =   "Kondisi Sesudah         BAIK                    KURANG                    RUSAK"
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
      Left            =   135
      TabIndex        =   36
      Top             =   3615
      Width           =   8865
   End
   Begin VB.Label Label7 
      Caption         =   "Kondisi Sebelum"
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
      Left            =   135
      TabIndex        =   37
      Top             =   3195
      Width           =   5265
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1710
      TabIndex        =   29
      Top             =   2295
      Width           =   6150
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1710
      TabIndex        =   28
      Top             =   1590
      Width           =   6150
   End
   Begin VB.Label Label3 
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
      Left            =   135
      TabIndex        =   27
      Top             =   2655
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
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   1230
      Width           =   1380
   End
   Begin VB.Label Label6 
      Caption         =   "Jenis Barang"
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
      Left            =   135
      TabIndex        =   15
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label5 
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
      Left            =   135
      TabIndex        =   14
      Top             =   450
      Width           =   1380
   End
   Begin VB.Label Label4 
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
      Height          =   315
      Left            =   135
      TabIndex        =   13
      Top             =   1980
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "No. Bukti"
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
      Left            =   135
      TabIndex        =   12
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "PKMBADD"
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

Private Sub AutoCompleteCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
PKMB.Show 1
End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Combo10 = "" Or Combo11 = "" Or Combo9 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "WARNING"
    Text1.SetFocus
    Exit Sub
End If

If Combo5.Visible = False Then
    If CCur(Text20) + CCur(Text21) + CCur(Text22) <> CCur(Text100) Then
        MsgBox "JUMLAH KONDISI SESUDAH SALAH", vbCritical, "WARNING"
        cmdSAVE.SetFocus
        Exit Sub
    End If
End If

TANYA = MsgBox("SIMPAN MUTASI " + Combo1, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Edit
Call HisBarang

ClearTextBoxes Me
Text1.SetFocus
Unload Me
PKMB.Show 1
End Sub

Private Sub Edit()
SEdit = "Select * From BRGINVENTARIS where KODEBRG = '" + Trim(Combo3) + "'"
Set REdit = RDCO.OpenResultset(SEdit, rdOpenDynamic, rdConcurRowVer)
If REdit.RowCount <> 0 Then
    REdit.Edit
    REdit("KODELOKASI") = Trim(Combo10)
    REdit("JENISLOKASI") = Trim(Combo11)
    REdit("RUANG") = Trim(Combo9)
    
    REdit("KODELOKASISESUDAH") = Trim(Combo10)
    REdit("JENISLOKASISESUDAH") = Trim(Combo11)
    REdit("RUANGSESUDAH") = Trim(Combo9)
    
    REdit("KODELOKASISEBELUM") = Trim(Combo2)
    REdit("JENISLOKASISEBELUM") = Trim(Label15)
    REdit("RUANGSEBELUM") = Trim(Text7)
    
    If Combo5.Visible = False Then
        REdit("KONDISI") = "BAIK." + Trim(Text20) + " KURANG." + Trim(Text21) + " RUSAK." + Trim(Text22)
        REdit("RBAIK") = CCur(Text20)
        REdit("RKURANG") = CCur(Text21)
        REdit("RRUSAK") = CCur(Text22)
    Else
        REdit("KONDISI") = Trim(Combo5)
    End If
    
    REdit.Update
End If
REdit.Close
Set REdit = Nothing
End Sub

Private Sub HisBarang()
SSave = "Select * from MUTASI"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NOBUKTI") = Trim(Text1)
    RSave("JENISBRG") = Trim(Combo1)
    RSave("TANGGAL") = DTPicker1
    RSave("KODELOKASI") = Trim(Combo2)
    RSave("JENISLOKASI") = Trim(Label15)
    RSave("KODEBARANG") = Trim(Combo3)
    RSave("NAMABARANG") = Trim(Label16)
    RSave("RUANG") = Trim(Text7)
    RSave("KETERANGAN") = Trim(Text3)
    
    RSave("KONDISISEBELUM") = Trim(Text4)
    
    If Combo5.Visible = False Then
        RSave("KONDISISESUDAH") = "BAIK." + Trim(Text20) + " KURANG." + Trim(Text21) + " RUSAK." + Trim(Text22)
    Else
        RSave("KONDISISESUDAH") = Trim(Combo5)
    End If
    
    RSave("KODELOKASISEBELUM") = Trim(Text5)
    RSave("JENISLOKASISEBELUM") = Trim(Text6)
    RSave("RUANGSEBELUM") = Trim(Text7)
    
    RSave("KODELOKASISESUDAH") = Trim(Combo10)
    RSave("JENISLOKASISESUDAH") = Trim(Combo11)
    RSave("RUANGSESUDAH") = Trim(Combo9)
   
RSave.Update
RSave.Close
Set RSave = Nothing

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub

If Combo1 = "RUANG DAN SUBDIN" Then
    Combo5.Visible = False
    
    Text20.Visible = True
    Text21.Visible = True
    Text22.Visible = True
    Label8.Caption = "Kondisi Sesudah         BAIK                    KURANG                    RUSAK"
Else
    Combo5.Visible = True
    
    Text20.Visible = False
    Text21.Visible = False
    Text22.Visible = False
    Label8.Caption = "Kondisi Sesudah"
End If

Call IsiCombo2
End Sub

Private Sub Combo10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo10_LostFocus()
If Combo10 = "" Then Exit Sub
Call IsiCombo11

If Text7 = "-" Then
    Combo9 = "-"
Else
    Call IsiCombo9
End If


End Sub

Private Sub Combo11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo11_LostFocus()
If Combo11 = "" Then Exit Sub

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo2_LostFocus()
If Combo2 = "" Then Exit Sub
Call IsiCombo3
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
Call IsiText2
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Label15 = ""
Label16 = ""

Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""
Combo5 = ""
Combo6 = ""
Combo7 = ""
Combo8 = ""
Combo9 = ""
Combo10 = ""
Combo11 = ""

Combo5.Visible = False
Text20.Visible = False
Text21.Visible = False
Text22.Visible = False

Label8.Caption = "Kondisi Sesudah"

DTPicker1 = Date

Me.Caption = INISIAL + " PERUBAHAN DAN MUTASI BARANG"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
        Call IsiCombo1
        Combo5.AddItem "BAIK"
        Combo5.AddItem "KURANG"
        Combo5.AddItem "RUSAK"
        Combo5.ListIndex = 0
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        'AutoCompleteCombo1 = GRIDKLIK
        Text1 = GRIDKLIK
        Call Cari
    End If

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With


End Sub

Private Sub Cari()
SCari = "Select * from MUTASI where NOBUKTI = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    DTPicker1 = RCari("tanggal")
    Combo1 = Trim(RCari("JENISBRG"))
    'Combo2 = Trim(RCari("KODELOKASI"))
    Combo3 = Trim(RCari("KODEBARANG"))
    Text3 = Trim(RCari("KETERANGAN"))
    Text4 = Trim(RCari("KONDISISEBELUM"))
    Combo5 = Trim(RCari("KONDISISESUDAH"))
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub IsiCombo1()
SCombo1 = "Select JNSBRG from JENISBRG order by KIB"
Set RCombo1 = RDCO.OpenResultset(SCombo1, rdOpenDynamic, rdConcurRowVer)
If RCombo1.RowCount <> 0 Then
    RCombo1.MoveFirst
    Do Until RCombo1.EOF
        Combo1.AddItem RCombo1("JNSBRG")
    RCombo1.MoveNext
    Loop
    Combo1.ListIndex = 0
End If
RCombo1.Close
Set RCombo1 = Nothing

SCombo10 = "Select * from LOKASI order by NoUrut"
Set RCombo10 = RDCO.OpenResultset(SCombo10, rdOpenDynamic, rdConcurRowVer)
If RCombo10.RowCount <> 0 Then
    RCombo10.MoveFirst
    Do Until RCombo10.EOF
        Combo10.AddItem RCombo10("KODELOKASI")
    RCombo10.MoveNext
    Loop
    Combo10.ListIndex = 0
End If
RCombo10.Close
Set RCombo10 = Nothing

End Sub

'LAMA TGL 27/10/2008

'Private Sub IsiCombo2()
'Combo2.Clear
'Combo10.Clear

'SCombo2 = "Select KODELOKASI from V_LOKASI order by KODELOKASI"
'Set RCombo2 = RDCO.OpenResultset(SCombo2, rdOpenDynamic, rdConcurRowVer)

'If RCombo2.RowCount <> 0 Then
'    RCombo2.MoveFirst
'    Do Until RCombo2.EOF
'        Combo2.AddItem RCombo2("KODELOKASI")
'        Combo10.AddItem RCombo2("KODELOKASI")
'    RCombo2.MoveNext
'    Loop
'    Combo2.ListIndex = 0
'    Combo10.ListIndex = 0
'End If
'RCombo2.Close
'Set RCombo2 = Nothing
'End Sub

Private Sub IsiCombo2()
Combo2.Clear
'Combo10.Clear

SCombo2 = "SELECT KODELOKASI From PKMBADD where JENISBRG = '" + Trim(Combo1) + "'"
Set RCombo2 = RDCO.OpenResultset(SCombo2, rdOpenDynamic, rdConcurRowVer)

If RCombo2.RowCount <> 0 Then
    RCombo2.MoveFirst
    Do Until RCombo2.EOF
        Combo2.AddItem RCombo2("KODELOKASI")
        'Combo10.AddItem RCombo2("KODELOKASI")
    RCombo2.MoveNext
    Loop
    Combo2.ListIndex = 0
    'Combo10.ListIndex = 0
End If
RCombo2.Close
Set RCombo2 = Nothing
End Sub

Private Sub IsiCombo3()
Combo3.Clear
SCombo3 = "Select * from BRGINVENTARIS where KODELOKASI = '" + Trim(Combo2) + "' and JENISBRG = '" + Trim(Combo1) + "'"
Set RCombo3 = RDCO.OpenResultset(SCombo3, rdOpenDynamic, rdConcurRowVer)

If RCombo3.RowCount <> 0 Then
    RCombo3.MoveFirst
    Do Until RCombo3.EOF
        Combo3.AddItem RCombo3("KODEBRG")
        Label15 = RCombo3("JENISLOKASI")
    RCombo3.MoveNext
    Loop
    Combo3.ListIndex = 0
End If
RCombo3.Close
Set RCombo3 = Nothing
End Sub

Private Sub IsiCombo11()
Combo11.Clear
SCombo11 = "Select NAMALOKASI from LOKASI where KODELOKASI ='" + Trim(Combo10) + "'"
Set RCombo11 = RDCO.OpenResultset(SCombo11, rdOpenDynamic, rdConcurRowVer)
If RCombo11.RowCount <> 0 Then
    RCombo11.MoveFirst
    Do Until RCombo11.EOF
        Combo11.AddItem RCombo11("NAMALOKASI")
    RCombo11.MoveNext
    Loop
End If
RCombo11.Close
Set RCombo11 = Nothing

Combo11.ListIndex = 0
End Sub

Private Sub IsiCombo9()
Combo9.Clear
SCombo9 = "Select RUANG from LOKASI where KODELOKASI ='" + Trim(Combo10) + "'"
Set RCombo9 = RDCO.OpenResultset(SCombo9, rdOpenDynamic, rdConcurRowVer)
If RCombo9.RowCount <> 0 Then
    RCombo9.MoveFirst
    Do Until RCombo9.EOF
        Combo9.AddItem RCombo9("RUANG")
    RCombo9.MoveNext
    Loop
End If
RCombo9.Close
Set RCombo9 = Nothing

Combo9.ListIndex = 0
End Sub

Private Sub IsiText2()
SCari = "Select * from BRGINVENTARIS where KODEBRG = '" + Trim(Combo3) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

On Error Resume Next
If RCari.RowCount <> 0 Then
    Text4 = RCari("KONDISI")
    Text5 = RCari("KODELOKASISEBELUM")
    Text6 = RCari("JENISLOKASISEBELUM")
    Text7 = RCari("RUANGSEBELUM")
    
    Label16 = RCari("NAMABRG")
    
    If Combo5.Visible = False Then
        Text20 = RCari("RBAIK")
        Text21 = RCari("RKURANG")
        Text22 = RCari("RRUSAK")
        
        Text100 = CCur(Text20) + CCur(Text21) + CCur(Text22)
    End If
    
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
If TIPE = 2 Then Exit Sub

'Text1 = Digit(10, Val(Text1))
STiket = "Select NOBUKTI from MUTASI where NOBUKTI = '" + Trim(Text1) + "'"
Set RTiket = RDCO.OpenResultset(STiket, rdOpenDynamic, rdConcurRowVer)

If RTiket.RowCount <> 0 Then
    MsgBox " NOMOR BUKTI TELAH DIGUNAKAN", vbOKOnly, "KONFIRMASI"
    Text1.SetFocus
    Text1 = ""
End If
RTiket.Close
Set RTiket = Nothing
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
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
