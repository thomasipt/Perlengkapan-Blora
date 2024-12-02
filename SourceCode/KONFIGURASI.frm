VERSION 5.00
Begin VB.Form KONFIGURASI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KONFIGURASI"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RESTORE DATA CABANG"
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
      Left            =   98
      TabIndex        =   23
      Top             =   7470
      Width           =   5625
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PERGANTIAN SEMESTER"
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
      Left            =   98
      TabIndex        =   22
      Top             =   6300
      Width           =   5625
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Keluar"
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
      Left            =   4583
      TabIndex        =   21
      Top             =   5265
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
      Left            =   2370
      TabIndex        =   20
      Top             =   5265
      Width           =   1080
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   4620
      Width           =   1170
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1695
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   4005
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1695
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   4200
      Width           =   1170
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1695
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   3690
      Width           =   4005
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1695
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   3270
      Width           =   4005
   End
   Begin VB.TextBox Text3 
      Height          =   1785
      Left            =   1695
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "KONFIGURASI.frx":0000
      Top             =   1380
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1695
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   960
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1695
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   540
      Width           =   4005
   End
   Begin VB.CommandButton cmdSAVE2 
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
      Height          =   495
      Left            =   158
      TabIndex        =   18
      Top             =   5265
      Width           =   1080
   End
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
      Height          =   495
      Left            =   165
      TabIndex        =   8
      Top             =   5265
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   -105
      ScaleHeight     =   675
      ScaleWidth      =   14865
      TabIndex        =   19
      Top             =   5145
      Width           =   14925
   End
   Begin VB.PictureBox Picture2 
      Height          =   870
      Left            =   -4552
      ScaleHeight     =   810
      ScaleWidth      =   14865
      TabIndex        =   24
      Top             =   7335
      Width           =   14925
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      TabIndex        =   17
      Top             =   4200
      Width           =   2430
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
      Left            =   120
      TabIndex        =   16
      Top             =   3270
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
      Left            =   120
      TabIndex        =   15
      Top             =   2115
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
      Left            =   120
      TabIndex        =   14
      Top             =   960
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
      Left            =   120
      TabIndex        =   13
      Top             =   540
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
      Left            =   120
      TabIndex        =   12
      Top             =   3690
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
      Left            =   120
      TabIndex        =   11
      Top             =   4200
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
      Left            =   120
      TabIndex        =   10
      Top             =   4620
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "KONFIGURASI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RST, RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private RSQL, SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String


Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private Lolos

Private Sub cmdEDIT_Click()
Call IsiCombo
Combo1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
cmdSAVE.Enabled = True
cmdEDIT.Enabled = False

Combo1.SetFocus
cmdSAVE.Visible = False
cmdSAVE2.Visible = True
End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Combo1, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If
    
Call Simpan

Unload Me
KONFIGURASI.Show 1
End Sub

Private Sub Simpan()
SCari = "Select * from KONFIG"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.AddNew
    RCari("KODELOKASI") = Combo1
    RCari("JENISLOKASI") = Text1
    RCari("NAMALOKASI") = Text2
    RCari("ALAMAT") = Text3
    RCari("PENGURUS") = Text4
    RCari("KEPALA_SKPD") = Text5
    RCari("SEMESTER") = Text6
    RCari("TAHUN") = Text7
    
    RCari("RESTORE") = Date
RCari.Update
RCari.Close
Set RCari = Nothing

SCari = "Select * from KONFIG_DATA"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
RCari.AddNew
    RCari("KODELOKASI") = Combo1
    RCari("JENISLOKASI") = Text1
    RCari("NAMALOKASI") = Text2
    RCari("ALAMAT") = Text3
    RCari("SEMESTER") = Text6
    RCari("TAHUN") = Text7
    RCari("TGL") = Now
RCari.Update
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Delete()
SEdit = "Delete * from KONFIG"
Set REdit = RDCO.OpenResultset(SEdit, rdOpenDynamic, rdConcurRowVer)
REdit.Close
Set REdit = Nothing
End Sub

Private Sub cmdSAVE2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Combo1, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If
    
Call Delete
Call Simpan

Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
MAINMENU.Show
End Sub

Private Sub Command2_Click()
Unload Me
EOD001.Show 1
End Sub

Private Sub Command3_Click()
ZBACKUP.Show 1
End Sub

Private Sub Command4_Click()
ZRESTORE.Show 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""
Label9 = ""
Call CEKKONFIG
'Call Cari
cmdSAVE2.Visible = False
End Sub

Private Sub CEKKONFIG()
SKNG = "Select * From KONFIG"
Set RKNG = RDCO.OpenResultset(SKNG, rdOpenKeyset, rdConcurRowVer)
If RKNG.RowCount <> 0 Then
    Combo1 = Format(RKNG("KODELOKASI"), ">")
    Text1 = Format(RKNG("JENISLOKASI"), ">")
    Text2 = Format(RKNG("NAMALOKASI"), ">")
    Text3 = Format(RKNG("ALAMAT"), ">")
    Text4 = Format(RKNG("PENGURUS"), ">")
    Text5 = Format(RKNG("KEPALA_SKPD"), ">")
    Text6 = Format(RKNG("SEMESTER"), ">")
    Text7 = RKNG("TAHUN")
    Call KUNCI
    cmdSAVE.Enabled = False
Else
    MsgBox "LAKUKAN KONFIGURASI PERTAMAKALI", vbInformation, "KONFIRMASI"
    Call IsiCombo
    cmdEDIT.Enabled = False
End If
RKNG.Close
Set RKNG = Nothing
End Sub

Private Sub KUNCI()
Combo1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
End Sub

Private Sub IsiCombo()
SCombo = "Select KODELOKASI from LOKASI order by KODELOKASI"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenKeyset, rdConcurRowVer)
If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo1.AddItem RCombo("KODELOKASI")
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

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SKode = "Select * From LOKASI where KODELOKASI = '" + Combo1 + "'"
Set RKode = RDCO.OpenResultset(SKode, rdOpenKeyset, rdConcurRowVer)
If RKode.RowCount <> 0 Then
    Text1 = Format(RKode("JENISLOKASI"), ">")
    Text2 = Format(RKode("NAMALOKASI"), ">")
Else
    Combo1.SetFocus
    MsgBox "KODE LOKASI BELUM TERDAFTAR", vbInformation, "KONFIRMASI"
End If
RKode.Close
Set RKode = Nothing
Combo1 = Format(Combo1, ">")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
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

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
If Text6 = "" Then Exit Sub
Label9 = "( " + Satuan(Text6) + " )"
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub
