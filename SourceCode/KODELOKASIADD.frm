VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form KODELOKASIADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODELOKASIADD"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   765
      Width           =   2955
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5670
      TabIndex        =   0
      Text            =   "11"
      Top             =   90
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1845
      TabIndex        =   7
      Text            =   "1"
      Top             =   90
      Width           =   3465
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
      TabIndex        =   5
      Top             =   2190
      Width           =   1080
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1845
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1605
      Width           =   4410
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1845
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1185
      Width           =   4410
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
      Height          =   495
      Left            =   255
      TabIndex        =   4
      Top             =   2190
      Width           =   1080
   End
   Begin VB.CommandButton cmdCANCEL 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
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
      Left            =   4785
      TabIndex        =   6
      Top             =   2190
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      Height          =   705
      Left            =   -105
      ScaleHeight     =   645
      ScaleWidth      =   14865
      TabIndex        =   11
      Top             =   2085
      Width           =   14925
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   2865
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3810
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3810
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3810
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
   Begin VB.Label Label2 
      Caption         =   "Cabang"
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
      Left            =   315
      TabIndex        =   13
      Top             =   765
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Lokasi                                                 ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   -1035
      TabIndex        =   8
      Top             =   150
      Width           =   7950
   End
   Begin VB.Label Label4 
      Caption         =   "Ruangan"
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
      Left            =   315
      TabIndex        =   10
      Top             =   1605
      Width           =   1380
   End
   Begin VB.Label Label3 
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
      Left            =   315
      TabIndex        =   9
      Top             =   1185
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   -90
      Top             =   -45
      Width           =   6900
   End
End
Attribute VB_Name = "KODELOKASIADD"
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

Private Sub Combo1_GotFocus()
    SendKeys "{F4}"
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text11.SetFocus
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text11_LostFocus()
'IPT = ""

If Text1.Text = "" Or Text11.Text = "" Then Exit Sub

'IPT = Trim(Text1) + "." + Trim(Text11)

'SLKS = "Select * from LOKASI where KODELOKASI = '" + Trim(IPT) + "'"
'Set RLKS = RDCO.OpenResultset(SLKS, rdOpenKeyset, rdConcurRowVer)
'If RLKS.RowCount <> 0 Then
'    Combo1 = RLKS("JENISLOKASI")
'    Text3 = RLKS("NAMALOKASI")
'    Text4 = RLKS("RUANG")
'    ColorTextBoxes Me
'End If
'RLKS.Close
'Set RLKS = Nothing

'Text1 = Format(Text1, ">")
'Text11 = Format(Text11, ">")

End Sub

Private Sub cmdCANCEL_Click()
Unload Me
KODELOKASI.Show 1
End Sub

Private Sub cmdEDIT_Click()
Call Edit
End Sub

Private Sub Edit()
TANYA = MsgBox("EDIT KODE LOKASI " + Text1, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SEdit = "Delete from LOKASI where KODELOKASI = '" + Trim(OYEN) + "' and RUANG ='" + Trim(GRIDKLIK2) + "' "
Set REdit = RDCO.OpenResultset(SEdit, rdOpenKeyset, rdConcurRowVer)
REdit.Close
Set REdit = Nothing


Call Simpan

ClearTextBoxes Me
Text1.SetFocus

MsgBox "UPDATE DATA", vbCritical, "KONFIRMASI"

Unload Me
KODELOKASI.Show 1

End Sub

Private Sub cmdSAVE_Click()
If Text1.Text = "" Or Combo1 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

Call Simpan
ClearTextBoxes Me
Text1.SetFocus

MsgBox "UPDATE DATA", vbCritical, "KONFIRMASI"

Unload Me
KODELOKASI.Show 1
End Sub

Private Sub Simpan()
SCari = "Select * from LOKASI"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.AddNew
    RCari("KODELOKASI") = Trim(Text1) + "." + Trim(Text11)
    RCari("JENISLOKASI") = Combo1
    RCari("NAMALOKASI") = Trim(Text3)
    RCari("RUANG") = Trim(Text4)
RCari.Update
RCari.Close
Set RCari = Nothing

SEdit = "Delete from TEMP"
Set REdit = RDCO.OpenResultset(SEdit, rdOpenDynamic, rdConcurRowVer)


SCari3 = "Select * from TEMP"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenDynamic, rdConcurRowVer)
RCari3.AddNew
    RCari3("KODE_LOKASI") = Trim(Text1) + "." + Trim(Text11)
RCari3.Update
RCari3.Close
Set RCari3 = Nothing

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Text1.Text = ""

Me.Caption = INISIAL + " NOMOR KODE LOKASI CABANG DINAS DIKNAS"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
        Call AutoKode
        Call IsiCombo
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text1 = Format(Left(GRIDKLIK, (Len(GRIDKLIK) - 3)), ">")
        Text11 = Right(GRIDKLIK, 2)
        Text4 = Format(GRIDKLIK2, ">")
        Call Cari
        Call IsiCombo
    End If
    
With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

End Sub

Private Sub IsiCombo()
SCombo1 = "Select * from TEMP_1 order by CABANG Asc"
Set RCombo1 = RDCO.OpenResultset(SCombo1, rdOpenDynamic, rdConcurRowVer)
If RCombo1.RowCount <> 0 Then
    RCombo1.MoveFirst
    Do Until RCombo1.EOF
        Combo1.AddItem RCombo1("CABANG")
    RCombo1.MoveNext
    Loop
    RCombo1.Close
    Set RCombo1 = Nothing
    Combo1.ListIndex = 0
End If
End Sub

Private Sub AutoKode()
Dim No As Double
SCari = "Select KODE_LOKASI from TEMP"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    No = Right(RCari("KODE_LOKASI"), 2) + 1
    Text1 = Left(RCari("KODE_LOKASI"), (Len(RCari("KODE_LOKASI")) - 3))
    Text11 = Digit(2, No)
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Cari()
OYEN = ""
OYEN = Trim(GRIDKLIK)
SCari = "Select * from LOKASI where KODELOKASI = '" + Trim(OYEN) + "' and RUANG ='" + Trim(Text4) + "' "
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Combo1.AddItem RCari("JENISLOKASI")
    Text3 = RCari("NAMALOKASI")
    Text4 = RCari("RUANG")
    Combo1.ListIndex = 0
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
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

