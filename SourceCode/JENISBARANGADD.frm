VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form JENISBARANGADD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JENIS BARANG"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2430
      TabIndex        =   7
      Top             =   1260
      Width           =   1080
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   630
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   210
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
      Left            =   4733
      TabIndex        =   3
      Top             =   1260
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
      Left            =   135
      TabIndex        =   2
      Top             =   1260
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   -210
      ScaleHeight     =   690
      ScaleWidth      =   14865
      TabIndex        =   6
      Top             =   1155
      Width           =   14925
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3440
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3440
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3440
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
   Begin VB.Label Label2 
      Caption         =   "K I B"
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
      Left            =   240
      TabIndex        =   5
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   1380
   End
End
Attribute VB_Name = "JENISBARANGADD"
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
JENISBARANG.Show 1
End Sub

Private Sub cmdEDIT_Click()
Call Edit
End Sub

Private Sub cmdSAVE_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
    ColorTextBoxes Me
    Exit Sub
End If

TANYA = MsgBox("SIMPAN " + Text1 + " & JENIS " + Text2, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call Simpan
ClearTextBoxes Me
Text1.SetFocus

MsgBox "UPDATE DATA", vbCritical, "KONFIRMASI"

Unload Me
JENISBARANG.Show 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""
Combo2 = ""
Me.Caption = INISIAL + " JENIS BARANG"
    If TIPE = 1 Then
        cmdSAVE.Enabled = True
        cmdEDIT.Enabled = False
    ElseIf TIPE = 2 Then
        cmdSAVE.Enabled = False
        cmdEDIT.Enabled = True
        Text1 = GRIDKLIK
        Text2 = GRIDKLIK2
        'Call Cari
    End If
    
With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

'Call AutoKode

End Sub

Private Sub AutoKode()
Dim No As Double
SCari = "Select KODEKELOMPOK from KELOMPOKBRG order by KODEKELOMPOK Desc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    No = Right(RCari("KODEKELOMPOK"), 5) + 1
    NoStr = "KL." + Digit(5, No)
    Text1 = NoStr
Else
    No = 1
    NoStr = "KL." + Digit(5, No)
    Text1 = NoStr
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Cari()
SCari = "Select KIB from JENISBRG where JNSBRG = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

If RCari.RowCount <> 0 Then
    Text2 = RCari("KIB")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Simpan()
SCari = "Select * from JENISBRG"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
RCari.AddNew
    RCari("JNSBRG") = Text1
    RCari("KIB") = Text2
RCari.Update
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Edit()
TANYA = MsgBox("EDIT " + Text1 + " & JENIS " + Text2, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SDel = "Delete from JENISBRG where JNSBRG = '" + Trim(GRIDKLIK) + "'"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing
    
Call Simpan

ClearTextBoxes Me
Text1.SetFocus

MsgBox "UPDATE DATA", vbCritical, "KONFIRMASI"

Unload Me
JENISBARANG.Show 1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then Exit Sub
SLKS = "Select * from JENISBRG where JNSBRG = '" + Trim(Text1.Text) + "'"
Set RLKS = RDCO.OpenResultset(SLKS, rdOpenKeyset, rdConcurRowVer)
If RLKS.RowCount <> 0 Then
    Text2 = RLKS("KIB")
End If
RLKS.Close
Set RLKS = Nothing

Text1 = Format(Text1, ">")

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TIPE = 1 Then
        cmdSAVE.SetFocus
    ElseIf TIPE = 2 Then
        cmdEDIT.SetFocus
    End If
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub
