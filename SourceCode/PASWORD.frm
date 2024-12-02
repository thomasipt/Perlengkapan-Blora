VERSION 5.00
Begin VB.Form PASWORD 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAINTENANCE PASSWORD"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3225
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1215
      Width           =   2700
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3225
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   90
      Width           =   2700
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3225
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   780
      Width           =   2700
   End
   Begin VB.CommandButton cmdEDIT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ganti"
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
      Left            =   390
      TabIndex        =   3
      Top             =   1785
      Width           =   1080
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
      Left            =   4725
      TabIndex        =   4
      Top             =   1785
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      Height          =   1320
      Left            =   -4365
      ScaleHeight     =   1260
      ScaleWidth      =   14865
      TabIndex        =   5
      Top             =   1665
      Width           =   14925
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Konfirmasi Password"
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
      Left            =   1290
      TabIndex        =   8
      Top             =   1215
      Width           =   1785
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID User"
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
      Left            =   1290
      TabIndex        =   7
      Top             =   90
      Width           =   1785
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
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
      Left            =   1290
      TabIndex        =   6
      Top             =   780
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   90
      Picture         =   "PASWORD.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1320
   End
End
Attribute VB_Name = "PASWORD"
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
If Text2 <> Text3 Then
    MsgBox "PASSWORD TIDAK SAMA", vbCritical, "KONFIRMASI"
    Text2.SetFocus
    Exit Sub
End If

If Text1 <> Trim(Operator) Then
    MsgBox "ID USER SALAH", vbCritical, "KONFIRMASI"
    End
End If

TANYA = MsgBox("ANDA YAKIN AKAN MERUBAH PASSWORD", vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

Call GantiPassword
End Sub

Private Sub GantiPassword()
SSave = "Select * From PASSWORD where Nama = 'ADMINISTRATOR'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.Edit
    RSave("Password") = Trim(Text2)
    RSave("UserCode") = Trim(Text1)
    RSave("Jam") = Now
RSave.Update
RSave.Close
Set RSave = Nothing
End
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
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
