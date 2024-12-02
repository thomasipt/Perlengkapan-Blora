VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ZBACKUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BACKUP DATABASE CABANG"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   3690
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   6795
      Width           =   1170
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   135
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   6705
      Width           =   1170
   End
   Begin VB.CommandButton cmdSAVE2 
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
      Left            =   240
      TabIndex        =   6
      Top             =   5318
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4665
      TabIndex        =   7
      Top             =   5325
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1762
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   165
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1762
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   585
      Width           =   4005
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   1785
      Left            =   1762
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "ZBACKUP.frx":0000
      Top             =   1425
      Width           =   4005
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1762
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   1005
      Width           =   4005
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   1762
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "Text6"
      Top             =   3345
      Width           =   1170
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   1762
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "Text7"
      Top             =   3765
      Width           =   1170
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   6090
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3466
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3466
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3466
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
      Height          =   735
      Left            =   -4470
      ScaleHeight     =   675
      ScaleWidth      =   14865
      TabIndex        =   15
      Top             =   5205
      Width           =   14925
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   52
      TabIndex        =   19
      Top             =   4680
      Width           =   5880
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   52
      TabIndex        =   17
      Top             =   4275
      Width           =   5880
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
      Left            =   195
      TabIndex        =   14
      Top             =   165
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
      Left            =   195
      TabIndex        =   13
      Top             =   3765
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
      Left            =   195
      TabIndex        =   12
      Top             =   3345
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
      Left            =   195
      TabIndex        =   11
      Top             =   585
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
      Left            =   195
      TabIndex        =   10
      Top             =   1005
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
      Left            =   195
      TabIndex        =   9
      Top             =   2160
      Width           =   1380
   End
End
Attribute VB_Name = "ZBACKUP"
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

Private Sub cmdSAVE2_Click()
SSave = "Select * From KONFIG_DATA where KODELOKASI='" + Trim(KODELOKASIc) + "' and SEMESTER='" + Trim(SEMESTERc) + "' and TAHUN='" + Trim(TAHUNc) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.Edit
    RSave("STATUS") = "1"
    RSave("TGL") = Now
RSave.Update
RSave.Close
Set RSave = Nothing

        ReturnValue = Shell("C:\Program Files\PERLENGKAPAN\BACKUP.EXE", 1)
        AppActivate ReturnValue
        End

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Label4 = ""
ClearTextBoxes Me
Call ISIDATA

If Text5 = "1" Then
    Label4 = "BACKUP TERAKHIR TANGGAL"
    Label8 = Trim(Text8)
Else
    Label4 = ""
    Label8 = ""
End If

End Sub

Private Sub ISIDATA()
SKNG = "Select * From KONFIG_DATA where KODELOKASI='" + Trim(KODELOKASIc) + "' and SEMESTER='" + Trim(SEMESTERc) + "' and TAHUN='" + Trim(TAHUNc) + "'"
Set RKNG = RDCO.OpenResultset(SKNG, rdOpenKeyset, rdConcurRowVer)
If RKNG.RowCount <> 0 Then
    Text1 = Format(RKNG("KODELOKASI"), ">")
    Text2 = Format(RKNG("JENISLOKASI"), ">")
    Text4 = Format(RKNG("NAMALOKASI"), ">")
    Text3 = Format(RKNG("ALAMAT"), ">")
    Text6 = Format(RKNG("SEMESTER"), ">")
    Text7 = RKNG("TAHUN")
    Text5 = RKNG("STATUS")
    Text8 = RKNG("TGL")
End If
RKNG.Close
Set RKNG = Nothing
End Sub

