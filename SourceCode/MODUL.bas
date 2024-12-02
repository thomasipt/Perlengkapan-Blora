Attribute VB_Name = "MODUL"
Global CR As ADODB.Connection
Public CN As String

Public INISIAL As String
Public TIPE As String
Public GRIDKLIK As String
Public GRIDKLIK2 As String
Public TANYA As String
Public GAMBAR As String
Public No, NoMax, NoMin As String
Public TOKET As String

Public KODELOKASIc As String
Public JENISLOKASIc As String
Public NAMALOKASIc As String
Public ALAMATc As String
Public PENGURUSc As String
Public KEPALA_SKPDc As String
Public SEMESTERc As String
Public TAHUNc As String

Public OYEN As String
Public NOVI As String

Public KodeC As String
Public Aplikasi As String
Public Operator As String
Public Super As String
Public Lolos As Integer
Public LolosOV As String
Public PesanOV As String
Public NomorNas As String
Public NoNas As String
Public NomorRek, NamaRek, NRek As String
Public Nobukti As String
Public CCab, NCab As String
Public Nourut As Boolean
Public Tglinput As String
Public NoBilyet As String
Public NoPinjaman As String
Public CodePinjaman As String
Public NomTrans As Currency
Public Status, User As String
Public CodeSl, NamaSl As String
Public Lunas As String
Public Sale As Currency

Sub Main()
CN = "DSN=PERLENGKAPAN;DRIVER={Microsoft Access Driver};Server=CENTRAL;UID= ;PWD= ;Database = PERLENGKAPAN.mdb;"
LOGON.Show
End Sub

Public Sub KoneksiPERLENGKAPAN()
If Status = "01" Then
    MAINMENU.Show
End If
End Sub

Public Sub ClearTextBoxes(frmClearMe As Form)
Dim txt As Control
For Each txt In frmClearMe
  If TypeOf txt Is TextBox Then txt.Text = ""
Next
End Sub

Public Sub ColorTextBoxes(frmColorMe As Form)
Dim txt As Control
For Each txt In frmColorMe
  If TypeOf txt Is TextBox Then txt.BackColor = &HC0E0FF
Next
End Sub

Public Function Digit(Panjang, Nilai As Double) As String
Dim Y, NilaiP As Double
Dim Kar, NilaiS As String

If Panjang <= 0 Then Panjang = 1

NilaiS = Trim(Str(Nilai))
NilaiP = Len(NilaiS)
If NilaiP >= Panjang Then Panjang = NilaiP

Kar = " "
For Y = 1 To Panjang
    Kar = Trim(Kar) + "0"
Next
If (Panjang - NilaiP) <= 0 Then
    Digit = NilaiS
Else
    Digit = Mid(Kar, 1, (Panjang - NilaiP)) + NilaiS
End If
End Function

Public Function Satuan(ByVal Nilai As Currency) As String
Select Case Nilai
    Case 1: Satuan = "SATU "
    Case 2: Satuan = "DUA "
    Case 3: Satuan = "TIGA "
    Case 4: Satuan = "EMPAT "
    Case 5: Satuan = "LIMA "
    Case 6: Satuan = "ENAM "
    Case 7: Satuan = "TUJUH "
    Case 8: Satuan = "DELAPAN "
    Case 9: Satuan = "SEMBILAN "
End Select
End Function

Public Function Ribuan(ByVal Bilangan As Currency) As String
Dim a, B As Currency
Dim C As String

C = ""
a = Bilangan \ 1000
B = Bilangan Mod 1000
If a > 1 Then C = C + Satuan(a) + "RIBU "
If a = 1 Then C = C + "SERIBU "

a = B \ 100
B = B Mod 100
If a > 1 Then C = C + Satuan(a) + "RATUS "
If a = 1 Then C = C + "SERATUS "

a = B \ 10
B = B Mod 10
If a > 1 Then C = C + Satuan(a) + "PULUH "
If a = 1 Then
    If B = 0 Then Ribuan = C + "SEPULUH "
    If B = 1 Then Ribuan = C + "SEBELAS "
    If B > 1 Then Ribuan = C + Satuan(B) + "BELAS "
Else
    Ribuan = C + Satuan(B)
End If
End Function

Public Function Terbilang(ByVal Bilangan As Currency) As String

Dim a, B As Currency
Dim C As String


a = Bilangan \ 1000000000
B = Bilangan Mod 1000000000
C = "#"
If a > 0 Then C = Ribuan(a) + "MILYAR "

a = B \ 1000000
B = B Mod 1000000
If a > 0 Then C = C + Ribuan(a) + "JUTA "

a = B \ 1000
B = B Mod 1000
If a > 1 Then C = C + Ribuan(a) + "RIBU "
If a = 1 Then C = C + "Seribu "
Terbilang = C + Ribuan(B) + "RUPIAH#"
End Function

Public Function SumHari(Dari, Ke As Date) As Integer
If Ke - Dari <= 1 Then
    SumHari = 1
Else
    SumHari = Ke - Dari
End If
End Function

Public Function Sisip(Kar As String, Posisi As Integer, Kar2 As String) As String
Dim Pj As Integer
Dim Akhir As String
Dim depan As String
Pj = Len(Kar)
If Len(Kar) < Len(Kar2) Then
    Sisip = Kar2
Else
    If Posisi = 1 Then Sisip = Kar2 + Mid(Kar, 2, Pj - 1)
    If Posisi > 1 And Posisi < Pj Then
        depan = Mid(Kar, 1, Posisi - 1)
        Akhir = Mid(Kar, Posisi + 1, Pj - Posisi)
        Sisip = depan + Kar2 + Akhir
    End If
    If Posisi = Pj Then Sisip = Mid(Kar, 1, Posisi - 1) + Kar2
End If
End Function

Public Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Public Function BulanStr(ByVal CBulan As Currency) As String
Select Case CBulan
    Case 1: BulanStr = " Jan. "
    Case 2: BulanStr = " Feb. "
    Case 3: BulanStr = " Mar. "
    Case 4: BulanStr = " Apr. "
    Case 5: BulanStr = " Mei "
    Case 6: BulanStr = " Juni "
    Case 7: BulanStr = " Juli "
    Case 8: BulanStr = " Agt. "
    Case 9: BulanStr = " Sept. "
    Case 10: BulanStr = " Okt. "
    Case 11: BulanStr = " Nov. "
    Case 12: BulanStr = " Des. "
End Select
BulanStr = BulanStr
End Function
Public Function BulanString(ByVal CBulan As Currency) As String
Select Case CBulan
    Case 1: BulanString = " JANUARI "
    Case 2: BulanString = " FEBRUARI "
    Case 3: BulanString = " MARET "
    Case 4: BulanString = " APRIL "
    Case 5: BulanString = " MEI "
    Case 6: BulanString = " JUNI "
    Case 7: BulanString = " JULI "
    Case 8: BulanString = " AGUSTUS "
    Case 9: BulanString = " SEPTEMBER "
    Case 10: BulanString = " OKTOBER "
    Case 11: BulanString = " NOVEMBER "
    Case 12: BulanString = " DESEMBER "
End Select
BulanString = BulanString
End Function


Public Function HariStr(ByVal CHari As Currency) As String
Select Case CHari
    Case 1: HariStr = " MINGGU "
    Case 2: HariStr = " SENIN "
    Case 3: HariStr = " SELASA "
    Case 4: HariStr = " RABU "
    Case 5: HariStr = " KAMIS "
    Case 6: HariStr = " JUMAT "
    Case 7: HariStr = " SABTU "
End Select
HariStr = HariStr
End Function

Public Function BlkKoma(Bilangan As Double) As String
Dim a, D As Double
Dim B, E, f As Double
Dim C As String
If Bilangan > 2000000000 Then
    C = ""
    
    D = Mid(Bilangan, 1, 7)
    a = D \ 1000000
    B = D Mod 1000000
    If a > 0 Then C = Ribuan(a) + "Milyar "
    
    E = Mid(Bilangan, 2, 10)
    a = E \ 1000000
    B = E Mod 1000000
    If a > 0 Then C = C + Ribuan(a) + "Juta "
    
    f = Mid(Bilangan, 5, 10)
    a = f \ 1000
    B = f Mod 1000
    If a > 0 Then C = C + Ribuan(a) + "Ribu "
    If a = 1 Then C = C + "Seribu "
BlkKoma = C + Ribuan(B)
Else
    C = ""
    a = Bilangan \ 1000000000
    B = Bilangan Mod 1000000000
    If a > 0 Then C = Ribuan(a) + "Milyar "
    
    a = B \ 1000000
    B = B Mod 1000000
    If a > 0 Then C = C + Ribuan(a) + "Juta "
    
    a = B \ 1000
    B = B Mod 1000
    If a > 1 Then C = C + Ribuan(a) + "Ribu "
    If a = 1 Then C = C + "Seribu "
BlkKoma = C + Ribuan(B)
End If
End Function


