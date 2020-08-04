VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Terbilang Desimal Inggris"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAngka 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblTerbilang 
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function SpellDigit(strNumeric As Integer)
 Dim cRet As String
 On Error GoTo Pesan
 cRet = ""
 Select Case strNumeric
        Case 0:     cRet = " zero"
        Case 1:     cRet = " one"
        Case 2:     cRet = " two"
        Case 3:     cRet = " three"
        Case 4:     cRet = " four"
        Case 5:     cRet = " five"
        Case 6:     cRet = " six"
        Case 7:     cRet = " seven"
        Case 8:     cRet = " eight"
        Case 9:     cRet = " nine"
        Case 10:    cRet = " ten"
        Case 11:    cRet = " eleven"
        Case 12:    cRet = " twelve"
        Case 13:    cRet = " thirteen"
        Case 14:    cRet = " fourteen"
        Case 15:    cRet = " fifteen"
        Case 16:    cRet = " sixteen"
        Case 17:    cRet = " seventeen"
        Case 18:    cRet = " eighteen"
        Case 19:    cRet = " ninetieen"
        Case 20:    cRet = " twenty"
        Case 30:    cRet = " thirty"
        Case 40:    cRet = " fourthy"
        Case 50:    cRet = " fifty"
        Case 60:    cRet = " sixty"
        Case 70:    cRet = " seventy"
        Case 80:    cRet = " eighty"
        Case 90:    cRet = " ninety"
        Case 100:   cRet = " one hundred"
        Case 200:   cRet = " two hundred"
        Case 300:   cRet = " three hundred"
        Case 400:   cRet = " four hundred"
        Case 500:   cRet = " five hundred"
        Case 600:   cRet = " six hundred"
        Case 700:   cRet = " seven hundred"
        Case 800:   cRet = " eight hundred"
        Case 900:   cRet = " nine hundred"
 End Select
 SpellDigit = cRet
Exit Function
Pesan:
  SpellDigit = "(maksimal 9 digit)"
End Function

Private Function SpellUnit(strNumeric As Integer)
 Dim cRet As String
 Dim n100 As Integer
 Dim n10 As Integer
 Dim n1 As Integer
 On Error GoTo Pesan
 cRet = ""
 n100 = Int(strNumeric / 100) * 100
 n10 = Int((strNumeric - n100) / 10) * 10
 n1 = (strNumeric - n100 - n10)
 If n100 > 0 Then
    cRet = SpellDigit(n100)
 End If
 If n10 > 0 Then
    If n10 = 10 Then
       cRet = cRet & SpellDigit(n10 + n1)
    Else
       cRet = cRet & SpellDigit(n10)
    End If
 End If
 If n1 > 0 And n10 <> 10 Then
    cRet = cRet & SpellDigit(n1)
 End If
 SpellUnit = cRet
 Exit Function
Pesan:
  SpellUnit = "(maksimal 9 digit)"
End Function

Public Function TerbilangInggris(strNumeric As String) As String
 Dim cRet As String
 Dim n1000000 As Long
 Dim n1000 As Long
 Dim n1 As Integer
 Dim n0 As Integer
   On Error GoTo Pesan
   Dim strValid As String, huruf As String * 1
   Dim i As Integer
   'Periksa setiap karakter masukan
   strValid = "1234567890.,"
   For i% = 1 To Len(strNumeric)
     huruf = Chr(Asc(Mid(strNumeric, i%, 1)))
     If InStr(strValid, huruf) = 0 Then
       MsgBox "Harus karakter angka!", _
              vbCritical, "Karakter Tidak Valid"
       Exit Function
     End If
   Next i%
 
 If strNumeric = "" Then Exit Function
 If Len(Trim(strNumeric)) > 9 Then GoTo Pesan
 
 cRet = ""
 n1000000 = Int(strNumeric / 1000000) * 1000000
 n1000 = Int((strNumeric - n1000000) / 1000) * 1000
 n1 = Int(strNumeric - n1000000 - n1000)
 n0 = (strNumeric - n1000000 - n1000 - n1) * 100
 If n1000000 > 0 Then
    cRet = SpellUnit(n1000000 / 1000000) & " million"
 End If
 If n1000 > 0 Then
    cRet = cRet & SpellUnit(n1000 / 1000) & " thousand"
 End If
 If n1 > 0 Then
    cRet = cRet & SpellUnit(n1)
 End If
 If n0 > 0 Then
    cRet = cRet & " and cents" & SpellUnit(n0)
 End If
 TerbilangInggris = cRet & " only"
 Exit Function
Pesan:
  TerbilangInggris = "(maximum 9 digit)"
End Function


Private Sub txtAngka_Change()
   lblTerbilang.Caption = TerbilangInggris(txtAngka.Text)
End Sub



