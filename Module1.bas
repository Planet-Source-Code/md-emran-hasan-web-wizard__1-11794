Attribute VB_Name = "Module1"
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList _
As Long, ByVal lpBuffer As String) As Long
Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long
Public Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
'Public Function LongToRGB(Value As Long, _
 '   Optional chooseDelimiter As String = ",") As String

'Dim Blue As Double, Green As Double, Red As Double
'Dim BlueS As Double, GreenS As Double, RGBs As String

 '   Blue = Fix((Value / 256) / 256)
'    BlueS = (Blue * 256) * 256
'    Green = Fix((Value - BlueS) / 256)
'    GreenS = Green * 256
'    Red = Fix(Value - BlueS - GreenS)
'    RGBs = (Red & chooseDelimiter & Green & chooseDelimiter & Blue)
'    LongToRGB = RGBs
'
'End Function
'Public Function RGB2HTMLColor(R As Byte, G As Byte, _
'   B As Byte) As String'
'
'Dim HexR, HexB, HexG As Variant
'Dim sTemp As String'
'
'On Error GoTo ErrorHandler

 'R
 'HexR = Hex(R)
 'If Len(HexR) < 2 Then HexR = "0" & HexR
 
 'Get Green Hex
' HexG = Hex(G)
'If Len(HexG) < 2 Then HexG = "0" & HexG

'HexB = Hex(B)
'If Len(HexB) < 2 Then HexB = "0" & HexB



'    RGB2HTMLColor = "#" & HexR & HexG & HexB
'ErrorHandler:
'End Function

Public Function RGBHexColor(lngColor As Long) As String
Dim sHex As String
sHex = Hex(lngColor)
sHex = Right$("000000" & sHex, 6)
sHex = Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
RGBHexColor = sHex
End Function
Public Sub Body()
Dim bg, txt, lnk, vis, BgImage, bgSound
Dim title, head, metaKey, metaDes
Dim hoverColor, h1, h3, h5
Dim h1Code, h2Code, h3Code
Dim isH1, isH3, isH5, isUnder
Dim dontUnder

'set all to 0 that means no
isH1 = 0
isH3 = 0
isH5 = 0
isUnder = 0

'extract the color codes
bg = RGBHexColor(f1.shpBg.FillColor)
txt = RGBHexColor(f1.shpTxt.FillColor)
lnk = RGBHexColor(f1.shpLnk.FillColor)
vis = RGBHexColor(f1.shpVis.FillColor)

'extracts the title,head and META tags information
title = f1.txtTitle.Text
head = f1.txtHead.Text
metaKey = f1.txtMetaKeywords.Text
metaDes = f1.txtMetaDescription.Text

'set the style sheet
hoverColor = RGBHexColor(f1.shpHoverHover.FillColor)
h1 = RGBHexColor(f1.shpH1.FillColor)
h3 = RGBHexColor(f1.shpH3.FillColor)
h5 = RGBHexColor(f1.shpH5.FillColor)

'check if any H1 or H3 or H5 checkbox is checked or not
'if checked, then generate the code
'else don't bother about it !
If f1.chkH1.Value = 1 Then
 If f1.chkH1Bold.Value = 1 Then
  isH1 = 1
  h1Code = "H1 {COLOR: " & h1 & "; FONT-WEIGHT: Bold; }"
 Else
  isH1 = 1
  h1Code = "H1 {COLOR: " & h1 & "; FONT-WEIGHT: none; }"
 End If
Else
 isH1 = 0
End If

If f1.chkH3.Value = 1 Then
 If f1.chkH3Bold.Value = 1 Then
  isH3 = 1
  h3Code = "H2 {COLOR: " & h3 & "; FONT-WEIGHT: Bold; }" & "H3 {COLOR: " & h3 & "; FONT-WEIGHT: Bold; }"
 Else
  isH3 = 1
  h3Code = "H2 {COLOR: " & h3 & "; FONT-WEIGHT: none; }" & "H3 {COLOR: " & h3 & "; FONT-WEIGHT: none; }"
 End If
Else
 isH3 = 0
End If

If f1.chkH5.Value = 1 Then
 If f1.chkH5Bold.Value = 1 Then
  isH5 = 1
  h5code = "H4 {COLOR: " & h5 & "; FONT-WEIGHT: Bold ; }" & "H5 {COLOR: " & h3 & "; FONT-WEIGHT: Bold; }"
 Else
  isH5 = 1
  h5code = "H4 {COLOR: " & h5 & "; FONT-WEIGHT: none; }" & "H5 {COLOR: " & h3 & "; FONT-WEIGHT: none; }"
 End If
Else
 isH5 = 0
End If

'check if the Don't Underline is checked or not
'if checked,then generate the code
'else don't mind that
If f1.chkDontUnder.Value = 1 Then
 isUnder = 1
 dontUnder = "A {TEXT-DECORATION: none}"
Else
 isUnder = 0
 dontUnder = ""
End If

'set the main page's source
f1.Text1.Text = "<HTML><HEAD>" & "<TITLE>" & title & "</TITLE></HEAD>"
f1.Text1.Text = f1.Text1.Text & "<META content = """ & f1.txtMetaDescription.Text & """ name=Description>"
f1.Text1.Text = f1.Text1.Text & "<META content = """ & f1.txtMetaKeywords.Text & """ name=keywords>"
f1.Text1.Text = f1.Text1.Text & "<STYLE type=text/css>"

'start checking the status of H1,H2,H3
If isH1 = 0 And isH3 = 0 And isH5 = 0 Then

ElseIf isH1 = 0 And isH3 = 0 And isH5 = 1 Then
f1.Text1.Text = f1.Text1.Text & h5code

ElseIf isH1 = 0 And isH3 = 1 And isH5 = 0 Then
f1.Text1.Text = f1.Text1.Text & h3Code

ElseIf isH1 = 0 And isH3 = 1 And isH5 = 1 Then
f1.Text1.Text = f1.Text1.Text & h3Code & " " & h5code

ElseIf isH1 = 1 And isH3 = 0 And isH5 = 0 Then
f1.Text1.Text = f1.Text1.Text & h1Code

ElseIf isH1 = 1 And isH3 = 0 And isH5 = 1 Then
f1.Text1.Text = f1.Text1.Text & h1Code & " " & h5code

ElseIf isH1 = 1 And isH3 = 1 And isH5 = 0 Then
f1.Text1.Text = f1.Text1.Text & h1Code & " " & h3Code

ElseIf isH1 = 1 And isH3 = 1 And isH5 = 1 Then
f1.Text1.Text = f1.Text1.Text & h1Code & " " & h3Code & " " & h5code

End If

'check the status of Don't Underline
If isUnder = 1 Then
f1.Text1.Text = f1.Text1.Text & " " & dontUnder
ElseIf isUnder = 0 Then

End If

f1.Text1.Text = f1.Text1.Text & "</STYLE>"

'generate the body code
If f1.Check1.Value = 1 And f1.txtBgImage.Text = "" Then
f1.Text1.Text = f1.Text1.Text & " " & "<BODY bgcolor=#" & bg & " text=#" & txt & " link=#" & lnk & " vlink=#" & vis & " >"

ElseIf f1.Check1.Value = 1 And f1.txtBgImage.Text <> "" Then
BgImage = f1.txtBgImage.Text
f1.Text1.Text = f1.Text1.Text & " " & "<BODY background=""" & BgImage & """ text=#" & txt & " link=#" & lnk & " vlink=#" & vis & " >"

Else
f1.Text1.Text = f1.Text1.Text & " " & "<BODY bgcolor=#" & bg & " text=#" & txt & " link=#" & lnk & " vlink=#" & vis & " >"

End If
End Sub

Public Sub Text()
Dim TextText
TextText = f1.txtText.Text
f1.Text1.Text = f1.Text1.Text & TextText
End Sub

Public Sub Contact()
Dim email, fax, phone, street
Dim lastEdit

street = f1.txtStreet.Text
phone = f1.txtPhone.Text
fax = f1.txtFax.Text
email = f1.txtEmail.Text
lastEdit = Date

f1.Text1.Text = f1.Text1.Text & "<br><br>"
f1.Text1.Text = f1.Text1.Text & "<hr size=5 noshade>"
f1.Text1.Text = f1.Text1.Text & "<center>"
f1.Text1.Text = f1.Text1.Text & "Copyright (c)" & f1.txtHead & "All rights reserved.<br>"
f1.Text1.Text = f1.Text1.Text & street & "<br>"
f1.Text1.Text = f1.Text1.Text & phone & " " & fax & " " & email
f1.Text1.Text = f1.Text1.Text & "</center>"

End Sub

