VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()
'===================================================================================================='
'Parent'
n1 = CreateRoundRectRgn(20, 320, 55, 330, 0, 0) 'Garis Atas'
'===================================================================================================='
'Huruf H'
n2 = CreateRoundRectRgn(20, 320, 35, 400, 0, 0) 'Garis Kiri'
n3 = CreateRoundRectRgn(20, 360, 70, 370, 0, 0) 'Garis Tengah'
n4 = CreateRoundRectRgn(60, 320, 75, 400, 0, 0) 'Garis Kanan'

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn n1, n2, n2, 2
CombineRgn n1, n1, n3, 2
CombineRgn n1, n1, n4, 2
'===================================================================================================='

'Huruf E'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n5 = CreateRoundRectRgn(80, 320, 125, 330, 0, 0) 'Garis Atas'
n6 = CreateRoundRectRgn(80, 360, 125, 370, 0, 0) 'Garis Tengah'
n7 = CreateRoundRectRgn(85, 320, 100, 400, 0, 0) 'Garis Kiri'
n9 = CreateRoundRectRgn(80, 390, 125, 400, 0, 0) 'Garis Bawah'

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn n1, n1, n5, 2
CombineRgn n1, n1, n6, 2
CombineRgn n1, n1, n7, 2
CombineRgn n1, n1, n8, 2
CombineRgn n1, n1, n9, 2

'===================================================================================================='

'Huruf R'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n10 = CreateRoundRectRgn(130, 320, 150, 400, 0, 0) 'Garis Kiri'
n11 = CreateRoundRectRgn(180, 360, 130, 370, 0, 0) 'Garis Tengah'
n12 = CreateRoundRectRgn(185, 320, 145, 330, 0, 0) 'Garis Atas'
n13 = CreateRoundRectRgn(175, 320, 190, 370, 0, 0) 'Garis Kanan'

n14 = CreateRoundRectRgn(145, 385, 165, 360, 0, 0) 'Garis Kanan'
n15 = CreateRoundRectRgn(160, 375, 180, 390, 0, 0) 'Garis Kanan'
n16 = CreateRoundRectRgn(175, 380, 190, 400, 0, 0) 'Garis Kanan'

CombineRgn n1, n1, n10, 2
CombineRgn n1, n1, n11, 2
CombineRgn n1, n1, n12, 2
CombineRgn n1, n1, n13, 2
CombineRgn n1, n1, n14, 2
CombineRgn n1, n1, n15, 2
CombineRgn n1, n1, n16, 2
'===================================================================================================='

'Huruf I'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n17 = CreateRoundRectRgn(195, 320, 210, 400, 0, 0)

CombineRgn n1, n1, n17, 2

'===================================================================================================='

'Huruf K'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n18 = CreateRoundRectRgn(260, 320, 245, 400, 0, 0)
n19 = CreateRoundRectRgn(270, 350, 255, 370, 0, 0)
n20 = CreateRoundRectRgn(268, 340, 290, 355, 0, 0)
n21 = CreateRoundRectRgn(268, 385, 290, 370, 0, 0)
n22 = CreateRoundRectRgn(285, 320, 300, 345, 0, 0)
n23 = CreateRoundRectRgn(285, 380, 300, 400, 0, 0)

CombineRgn n1, n1, n18, 2
CombineRgn n1, n1, n19, 2
CombineRgn n1, n1, n20, 2
CombineRgn n1, n1, n21, 2
CombineRgn n1, n1, n22, 2
CombineRgn n1, n1, n23, 2

'===================================================================================================='
'Huruf H'
n24 = CreateRoundRectRgn(305, 320, 320, 400, 0, 0) 'Garis Kiri'
n25 = CreateRoundRectRgn(305, 360, 355, 370, 0, 0) 'Garis Tengah'
n26 = CreateRoundRectRgn(345, 320, 360, 400, 0, 0) 'Garis Kanan'

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn n1, n1, n24, 2
CombineRgn n1, n1, n25, 2
CombineRgn n1, n1, n26, 2

'===================================================================================================='

'Huruf A'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n27 = CreateRoundRectRgn(365, 320, 405, 330, 0, 0) 'Garis Atas'
n28 = CreateRoundRectRgn(365, 350, 405, 360, 0, 0) 'Garis Tengah'
n29 = CreateRoundRectRgn(365, 320, 380, 400, 0, 0) 'Garis Kiri'
n30 = CreateRoundRectRgn(400, 320, 405, 400, 0, 0) 'Garis Kanan Luar'
n31 = CreateRoundRectRgn(405, 325, 410, 390, 0, 0) 'Garis Kanan Dalam'

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn n1, n1, n27, 2
CombineRgn n1, n1, n28, 2
CombineRgn n1, n1, n29, 2
CombineRgn n1, n1, n30, 2
CombineRgn n1, n1, n31, 2

'===================================================================================================='

'Huruf R'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n32 = CreateRoundRectRgn(415, 320, 430, 400, 0, 0) 'Garis Kiri'
n33 = CreateRoundRectRgn(420, 360, 460, 370, 0, 0) 'Garis Tengah'
n34 = CreateRoundRectRgn(420, 320, 460, 330, 0, 0) 'Garis Atas'
n35 = CreateRoundRectRgn(455, 320, 470, 370, 0, 0) 'Garis Kanan'

n36 = CreateRoundRectRgn(425, 380, 445, 360, 0, 0) 'Garis Kanan'
n37 = CreateRoundRectRgn(440, 375, 460, 390, 0, 0) 'Garis Kanan'
n38 = CreateRoundRectRgn(455, 380, 470, 400, 0, 0) 'Garis Kanan'

CombineRgn n1, n1, n32, 2
CombineRgn n1, n1, n33, 2
CombineRgn n1, n1, n34, 2
CombineRgn n1, n1, n35, 2
CombineRgn n1, n1, n36, 2
CombineRgn n1, n1, n37, 2
CombineRgn n1, n1, n38, 2

'===================================================================================================='

'Huruf I'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n39 = CreateRoundRectRgn(480, 320, 495, 400, 0, 0)

CombineRgn n1, n1, n39, 2

'===================================================================================================='

'Huruf O'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n40 = CreateRoundRectRgn(505, 350, 520, 370, 0, 0) 'garis kiri'
n41 = CreateRoundRectRgn(520, 370, 535, 390, 0, 0) 'garis kiri bawah'
n42 = CreateRoundRectRgn(520, 350, 535, 330, 0, 0) 'garis kiri atas'
n43 = CreateRoundRectRgn(535, 380, 555, 400, 0, 0) 'garis tengah bawah'
n44 = CreateRoundRectRgn(535, 340, 555, 320, 0, 0) 'garis tengah atas'
n45 = CreateRoundRectRgn(555, 370, 570, 390, 0, 0) 'garis kanan bawah'
n46 = CreateRoundRectRgn(555, 350, 570, 330, 0, 0) 'garis kanan atas'
n47 = CreateRoundRectRgn(570, 350, 585, 370, 0, 0) 'garis kanan'

CombineRgn n1, n1, n40, 2
CombineRgn n1, n1, n41, 2
CombineRgn n1, n1, n42, 2
CombineRgn n1, n1, n43, 2
CombineRgn n1, n1, n44, 2
CombineRgn n1, n1, n45, 2
CombineRgn n1, n1, n46, 2
CombineRgn n1, n1, n47, 2

'===================================================================================================='

'Huruf N'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n39 = CreateRoundRectRgn(595, 320, 610, 400, 0, 0) 'garis kiri'
n40 = CreateRoundRectRgn(605, 330, 620, 345, 0, 0)
n41 = CreateRoundRectRgn(620, 345, 635, 360, 0, 0)
n42 = CreateRoundRectRgn(635, 360, 650, 375, 0, 0)
n43 = CreateRoundRectRgn(650, 375, 665, 390, 0, 0)
n44 = CreateRoundRectRgn(660, 320, 675, 400, 0, 0) 'garis kanan'

CombineRgn n1, n1, n39, 2
CombineRgn n1, n1, n40, 2
CombineRgn n1, n1, n41, 2
CombineRgn n1, n1, n42, 2
CombineRgn n1, n1, n43, 2
CombineRgn n1, n1, n44, 2

'===================================================================================================='

'Huruf O'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n45 = CreateRoundRectRgn(685, 350, 700, 370, 0, 0) 'garis kiri'
n46 = CreateRoundRectRgn(700, 370, 715, 390, 0, 0) 'garis kiri bawah'
n47 = CreateRoundRectRgn(700, 350, 715, 330, 0, 0) 'garis kiri atas'
n48 = CreateRoundRectRgn(715, 380, 735, 400, 0, 0) 'garis tengah bawah'
n49 = CreateRoundRectRgn(715, 340, 735, 320, 0, 0) 'garis tengah atas'
n50 = CreateRoundRectRgn(735, 370, 750, 390, 0, 0) 'garis kanan bawah'
n51 = CreateRoundRectRgn(735, 350, 750, 330, 0, 0) 'garis kanan atas'
n52 = CreateRoundRectRgn(750, 350, 765, 370, 0, 0) 'garis kanan'

CombineRgn n1, n1, n45, 2
CombineRgn n1, n1, n46, 2
CombineRgn n1, n1, n47, 2
CombineRgn n1, n1, n48, 2
CombineRgn n1, n1, n49, 2
CombineRgn n1, n1, n50, 2
CombineRgn n1, n1, n51, 2
CombineRgn n1, n1, n52, 2

SetWindowRgn Form1.hwnd, n1, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Form1.hwnd, &HA1, 2, 0&
End Sub
