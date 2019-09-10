VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nostalji v2.0"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   3090
   ClientWidth     =   5205
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   4031.951
   ScaleMode       =   0  'User
   ScaleWidth      =   3343.609
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox roket 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   1320
      Picture         =   "frmMain.frx":3EFBC
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox roket 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   2040
      Picture         =   "frmMain.frx":3FBFE
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox roket 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2760
      Picture         =   "frmMain.frx":40840
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox roket 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   3480
      Picture         =   "frmMain.frx":41482
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3480
      Picture         =   "frmMain.frx":420C4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   2760
      Picture         =   "frmMain.frx":42D06
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   2040
      Picture         =   "frmMain.frx":43948
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   1320
      Picture         =   "frmMain.frx":4458A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBoom 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":451CC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B&aþlat"
      Height          =   315
      Left            =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Küny&e"
      Height          =   340
      Left            =   1920
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1275
   End
   Begin VB.PictureBox roket 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "frmMain.frx":45E0E
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   1320
      Picture         =   "frmMain.frx":46250
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1320
      Picture         =   "frmMain.frx":46E92
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1320
      Picture         =   "frmMain.frx":47AD4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   1320
      Picture         =   "frmMain.frx":48716
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   2040
      Picture         =   "frmMain.frx":49358
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   2040
      Picture         =   "frmMain.frx":49F9A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2040
      Picture         =   "frmMain.frx":4ABDC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2040
      Picture         =   "frmMain.frx":4B81E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   2760
      Picture         =   "frmMain.frx":4C460
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   2760
      Picture         =   "frmMain.frx":4D0A2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2760
      Picture         =   "frmMain.frx":4DCE4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2760
      Picture         =   "frmMain.frx":4E926
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   3480
      Picture         =   "frmMain.frx":4F568
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   3480
      Picture         =   "frmMain.frx":501AA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   3480
      Picture         =   "frmMain.frx":50DEC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox gemi4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   3480
      Picture         =   "frmMain.frx":51A2E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Çý&kýþ"
      Height          =   340
      Left            =   3600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4440
      Top             =   1680
   End
   Begin VB.PictureBox roket 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   4200
      Picture         =   "frmMain.frx":52670
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox bayrak1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      Picture         =   "frmMain.frx":52AB2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox bayrak2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4200
      Picture         =   "frmMain.frx":52EF4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   114.987
      X2              =   3199.072
      Y1              =   3304.878
      Y2              =   3304.878
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   115.629
      X2              =   3199.715
      Y1              =   3354.451
      Y2              =   3354.451
   End
   Begin VB.Label life 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   450
      Left            =   600
      TabIndex        =   23
      Top             =   435
      Width           =   255
   End
   Begin VB.Label bonus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4170
      TabIndex        =   22
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Sol = 52              'Sol yön tuþu
Const Sag = 54              'Sað yön tuþu

Dim RoketYeri As Integer    'Roketin konumu 0..1.2.3.4..5
Dim Bayrak As Integer       'Bayraðýn yeri 1 sol | 2 sað
Dim Yay As Integer          'Fýrlatma bandý
Dim Puan As Integer         'Puan durumu (sayac 10-bayrak 50)
Dim can As Integer          'Hak
Dim speed As Integer        'oyunun akýþ hýzý (ms)
Dim gemiler(1 To 4) As Integer


Private Sub BayrakSondur()
    Select Case Bayrak
           Case 1
                Bayrak = 2
                Kenar (2)
                RoketYeriKapa (1)
                RoketYeri = 0
                RoketYeriAc (0)
           Case 2
                Bayrak = 1
                Kenar (1)
                RoketYeriKapa (4)
                RoketYeri = 5
                RoketYeriAc (5)
    End Select
End Sub

Private Sub Command1_Click()
'flip & flop
    Timer1.Enabled = Not Timer1.Enabled
    If Timer1.Enabled Then
       Command1.Caption = "B&aþlat"
    Else
       Command1.Caption = "Durduruldu"
    End If
End Sub

Private Sub Command2_Click()
    frmKunye.Show 1
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case Sol
             If RoketYeri > 1 Then
                RoketYeriKapa (RoketYeri)
                RoketYeri = RoketYeri - 1
                RoketYeriAc (RoketYeri)
             ElseIf RoketYeri = 1 And Bayrak = 1 Then
                Puan = Puan + 50
                RoketYeri = 0
                BayrakSondur
             End If
           Case Sag
             If RoketYeri < 4 Then
                RoketYeriKapa (RoketYeri)
                RoketYeri = RoketYeri + 1
                RoketYeriAc (RoketYeri)
             ElseIf RoketYeri = 4 And Bayrak = 2 Then
                Puan = Puan + 50
                RoketYeri = 5
                BayrakSondur
             End If
    End Select
End Sub

Private Sub Form_Load()
   'Formu Ortala
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.

   'Ýlk Deðerler
    RoketYeri = 0
    Bayrak = 2
    can = 5
    Yay = 1
    Puan = 0
   'Ýlk Atamalar
    Randomize
    RoketYeriAc (RoketYeri)
    Kenar (Bayrak)
End Sub

Private Sub Kenar(yer As Integer)
    Select Case yer
           Case 0: bayrak1.Visible = False: bayrak2.Visible = False
           Case 1: bayrak1.Visible = True:  bayrak2.Visible = False
           Case 2: bayrak1.Visible = False: bayrak2.Visible = True
    End Select
End Sub

Private Sub rip()
    
    RoketYeriKapa (RoketYeri)
    
    can = can - 1
    life = can

    'Select Case Bayrak
    '       Case 1: RoketYeri = 0
    '       Case 2: RoketYeri = 5
    'End Select
   
    Bayrak = 2
    Kenar (Bayrak)

    RoketYeri = 0
    RoketYeriAc (RoketYeri)

    For sayac = 1 To 5
        gemi1(sayac).Visible = False
        gemi2(sayac).Visible = False
        gemi3(sayac).Visible = False
        gemi4(sayac).Visible = False
    Next
    'picBoom.Visible = islev
    
    If can = 0 Then
       MsgBox "Oyun Bitti", 0, "Toplam " & Puan & " puan yaptýn..."
       Form_Load
       Timer1.Enabled = False
    End If
End Sub

Private Sub RoketYeriAc(yer As Integer)
    roket(yer).Visible = True
End Sub

Private Sub RoketYeriKapa(yer As Integer)
    roket(yer).Visible = False
End Sub

Private Sub Timer1_Timer()
   'Puan +++
    Puan = Puan + 10
    bonus = Puan

   'Yay Mekanizmasý \/
    If RoketYeri = 0 Xor RoketYeri = 5 Then Yay = Yay + 1 Else Yay = 1
    If Yay = 4 Then
       If RoketYeri = 0 Then Form_KeyPress (Sag)
       If RoketYeri = 5 Then Form_KeyPress (Sol)
    End If

   'Gemi Ýndirme \/
    For sayac = 4 To 1 Step -1
        gemi1(sayac + 1).Visible = gemi1(sayac).Visible
        gemi2(sayac + 1).Visible = gemi2(sayac).Visible
        gemi3(sayac + 1).Visible = gemi3(sayac).Visible
        gemi4(sayac + 1).Visible = gemi4(sayac).Visible
    Next

    If gemi1(5).Visible And RoketYeri = 1 Then rip
    If gemi2(5).Visible And RoketYeri = 2 Then rip
    If gemi3(5).Visible And RoketYeri = 3 Then rip
    If gemi4(5).Visible And RoketYeri = 4 Then rip

   'Gemi Yerleþtirme +++  oran 3/1  -  5/2 yapýlcak
    gemiler(1) = Fix(Rnd * 3) \ 2
    gemiler(2) = Fix(Rnd * 3) \ 2
    gemiler(3) = Fix(Rnd * 3) \ 2
    gemiler(4) = Fix(Rnd * 3) \ 2

   'Gemi Hesabý \/
    For sayac = 1 To 4
        toplam = toplam + gemiler(sayac)
    Next

    If toplam = 4 Then
        gemiler(Int(Rnd * 4 + 1)) = 0
    End If

   'Yeni gemiler
    If gemiler(1) = 1 Then gemi1(1).Visible = True Else gemi1(1).Visible = False
    If gemiler(2) = 1 Then gemi2(1).Visible = True Else gemi2(1).Visible = False
    If gemiler(3) = 1 Then gemi3(1).Visible = True Else gemi3(1).Visible = False
    If gemiler(4) = 1 Then gemi4(1).Visible = True Else gemi4(1).Visible = False
End Sub


Private Sub tmrFF_Timer()

End Sub
