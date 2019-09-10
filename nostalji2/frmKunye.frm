VERSION 5.00
Begin VB.Form frmKunye 
   BackColor       =   &H008080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nostalji v2.0 - Künye"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00404040&
      Height          =   1065
      ItemData        =   "frmKunye.frx":0000
      Left            =   240
      List            =   "frmKunye.frx":0002
      TabIndex        =   5
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   240
      Picture         =   "frmKunye.frx":0004
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      ToolTipText     =   "Bana benziyo diye koydum baþka bir niyetim yoktu..."
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "elmek6@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1545
      MousePointer    =   2  'Cross
      TabIndex        =   4
      ToolTipText     =   "Bana e-mail atmak isterseniz, bir kez click'leyiniz."
      Top             =   3000
      Width           =   1725
   End
   Begin VB.Label lblIcerik 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmKunye.frx":0C46
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lblBaslik 
      BackStyle       =   0  'Transparent
      Caption         =   "Nostalji v2.0  05.02.2001, 18:12 build a3 - Freeware - (R) (C) 1994-2001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmKunye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    List1.AddItem "   Bu program visual basic 5.0 ile yazýlmýþtýr."
    List1.AddItem "programýn kodu oldukça amatörcedir vede herhangi bir"
    List1.AddItem "(yazým) standart(ýn)a uyulmamýþtýr."
    List1.AddItem "   Programý istediðiniz gibi deðiþtirebilirsiniz ama"
    List1.AddItem "yazarýn ismine lütfen saygý duyununuz....."
    List1.AddItem ""
    List1.AddItem "Programýn Oynanýþý"
    List1.AddItem "------------------"
    List1.AddItem "Çok basit!!!"
    List1.AddItem ""
    List1.AddItem "Strateji ve Ýpuçlarý"
    List1.AddItem "---------------------"
    List1.AddItem ""
End Sub

