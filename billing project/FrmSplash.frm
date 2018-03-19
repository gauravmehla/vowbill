VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   4740
   ClientTop       =   1500
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   Picture         =   "FrmSplash.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   195
      Top             =   270
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(C) The Vowers ltd. , 2012-14"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1500
      TabIndex        =   3
      Top             =   2820
      Width           =   2790
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is Lienceced to : Kitchen Castle 17"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   660
      TabIndex        =   2
      Top             =   3150
      Width           =   4485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version : 1.0.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VowBill"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   915
      Left            =   22
      Top             =   2670
      Width           =   5760
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Mr. Atanu Maity
'          Date : 21-Aug-2006
'*************************************
'             Splash Screen
'      Used Table : NA
'Module to show startup screen
'*************************************

Option Explicit

Dim r As Integer
Dim i As Integer
Private Sub Form_Load()
    '>>> center the form
    Me.Left = (Screen.Height - Me.Height)
    Me.Top = (Screen.Width - Me.Width) / 4
    
    '>>> get a random value to decide how many seconds
    '>>> startup screen should be displayed
    r = Rnd * 5 + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '>>> release all the references
    Set FrmSplash = Nothing
End Sub

Private Sub Timer1_Timer()
    '>>> check the ellapsed time
    '>>> if the ellapsed time greater then random value
    '>>> stored in form load, stop the timer
    '>>> show main from and close the startup screen
    i = i + 1
    If r >= i Then
        i = 0
        Timer1.Interval = 0
        Unload Me

        Load FrmMain
        FrmMain.Show
    End If
End Sub

