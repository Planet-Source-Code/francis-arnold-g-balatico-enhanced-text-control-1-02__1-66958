VERSION 5.00
Begin VB.Form frmSample 
   BackColor       =   &H00F3DECF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enhanced Text Control 1.02"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMultiline 
      Caption         =   "Multiline / Single Line"
      Height          =   420
      Left            =   135
      TabIndex        =   5
      Top             =   1800
      Width           =   1980
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable / Enable"
      Height          =   420
      Left            =   2865
      TabIndex        =   4
      Top             =   1305
      Width           =   1980
   End
   Begin VB.CommandButton cmdGroove 
      Caption         =   "Change Groove Color"
      Height          =   420
      Left            =   135
      TabIndex        =   3
      Top             =   1320
      Width           =   1980
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   4935
      TabIndex        =   1
      Top             =   0
      Width           =   4965
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSample.frx":0000
         ForeColor       =   &H00C66237&
         Height          =   660
         Left            =   105
         TabIndex        =   2
         Top             =   45
         Width           =   4620
      End
   End
   Begin Project1.EnhancedText txtSample 
      Height          =   360
      Left            =   135
      TabIndex        =   0
      Top             =   855
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   635
      BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisabledGrooveBackColor=   14737632
      NormalBorderColor=   8421504
      DisabledBorderColor=   16777215
      FocusBorderColor=   33023
      PasswordChar    =   ""
      Object.Tag             =   ""
      TextFormat      =   ""
      MaxLength       =   0
      SpecialCharacter=   ""
   End
   Begin Project1.EnhancedText txtNumeric 
      Height          =   360
      Left            =   135
      TabIndex        =   6
      Top             =   2700
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   635
      InputType       =   2
      Alignment       =   1
      NormalBackColor =   13001271
      NormalGrooveBackColor=   16776960
      NormalFontColor =   16777215
      BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      DisabledGrooveBackColor=   0
      NormalBorderColor=   16752456
      FocusBorderColor=   33023
      PasswordChar    =   ""
      Object.Tag             =   ""
      TextFormat      =   ""
      MaxLength       =   0
      SpecialCharacter=   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Now, numeric input type accepts only one decimal point"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   2460
      Width           =   3945
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bGroove As Boolean
Dim bDisabled As Boolean
Dim bMultiline As Boolean

Private Sub cmdDisable_Click()
    bDisabled = Not bDisabled
    
    If bDisabled = False Then
        txtSample.Enabled = True
        cmdGroove.Enabled = True
        cmdMultiline.Enabled = True
    Else
        txtSample.Enabled = False
        cmdGroove.Enabled = False
        cmdMultiline.Enabled = False
    End If
End Sub

Private Sub cmdGroove_Click()
    bGroove = Not bGroove
    
    If bGroove = False Then
        txtSample.NormalGrooveBackColor = &H8000000F
    Else
        txtSample.NormalGrooveBackColor = &HEBB076
    End If
End Sub

Private Sub cmdMultiline_Click()
    bMultiline = Not bMultiline
    
    If bMultiline = False Then
        txtSample.MultiLiner = False
    Else
        txtSample.MultiLiner = True
    End If
End Sub

Private Sub Form_Load()
    'set initial values
    bGroove = True
    bDisabled = False
    bMultiline = False
End Sub
