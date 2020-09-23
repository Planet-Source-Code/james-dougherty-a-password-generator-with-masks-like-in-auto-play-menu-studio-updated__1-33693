VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Generator"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1500
      Left            =   0
      TabIndex        =   3
      Top             =   -45
      Width           =   4695
      Begin VB.CommandButton cmdSet 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set"
         Height          =   285
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtCustom 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   400
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Mask"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Mask"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox txtNumPass 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Text            =   "100"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdGen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generate"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2480
         Width           =   1095
      End
      Begin VB.ListBox lstPass 
         Height          =   2595
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   1
         Top             =   260
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Passwords To Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGen_Click()
'Generate the passwords
GeneratePasswords lstPass, CLng(txtNumPass), cmbMask.Text
End Sub

Private Sub cmdSet_Click()
'Add the custom mask to the list box and select it
If txtCustom <> "" Then
 cmbMask.AddItem txtCustom
 cmbMask.ListIndex = cmbMask.ListCount - 1
 txtCustom = ""
End If
End Sub

Private Sub Form_Load()

'For our password mask we are going to set it up like this:
'
'# = Numbers
'X = Letters
'
'This can be whatever you want

cmbMask.AddItem "####-####-####-####", 0
cmbMask.AddItem "XXXX-XXXX-XXXX-XXXX", 1
cmbMask.AddItem "XXXX-XXXX-####-####", 2
cmbMask.AddItem "####-####-XXXX-XXXX", 3
cmbMask.AddItem "####-XXXX-####-XXXX", 4
cmbMask.AddItem "XXXX-####-XXXX-####", 5
cmbMask.AddItem "CDKEY - ####-####-####-####", 6
cmbMask.AddItem "CDKEY - XXXX-XXXX-XXXX-XXXX", 7
cmbMask.AddItem "CDKEY - ####-####-XXXX-XXXX", 8
cmbMask.AddItem "CDKEY - XXXX-XXXX-####-####", 9
cmbMask.AddItem "CDKEY - ####-XXXX-####-XXXX", 10
cmbMask.AddItem "CDKEY - XXXX-####-XXXX-####", 11
cmbMask.AddItem "KEY - ####-####-####-####", 12
cmbMask.AddItem "KEY - XXXX-XXXX-XXXX-XXXX", 13
cmbMask.AddItem "KEY - ####-####-XXXX-XXXX", 14
cmbMask.AddItem "KEY - XXXX-XXXX-####-####", 15
cmbMask.AddItem "KEY - ####-XXXX-####-XXXX", 16
cmbMask.AddItem "KEY - XXXX-####-XXXX-####", 17
cmbMask.ListIndex = 0
End Sub
