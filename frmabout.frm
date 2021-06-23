VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About "
   ClientHeight    =   2415
   ClientLeft      =   1605
   ClientTop       =   1695
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2123
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lbltext 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strtext As String
strtext = "This is a Internet Explorer Skin Browser." & vbCrLf
strtext = strtext & "Click On Browse Button To Select the picture." & vbCrLf
strtext = strtext & "Take a Preview of the Selected Picture." & vbCrLf
strtext = strtext & "Click Apply to Set that Picture as IE Skin" & vbCrLf
strtext = strtext & "Mail me @ Multiplesoftware@hotmail.com"
lbltext.Caption = strtext
End Sub

