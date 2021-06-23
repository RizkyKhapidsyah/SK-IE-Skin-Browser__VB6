VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Skin Browser"
   ClientHeight    =   2115
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4260
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
   Begin VB.PictureBox picpreview 
      Height          =   1212
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   720
      Width           =   2532
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4800
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DialogTitle     =   "Select a Bitmap File"
      Filter          =   "Bitmap Files | *.bmp"
   End
   Begin VB.TextBox txtpath 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2532
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&Browse"
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuhelpabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const REG_DWORD = 4
Private Const REG_SZ = 1
Const regkey As String = "software\microsoft\internet explorer\toolbar\"
Dim strfile As String

Private Sub cmdapply_Click()
Dim reponse As Variant
Dim retvalue As Long
Dim keyid As Long
Dim subkey As String
Dim keyvalue As String
Dim buffer As Long

subkey = "BackBitmapIE5"
retvalue = RegCreateKey(HKEY_CURRENT_USER, regkey, keyid)
retvalue = RegQueryValueEx(keyid, subkey, 0&, REG_SZ, keyvalue, buffer)
If keyvalue = "" Then
keyvalue = String(buffer + 1, " ")
keyvalue = strfile
retvalue = RegSetValueEx(keyid, subkey, 0&, REG_SZ, ByVal keyvalue, Len(buffer) + 1)
End If
txtpath.Text = ""
cmdapply.Enabled = False
End Sub

Private Sub cmdbrowse_Click()
cd.ShowOpen
If cd.FileName <> "" Then strfile = cd.FileName: txtpath.Text = strfile: picpreview.Picture = LoadPicture(strfile): cmdapply.Enabled = True
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub mnuhelpabout_Click()
frmabout.Show
End Sub
