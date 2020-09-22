VERSION 5.00
Begin VB.UserControl fbTrayIcon 
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   ScaleHeight     =   1155
   ScaleWidth      =   1350
   Begin VB.PictureBox pichook 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Developed By Federico Bridger (Rosario, Arg.)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1395
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   390
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "fbTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'APIs
'*******************************************************************
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'*******************************************************************

'Types
'*******************************************************************
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'*******************************************************************

'Constants
'*******************************************************************
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
'*******************************************************************

'Enumerations
'*******************************************************************
Public Enum EnumFBButtonConstants
  FB_LEFT_BUTTON_DBLCLK = &H203
  FB_LEFT_BUTTON_DOWN = &H201
  FB_LEFT_BUTTON_UP = &H202
  FB_MIDDLE_BUTTON_DBLCLK = &H209
  FB_MIDDLE_BUTTON_DOWN = &H207
  FB_MIDDLE_BUTTON_UP = &H208
  FB_RIGHT_BUTTOND_BLCLK = &H206
  FB_RIGHT_BUTTON_DOWN = &H204
  FB_RIGHT_BUTTON_UP = &H205
End Enum
'*******************************************************************

'Events
'*******************************************************************
Event MouseClick(ByVal FBButton As EnumFBButtonConstants)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************************


'Variables
'*******************************************************************
Private TrayI As NOTIFYICONDATA
'*******************************************************************

'Properties
'*******************************************************************
Public Property Get Tip() As String
  
  Tip = TrayI.szTip
  
End Property

Public Property Let Tip(ByVal Value As String)

  TrayI.szTip = Value & Chr$(0)
  
End Property
'*******************************************************************

'Subs
'*******************************************************************
Public Sub AddTrayIcon(ByVal ImageFile As String, Tip As String)

  'Load the Picture in the Image Object
  imgIcon(0).Picture = LoadPicture(ImageFile)
  
  TrayI.cbSize = Len(TrayI)
  TrayI.szTip = Tip & Chr$(0)
  TrayI.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
  TrayI.uId = 1&
  TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  TrayI.ucallbackMessage = FB_LEFT_BUTTON_DOWN
  TrayI.hIcon = imgIcon(0).Picture
  
  'Create the icon
  Shell_NotifyIcon NIM_ADD, TrayI
  
End Sub

Public Sub RemoveTrayIcon()
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd
    TrayI.uId = 1&
    'Delete the icon
    Shell_NotifyIcon NIM_DELETE, TrayI
End Sub

Public Sub ChangeTrayIcon(ByVal ImageFile As String, Tip As String)
    
    'Load the Picture in the Image Object
    imgIcon(0).Picture = LoadPicture(ImageFile)
    
    TrayI.hIcon = imgIcon(0).Picture
    TrayI.szTip = Tip & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, TrayI
    
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
  RaiseEvent MouseDown(Button, Shift, X, Y)
  
  RaiseEvent MouseClick(X / Screen.TwipsPerPixelX)
  
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  RaiseEvent MouseMove(Button, Shift, X, Y)
  
End Sub

Private Sub pichook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  RaiseEvent MouseUp(Button, Shift, X, Y)
  
  RaiseEvent MouseClick(X / Screen.TwipsPerPixelX)
  
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = 1160
  UserControl.Width = 1350
End Sub
'*******************************************************************
