VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trans Tray"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   1590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   75
      Top             =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Declarations
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Type Declarations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Private Variables
Private MousePos As POINTAPI
Private TrayRect As RECT
Private hWndTray As Long
Private Transparent As Boolean

'Unload form when user clicks Exit button
Private Sub cmdExit_Click()
    Unload Me
End Sub

'When form is unloaded, make sure Tray is visible
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MakeNotTransparent hWndTray
End Sub

'If cursor is over Tray, make it visible.
'Otherwise make the Tray transparent.
Private Sub Timer1_Timer()
    hWndTray = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect hWndTray, TrayRect
    GetCursorPos MousePos
    If PtInRect(TrayRect, MousePos.x, MousePos.y) Then
        If Transparent = True Then
            MakeNotTransparent hWndTray
            Transparent = False
        End If
    Else
        If Transparent = False Then
            MakeTransparent 100, hWndTray
            Transparent = True
        End If
    End If
End Sub
