VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmLogin 
   Caption         =   "User Credentials"
   ClientHeight    =   5664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6876
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5664
   ScaleWidth      =   6876
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3456
      Left            =   336
      TabIndex        =   0
      Top             =   252
      Width           =   5136
      ExtentX         =   9059
      ExtentY         =   6096
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oOAuth     As cGcpOAuth
Attribute m_oOAuth.VB_VarHelpID = -1
Private m_bConfirm              As Boolean

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit( _
            sClientId As String, _
            sClientSecret As String, _
            sRefreshToken As String, _
            sUserEmail As String, _
            oOwnerForm As Object) As Boolean
    Load Me
    Set m_oOAuth = New cGcpOAuth
    m_oOAuth.Init WebBrowser1, sClientId, sClientSecret
    Show vbModal, oOwnerForm
    If m_bConfirm Then
        sRefreshToken = m_oOAuth.RefreshToken
        sUserEmail = m_oOAuth.UserEmail
        '--- success
        frInit = True
    End If
    Unload Me
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Visible = False
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub m_oOAuth_Complete(ByVal Allowed As Boolean, DenyReason As String)
    m_bConfirm = True ' Allowed
    Visible = False
End Sub
