VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGcpOAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cGcpOAuth"

'=========================================================================
' Public events
'=========================================================================

Event Complete(ByVal Allowed As Boolean, DenyReason As String)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_SUCCESS           As String = "success code="
Private Const STR_DENIED            As String = "denied error="
Private Const STR_CSS               As String = "font-family: arial, helvetica, sans-serif; font-size: 16px;"
Private Const URL_BLANK             As String = "about:<body style='{style}'>Please wait...</body>"
Private Const URL_ERROR             As String = "about:<body style='{style}'>Error: {0}<p><a href=""{1}"">Retry</a></body>"
Private Const URL_AUTH              As String = "https://accounts.google.com/o/oauth2/auth?scope={1}&redirect_uri=urn:ietf:wg:oauth:2.0:oob&response_type=code&client_id={0}"
Private Const DEF_SCOPE             As String = "https://www.googleapis.com/auth/cloudprint https://www.googleapis.com/auth/userinfo.email"

Private WithEvents m_oCtl       As WebBrowser
Attribute m_oCtl.VB_VarHelpID = -1
Private m_sClientId             As String
Private m_sClientSecret         As String
Private m_sScope                As String
Private m_sCode                 As String
Private m_sRefreshToken         As String
Private m_sUserEmail            As String

'=========================================================================
' Properties
'=========================================================================

Property Get Code() As String
    Code = m_sCode
End Property

Property Get RefreshToken() As String
    RefreshToken = m_sRefreshToken
End Property

Property Get UserEmail() As String
    UserEmail = m_sUserEmail
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init( _
            oCtl As WebBrowser, _
            sClientId As String, _
            sClientSecret As String, _
            Optional sScope As String) As Boolean
    Set m_oCtl = oCtl
    m_sClientId = sClientId
    m_sClientSecret = sClientSecret
    m_sScope = IIf(LenB(sScope) <> 0, sScope, DEF_SCOPE)
    '--- show "wait..." and navigate to logon form
    m_oCtl.Navigate Replace(URL_BLANK, "{style}", STR_CSS)
    Do While m_oCtl.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    m_oCtl.Navigate Replace(Replace(URL_AUTH, "{0}", m_sClientId), "{1}", m_sScope)
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub m_oCtl_TitleChange(ByVal Text As String)
    If LenB(m_sCode) = 0 And LCase$(Left$(Text, Len(STR_SUCCESS))) = STR_SUCCESS Then
        m_sCode = Mid$(Text, Len(STR_SUCCESS) + 1)
        m_oCtl.Stop
        m_oCtl.Navigate Replace(URL_BLANK, "{style}", STR_CSS)
    ElseIf LCase$(Left$(Text, Len(STR_DENIED))) = STR_DENIED Then
        RaiseEvent Complete(False, Mid$(Text, Len(STR_DENIED) + 1))
    End If
End Sub

Private Sub m_oCtl_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim sError          As String
    
    On Error GoTo EH
    If URL Like "about:*" And LenB(m_sCode) <> 0 Then
        With New cGcpService
             If .frGetRefreshToken(m_sCode, m_sClientId, m_sClientSecret, m_sRefreshToken, m_sUserEmail, sError) Then
                RaiseEvent Complete(True, vbNullString)
                Exit Sub
             End If
        End With
NavigateToRetry:
        On Error Resume Next
        m_sCode = vbNullString
        m_sRefreshToken = vbNullString
        m_sUserEmail = vbNullString
        m_oCtl.Navigate Replace(Replace(Replace(URL_ERROR, "{0}", sError), "{1}", Replace(Replace(URL_AUTH, "{0}", m_sClientId), "{1}", m_sScope)), "{style}", STR_CSS)
    End If
    Exit Sub
EH:
    sError = Error
    Resume NavigateToRetry
End Sub

