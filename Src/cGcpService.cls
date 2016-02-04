VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGcpService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: $
'
'   VB6 Google Cloud Print proxy
'   Copyright (c) 2012 Unicontsoft
'
'   Google Cloud Print service helper
'
' $Log: $
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cGcpService"

#Const ASYNC_SUPPORT = True

'=========================================================================
' Public events
'=========================================================================

#If ASYNC_SUPPORT Then
    Event Complete(Callback As cGcpCallback)
#End If

'=========================================================================
' Public enums
'=========================================================================

Public Enum GcpCredentialsTypeEnum
    gcpCrtGoogleLogin = 1
    gcpCrtOAuthRefreshToken
End Enum

'=========================================================================
' API
'=========================================================================

'--- for WideCharToMultiByte
Private Const CP_UTF8                               As Long = 65001

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbByte, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Byte()
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_BOUNDARY          As String = "4868eafe-ea88-4b50-ad89-e134e2de80eb"
Private Const STR_SOURCE            As String = "vb-gcp-proxy-v1"
Private Const URL_SUBMIT            As String = "/cloudprint/submit"
Private Const URL_SEARCH            As String = "/cloudprint/search"
Private Const URL_PRINTER           As String = "/cloudprint/printer"
Private Const URL_JOBS              As String = "/cloudprint/jobs"
Private Const URL_DELETEJOB         As String = "/cloudprint/deletejob"
Private Const URL_CLIENTLOGIN       As String = "https://www.google.com/accounts/ClientLogin?accountType=HOSTED_OR_GOOGLE&Email={0}&Passwd={1}&service=cloudprint&source=" & STR_SOURCE
Private Const URL_OAUTH_TOKEN       As String = "https://accounts.google.com/o/oauth2/token"
Private Const URL_OAUTH_USERINFO    As String = "https://www.googleapis.com/oauth2/v1/userinfo"
Private Const BOM_UTF               As String = "ï»¿" '--- "\xEF\xBB\xBF"
Private Const DEF_TIMEOUT           As Long = 10 '--- seconds

Private m_sHost                 As String
Private m_sCredentials          As String
Private m_eCredType             As GcpCredentialsTypeEnum
Private m_sAuthorization        As String
Private m_sLastError            As String
Private m_lTimeout              As Long
#If ASYNC_SUPPORT Then
    Private m_bAsyncOperations  As Boolean
    Private m_cAsyncCallbacks   As Collection
    Private m_lAsyncCookie      As Long
#End If

'=========================================================================
' Properties
'=========================================================================

Property Get Timeout() As Long
    Timeout = IIf(m_lTimeout > 0, m_lTimeout, DEF_TIMEOUT)
End Property

Property Let Timeout(ByVal lValue As Long)
    m_lTimeout = lValue
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

#If ASYNC_SUPPORT Then
    Property Get AsyncOperations() As Boolean
        AsyncOperations = m_bAsyncOperations
    End Property
    
    Property Let AsyncOperations(ByVal bValue As Boolean)
        m_bAsyncOperations = bValue
    End Property
#End If

'=========================================================================
' Methods
'=========================================================================

Public Function Init( _
            sHost As String, _
            sCredentials As String, _
            ByVal eType As GcpCredentialsTypeEnum) As Boolean
    m_sHost = sHost
    m_sCredentials = sCredentials
    m_eCredType = eType
    m_sAuthorization = vbNullString
    '--- success
    Init = True
End Function

Public Function PrintDocument( _
            sPrinterId As String, _
            sFile As String, _
            Optional Title As String, _
            Optional ContentType As String, _
            Optional Capabilities As String) As Object
    #Const USE_DATA_URL = False
    #Const USE_BASE64 = False
    Dim sPostData       As String
    
    If Not pvAuthorize(PrintDocument) Then
        Exit Function
    End If
    If LenB(ContentType) = 0 Then
        ContentType = pvMatchContentType(sFile)
    End If
    pvRestAddParam sPostData, "printerid", sPrinterId
    If ContentType Like "text/*" And False Then
        pvRestAddParam sPostData, "contentType", ContentType
    Else
        #If USE_DATA_URL Then
            pvRestAddParam sPostData, "contentType", "dataUrl"
        #Else
            pvRestAddParam sPostData, "contentType", ContentType
            #If USE_BASE64 Then
                pvRestAddParam sPostData, "contentTransferEncoding", "base64"
            #End If
        #End If
    End If
    pvRestAddParam sPostData, "title", Title
    pvRestAddParam sPostData, "capabilities", Capabilities
    If ContentType Like "text/*" And False Then
        pvRestAddParam sPostData, "content", BOM_UTF & pvToUtf8(pvReadTextFile(sFile)), ContentType
    Else
        #If USE_DATA_URL Then
            pvRestAddParam sPostData, "content", "data:" & ContentType & ";base64," & pvToBase64(pvReadBinaryFile(sFile))
        #ElseIf USE_BASE64 Then
            pvRestAddParam sPostData, "content", pvToBase64(pvReadBinaryFile(sFile)), ContentType
        #Else
            pvRestAddParam sPostData, "content""; filename=""" & Mid$(sFile, InStrRev(sFile, "\") + 1), StrConv(pvReadBinaryFile(sFile), vbUnicode), ContentType
        #End If
    End If
    pvRestInvoke URL_SUBMIT, sPostData, PrintDocument, Timeout:=Timeout * 3
End Function

Public Function GetPrinters(Optional Pattern As String) As Object
    Dim sPostData       As String
    
    If Not pvAuthorize(GetPrinters) Then
        Exit Function
    End If
    pvRestAddParam sPostData, "q", Pattern
    pvRestAddParam sPostData, "connection_status", "ALL"
    pvRestInvoke URL_SEARCH, sPostData, GetPrinters
End Function

Public Function GetPrinterInfo(sPrinterId As String) As Object
    Dim sPostData       As String

    If Not pvAuthorize(GetPrinterInfo) Then
        Exit Function
    End If
    pvRestAddParam sPostData, "printerid", sPrinterId
    pvRestAddParam sPostData, "use_cdd", "false"
    pvRestInvoke URL_PRINTER, sPostData, GetPrinterInfo
End Function

Public Function GetJobs(Optional PrinterId As String, Optional ByVal Limit As Long = 10) As Object
    Dim sPostData       As String

    If Not pvAuthorize(GetJobs) Then
        Exit Function
    End If
    pvRestAddParam sPostData, "printerid", PrinterId
    pvRestAddParam sPostData, "limit", CStr(Limit)
    pvRestInvoke URL_JOBS, sPostData, GetJobs
End Function

Public Function DeleteJob(sJobId As String) As Object
    Dim sPostData       As String

    If Not pvAuthorize(DeleteJob) Then
        Exit Function
    End If
    pvRestAddParam sPostData, "jobid", sJobId
    pvRestInvoke URL_DELETEJOB, sPostData, DeleteJob
End Function

'--- Used by cGcpOAuth to retrieve rehresh_token from OAuth code. On success retrieves user's email too.
Friend Function frGetRefreshToken( _
            sCode As String, _
            sClientId As String, _
            sClientSecret As String, _
            sRefreshToken As String, _
            sUserEmail As String, _
            sError As String) As Boolean
    Dim sPostData       As String
    Dim oToken          As Object
    Dim sAuthorization  As String
    Dim oUserInfo       As Object
    
    pvRestAddParam sPostData, "code", sCode
    pvRestAddParam sPostData, "client_id", sClientId
    pvRestAddParam sPostData, "client_secret", sClientSecret
    pvRestAddParam sPostData, "redirect_uri", "urn:ietf:wg:oauth:2.0:oob"
    pvRestAddParam sPostData, "grant_type", "authorization_code"
    pvRestAddParam sPostData
    With pvInitRequest("POST", URL_OAUTH_TOKEN)
        .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .Send pvToByteArray(sPostData)
        If pvJsonParse(.ResponseText, 1, oToken, sError) Then
            If LenB(oToken!refresh_token) <> 0 Then
                sRefreshToken = oToken!refresh_token
                sAuthorization = "OAuth " & oToken!access_token
            Else
                sError = oToken!Error
            End If
        End If
    End With
    If LenB(sAuthorization) Then
        With pvInitRequest("GET", URL_OAUTH_USERINFO)
            .SetRequestHeader "Authorization", sAuthorization
            .Send
            If pvJsonParse(.ResponseText, 1, oUserInfo, sError) Then
                If oUserInfo.Exists("error") Then
                    sError = oUserInfo!Error!message
                Else
                    sUserEmail = oUserInfo!email
                    '--- success
                    frGetRefreshToken = True
                End If
            End If
        End With
    End If
End Function

#If ASYNC_SUPPORT Then
    '--- Used by cGcpCallback to call back in when underlying XHR ready-state changes
    Friend Sub frOnReadyStateChange(oCallback As cGcpCallback)
        Dim oResult             As Object
        
        If oCallback.Request.ReadyState = 4 Then
            If Not pvJsonParse(oCallback.Request.ResponseText, 1, oResult, m_sLastError) Then
                Set oResult = pvInitDictionary()
                oResult!success = False
                oResult!message = "Invalid response"
                oResult!response = oCallback.Request.ResponseText
                oResult!parseError = m_sLastError
            End If
            oCallback.frSetResult oResult
            RaiseEvent Complete(oCallback)
            oCallback.frFireComplete
            m_cAsyncCallbacks.Remove oCallback.Cookie
            oCallback.frTerminate
        End If
    End Sub
#End If

'= private ===============================================================

Private Function pvAuthorize(oResult As Object) As Boolean
    Dim vElem           As Variant
    Dim sUrl            As String
    Dim lPos            As Long
    Dim sPostData       As String
    Dim oToken          As Object
    Dim sError          As String
    Dim sResponse       As String
    
    If LenB(m_sAuthorization) <> 0 Then
        pvAuthorize = True
    ElseIf m_eCredType = gcpCrtGoogleLogin Then
        '--- split username & password
        lPos = InStr(m_sCredentials, ":")
        If lPos = 0 Then
            lPos = Len(m_sCredentials) + 1
        End If
        sUrl = Replace(Replace(URL_CLIENTLOGIN, "{0}", Left$(m_sCredentials, lPos - 1)), "{1}", Mid$(m_sCredentials, lPos + 1))
        With pvInitRequest("GET", sUrl)
            .Send
            For Each vElem In Split(.ResponseText, vbLf)
                If LCase$(Left$(vElem, 5)) = "auth=" Then
                    m_sAuthorization = "GoogleLogin auth=" & Mid$(vElem, 6)
                    pvAuthorize = True
                    Exit For
                End If
            Next
            sResponse = .ResponseText
        End With
    Else ' If m_eCredType = gcpCrtOAuthRefreshToken Then
        vElem = Split(m_sCredentials, ":")
        pvRestAddParam sPostData, "refresh_token", At(vElem, 0)
        pvRestAddParam sPostData, "client_id", At(vElem, 1)
        pvRestAddParam sPostData, "client_secret", At(vElem, 2)
        pvRestAddParam sPostData, "grant_type", "refresh_token"
        pvRestAddParam sPostData
        With pvInitRequest("POST", URL_OAUTH_TOKEN)
            .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
            .Send pvToByteArray(sPostData)
            If pvJsonParse(.ResponseText, 1, oToken, sError) Then
                If LenB(oToken!access_token) <> 0 Then
                    m_sAuthorization = "OAuth " & oToken!access_token
                    pvAuthorize = True
                Else
                    sError = oToken!Error
                End If
            End If
            sResponse = .ResponseText
       End With
    End If
    If Not pvAuthorize Then
        m_sLastError = "Authorization failure" & IIf(LenB(sError) <> 0, ": ", vbNullString) & sError
        If Not AsyncOperations Then
            Set oResult = pvInitDictionary()
            oResult!success = False
            oResult!message = m_sLastError
            oResult!response = sResponse
        End If
    End If
End Function

Private Sub pvRestAddParam(sPostData As String, Optional Name As String, Optional Value As String, Optional ContentType As String)
    If LenB(Name) <> 0 Then
        If LenB(Value) <> 0 Then
            sPostData = sPostData & "--" & STR_BOUNDARY & vbCrLf & _
                "Content-Disposition: form-data; name=""" & Name & """"
            If LenB(ContentType) <> 0 Then
                sPostData = sPostData & vbCrLf & _
                    "Content-Type: " & ContentType & vbCrLf & vbCrLf & _
                    Value & vbCrLf
            Else
                sPostData = sPostData & vbCrLf & vbCrLf & _
                    pvToUtf8(Value) & vbCrLf
            End If
        End If
    ElseIf LenB(sPostData) <> 0 And Right$(sPostData, 2) <> "--" Then
        sPostData = sPostData & "--" & STR_BOUNDARY & "--"
    End If
End Sub

Private Function pvRestInvoke(sUrl As String, sPostData As String, oResult As Object, Optional ByVal Timeout As Long) As Boolean
    Dim bAsync          As Boolean
    Dim oRequest        As Object
    
    On Error GoTo EH
    '--- append final boundary (if necessary)
    pvRestAddParam sPostData
    #If ASYNC_SUPPORT Then
        bAsync = AsyncOperations
    #End If
    With pvInitRequest("POST", IIf(Left$(sUrl, 1) = "/", m_sHost, vbNullString) & sUrl, Timeout:=Timeout, Async:=bAsync, RetVal:=oRequest)
        If LenB(sPostData) <> 0 Then
            .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        End If
        .SetRequestHeader "X-CloudPrint-Proxy", STR_SOURCE
        .SetRequestHeader "Authorization", m_sAuthorization
        '--- use pvToByteArray if sPostData contains '\0' chars from binary files
        .Send pvToByteArray(sPostData)
        #If ASYNC_SUPPORT Then
            If bAsync Then
                pvInitCallback oRequest, RetVal:=oResult
                GoTo QH
            End If
        #End If
        pvRestInvoke = pvJsonParse(.ResponseText, 1, oResult, m_sLastError)
        If Not pvRestInvoke Then
            Set oResult = pvInitDictionary()
            oResult!success = False
            oResult!message = "Invalid response"
            oResult!response = .ResponseText
            oResult!parseError = m_sLastError
        End If
    End With
QH:
    Exit Function
EH:
    m_sLastError = Err.Description
    Set oResult = pvInitDictionary()
    oResult!success = False
    oResult!message = m_sLastError
End Function

Friend Function pvJsonParse(sText As String, lPos As Long, vResult As Variant, Optional Error As String) As Boolean
    Dim vToken          As Variant
    Dim sName           As String
    Dim vValue          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    Error = vbNullString
    vToken = pvJsonGetToken(sText, lPos)
    If VarType(vToken) = vbString Then
        Select Case Left$(vToken, 1)
        Case "{"
            Set vResult = CreateObject("Scripting.Dictionary")
            vResult.CompareMode = 1 ' TextCompare
            Do
                vToken = pvJsonGetToken(sText, lPos)
                If VarType(vToken) <> vbString Then
                    GoTo InvalidToken
                ElseIf vToken = "}" Then
                    Exit Do
                ElseIf Left$(vToken, 1) <> """" Or Len(vToken) < 2 Then
                    Error = "Expected name at position " & lPos - 1
                    GoTo QH
                End If
                sName = Mid$(vToken, 2, Len(vToken) - 2)
                vToken = pvJsonGetToken(sText, lPos)
                If VarType(vToken) <> vbString Then
                    GoTo InvalidToken
                ElseIf vToken <> ":" Then
                    Error = "Expected ':' at position " & lPos - 1
                    GoTo QH
                End If
                If Not pvJsonParse(sText, lPos, vValue, Error) Then
                    GoTo QH
                End If
                vResult.Add sName, vValue
                vToken = pvJsonGetToken(sText, lPos)
                If VarType(vToken) <> vbString Then
                    GoTo InvalidToken
                End If
                Select Case vToken
                Case "}"
                    Exit Do
                Case ","
                Case Else
                    GoTo InvalidToken
                End Select
            Loop
        Case "["
            Set vResult = CreateObject("Scripting.Dictionary")
            vResult.CompareMode = 0 ' BinaryCompare
            '--- peek next token and check for empty array
            lIdx = lPos
            vToken = pvJsonGetToken(sText, lPos)
            If VarType(vToken) = vbString Then
                If vToken = "]" Then
                    '--- success
                    pvJsonParse = True
                    GoTo QH
                End If
            End If
            lPos = lIdx
            '--- non-empty array
            For lIdx = 0 To &H7FFFFFFF
                If Not pvJsonParse(sText, lPos, vValue, Error) Then
                    GoTo QH
                End If
                vResult.Add lIdx, vValue
                vToken = pvJsonGetToken(sText, lPos)
                If VarType(vToken) <> vbString Then
                    GoTo InvalidToken
                End If
                Select Case vToken
                Case "]"
                    Exit For
                Case ","
                Case Else
                    GoTo InvalidToken
                End Select
            Next
        Case """"
            If Len(vToken) < 2 Then
                GoTo InvalidToken
            End If
            vResult = Mid$(vToken, 2, Len(vToken) - 2)
        Case Else
            GoTo InvalidToken
        End Select
    Else
        vResult = vToken
    End If
    '--- success
    pvJsonParse = True
QH:
    Exit Function
InvalidToken:
    Error = "Invalid token " & Switch(VarType(vToken) = vbEmpty, "Empty", VarType(vToken) = vbNull, "Null", _
        VarType(vToken) = vbString, "'" & vToken & "'", True, vToken & "") & " at position " & lPos - 1
    Exit Function
EH:
    Debug.Print Error
    Resume Next
End Function

Private Function pvJsonGetToken(sText As String, lPos As Long) As Variant
    Dim sChar           As String
    
    On Error GoTo EH
    '--- skip white-space
    Do
        sChar = Mid$(sText, lPos, 1)
        lPos = lPos + 1
        Select Case sChar
        Case " ", vbTab, vbCr, vbLf
        Case Else
            Exit Do
        End Select
    Loop
    Select Case LCase$(sChar)
    Case vbNullString
        '--- return empty
        GoTo QH
    Case "t"
        If "rue" = LCase$(Mid$(sText, lPos, 3)) Then
            lPos = lPos + 3
            pvJsonGetToken = True
            GoTo QH
        End If
    Case "f"
        If "alse" = LCase$(Mid$(sText, lPos, 4)) Then
            lPos = lPos + 4
            pvJsonGetToken = False
            GoTo QH
        End If
    Case "n"
        If "ull" = LCase$(Mid$(sText, lPos, 3)) Then
            lPos = lPos + 3
            pvJsonGetToken = Null
            GoTo QH
        End If
    Case """"
        pvJsonGetToken = sChar
        Do
            sChar = Mid$(sText, lPos, 1)
            lPos = lPos + 1
            Select Case sChar
            Case "\"
                sChar = Mid$(sText, lPos, 1)
                lPos = lPos + 1
                Select Case sChar
                Case "b"
                    pvJsonGetToken = pvJsonGetToken & Chr$(8)
                Case "f"
                    pvJsonGetToken = pvJsonGetToken & Chr$(12)
                Case "n"
                    pvJsonGetToken = pvJsonGetToken & vbLf
                Case "r"
                    pvJsonGetToken = pvJsonGetToken & vbCr
                Case "t"
                    pvJsonGetToken = pvJsonGetToken & vbTab
                Case "u"
                    pvJsonGetToken = pvJsonGetToken & ChrW$(CLng("&H" & Mid$(sText, lPos, 4)))
                    lPos = lPos + 4
                Case Else ' "\", "'", """"
                    pvJsonGetToken = pvJsonGetToken & sChar
                End Select
            Case """"
                pvJsonGetToken = pvJsonGetToken & sChar
                Exit Do
            Case Else
                pvJsonGetToken = pvJsonGetToken & sChar
            End Select
        Loop
        GoTo QH
    Case Else
        If sChar Like "[0-9-]" Then '
            pvJsonGetToken = sChar
            Do
                sChar = Mid$(sText, lPos, 1)
                If sChar Like "[0-9eE.]" Then
                    lPos = lPos + 1
                    pvJsonGetToken = pvJsonGetToken & sChar
                Else
                    Exit Do
                End If
            Loop
            pvJsonGetToken = CDbl(pvJsonGetToken)
            GoTo QH
        End If
    End Select
    pvJsonGetToken = sChar
QH:
    Exit Function
EH:
    Debug.Print Error
    Resume Next
End Function

Private Function pvInitRequest( _
            sType As String, _
            sUrl As String, _
            Optional ByVal Timeout As Long, _
            Optional ByVal Async As Boolean, _
            Optional RetVal As Object) As Object
    '--- first try server-side XMLHTTP because it has timeouts
    On Error Resume Next
    Set RetVal = CreateObject("MSXML2.ServerXMLHTTP")
    RetVal.SetTimeouts 5000, 5000, 5000, IIf(Timeout > 0, Timeout, Me.Timeout) * 1000
    On Error GoTo 0
    If RetVal Is Nothing Then
        Set RetVal = CreateObject("MSXML2.XMLHTTP")
    End If
    RetVal.Open sType, sUrl, Async
    Set pvInitRequest = RetVal
End Function

#If ASYNC_SUPPORT Then
    Private Function pvInitCallback(oRequest As Object, Optional RetVal As cGcpCallback) As cGcpCallback
        Set RetVal = New cGcpCallback
        If RetVal.frInit(Me, oRequest, "#" & m_lAsyncCookie) Then
            If m_cAsyncCallbacks Is Nothing Then
                Set m_cAsyncCallbacks = New Collection
            End If
            m_cAsyncCallbacks.Add RetVal, "#" & m_lAsyncCookie
            m_lAsyncCookie = m_lAsyncCookie + 1
            Set pvInitCallback = RetVal
        End If
    End Function
#End If

Private Function pvInitDictionary() As Object
    Set pvInitDictionary = CreateObject("Scripting.Dictionary")
    pvInitDictionary.CompareMode = 1 ' TextCompare
End Function

Private Function pvToUtf8(sText As String) As String
    Dim lSize           As Long
    
    pvToUtf8 = String(4 * Len(sText), 0)
    lSize = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(sText), Len(sText), ByVal pvToUtf8, Len(pvToUtf8), 0, 0)
    pvToUtf8 = Left$(pvToUtf8, lSize)
End Function

Private Function pvToBase64(baData() As Byte) As String
    With CreateObject("MSXML2.DOMDocument").createElement("dummy")
        .DataType = "bin.base64"
        .NodeTypedValue = baData
        pvToBase64 = .Text
    End With
End Function

Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function

Private Function pvMatchContentType(sFile As String) As String
    Const STR_MIME      As String = "|application/pdf|.xls.xlsx.xlsb.|application/vnd.ms-excel|.ppt.pptx.pptb.|application/vnd.ms-powerpoint.doc.docx.docb.|application/vnd.ms-word|.xps.|application/vnd.ms-xpsdocument|.pdf.|application/pdf|" & _
        ".ps.prn.|application/postscript|.rtf.|application/rtf|.bmp.|image/bmp|.gif.|image/gif|.jpg.jpeg.|image/jpeg|.png.|image/png|.psd.|image/photoshop|.tif.tiff.|image/tiff|.htm.html.|text/html|.txt.|text/plain|.xml.|text/xml"
    pvMatchContentType = Split(Mid$(STR_MIME, InStr(1, STR_MIME, "." & Mid$(sFile, InStrRev(sFile, ".") + 1) & ".", vbTextCompare) + 1), "|")(1)
End Function

Private Function pvReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    
    On Error GoTo EH
    baBuffer = ApiEmptyByteArray()
    If GetAttr(sFile) Or True Then
        nFile = FreeFile
        Open sFile For Binary Access Read As nFile
        If LOF(nFile) > 0 Then
            ReDim baBuffer(0 To LOF(nFile) - 1)
            Get nFile, , baBuffer
        End If
        Close nFile
    End If
    pvReadBinaryFile = baBuffer
EH:
End Function

Private Function pvReadTextFile(sFile As String) As String
    Const BOM_UTF       As String = "ï»¿" '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "ÿþ"  '--- "\xFF\xFE"
    Const ForReading    As Long = 1
    Dim lSize           As Long
    Dim sPrefix         As String
    
    With CreateObject("Scripting.FileSystemObject")
        lSize = .GetFile(sFile).Size
        If lSize = 0 Then
            Exit Function
        End If
        sPrefix = .OpenTextFile(sFile, ForReading).Read(IIf(lSize < 50, lSize, 50))
        If Left$(sPrefix, Len(BOM_UTF)) <> BOM_UTF And Left$(sPrefix, Len(BOM_UNICODE)) <> BOM_UNICODE Then
            '--- special xml encoding test
            If InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
                sPrefix = BOM_UTF
            End If
        End If
        If Left$(sPrefix, Len(BOM_UTF)) <> BOM_UTF Then
            On Error Resume Next
            pvReadTextFile = .OpenTextFile(sFile, ForReading, False, Left$(sPrefix, Len(BOM_UNICODE)) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0).ReadAll()
            On Error GoTo 0
        Else
            With CreateObject("ADODB.Stream")
                .Open
                If Left$(sPrefix, Len(BOM_UNICODE)) = BOM_UNICODE Then
                    .Charset = "Unicode"
                ElseIf Left$(sPrefix, Len(BOM_UTF)) = BOM_UTF Then
                    .Charset = "UTF-8"
                Else
                    .Charset = "_autodetect_all"
                End If
                .LoadFromFile sFile
                pvReadTextFile = .ReadText
            End With
        End If
    End With
End Function

Private Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error Resume Next
    At = Default
    At = Data(Index)
    On Error GoTo 0
End Function

'=========================================================================
' Base class events
'=========================================================================

#If ASYNC_SUPPORT Then
    Private Sub Class_Terminate()
        Dim oCallback       As cGcpCallback
        
        If Not m_cAsyncCallbacks Is Nothing Then
            For Each oCallback In m_cAsyncCallbacks
                oCallback.frTerminate
            Next
        End If
'        Debug.Print MODULE_NAME & " Terminate"
    End Sub
#End If
