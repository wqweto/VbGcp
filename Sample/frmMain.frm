VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Google Cloud Print sample"
   ClientHeight    =   7128
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8076
   LinkTopic       =   "Form1"
   ScaleHeight     =   7128
   ScaleWidth      =   8076
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Job"
      Height          =   348
      Left            =   6468
      TabIndex        =   24
      Top             =   3276
      Width           =   1356
   End
   Begin VB.Frame fraJob 
      Caption         =   "Job settings"
      Height          =   1188
      Left            =   252
      TabIndex        =   19
      Top             =   1512
      Visible         =   0   'False
      Width           =   6060
      Begin VB.TextBox txtCopies 
         Height          =   288
         Left            =   1008
         TabIndex        =   4
         Text            =   "1"
         Top             =   336
         Width           =   516
      End
      Begin VB.PictureBox picTab1 
         BorderStyle     =   0  'None
         Height          =   684
         Left            =   2016
         ScaleHeight     =   684
         ScaleWidth      =   1272
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   336
         Width           =   1272
         Begin VB.OptionButton optPortrait 
            Caption         =   "Portrait"
            Height          =   264
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   1188
         End
         Begin VB.OptionButton optLandscape 
            Caption         =   "Landscape"
            Height          =   264
            Left            =   0
            TabIndex        =   7
            Top             =   336
            Width           =   1188
         End
      End
      Begin VB.ComboBox cobPaper 
         Height          =   288
         Left            =   4284
         TabIndex        =   8
         Top             =   336
         Width           =   1608
      End
      Begin VB.ComboBox cobResolution 
         Height          =   288
         Left            =   4284
         TabIndex        =   9
         Top             =   756
         Width           =   1608
      End
      Begin VB.CheckBox chkCollate 
         Caption         =   "Collate"
         Height          =   264
         Left            =   1008
         TabIndex        =   5
         Top             =   672
         Width           =   1356
      End
      Begin VB.Label Label1 
         Caption         =   "Copies:"
         Height          =   264
         Left            =   168
         TabIndex        =   23
         Top             =   336
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "Paper size:"
         Height          =   264
         Left            =   3360
         TabIndex        =   22
         Top             =   336
         Width           =   1356
      End
      Begin VB.Label Label3 
         Caption         =   "Resolution:"
         Height          =   264
         Left            =   3360
         TabIndex        =   21
         Top             =   756
         Width           =   1356
      End
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties..."
      Height          =   348
      Left            =   6468
      TabIndex        =   10
      Top             =   1596
      Width           =   1356
   End
   Begin VB.ListBox lstJobs 
      Height          =   1968
      Left            =   252
      TabIndex        =   14
      Top             =   3780
      Width           =   7572
   End
   Begin VB.Frame Frame1 
      Caption         =   "Google Cloud Print setttings"
      Height          =   852
      Left            =   252
      TabIndex        =   17
      Top             =   168
      Width           =   6060
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Setup"
         Height          =   348
         Left            =   4536
         TabIndex        =   0
         Top             =   336
         Width           =   1356
      End
      Begin VB.Label labOAuth 
         Height          =   348
         Left            =   252
         TabIndex        =   18
         Top             =   336
         Width           =   4044
      End
   End
   Begin VB.ComboBox cobFile 
      Height          =   288
      Left            =   252
      TabIndex        =   12
      Top             =   2856
      Width           =   6060
   End
   Begin VB.CommandButton cmdPrinterInfo 
      Caption         =   "Printer Info"
      Height          =   348
      Left            =   6468
      TabIndex        =   3
      Top             =   1176
      Width           =   1356
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1356
      Left            =   84
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "frmMain.frx":0000
      Top             =   5880
      Width           =   7824
   End
   Begin VB.CommandButton cmdJobs 
      Caption         =   "Jobs"
      Height          =   348
      Left            =   6468
      TabIndex        =   11
      Top             =   2016
      Width           =   1356
   End
   Begin VB.ComboBox cobPrinter 
      Height          =   288
      Left            =   252
      TabIndex        =   2
      Top             =   1176
      Width           =   6060
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   348
      Left            =   6468
      TabIndex        =   13
      Top             =   2856
      Width           =   1356
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   348
      Left            =   6468
      TabIndex        =   1
      Top             =   504
      Width           =   1356
   End
   Begin VB.Label labJob 
      Height          =   264
      Left            =   252
      TabIndex        =   16
      Top             =   3360
      Width           =   6060
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' $Header: $
'
'   VB6 Google Cloud Print proxy
'   Copyright (c) 2012 Unicontsoft
'
'   Sample print document form
'
' $Log: $
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CLIENT_ID         As String = "270393284181.apps.googleusercontent.com"
Private Const STR_CLIENT_SECRET     As String = "BjX4X3EQDN-MRpc0F4sETz83"
Private Const REG_GCP_SECTION       As String = "GoogleCloudPrint"
Private Const URL_HOST              As String = "https://www.google.com"

Private m_bActivated            As Boolean
Private WithEvents m_oService   As cGcpService
Attribute m_oService.VB_VarHelpID = -1
Private m_oPrinters             As Object
Private m_oPrinterCaps          As cGcpPrinterCaps
Private m_oJobs                 As Object
Private WithEvents m_oAsyncConnect As cGcpCallback
Attribute m_oAsyncConnect.VB_VarHelpID = -1
Private WithEvents m_oAsyncPrinterInfo As cGcpCallback
Attribute m_oAsyncPrinterInfo.VB_VarHelpID = -1
Private WithEvents m_oAsyncJobs As cGcpCallback
Attribute m_oAsyncJobs.VB_VarHelpID = -1
Private WithEvents m_oAsyncPrint As cGcpCallback
Attribute m_oAsyncPrint.VB_VarHelpID = -1
Private WithEvents m_oAsyncDelete As cGcpCallback
Attribute m_oAsyncDelete.VB_VarHelpID = -1

'=========================================================================
' Methods
'=========================================================================

Private Function pvInitService(ByVal Async As Boolean, Optional ByVal Clear As Boolean = True) As cGcpService
    If m_oService Is Nothing Then
        Set m_oService = New cGcpService
'        m_oService.Init URL_HOST, "{user}:{pass}", gcpCrtGoogleLogin
    End If
    Set pvInitService = m_oService
    pvInitService.AsyncOperations = Async
    If Clear Then
        txtLog.Text = vbNullString
    End If
End Function

Private Function pvLogResult(oResult As Object, sMethod As String) As Object
    txtLog.Text = Right$(txtLog.Text & Format$(Timer, "0.00") & " " & sMethod & vbCrLf & pvJsonDump(oResult) & vbCrLf & vbCrLf, &HFFFF&)
    Set pvLogResult = oResult
    If oResult!success Then
    Else
        MsgBox oResult!message, vbExclamation, "Error"
    End If
End Function

Private Function pvJsonDump(vJson As Variant) As String
    With New cGcpPrinterCaps
        pvJsonDump = .pvJsonDump(vJson)
    End With
End Function

Private Sub pvSetupOAuth()
    Dim sRefreshToken       As String
    
    sRefreshToken = GetSetting(App.Title, REG_GCP_SECTION, "RefreshToken", vbNullString)
    If LenB(sRefreshToken) <> 0 Then
        labOAuth.Caption = "Current user: " & GetSetting(App.Title, REG_GCP_SECTION, "Email", "<unknown>")
        pvInitService(False).Init URL_HOST, sRefreshToken & ":" & STR_CLIENT_ID & ":" & STR_CLIENT_SECRET, gcpCrtOAuthRefreshToken
    Else
        labOAuth.Caption = "Not setup"
        Set m_oService = Nothing
    End If
    Me.Refresh
    If Not m_oService Is Nothing Then
        cmdConnect.Value = True
    End If
End Sub

Private Sub pvFillPrinterCaps(oPrinterCaps As cGcpPrinterCaps)
    With oPrinterCaps
        pvFillCombo cobPaper, .SortList(.PropOptionDisplay(gcpCapPaperSize)), .PropValue(gcpCapPaperSize)
        pvFillCombo cobResolution, .SortList(.PropOptionDisplay(gcpCapResolution)), .PropValue(gcpCapResolution)
        If .PropValueByName(gcpCapOrientation) = "portrait" Then
            optPortrait.Value = True
        Else
            optLandscape.Value = True
        End If
        If .PropValueByName(gcpCapCollate) = "collate" Then
            chkCollate.Value = vbChecked
        Else
            chkCollate.Value = vbUnchecked
        End If
        If Not IsEmpty(.PropValue(gcpCapCopies)) Then
            txtCopies.Text = .PropValue(gcpCapCopies)
        Else
            txtCopies.Text = "1"
        End If
    End With
    fraJob.Visible = True
    Me.Refresh
    fraJob.Refresh
End Sub

Private Sub pvUploadPrinterCaps(oPrinterCpas As cGcpPrinterCaps)
    Dim sValue          As String
    
    With oPrinterCpas
        sValue = CStr(Val(txtCopies.Text))
        If .PropValue(gcpCapCopies) <> sValue Then
            .PropValue(gcpCapCopies) = sValue
        End If
        sValue = IIf(optPortrait.Value, "portrait", "landscape")
        If .PropValueByName(gcpCapOrientation) <> sValue Then
            .PropValueByName(gcpCapOrientation) = sValue
        End If
        sValue = IIf(chkCollate.Value = vbChecked, "collate", "no-collate")
        If .PropValueByName(gcpCapCollate) <> sValue Then
            .PropValueByName(gcpCapCollate) = sValue
        End If
        If cobPaper.ListIndex >= 0 Then
            If .PropValue(gcpCapPaperSize) <> cobPaper.ItemData(cobPaper.ListIndex) Then
                .PropValue(gcpCapPaperSize) = cobPaper.ItemData(cobPaper.ListIndex)
            End If
        End If
        If cobResolution.ListIndex >= 0 Then
            If .PropValue(gcpCapResolution) <> cobResolution.ItemData(cobResolution.ListIndex) Then
                .PropValue(gcpCapResolution) = cobResolution.ItemData(cobResolution.ListIndex)
            End If
        End If
    End With
End Sub

Private Sub pvFillFiles(sFolder As String)
    Dim sFile           As String
    
    cobFile.Clear
    sFile = Dir(sFolder & "\*.*")
    Do While LenB(sFile) <> 0
        If sFile Like "*.txt" Or sFile Like "*.pdf" Or sFile Like "*.gif" Or sFile Like "*.jpg" Or sFile Like "*.doc*" Or sFile Like "*.xls*" Or sFile Like "*.xml" Then
            cobFile.AddItem sFolder & "\" & sFile
        End If
        sFile = Dir
    Loop
    If cobFile.ListCount > 0 Then
        cobFile.ListIndex = 0
    End If
End Sub

Private Sub pvFillCombo(cobCombo As ComboBox, oList As Object, ByVal lCurrent As Long)
    Dim lIdx            As Long
    
    cobCombo.Clear
    For lIdx = 0 To oList.Count - 1
        If LenB(oList(lIdx)!Name) <> 0 Then
            cobCombo.AddItem oList(lIdx)!Name
            cobCombo.ItemData(cobCombo.NewIndex) = oList(lIdx)!ID
            If oList(lIdx)!ID = lCurrent Then
                cobCombo.ListIndex = cobCombo.ListCount - 1
            End If
        End If
    Next
End Sub

Private Function pvInitCaps(oPrinterInfo As Object, Optional RetVal As cGcpPrinterCaps) As cGcpPrinterCaps
    Set RetVal = New cGcpPrinterCaps
    If RetVal.Init(oPrinterInfo) Then
        Set pvInitCaps = RetVal
    End If
End Function

Private Function IsKeyPressed(ByVal lVirtKey As KeyCodeConstants) As Boolean
    IsKeyPressed = ((GetAsyncKeyState(lVirtKey) And &H8000) = &H8000)
End Function

Private Function GetShiftState() As ShiftConstants
    GetShiftState = vbShiftMask * -IsKeyPressed(vbKeyShift) _
                Or vbCtrlMask * -IsKeyPressed(vbKeyControl) _
                Or vbAltMask * -IsKeyPressed(vbKeyMenu)
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub cmdSetup_Click()
    Dim sRefreshToken   As String
    Dim sUserEmail      As String
    
    If (GetShiftState() And vbCtrlMask) <> 0 Then
        On Error Resume Next
        DeleteSetting App.Title, REG_GCP_SECTION, "Email"
        DeleteSetting App.Title, REG_GCP_SECTION, "RefreshToken"
        On Error GoTo 0
        pvSetupOAuth
    Else
        With New frmLogin
            If .frInit(STR_CLIENT_ID, STR_CLIENT_SECRET, sRefreshToken, sUserEmail, Me) Then
                SaveSetting App.Title, REG_GCP_SECTION, "Email", sUserEmail
                SaveSetting App.Title, REG_GCP_SECTION, "RefreshToken", sRefreshToken
                pvSetupOAuth
            End If
        End With
    End If
End Sub

Private Sub cmdConnect_Click()
    If m_oService Is Nothing Then
        cmdSetup.Value = True
    End If
    If Not m_oService Is Nothing Then
        Set m_oAsyncConnect = pvInitService(Async:=True).GetPrinters()
    End If
End Sub

Private Sub m_oAsyncConnect_Complete(oResult As Object)
    Dim vElem           As Variant
    
    With pvLogResult(oResult, "m_oAsyncConnect_Complete")
        If !success Then
            cobPrinter.Clear
            Set m_oPrinters = !printers
            For Each vElem In m_oPrinters.Items
                cobPrinter.AddItem vElem!DisplayName & IIf(LenB(vElem!connectionStatus) <> 0 And vElem!connectionStatus <> "ONLINE", " (" & vElem!connectionStatus & ")", vbNullString)
            Next
            If cobPrinter.ListCount > 0 Then
                cobPrinter.ListIndex = 0
            End If
        End If
    End With
End Sub

Private Sub cmdPrinterInfo_Click()
    Dim sPrinterId      As String
    
    fraJob.Visible = False
    If cobPrinter.ListIndex >= 0 Then
        sPrinterId = m_oPrinters(cobPrinter.ListIndex)!ID
    End If
    Set m_oAsyncPrinterInfo = pvInitService(Async:=True).GetPrinterInfo(sPrinterId)
    cmdJobs.Value = True
End Sub

Private Sub m_oAsyncPrinterInfo_Complete(oResult As Object)
    Dim oPrinterInfo    As Object
        
    With pvLogResult(oResult, "m_oAsyncPrinterInfo_Complete")
        If !success Then
            Set oPrinterInfo = !printers(0)
        End If
    End With
    If Not oPrinterInfo Is Nothing Then
        pvFillPrinterCaps pvInitCaps(oPrinterInfo, RetVal:=m_oPrinterCaps)
    End If
End Sub

Private Sub cmdProperties_Click()
    Dim sPrinterId      As String
    Dim sPrinterName    As String
    
    If Not m_oPrinterCaps Is Nothing Then
        pvUploadPrinterCaps m_oPrinterCaps
    Else
        If cobPrinter.ListIndex >= 0 Then
            sPrinterId = m_oPrinters(cobPrinter.ListIndex)!ID
        End If
        m_oAsyncPrinterInfo_Complete pvInitService(Async:=False).GetPrinterInfo(sPrinterId)
    End If
    If Not m_oPrinterCaps Is Nothing Then
        If cobPrinter.ListIndex >= 0 Then
            sPrinterName = m_oPrinters(cobPrinter.ListIndex)!DisplayName
        End If
        With New frmSetup
            If .frInit(sPrinterName, m_oPrinterCaps, Me) Then
                pvFillPrinterCaps m_oPrinterCaps
            End If
        End With
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim sPrinterId      As String
    Dim lSize           As Long
    Dim sTitle          As String
    Dim sCapabilities   As String
    
    If cobPrinter.ListIndex >= 0 Then
        sPrinterId = m_oPrinters(cobPrinter.ListIndex)!ID
    End If
    If (GetAttr(cobFile.Text) And vbDirectory) <> 0 Then
        pvFillFiles cobFile.Text
    Else
        On Error Resume Next
        lSize = FileLen(cobFile.Text)
        On Error GoTo 0
        If Not m_oPrinterCaps Is Nothing Then
            pvUploadPrinterCaps m_oPrinterCaps
            sCapabilities = m_oPrinterCaps.FormatCapabilties()
        End If
        sTitle = App.Title & " - " & Mid$(cobFile.Text, InStrRev(cobFile.Text, "\") + 1) & " @ " & Now
        labJob.Caption = "Submitting " & Format$((lSize + 1023) \ 1024, "#,0") & "KB ..."
        Me.Refresh
        With pvInitService(Async:=True)
            Set m_oAsyncPrint = .PrintDocument(sPrinterId, cobFile.Text, Title:=sTitle, Capabilities:=sCapabilities)
        End With
    End If
End Sub

Private Sub m_oAsyncPrint_Complete(oResult As Object)
    Dim oJob            As Object
    
    With pvLogResult(oResult, "m_oAsyncPrint_Complete")
        If !success Then
            Set oJob = !job
        End If
    End With
    If Not oJob Is Nothing Then
        labJob.Caption = oJob!Title
        Me.Refresh
        cmdJobs.Value = True
    End If
End Sub

Private Sub cmdJobs_Click()
    Dim sPrinterId      As String
    
    If cobPrinter.ListIndex >= 0 Then
        sPrinterId = m_oPrinters(cobPrinter.ListIndex)!ID
    End If
    Set m_oAsyncJobs = pvInitService(Async:=True, Clear:=False).GetJobs(sPrinterId, 20)
End Sub

Private Sub m_oAsyncJobs_Complete(oResult As Object)
    Dim oJobs           As Object
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sText           As String
    
    With pvLogResult(oResult, "m_oAsyncJobs_Complete")
        If !success Then
            Set oJobs = !jobs
        End If
    End With
    If Not oJobs Is Nothing Then
        For Each vElem In oJobs.Items
            sText = vElem!Status & vbTab & vElem!numberOfPages & vbTab & vElem!Title & vbTab & vElem!printerName
            If lIdx < lstJobs.ListCount Then
                lstJobs.List(lIdx) = sText
            Else
                lstJobs.AddItem sText
            End If
            lIdx = lIdx + 1
        Next
        Do While lIdx < lstJobs.ListCount
            lstJobs.RemoveItem lIdx
        Loop
        Set m_oJobs = oJobs
    End If
    If m_oJobs.Count > 0 Then
        If m_oJobs(0)!Status = "QUEUED" Or m_oJobs(0)!Status = "IN_PROGRESS" Then
            cmdJobs.Value = True
        End If
    End If
End Sub

Private Sub cobPrinter_Click()
    cobPaper.Clear
    cobResolution.Clear
    Set m_oPrinterCaps = Nothing
    fraJob.Visible = False
    lstJobs.Clear
    cmdPrinterInfo.Value = True
End Sub

Private Sub cobFile_Click()
    Dim lSize As Long
    
    lSize = -1
    On Error Resume Next
    lSize = FileLen(cobFile.Text)
    On Error GoTo 0
    If lSize >= 0 Then
        labJob.Caption = "Size of " & Mid$(cobFile.Text, InStrRev(cobFile.Text, "\") + 1) & ": " & Format$((lSize + 1023) \ 1024, "#,0") & "KB"
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim sJobId          As String
    
    If lstJobs.ListIndex >= 0 Then
        sJobId = m_oJobs(lstJobs.ListIndex)!ID
    End If
    Set m_oAsyncDelete = pvInitService(Async:=True).DeleteJob(sJobId)
End Sub

Private Sub m_oAsyncDelete_Complete(oResult As Object)
    With pvLogResult(oResult, "m_oAsyncDelete_Complete")
        If !success Then
            
        End If
    End With
    cmdJobs.Value = True
End Sub

Private Sub Form_Activate()
    If Not m_bActivated Then
        m_bActivated = True
        Me.Refresh
        pvFillFiles Environ$("TEMP")
        pvSetupOAuth
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtLog.Move 0, txtLog.Top, ScaleWidth, ScaleHeight - txtLog.Top
End Sub

