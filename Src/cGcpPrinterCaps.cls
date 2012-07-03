VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGcpPrinterCaps"
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
'   Google Cloud Print printer capabilities helper
'
' $Log: $
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cGcpPrinterCaps"

'=========================================================================
' Public enums
'=========================================================================

Public Enum GcpCapsPropertyEnum
    gcpCapCopies
    gcpCapCollate
    gcpCapOrientation
    gcpCapPaperSize
    gcpCapResolution
    gcpCapColor
    gcpCapDuplex
    [_gcpCapLast]
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ERR_UNKNOWN_FORMAT        As String = "Unknown '{0}' capabilites format"
Private Const FORMAT_XPS                As String = "xps"
Private Const FORMAT_PPD                As String = "ppd"
Private Const REX_COLLATE_COLLATE       As String = "/(.*:collated.*|true)/i"
Private Const REX_COLLATE_NO_COLLATE    As String = "/(.*:uncollated.*|false)/i"
Private Const REX_ORIENTATION_PORTRAIT  As String = "/psk:Portrait/i"
Private Const REX_ORIENTATION_LANDSCAPE As String = "/(psk:Landscape|psk:ReverseLandscape)/i"
Private Const REX_DUPLEX_LONG_EDGE      As String = "/(psk:TwoSidedLongEdge|DuplexNoTumble)/i"
Private Const REX_DUPLEX_OFF            As String = "/(psk:OneSided|None)/i"
Private Const REX_DUPLEX_SHORT_EDGE     As String = "/(psk:TwoSidedShortEdge|DuplexTumble)/i"

Private m_oPrinterInfo          As Object
Private m_sFormatType           As String
Private m_sFormatDisplayName    As String
Private m_oCaps                 As Object
Private m_vValues               As Variant
Private m_vDefaultValues        As Variant
Private m_aProps()              As Long

'=========================================================================
' Properties
'=========================================================================

Property Get PrinterInfo() As Object
    Set PrinterInfo = m_oPrinterInfo
End Property

Property Get Count() As Long
    Count = m_oCaps.Count
End Property

Property Get Name(ByVal Index As Long) As String
    Name = m_oCaps(Index)!Name
End Property

Property Get DisplayName(ByVal Index As Long) As String
    DisplayName = m_oCaps(Index)(m_sFormatDisplayName)
End Property

Property Get DefaultValue(ByVal Index As Long) As Variant
    DefaultValue = m_vDefaultValues(Index)
End Property

Property Get Value(ByVal Index As Long) As Variant
    Value = m_vValues(Index)
    If IsEmpty(Value) Then
        Value = m_vDefaultValues(Index)
    End If
End Property

Property Let Value(ByVal Index As Long, vValue As Variant)
    If m_oCaps(Index).Exists("options") Then
        m_vValues(Index) = CLng(vValue)
    Else
        m_vValues(Index) = CStr(vValue)
    End If
End Property

Property Get ValueByName(ByVal Index As Long) As String
    With m_oCaps(Index)
        If .Exists("options") Then
            ValueByName = !Options(CLng(Value(Index)))!Name
        End If
    End With
End Property

Property Let ValueByName(ByVal Index As Long, sText As String)
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim oRegExp         As Object
    
    If Left$(sText, 1) = "/" Then
        Set oRegExp = pvInitRegExp(sText)
        For Each vElem In OptionNames(Index)
            If oRegExp.test(vElem) Then
                m_vValues(Index) = lIdx
                Exit For
            End If
            lIdx = lIdx + 1
        Next
    Else
        For Each vElem In OptionNames(Index)
            If LCase$(sText) = LCase$(vElem) Then
                m_vValues(Index) = lIdx
                Exit For
            End If
            lIdx = lIdx + 1
        Next
    End If
End Property

Property Get HasOptions(ByVal Index As Long) As Boolean
    HasOptions = m_oCaps(Index).Exists("options")
End Property

Property Get OptionDisplay(ByVal Index As Long) As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    With m_oCaps(Index)
        If .Exists("options") Then
            vRetVal = !Options.Keys
            For lIdx = 0 To UBound(vRetVal)
                vRetVal(lIdx) = !Options(vRetVal(lIdx))(m_sFormatDisplayName)
            Next
        End If
    End With
    OptionDisplay = vRetVal
End Property

Property Get OptionNames(ByVal Index As Long) As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    With m_oCaps(Index)
        If .Exists("options") Then
            vRetVal = !Options.Keys
            For lIdx = 0 To UBound(vRetVal)
                vRetVal(lIdx) = !Options(vRetVal(lIdx))!Name
            Next
        End If
    End With
    OptionNames = vRetVal
End Property

Property Get PropDefaultValue(ByVal Index As GcpCapsPropertyEnum) As Variant
    If m_aProps(Index) >= 0 Then
        PropDefaultValue = DefaultValue(m_aProps(Index))
    End If
End Property

Property Get PropValue(ByVal Index As GcpCapsPropertyEnum) As Variant
    If m_aProps(Index) >= 0 Then
        PropValue = Value(m_aProps(Index))
    End If
End Property

Property Let PropValue(ByVal Index As GcpCapsPropertyEnum, vValue As Variant)
    If m_aProps(Index) >= 0 Then
        Value(m_aProps(Index)) = vValue
    End If
End Property

Property Get PropValueByName(ByVal Index As GcpCapsPropertyEnum) As Variant
    If m_aProps(Index) >= 0 Then
        PropValueByName = ValueByName(m_aProps(Index))
        Select Case Index
        Case gcpCapCollate
            If pvInitRegExp(REX_COLLATE_COLLATE).test(PropValueByName) Then
                PropValueByName = "collate"
            ElseIf pvInitRegExp(REX_COLLATE_NO_COLLATE).test(PropValueByName) Then
                PropValueByName = "no-collate"
            End If
        Case gcpCapOrientation
            If pvInitRegExp(REX_ORIENTATION_PORTRAIT).test(PropValueByName) Then
                PropValueByName = "portrait"
            ElseIf pvInitRegExp(REX_ORIENTATION_LANDSCAPE).test(PropValueByName) Then
                PropValueByName = "landscape"
            End If
        Case gcpCapDuplex
            If pvInitRegExp(REX_DUPLEX_LONG_EDGE).test(PropValueByName) Then
                PropValueByName = "long-edge"
            ElseIf pvInitRegExp(REX_DUPLEX_OFF).test(PropValueByName) Then
                PropValueByName = "off"
            ElseIf pvInitRegExp(REX_DUPLEX_SHORT_EDGE).test(PropValueByName) Then
                PropValueByName = "short-edge"
            End If
        End Select
    End If
End Property

Property Let PropValueByName(ByVal Index As GcpCapsPropertyEnum, vValue As Variant)
    If m_aProps(Index) >= 0 Then
        Select Case Index
        Case gcpCapCollate
            Select Case vValue
            Case "collate"
                ValueByName(m_aProps(Index)) = REX_COLLATE_COLLATE
            Case "no-collate"
                ValueByName(m_aProps(Index)) = REX_COLLATE_NO_COLLATE
            Case Else
                ValueByName(m_aProps(Index)) = vValue
            End Select
        Case gcpCapOrientation
            Select Case vValue
            Case "portrait"
                ValueByName(m_aProps(Index)) = REX_ORIENTATION_PORTRAIT
            Case "landscape"
                ValueByName(m_aProps(Index)) = REX_ORIENTATION_LANDSCAPE
            Case Else
                ValueByName(m_aProps(Index)) = vValue
            End Select
        Case gcpCapDuplex
            Select Case vValue
            Case "long-edge"
                ValueByName(m_aProps(Index)) = REX_DUPLEX_LONG_EDGE
            Case "off"
                ValueByName(m_aProps(Index)) = REX_DUPLEX_OFF
            Case "short-edge"
                ValueByName(m_aProps(Index)) = REX_DUPLEX_SHORT_EDGE
            Case Else
                ValueByName(m_aProps(Index)) = vValue
            End Select
        Case Else ' gcpCapCopies, gcpCapPaperSize, gcpCapResolution, gcpCapColor
            Value(m_aProps(Index)) = vValue
        End Select
    End If
End Property

Property Get PropOptionDisplay(ByVal Index As GcpCapsPropertyEnum) As Variant
    If m_aProps(Index) >= 0 Then
        PropOptionDisplay = OptionDisplay(m_aProps(Index))
    End If
End Property

Property Get PropOptionNames(ByVal Index As GcpCapsPropertyEnum) As Variant
    If m_aProps(Index) >= 0 Then
        PropOptionNames = OptionNames(m_aProps(Index))
    End If
End Property

Property Get PropOptions(ByVal Index As GcpCapsPropertyEnum) As Object
    If m_aProps(Index) >= 0 Then
        With m_oCaps(m_aProps(Index))
            If .Exists("options") Then
                Set PropOptions = !Options
            End If
        End With
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oPrinterInfo As Object) As Boolean
    Const XPS_INDEXES   As String = "psk:JobCopiesAllDocuments|psk:DocumentCollate|psk:PageOrientation|psk:PageMediaSize|psk:PageResolution|psk:PageOutputColor|psk:JobDuplexAllDocumentsContiguously"
    Const PPD_INDEXES   As String = "Copies|Collate|Orientation|PageSize|Resolution|ColorModel|Duplex"
    Dim lIdx            As Long
    Dim sIndexes        As String
    
    Select Case LCase$(oPrinterInfo!capsFormat)
    Case FORMAT_XPS
        m_sFormatType = FORMAT_XPS
        m_sFormatDisplayName = "psk:DisplayName"
        sIndexes = XPS_INDEXES
    Case FORMAT_PPD
        m_sFormatType = FORMAT_PPD
        m_sFormatDisplayName = "displayName"
        sIndexes = PPD_INDEXES
    Case Else
        Err.Raise vbObjectError, , Replace(ERR_UNKNOWN_FORMAT, "{0}", oPrinterInfo!capsFormat)
    End Select
    Set m_oPrinterInfo = oPrinterInfo
    Set m_oCaps = oPrinterInfo!Capabilities
    If Count > 0 Then
        ReDim m_vValues(0 To Count - 1) As Variant
        ReDim m_vDefaultValues(0 To Count - 1) As Variant
    End If
    '--- setup known props indexes
    ReDim m_aProps(0 To [_gcpCapLast] - 1) As Long
    For lIdx = 0 To UBound(m_aProps)
        m_aProps(lIdx) = -1
    Next
    For lIdx = 0 To Count - 1
        m_vDefaultValues(lIdx) = pvGetDefaultValue(lIdx)
        If InStr(1, sIndexes, Name(lIdx)) > 0 Then
            m_aProps(UBound(Split(Left$(sIndexes, InStr(1, sIndexes, Name(lIdx))), "|"))) = lIdx
        End If
    Next
    '--- success
    Init = True
End Function

Public Function SortList(vList As Variant) As Object
    Const adInteger     As Long = 3
    Const adDouble      As Long = 5
    Const adVarChar     As Long = 200
    Const adFldIsNullable As Long = 32
    Dim bNumericOrder   As Boolean
    Dim vElem           As Variant
    Dim sText           As String
    Dim lIdx            As Long
    Dim rsSort          As Object
    Dim oNewItem        As Object
    
    Set rsSort = CreateObject("ADODB.Recordset")
    rsSort.Fields.Append "ID", adInteger
    rsSort.Fields.Append "Name", adVarChar, 1000, adFldIsNullable
    rsSort.Fields.Append "Num", adDouble, , adFldIsNullable
    rsSort.Open
    If IsArray(vList) Then
        bNumericOrder = True
        For Each vElem In vList
            sText = Left$(Trim$(vElem), 1)
            If Len(sText) > 0 And Not IsNumeric(sText) Then
                bNumericOrder = False
            End If
            If Len(sText) > 0 Then
                rsSort.AddNew Array(0, 1, 2), Array(lIdx, vElem, Val(vElem))
            Else
                rsSort.AddNew Array(0, 1, 2), Array(lIdx, Null, Null)
            End If
            lIdx = lIdx + 1
        Next
    End If
    rsSort.Sort = IIf(bNumericOrder, "Num", "Name")
    Set SortList = pvInitJson(Arr:=True)
    lIdx = 0
    Do While Not rsSort.EOF
        With pvInitJson(RetVal:=oNewItem)
            !ID = rsSort!ID.Value
            !Name = rsSort!Name.Value & vbNullString
        End With
        Set SortList.Item(lIdx) = oNewItem
        lIdx = lIdx + 1
        rsSort.MoveNext
    Loop
End Function

Public Function FormatCapabilties() As String
    Dim lIdx            As Long
    Dim oResult         As Object
    Dim oNewItem        As Object
    
    Set oResult = pvInitJson()
    Set oResult("capabilities") = pvInitJson(Arr:=True)
    For lIdx = 0 To Count - 1
        If Not IsEmpty(m_vValues(lIdx)) Then
            Set oNewItem = pvInitJson()
            With m_oCaps(lIdx)
                oNewItem("name") = !Name
                oNewItem("type") = !Type
                If .Exists("options") Then
                    If IsNumeric(m_vValues(lIdx)) Then
                        Set oNewItem("options") = pvInitJson(Arr:=True)
                        Set oNewItem!Options(0) = !Options(CLng(m_vValues(lIdx)))
                    Else
                        '--- wrong index -> skip value
                        Set oNewItem = Nothing
                    End If
                Else
                    oNewItem("value") = CStr(m_vValues(lIdx))
                End If
            End With
            If Not oNewItem Is Nothing Then
                Set oResult!Capabilities(oResult!Capabilities.Count) = oNewItem
            End If
        End If
    Next
    FormatCapabilties = pvJsonDump(oResult)
End Function

Private Function pvGetDefaultValue(ByVal Index As Long) As Variant
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    With m_oCaps(Index)
        If .Exists("options") Then
            For Each vElem In !Options.Items
                If vElem.Exists("default") Then
                    If vElem!Default Then
                        pvGetDefaultValue = lIdx
                        Exit For
                    End If
                End If
                lIdx = lIdx + 1
            Next
        Else
            pvGetDefaultValue = ![psf:DefaultValue]
        End If
    End With
End Function

Friend Function pvJsonDump(vJson As Variant, Optional ByVal lLevel As Long) As String
    Const STR_CODES     As String = "\u0000|\u0001|\u0002|\u0003|\u0004|\u0005|\u0006|\u0007|\b|\t|\n|\u000B|\f|\r|\u000E|\u000F|\u0010|\u0011|\u0012|\u0013|\u0014|\u0015|\u0016|\u0017|\u0018|\u0019|\u001A|\u001B|\u001C|\u001D|\u001E|\u001F"
    Const INDENT        As Long = 4
    Static vTranscode   As Variant
    Dim vKeys           As Variant
    Dim vItems          As Variant
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim sCompound       As String
    Dim lAsc            As Long
    
    On Error GoTo EH
    Select Case VarType(vJson)
    Case vbObject
        sCompound = IIf(vJson.CompareMode = 0, "[]", "{}")
        If vJson.Count = 0 Then
            pvJsonDump = sCompound
        Else
            vKeys = vJson.Keys
            vItems = vJson.Items
            For lIdx = 0 To vJson.Count - 1
                vItems(lIdx) = pvJsonDump(vItems(lIdx), lLevel + 1)
                If VarType(vKeys(lIdx)) = vbString Then
                    vItems(lIdx) = pvJsonDump(vKeys(lIdx)) & ": " & vItems(lIdx)
                End If
                lSize = lSize + Len(vItems(lIdx))
            Next
            If lSize > 100 Then
                pvJsonDump = Left$(sCompound, 1) & vbCrLf & _
                    Space((lLevel + 1) * INDENT) & Join(vItems, "," & vbCrLf & Space((lLevel + 1) * INDENT)) & vbCrLf & _
                    Space(lLevel * INDENT) & Right$(sCompound, 1)
            Else
                pvJsonDump = Left$(sCompound, 1) & " " & Join(vItems, ", ") & " " & Right$(sCompound, 1)
            End If
        End If
    Case vbNull
        pvJsonDump = "Null"
    Case vbEmpty
        pvJsonDump = "Empty"
    Case vbString
        '--- one-time initialization of transcoding array
        If IsEmpty(vTranscode) Then
            vTranscode = Split(STR_CODES, "|")
        End If
        For lIdx = 1 To Len(vJson)
            lAsc = AscW(Mid$(vJson, lIdx, 1))
            If lAsc = 92 Or lAsc = 34 Then '--- \ and "
                pvJsonDump = pvJsonDump & "\" & Chr$(lAsc)
            ElseIf lAsc >= 32 And lAsc < 256 Then
                pvJsonDump = pvJsonDump & Chr$(lAsc)
            ElseIf lAsc >= 0 And lAsc < 32 Then
                pvJsonDump = pvJsonDump & vTranscode(lAsc)
            ElseIf Asc(Mid$(vJson, lIdx, 1)) <> 63 Then '--- ?
                pvJsonDump = pvJsonDump & Chr$(Asc(Mid$(vJson, lIdx, 1)))
            Else
                pvJsonDump = pvJsonDump & "\u" & Right$("0000" & Hex(lAsc), 4)
            End If
        Next
        pvJsonDump = """" & pvJsonDump & """"
    Case Else
        pvJsonDump = vJson & ""
    End Select
    Exit Function
EH:
    Debug.Print Error
    Resume Next
End Function

Private Function pvInitRegExp(sPattern As String) As Object
    Dim lIdx            As Long
    
    Set pvInitRegExp = CreateObject("VBScript.RegExp")
    lIdx = InStrRev(sPattern, "/")
    If lIdx > 0 Then
        pvInitRegExp.IgnoreCase = InStr(1, Mid$(sPattern, lIdx), "i") > 0
        pvInitRegExp.Pattern = Mid$(sPattern, 2, lIdx - 2)
    Else
        pvInitRegExp.Pattern = sPattern
    End If
End Function

Private Function pvInitJson(Optional ByVal Arr As Boolean, Optional RetVal As Object) As Object
    Set RetVal = CreateObject("Scripting.Dictionary")
    RetVal.CompareMode = IIf(Arr, 0, 1)
    Set pvInitJson = RetVal
End Function

