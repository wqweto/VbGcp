VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGcpCallback"
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
'   Async REST service callback
'
' $Log: $
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cGcpCallback"

#Const SERVICE_WEAK_REF = True

'=========================================================================
' Public events
'=========================================================================

Public Event Complete(oResult As Object)

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oService              As cGcpService
Private m_oRequest              As Object
Private m_sCookie               As String
Private m_oResult               As Object

'=========================================================================
' Properties
'=========================================================================

Property Get Cookie() As String
    Cookie = m_sCookie
End Property

Property Get Request() As Object
    Set Request = m_oRequest
End Property

Property Get Result() As Object
    Set Result = m_oResult
End Property

'=========================================================================
' Methods
'=========================================================================

'--- for Callback routine set Procedure ID=(Default) and Hidden=[x] in menu Tools->Procedure Attributes
Public Sub Callback()
Attribute Callback.VB_UserMemId = 0
Attribute Callback.VB_MemberFlags = "40"
    If Not m_oService Is Nothing Then
        m_oService.frOnReadyStateChange Me
    End If
End Sub

Public Sub Abort()
    If Not m_oRequest Is Nothing Then
        m_oRequest.OnReadyStateChange = Nothing
        m_oRequest.Abort
        Set m_oRequest = Nothing
    End If
End Sub

Friend Function frInit(oService As cGcpService, oRequest As Object, sCookie As String) As Boolean
    #If SERVICE_WEAK_REF Then
        Call CopyMemory(m_oService, oService, 4)
    #Else
        Set m_oService = oService
    #End If
    Set m_oRequest = oRequest
    m_oRequest.OnReadyStateChange = Me
    m_sCookie = sCookie
    '--- success
    frInit = True
End Function

Friend Sub frSetResult(oResult As Object)
    Set m_oResult = oResult
End Sub

Friend Sub frFireComplete()
    RaiseEvent Complete(m_oResult)
End Sub

Friend Sub frTerminate()
    '--- don't clear cookie and result
    If Not m_oRequest Is Nothing Then
        m_oRequest.OnReadyStateChange = Nothing
        Set m_oRequest = Nothing
    End If
    #If SERVICE_WEAK_REF Then
        Call CopyMemory(m_oService, 0&, 4)
    #Else
        Set m_oService = Nothing
    #End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
'    Debug.Print MODULE_NAME & " Terminate: " & m_sCookie
End Sub
