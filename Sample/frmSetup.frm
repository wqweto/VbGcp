VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   4248
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5544
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4248
   ScaleWidth      =   5544
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue 
      Height          =   288
      Index           =   0
      Left            =   2604
      TabIndex        =   4
      Top             =   672
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   3948
      TabIndex        =   3
      Top             =   1260
      Width           =   1272
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   348
      Left            =   2604
      TabIndex        =   2
      Top             =   1260
      Width           =   1272
   End
   Begin VB.ComboBox cobValue 
      Height          =   288
      Index           =   0
      Left            =   2604
      TabIndex        =   1
      Top             =   252
      Visible         =   0   'False
      Width           =   2616
   End
   Begin VB.Label labName 
      AutoSize        =   -1  'True
      Caption         =   "Template"
      Height          =   348
      Index           =   0
      Left            =   252
      TabIndex        =   0
      Top             =   252
      Visible         =   0   'False
      Width           =   2364
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSetup"
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
'   Sample print job settings form
'
' $Log: $
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const GRID_SIZE             As Double = 84
Private Const DBL_STEP              As Double = 348 ' 432

Private m_bConfirm              As Boolean
Private m_bInSet                As Boolean

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit(sPrinterName As String, oPrinterCaps As cGcpPrinterCaps, oOwnerForm As Object) As Boolean
    Dim lIdx            As Long
    Dim dblTop          As Double
    
    '--- contruct UI
    m_bInSet = True
    dblTop = labName(0).Top
    For lIdx = 0 To oPrinterCaps.Count - 1
        If LenB(oPrinterCaps.DisplayName(lIdx)) <> 0 Then
            Load labName(lIdx + 1)
            With labName(lIdx + 1)
                .Top = dblTop
                .Left = labName(0).Left
                .Caption = Replace(oPrinterCaps.DisplayName(lIdx), "&", "&&")
                .Visible = True
            End With
            If oPrinterCaps.HasOptions(lIdx) Then
                Load cobValue(lIdx + 1)
                With cobValue(lIdx + 1)
                    .Top = dblTop
                    .Left = cobValue(0).Left
                    pvFillCombo cobValue(lIdx + 1), oPrinterCaps.SortList(oPrinterCaps.OptionDisplay(lIdx)), oPrinterCaps.Value(lIdx)
                    .Visible = True
                End With
            Else
                Load txtValue(lIdx + 1)
                With txtValue(lIdx + 1)
                    .Top = dblTop
                    .Left = txtValue(0).Left
                    .Text = oPrinterCaps.Value(lIdx)
                    .Visible = True
                End With
            End If
            dblTop = dblTop + DBL_STEP
            If dblTop > Screen.Height * 0.7 Then
                '--- expand form height & show
                dblTop = dblTop + (Height - ScaleHeight) + GRID_SIZE + cmdOk.Height + 2 * GRID_SIZE
                Height = dblTop
                dblTop = cobValue(0).Left + cobValue(0).Width + DBL_STEP - labName(0).Left
                labName(0).Left = labName(0).Left + dblTop
                cobValue(0).Left = cobValue(0).Left + dblTop
                txtValue(0).Left = txtValue(0).Left + dblTop
                Width = Width + dblTop
                dblTop = labName(0).Top
            End If
        End If
    Next
    m_bInSet = False
    dblTop = dblTop + (Height - ScaleHeight) + GRID_SIZE + cmdOk.Height + 2 * GRID_SIZE
    If Height < dblTop Then
        Height = dblTop
    End If
    cmdOk.Move ScaleWidth - cmdOk.Width - GRID_SIZE - cmdCancel.Width - 4 * GRID_SIZE, ScaleHeight - cmdOk.Height - 2 * GRID_SIZE
    cmdCancel.Move ScaleWidth - cmdCancel.Width - 4 * GRID_SIZE, ScaleHeight - cmdCancel.Height - 2 * GRID_SIZE
    Caption = sPrinterName & " - Properties"
    Show vbModal, oOwnerForm
    If m_bConfirm Then
        '--- apply modifications
        For lIdx = 0 To oPrinterCaps.Count - 1
            If pvIsModified(txtValue, lIdx + 1) Then
                oPrinterCaps.Value(lIdx) = txtValue(lIdx + 1).Text
            ElseIf pvIsModified(cobValue, lIdx + 1) Then
                If cobValue(lIdx + 1).ListIndex >= 0 Then
                    oPrinterCaps.Value(lIdx) = cobValue(lIdx + 1).ItemData(cobValue(lIdx + 1).ListIndex)
                End If
            End If
        Next
        '--- success
        frInit = True
    End If
    Unload Me
End Function

Private Function pvIsModified(oCtl As Object, ByVal lIdx As Long) As Boolean
    On Error Resume Next
    pvIsModified = LenB(oCtl(lIdx).Tag) <> 0
    On Error GoTo 0
End Function

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

'=========================================================================
' Control events
'=========================================================================

Private Sub cmdCancel_Click()
    Visible = False
End Sub

Private Sub cmdOk_Click()
    m_bConfirm = True
    Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Visible = False
        Cancel = 1
    End If
End Sub

Private Sub cobValue_Click(Index As Integer)
    If Not m_bInSet Then
        cobValue(Index).Tag = "Changed"
    End If
End Sub

Private Sub txtValue_Change(Index As Integer)
    If Not m_bInSet Then
        txtValue(Index).Tag = "Changed"
    End If
End Sub
