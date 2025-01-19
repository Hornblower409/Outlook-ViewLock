VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ViewLockForm 
   Caption         =   "ViewLock"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9105
   OleObjectBlob   =   "ViewLockForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ViewLockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' v2025-01-18
'
' =====================================================================
'   ViewLock - Outlook Lock and Unlock Views Module
'
'   Github Home
'   https://github.com/Hornblower409/Outlook-ViewLock
'
'   Github Releases
'   https://github.com/Hornblower409/Outlook-ViewLock/releases
'
' =====================================================================
'
'   Copyright (C) 2024, 2025 Lycon Of Texas
'
'   This program is free software: you can redistribute it
'   and/or modify it under the terms of the GNU General Public
'   License Version 3 as published by the Free Software
'   Foundation.
'
'   This program is distributed in the hope that it will be
'   useful, but WITHOUT ANY WARRANTY; without even the implied
'   warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR
'   PURPOSE.  See the GNU General Public License for more
'   details.
'
'   You should have received a copy of the GNU General Public
'   License along with this program.  If not, see
'   <https://www.gnu.org/licenses/>.
'
'==============================================================

' ---------------------------------------------------------------------
'   Pseudo ENums shared between the Module and the Form
' ---------------------------------------------------------------------

    '   Actions
    '
    Private Const Actions_First     As Long = 0
    Private Const Actions_Unlock    As Long = 0
    Private Const Actions_Lock      As Long = 1
    Private Const Actions_State     As Long = 2
    Private Const Actions_Save      As Long = 3
    Private Const Actions_Status    As Long = 4
    Private Const Actions_Form      As Long = 5
    Private Const Actions_Last      As Long = 5
    
    '   Scopes
    '
    Private Const Scopes_First      As Long = 0
    Private Const Scopes_Stores     As Long = 0
    Private Const Scopes_Store      As Long = 1
    Private Const Scopes_Folders    As Long = 2
    Private Const Scopes_Folder     As Long = 3
    Private Const Scopes_View       As Long = 4
    Private Const Scopes_Last       As Long = 4
    
    '   View Types
    '
    Private Const VTypes_First      As Long = 0
    Private Const VTypes_None       As Long = 0
    Private Const VTypes_Shared     As Long = 1
    Private Const VTypes_Public     As Long = 2
    Private Const VTypes_Private    As Long = 4
    Private Const VTypes_All        As Long = 7
    Private Const VTypes_Last       As Long = 7
    
' ---------------------------------------------------------------------
'   Form Globals
' ---------------------------------------------------------------------

    Private Scope                   As Long
    Private LastScope               As Long
    Private Action                  As Long
'   Private VTypes                  As Long     '   VTypes is calculated from the Checkboxes by Function VTypes() when needed.
    
' =====================================================================
Private Sub UserForm_Initialize()

    LastScope = -1
    ScopeViewButton_Click
    
End Sub

Private Sub UserForm_Activate()
    OnStatusChange
End Sub

' ---------------------------------------------------------------------
'   Scope
' ---------------------------------------------------------------------
'
Private Sub ScopeStoresButton_Click()
    Scope = Scopes_Stores
    OnScopeChange
End Sub

Private Sub ScopeStoreButton_Click()
    Scope = Scopes_Store
    OnScopeChange
End Sub

Private Sub ScopeFoldersButton_Click()
    Scope = Scopes_Folders
    OnScopeChange
End Sub

Private Sub ScopeFolderButton_Click()
    Scope = Scopes_Folder
    OnScopeChange
End Sub

Private Sub ScopeViewButton_Click()
    Scope = Scopes_View
    OnScopeChange
End Sub

Private Sub OnScopeChange()

    '   If not (changing FROM View Scope or TO View Scope) - Done
    '
    If (LastScope <> Scopes_View) And (Scope <> Scopes_View) Then
        LastScope = Scope
        OnStatusChange
        Exit Sub
    End If
    
    '   Changing TO View Scope
    '
    If ScopeViewButton.Value = True Then
    
        '   Enable the View buttons and Lock down the Types
        '   selection to only the VType of the current View
        '
        ViewStateButton.Enabled = True
        ViewSaveButton.Enabled = True
        VTypeAllCheckbox.Value = False
        VTypeAllCheckbox.Enabled = False
        VTypeAllValues False
        VTypeAllEnabled False
        
        '   Get the VType of the Current View and set the
        '   equivalent VType Type Checkbox
        '
        Select Case ViewVType()
            Case VTypes_Shared
                VTypeSharedCheckbox.Value = True
            Case VTypes_Public
                VTypePublicCheckBox.Value = True
            Case VTypes_Private
                VTypePrivateCheckBox.Value = True
            Case Else
                '   Oops
                Stop: Exit Sub
        End Select
        
    '   Changing FROM View Scope
    '
    Else
    
        '   Disable the View buttons and unlock
        '   the Types selection
        '
        ViewStateButton.Enabled = False
        ViewSaveButton.Enabled = False
        VTypeAllCheckbox.Enabled = True
        VTypeAllCheckbox.Value = True
        
    End If
    
    LastScope = Scope
    OnStatusChange

End Sub

'   Get the VTypes of the Current View
'
Private Function ViewVType() As Long
    
    Select Case ActiveWindow.CurrentView.SaveOption
        Case olViewSaveOptionAllFoldersOfType
            ViewVType = VTypes_Shared
        Case olViewSaveOptionThisFolderEveryone
            ViewVType = VTypes_Public
        Case olViewSaveOptionThisFolderOnlyMe
            ViewVType = VTypes_Private
        Case Else
            '   Ooops
            Stop: Exit Function
    End Select

End Function

' =====================================================================
Private Sub VTypeAllCheckbox_Click()

    VTypeAllValues VTypeAllCheckbox.Value
    VTypeAllEnabled (Not VTypeAllCheckbox.Value)
    
End Sub

Private Sub VTypeSharedCheckbox_Click()
    onVTypeChange
End Sub

Private Sub VTypePublicCheckBox_Click()
    onVTypeChange
End Sub

Private Sub VTypePrivateCheckBox_Click()
    onVTypeChange
End Sub

Private Sub onVTypeChange()

    LockButton.Enabled = (VTypes <> 0)
    UnlockButton.Enabled = (VTypes <> 0)
    OnStatusChange

End Sub
Private Sub VTypeAllEnabled(ByVal Enabled As Boolean)

    VTypeSharedCheckbox.Enabled = Enabled
    VTypePublicCheckBox.Enabled = Enabled
    VTypePrivateCheckBox.Enabled = Enabled

End Sub

Private Sub VTypeAllValues(ByVal Value As Boolean)

    VTypeSharedCheckbox.Value = Value
    VTypePublicCheckBox.Value = Value
    VTypePrivateCheckBox.Value = Value

End Sub

Private Function VTypes() As Long

    VTypes = _
           IIf(VTypeSharedCheckbox.Value, 1, 0) _
        + (IIf(VTypePublicCheckBox.Value, 1, 0) * 2) _
        + (IIf(VTypePrivateCheckBox.Value, 1, 0) * 4)

End Function

Private Sub OnStatusChange()

    ViewLock_Xeq _
        ActionEnum:=Actions_Status, _
        ScopeENum:=Scope, _
        CallerName:="ViewLock_FormInput", _
        VTypesENum:=VTypes, _
        WhatIfBool:=WhatIfCheckBox.Value

End Sub

' =====================================================================
Private Sub LockButton_Click()
    Action = Actions_Lock
    OnCommand
End Sub

Private Sub UnlockButton_Click()
    Action = Actions_Unlock
    OnCommand
End Sub

Private Sub ViewStateButton_Click()
    Action = Actions_State
    OnCommand
End Sub

Private Sub ViewSaveButton_Click()
    Action = Actions_Save
    OnCommand
End Sub

Private Sub OnCommand()
    
    ViewLock_Xeq _
        ActionEnum:=Action, _
        ScopeENum:=Scope, _
        CallerName:="ViewLock_FormInput", _
        VTypesENum:=VTypes, _
        WhatIfBool:=WhatIfCheckBox.Value
        
    ExitButton.SetFocus
        
End Sub

' ---------------------------------------------------------------------
'   Exit/Close
' ---------------------------------------------------------------------
'
Private Sub ExitButton_Click()
    OnExit
End Sub
 
'   From : https://rubberduckvba.blog/2017/10/25/userform1-show/
'
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnExit
    End If
    
End Sub
 
Private Sub OnExit()
    Hide
End Sub

' ---------------------------------------------------------------------
'   Update the Status Display text
' ---------------------------------------------------------------------
'
'   2025-01-18 - RubberDuck
'@Ignore WriteOnlyProperty
Public Property Let StatusDisplay(ByVal Text As String)
    StatusLabel.Caption = Text
End Property

