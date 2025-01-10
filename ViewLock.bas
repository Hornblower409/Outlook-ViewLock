Attribute VB_Name = "ViewLock"
Option Explicit

'2025-01-09_182059086
'
'==============================================================
'- ViewLock - Outlook Lock and Unlock Views Module
'
'    Github Home Page
'    https://github.com/Hornblower409/Outlook-ViewLock
'
'- Purpose ----------------------------------------------------
'
'    To compensate for the lack of a "Save Changes?" step after
'    modifying the Settings of an Outlook View.
'
'    ViewLock can lock a View, preventing any accidental changes
'    from being saved when you close the Explorer, and allows
'    you to Edit a View with the option to not save your
'    changes.
'
'    When you make changes to an Outlook View there is no way to
'    cancel or undo them. If you accidentally click on a column
'    heading it permanently changes the Sort order for that
'    View. If you have modified any of the Advanced View
'    Settings (Columns, Group By, Sort, etc.) and don't like the
'    new design, there's no way to undo your changes once you
'    hit "OK" on the Settings dialog. There are ways to backup
'    and restore a View, but they are manual and tedious.
'
'- Install ----------------------------------------------------
'
'    This is a standalone Module with no external references and
'    one Form. To install from the VBA Editor do:
'
'    File -> Import: ViewLock.bas
'    File -> Import: ViewLockForm.frm
'
'    The Module has five Macros:
'
'    ViewLock_Lock
'    ViewLock_Unlock
'    ViewLock_State
'    ViewLock_Save
'    ViewLock_Form
'
'    The first four operate only on the current View.
'    ViewLock_Form opens a User Form that allows you Lock/Unlock
'    Views in the current Folder, the current Folder and all
'    it 's subfolders, the current Store (.pst file) or the
'    entire system.
'
'- VBA Help ---------------------------------------------------
'
'    For help on using the VBA Editor, running Macros, or adding
'    Macros to your Quick Access Toolbar or Ribbon see the
'    Slipstick Systems web site article:
'
'    How to use Outlook's VBA Editor
'    https://www.slipstick.com/developer/how-to-use-outlooks-
'    vba-editor/
'
'- Step 1 - Lock All Views ------------------------------------
'
'    Run the "ViewLock_Form" macro. Select the "System" Scope,
'    "All" Type, check the "What If?" box, and click the "Lock"
'    button.
'
'    If everything runs OK, then uncheck the "What If?" box and
'    click "Lock" again. This locks all the Views on your
'    system.
'
'    Now, if you inadvertently make changes to a locked View,
'    just close any Explorers using that View. When you reopen
'    the View, it will have reverted to the unmodified version.
'
'- Step 2 - Making Changes ------------------------------------
'
'    When you need to edit a Locked View, there are two methods:
'
'    Lock, Edit, Save - Make sure the View is locked and leave
'    it locked, make your changes, and then run Save. The
'    disadvantage of this method is that if you close the
'    Explorer without running Save your changes are lost.
'
'    Unlock, Edit, Save - Unlock the View, make your changes,
'    and Save or Lock it. The disadvantage of this method is
'    that you can't discard any changes because your changes
'    become permanent when you close the Explorer.
'
'- Side Effects on Open Views ---------------------------------
'
'    Changing the state of an open View may sometimes hoark up
'    it 's appearance. Don't Panic. Just close and reopen the
'    Explorer or switch to a different View and back again.
'
'- Lock Pickers -----------------------------------------------
'
'    I have found the following situations where the Locked
'    state of a View is ignored. I'm sure there are more.
'
'    Standard Views - If you have modified and Locked any of the
'    Outlook Standard Views (e.g. Compact, Single, Preview) then
'    doing a "Reset" in the Advanced View Settings dialog
'    ignores any locks.
'
'    Shared Views - These are Views that have "All Xxxx folders"
'    in the "Can Be Used On" column of the Manage All Views
'    dialog. (e.g. "All Mail and Post folders"). There is really
'    only one copy of these Views on your system, even though
'    they appear in multiple Folders. Making a change to the
'    Locked state on any one of them changes all occurrences of
'    that View for that Folder type system wide.
'
'- Legal ------------------------------------------------------
'
'    Copyright (C) 2024 Lycon Of Texas
'
'    This program is free software: you can redistribute it
'    and/or modify it under the terms of the GNU General Public
'    License Version 3 as published by the Free Software
'    Foundation.
'
'    This program is distributed in the hope that it will be
'    useful, but WITHOUT ANY WARRANTY; without even the implied
'    warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR
'    PURPOSE.  See the GNU General Public License for more
'    details.
'
'    You should have received a copy of the GNU General Public
'    License along with this program.  If not, see
'    <https://www.gnu.org/licenses/>.
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
'   Module Level Pseudo ENums
' ---------------------------------------------------------------------
    
    '   WhatIfMsgs - Controls the WhatIf line in MsgBox
    '
    Private Const WhatIfMsgs_First  As Long = 0
    Private Const WhatIfMsgs_None   As Long = 0
    Private Const WhatIfMsgs_Before As Long = 1
    Private Const WhatIfMsgs_After  As Long = 2
    Private Const WhatIfMsgs_Last   As Long = 2
    
    '   Index into the MsgHdr Array
    '
    Private Const MsgHdrs_First     As Long = 0
    Private Const MsgHdrs_Macro     As Long = 0
    Private Const MsgHdrs_Proc      As Long = 1
    Private Const MsgHdrs_Action    As Long = 2
    Private Const MsgHdrs_Folder    As Long = 3
    Private Const MsgHdrs_Scope     As Long = 4
    Private Const MsgHdrs_WhatIf    As Long = 5
    Private Const MsgHdrs_ErrNum    As Long = 6
    Private Const MsgHdrs_ErrDesc   As Long = 7
    Private Const MsgHdrs_Warning   As Long = 8
    Private Const MsgHdrs_Counts    As Long = 9
    Private Const MsgHdrs_ViewName  As Long = 10
    Private Const MsgHdrs_Text      As Long = 11
    Private Const MsgHdrs_Last      As Long = 11
    
    '   Counts List
    '
    Private Const Counts_First                              As Long = 0
    
        Private Const Counts_ActionFirst                    As Long = 0
            Private Const Counts_ViewsUnlocked              As Long = 0
            Private Const Counts_ViewsLocked                As Long = 1
            Private Const Counts_ViewsSaved                 As Long = 2
        Private Const Counts_ActionLast                     As Long = 2
        
        Private Const Counts_SkippedFirst                   As Long = 3
            Private Const Counts_ViewsSkipped_NoChange      As Long = 3
            Private Const Counts_ViewsSkipped_TypeFilter    As Long = 4
            Private Const Counts_ViewsSkipped_SharedSeen    As Long = 5
            Private Const Counts_ViewsSkipped_Error         As Long = 6
            Private Const Counts_ViewsSkipped_CacheGhost    As Long = 7
            Private Const Counts_FoldersSkipped_Pooled      As Long = 8
            Private Const Counts_FoldersSkipped_SharePoint  As Long = 9
            Private Const Counts_FoldersSkipped_Config      As Long = 10
        Private Const Counts_SkippedLast                    As Long = 10
        
        Private Const Counts_ScannedFirst                   As Long = 11
            Private Const Counts_ViewCount                  As Long = 11
            Private Const Counts_FolderCount                As Long = 12
            Private Const Counts_StoreCount                 As Long = 13
        Private Const Counts_ScannedLast                    As Long = 13
        
    Private Const Counts_Last                               As Long = 13
    
    '   Counts Type and Desc Table
    '
    Private Const CountNames_First                  As Long = 0
    Private Const CountNames_Type                   As Long = 0
    Private Const CountNames_Desc                   As Long = 1
    Private Const CountNames_Last                   As Long = 1
        
' ---------------------------------------------------------------------
'   Module Level Constants
' ---------------------------------------------------------------------
    
    Private Const ThisModule                As String = "ViewLock"          '   Name of this Module
    
    '   Hidden Pooled Search Folders
    '
    Private Const PooledNamePrefix          As String = "MS-OLK"
    
    '   PropTags
    '
    Private Const PR_CONTAINER_CLASS        As String = "http://schemas.microsoft.com/mapi/proptag/0x3613001E"
    Private Const PR_EXTENDED_FOLDER_FLAGS  As String = "http://schemas.microsoft.com/mapi/proptag/0x36DA0102"
    
    '   Misc
    '
    Private Const BlankLine                 As String = vbNewLine & vbNewLine       '   Blank Line for MsgBox
    Private Const PAPropertyNotFound        As Long = -2147221233                   '   PropertyAccessor Error - Does not have Property
    Private Const ErrViewNotFound           As String = "The view cannot be found." '   Error - Cached View renamed or deleted

    '   View XML
    '
    '       ViewReadOnly XML elements
    '       ViewTime XML element end tag
    '
    Private Const XMLReadOnlyLocked As String = "<viewreadonly>1</viewreadonly>"
    Private Const XMLReadOnlyUnlocked As String = "<viewreadonly>0</viewreadonly>"
    Private Const XMLTimeEnd As String = "</viewtime>" & vbNewLine
        
' ---------------------------------------------------------------------
'   Module Level Globals
' ---------------------------------------------------------------------

    '   Passed Param and Derived
    '
    Private Action          As Long                     '   ENum Value
    Private ActionBool      As Boolean                  '   Only Lock and Unlock as Boolean
    Private ActionName      As String                   '   From ViewLock_ActionNames(Action)
    
    Private Scope           As Long                     '   ENum Value
    Private ScopeName       As String                   '   From ViewLock_ScopeNames
    Private ScopeShortName  As String                   '   From ViewLock_ScopeShortNames
    Private ScopeString     As String                   '   From ViewLock_ScopeString
    
    Private WhatIf          As Boolean                  '   Value as a Boolean
    Private Form            As ViewLockForm             '   Form (If Open)
    Private Caller          As String                   '   Calling Macro Name
    
    Private VTypes          As Long                     '   VTypes ENum Value
    Private VTypesList      As String                   '   VTypes From ViewLock_VTypesList
    
    '   Current View
    '
    Private CurrentFolder   As Outlook.Folder           '   Current View Folder Object
    Private CurrentView     As Outlook.View             '   Current View Object
    
    '   Collections
    '
    
    '   Share Views I've already Seen. Don't need to be updated.
    '   Item = "", Key = View.Name & vbFormFeed & Folder.IPFRoot
    '
    Private SharedSeen              As VBA.Collection
    
    '   Arrays
    '
    Private Counts()        As Long                     '   Counters Count
    Private CountNames()    As String                   '   Counters Type and Desc
    
' ---------------------------------------------------------------------
'   Public - Macros
' ---------------------------------------------------------------------
    
    Public Sub ViewLock_Lock():     ViewLock_Xeq Actions_Lock, Scopes_View, "ViewLock_Lock":        End Sub
    Public Sub ViewLock_Unlock():   ViewLock_Xeq Actions_Unlock, Scopes_View, "ViewLock_Unlock":    End Sub
    Public Sub ViewLock_State():    ViewLock_Xeq Actions_State, Scopes_View, "ViewLock_State":      End Sub
    Public Sub ViewLock_Save():     ViewLock_Xeq Actions_Save, Scopes_View, "ViewLock_Save":        End Sub
    Public Sub ViewLock_Form():     ViewLock_Xeq Actions_Form, Scopes_View, "ViewLock_Form":        End Sub
    
' ---------------------------------------------------------------------
'   Public Main - Called from Macros and the Form
' ---------------------------------------------------------------------
'
Public Function ViewLock_Xeq( _
    ByVal ActionEnum As Long, _
    ByVal ScopeENum As Long, _
    ByVal CallerName As String, _
    Optional ByVal VTypesENum As Long = -1, _
    Optional ByVal WhatIfBool As Boolean = False _
    )
Const ThisProc = "ViewLock_Xeq"
    
    '   Check/Set the current Environment
    '
    If Not ViewLock_CurrentEnv() Then Exit Function
    
    '   Init Globals
    
    Action = ActionEnum
    ActionBool = IIf(Action = Actions_Lock, True, False)
    ActionName = ViewLock_ActionNames()
    
    Scope = ScopeENum
    ScopeName = ViewLock_ScopeNames()
    ScopeShortName = ViewLock_ScopeShortNames()
    ScopeString = ViewLock_ScopeString()

    VTypes = VTypesENum
    If VTypes = -1 Then VTypes = ViewLock_ViewVType(CurrentView)        '   If no VTypes from Caller - Default to Current View
    VTypesList = ViewLock_VTypesList(VTypes)
    
    Caller = CallerName
    WhatIf = WhatIfBool
    
    ReDim Counts(Counts_First To Counts_Last)
    ViewLock_CountNames
    Counts(Counts_StoreCount) = IIf(Scope <> Scopes_Stores, 1, 0)
    Counts(Counts_FolderCount) = IIf(Scope <> Scopes_View, 0, 1)
    
    Set SharedSeen = New VBA.Collection
    
    '   Dispatch to the appropriate Proc
    '
    ViewLock_XeqDispatch
    
End Function

'   Dispatch based on Action and Scope
'
Private Function ViewLock_XeqDispatch()

    '   Called from the Form to update the Default Form
    '   Status Display after any possible updates.
    '
    If Action = Actions_Status Then
        '   Continue
    
    '   Create and Show the Form. Wait for Hide.
    '   - Form calls back me to execute commands.
    '
    ElseIf Action = Actions_Form Then
        Set Form = New ViewLockForm
        Form.Show
        Set Form = Nothing
        
    '   Do any View Scope commands.
    '
    ElseIf Scope = Scopes_View Then
        ViewLock_ViewScope

    '   Do any Store, Folders, Folder Scope commands
    '
    Else
        ViewLock_StoreFolderScope
        
    End If
    
    '   Update the Default Form Status Display
    '
    ViewLock_FormStatusDisplayDefault

End Function

' ---------------------------------------------------------------------
'   Wide Scope (Folder, Folders, Store, Stores)
' ---------------------------------------------------------------------
'
Private Function ViewLock_StoreFolderScope()
Const ThisProc = "ViewLock_StoreFolderScope"

    '   Warn about changes and get permission

    Dim WarningMsg As String
    
    '   Shared Views Warning
    '
    If (VTypes And VTypes_Shared) <> 0 Then
        WarningMsg = "Shared Views - Changes to a Shared View in the Current Scope will affect" & _
                     " all apperances of that View across the entire system."
    End If
    
    '   Current Scope Warning
    '
    WarningMsg = WarningMsg & BlankLine & "You are about to " & ActionName & " all Views in the Current Scope. "
        
    Select Case ViewLock_MsgBox(ThisProc, _
            WhatIfMsg:=WhatIfMsgs_Before, _
            Warning:=WarningMsg, _
            Text:="Continue?", _
            Buttons:=vbOKCancel, Default:=vbDefaultButton2, _
            Icon:=vbQuestion)
        Case vbOK
            '   Continue
        Case vbCancel
            Exit Function
        Case Else
            Stop: Exit Function
    End Select
    
    '   And awaaaaaay we go!
    '
    ViewLock_FormStatusDisplayCounts
    Select Case Scope
    
        Case Scopes_Stores
            If Not ViewLock_Stores() Then Exit Function
        Case Scopes_Store
            If Not ViewLock_Store(CurrentFolder.Store) Then Exit Function
        Case Scopes_Folders
            If Not ViewLock_Folders(CurrentFolder) Then Exit Function
        Case Scopes_Folder
            If Not ViewLock_Folder(CurrentFolder) Then Exit Function
        Case Else
            Stop: Exit Function
            
    End Select

    '   Show the Wide Scope Show and Tell Msg
    '
    ViewLock_ShowAndTell
    
End Function

'   Walk  all Stores
'
Private Function ViewLock_Stores() As Boolean
Const ThisProc = "ViewLock_Stores"
ViewLock_Stores = False

    Counts(Counts_StoreCount) = 0
    Dim oStore As Outlook.Store
    For Each oStore In Session.Stores
        If Not ViewLock_Store(oStore) Then Exit Function
        Counts(Counts_StoreCount) = Counts(Counts_StoreCount) + 1
    Next oStore
    
ViewLock_Stores = True
End Function

'   Walk all Folders and Search Folders in a Store
'
Private Function ViewLock_Store(ByVal oStore As Outlook.Store) As Boolean
Const ThisProc = "ViewLock_Store"
ViewLock_Store = False

    '   Do the Normal Folders
    '   Do the Search Folders Collection
    '
    If Not ViewLock_Folders(oStore.GetRootFolder) Then Exit Function
    If Not ViewLock_SearchFolders(oStore) Then Exit Function

ViewLock_Store = True
End Function

'   Do a Store level Search Folders Collection
'
Private Function ViewLock_SearchFolders(oStore) As Boolean
ViewLock_SearchFolders = False

    Dim oFolders As Outlook.Folders
    Set oFolders = oStore.GetSearchFolders
    Dim oFolder As Outlook.Folder
    For Each oFolder In oFolders: Do
    
        '   Skip any Pooled Search Folders
        '
        '       Must meet all three criteria.
        '       (All I could find that differ from a normal Search Folder)
        '
        '       1) No PR_CONTAINER_CLASS
        '       2) No PR_EXTENDED_FOLDER_FLAGS
        '       3) Name prefix is "MS-OLK"
        '
        Dim Property As Variant
        If Not ViewLock_GetProperty(oFolder, PR_CONTAINER_CLASS, Property) Then
            If Not ViewLock_GetProperty(oFolder, PR_EXTENDED_FOLDER_FLAGS, Property) Then
                If Left(oFolder.Name, 6) = PooledNamePrefix Then
                    Counts(Counts_FoldersSkipped_Pooled) = Counts(Counts_FoldersSkipped_Pooled) + 1
                    Exit Do ' Next oFolder
                End If
            End If
        End If
        
        '   Search Folders have no subfolders. Go straight to ViewLock_Folder.
        '
        If Not ViewLock_Folder(oFolder) Then Exit Function
        
    Loop While False: Next oFolder

ViewLock_SearchFolders = True
End Function

'   Do the current Folder and then a Recursive Descent into all SubFolders
'
Private Function ViewLock_Folders(ByVal oFolder As Outlook.Folder) As Boolean
Const ThisProc = "ViewLock_Folders"
ViewLock_Folders = False

    '   Skip any SharePoint Folders
    '
    If oFolder.IsSharePointFolder Then
        Counts(Counts_FoldersSkipped_SharePoint) = Counts(Counts_FoldersSkipped_SharePoint) + 1
        ViewLock_Folders = True
        Exit Function
    End If
    
    '   Skip any Config Folders
    '
    If ViewLock_FolderIPFRoot(oFolder) = "Configuration" Then
        Counts(Counts_FoldersSkipped_Config) = Counts(Counts_FoldersSkipped_Config) + 1
        ViewLock_Folders = True
        Exit Function
    End If
    
    '   Do the current folder
    '
    If Not ViewLock_Folder(oFolder) Then Exit Function
    
    '   If has no Folders Collection - Done
    '
    Dim Dummy As Outlook.Folders
    On Error Resume Next
        Set Dummy = oFolder.Folders
        If Err.Number <> 0 Then
            ViewLock_Folders = True
            Exit Function
        End If
    On Error GoTo 0
    
    '   Call myself for any subfolders
    '
    For Each oFolder In oFolder.Folders
        If Not ViewLock_Folders(oFolder) Then Exit Function
    Next oFolder

ViewLock_Folders = True
End Function

'   Do all Views in a Folder
'
Private Function ViewLock_Folder(ByVal oFolder As Outlook.Folder) As Boolean
Const ThisProc = "ViewLock_Folder"
ViewLock_Folder = False

    '   Skip any SharePoint Folders
    '
    If oFolder.IsSharePointFolder Then
        Counts(Counts_FoldersSkipped_SharePoint) = Counts(Counts_FoldersSkipped_SharePoint) + 1
        ViewLock_Folder = True
        Exit Function
    End If
    
    Counts(Counts_FolderCount) = Counts(Counts_FolderCount) + 1
    
    '   Call StateChange for each View in the Folder
    '
    Dim oView As Outlook.View
    For Each oView In oFolder.Views
        ViewLock_StateChange oView
    Next oView

ViewLock_Folder = True
End Function

'   Show the Wide Scope Show and Tell
'
Private Function ViewLock_ShowAndTell()
Const ThisProc = "ViewLock_ShowAndTell"

    '   Put back the Default Form Status Diaplay
    '
    ViewLock_FormStatusDisplayDefault
    
    '   Show Counts and an Information Icon
    '
    ViewLock_MsgBox ThisProc, _
        WhatIfMsg:=WhatIfMsgs_After, _
        Counts:=ViewLock_ShowAndTellCounts(), _
        Icon:=vbInformation

End Function

'   Build the Counters section of the Wide Scope Show and Tell
'
Private Function ViewLock_ShowAndTellCounts() As String

    Dim Inx As Long
    
    Dim CountsLit(Counts_First To Counts_Last) As String
    For Inx = Counts_First To Counts_Last
        CountsLit(Inx) = _
            CStr(Counts(Inx)) & " " & _
            CountNames(Inx, CountNames_Type) & IIf(Counts(Inx) <> 1, "s", "") & " " & _
            CountNames(Inx, CountNames_Desc)
    Next Inx
    
    Dim CountsLine(Counts_First To Counts_Last) As String
    For Inx = Counts_First To Counts_Last
    
        Select Case Inx
        
            '   Either Unlocked or Locked. Even if Zero.
            '
            Case Counts_ActionFirst To Counts_ActionLast
                CountsLine(Inx) = _
                    IIf(Inx = Action, CountsLit(Inx), "")
                
            '   Skipped Lines - Only if Non Zero
            '
            Case Counts_SkippedFirst To Counts_SkippedLast
                If Counts(Inx) <> 0 Then CountsLine(Inx) = CountsLit(Inx)
            
            '   Last Line - Scanned Line of concats
            '
            Case Counts_ScannedFirst To Counts_ScannedLast
                CountsLine(Counts_Last) = CountsLine(Counts_Last) & CountsLit(Inx) & " "
            
            Case Else
                '   Oops
                Stop: Exit Function
                
        End Select
    
    Next Inx
    
    '   Remove trailing space from the Last (Scanned) line
    '
    CountsLine(Counts_Last) = Left(CountsLine(Counts_Last), Len(CountsLine(Counts_Last)) - 1)
    
    Dim CountsBlock As String
    For Inx = Counts_First To Counts_Last
        If CountsLine(Inx) <> "" Then CountsBlock = CountsBlock & CountsLine(Inx) & "." & vbNewLine
        If Inx = Action Then CountsBlock = CountsBlock & vbNewLine
    Next Inx
    
    '   Remove trailing vbNewLine from the Block and Return
    '
    ViewLock_ShowAndTellCounts = Left(CountsBlock, Len(CountsBlock) - 2)
    
End Function

' =====================================================================
'   View Scope
' =====================================================================
'
Private Function ViewLock_ViewScope()

    Select Case Action
        Case Actions_State
            ViewLock_ViewState
        Case Actions_Lock, Actions_Unlock
            ViewLock_ViewSet
        Case Actions_Save
            ViewLock_ViewSave
        Case Else
            Stop: Exit Function
    End Select

End Function

'   Save the Current View
'
Private Function ViewLock_ViewSave()
Const ThisProc = "ViewLock_ViewSave"

    '   Warnings
    '
    Dim WarningMsg As String
    If (VTypes And VTypes_Shared) <> 0 Then
            WarningMsg = "Current View is Shared - Changes to a Shared View will affect" & _
                         " all apperances of that View across the entire system."
    End If
    
    '   Ask Permission to Save
    '
    Select Case ViewLock_MsgBox(ThisProc, _
            WhatIfMsg:=WhatIfMsgs_Before, _
            ViewName:=CurrentView.Name, _
            Warning:=WarningMsg, _
            Text:= _
                "You are about to " & ActionName & " any changes to the Current View and Lock it." & BlankLine & _
                "Continue?", _
            Buttons:=vbOKCancel, Default:=vbDefaultButton2, _
            Icon:=vbQuestion)
        Case vbOK
            '   Continue
        Case vbCancel
            Exit Function
        Case Else
            '   Oops
            Stop: Exit Function
    End Select
    
    '   Ignore Current State and just Lock and Save
    '   (oView.Save ignores oView.LockUserChanges)
    '
    ActionBool = True
    If Not ViewLock_StateSave(oView:=CurrentView, IncChangedCount:=False, SetLock:=True, SetXML:=False) Then Stop: Exit Function

    '   Show and Tell
    '
    ViewLock_MsgBox Proc:=ThisProc, _
        WhatIfMsg:=WhatIfMsgs_After, _
        ViewName:=CurrentView.Name, _
        Warning:="If the View appears hoarked - Just close and reopen the current Explorer or switch to a different View and back again.", _
        Text:="View is now Saved and " & ViewLock_LockStateName(CurrentView) & ".", _
        Icon:=vbInformation

End Function

'   Show the State of the Current View
'
Private Function ViewLock_ViewState()
Const ThisProc = "ViewLock_ViewState"
    
    '   Show and Tell (SAT)
    '
    ViewLock_MsgBox Proc:=ThisProc, _
        ViewName:=CurrentView.Name, _
        Text:="View is " & ViewLock_LockStateName(CurrentView) & ".", _
        Icon:=vbInformation

End Function

'   Set the State of the Current View
'
Private Function ViewLock_ViewSet()
Const ThisProc = "ViewLock_ViewSet"
    
    '   If already in the requested State - done
    '
    If CurrentView.LockUserChanges = ActionBool Then
        ViewLock_MsgBox Proc:=ThisProc, _
            ViewName:=CurrentView.Name, _
            Text:="View is already " & ViewLock_LockStateName(CurrentView) & ".", _
            Icon:=vbInformation
        ViewLock_ViewSet = True
        Exit Function
    End If
    
    '   Warning if a Shared View and get Permisison
    '
    If Not ViewLock_SharedWarning(ThisProc) Then Exit Function
    
    '   Set it
    '
    ViewLock_StateChange CurrentView
    
    '   Setup the current Environment
    '   Show and Tell
    '
    ViewLock_MsgBox Proc:=ThisProc, _
        WhatIfMsg:=WhatIfMsgs_After, _
        ViewName:=CurrentView.Name, _
        Warning:="If the View appears hoarked - Just close and reopen the current Explorer or switch to a different View and back again.", _
        Text:="View is now " & ViewLock_LockStateName(CurrentView) & ".", _
        Icon:=vbInformation

End Function

'   Show a View Scope Shared View Lock/Unlock Warning and get Permission
'
Private Function ViewLock_SharedWarning(ByVal Caller As String) As Boolean
ViewLock_SharedWarning = False

    '   If not a Shared View - True and Done
    '
    If (VTypes And VTypes_Shared) = 0 Then
        ViewLock_SharedWarning = True
        Exit Function
    End If
    
    '   Ask for Permission. Return True if OK.
    '
    Dim WarningMsg As String
        WarningMsg = "Current View is Shared - Changes to a Shared View will affect" & _
                     " all apperances of that View across the entire system."
                     
    Select Case ViewLock_MsgBox(Caller, _
            WhatIfMsg:=WhatIfMsgs_Before, _
            ViewName:=CurrentView.Name, _
            Warning:=WarningMsg, _
            Text:="Continue?", _
            Buttons:=vbOKCancel, Default:=vbDefaultButton2, _
            Icon:=vbQuestion)
        Case vbOK
            '   Continue
        Case vbCancel
            Exit Function
        Case Else
            Stop: Exit Function
    End Select

ViewLock_SharedWarning = True
End Function

' =====================================================================
'   State Change
' =====================================================================

'   Handle a possible View State Change
'
Private Function ViewLock_StateChange(ByVal oView As Outlook.View)
Const ThisProc = "ViewLock_StateChange"

    DoEvents
    Counts(Counts_ViewCount) = Counts(Counts_ViewCount) + 1
    ViewLock_FormStatusDisplayCounts

    '   If it doesn't pass the VTypes filter - Done
    '
    If Not ViewLock_StateVTypeFilter(oView) Then Exit Function
    
    ' ---------------------------------------------------------------------
    '   Break out Explorer.CurrentView handling because it's so
    '   different from the normal flow.
    '
    Select Case ViewLock_StateExplorerCurrent(oView)
        Case True
            '   Continue
        
        Case False
        
            '   If it's a Shared View I've seen before - Inc Counter and Done
            '   If there is no State Change - Inc Counter and Done
            '   Save the View
            '
            If ViewLock_StateSharedSeen(oView) Then Exit Function
            If Not ViewLock_StateChanging(oView) Then Exit Function
            If Not ViewLock_StateSave(oView:=oView, IncChangedCount:=True, SetLock:=False, SetXML:=True) Then Stop: Exit Function
            
        Case Else
            ' Oops
            Stop: Exit Function
    End Select
    ' ---------------------------------------------------------------------
    
End Function

'   Update the View.XML
'
'       SPOS - Testing has shown that just setting oView.LockUserChanges is not
'       reliable for Views other than Explorer.CurrentView, so we update the XML as well.
'
Private Function ViewLock_StateXML(ByVal oView As Outlook.View) As Boolean
ViewLock_StateXML = False

    '   SPOS - Stupid doesn't look at the value of <viewreadonly>.
    '   Only if it exist. If it exist then LockUserChanges is True Else False.
    '
    '   Read in the View's XML.
    '   Remove any Read Only True and any spurious Read Only False elements
    '   If Action is Lock - Insert a Read Only element just after the </viewtime> tag
    '
    Dim ViewXML As String
    ViewXML = oView.XML
    ViewXML = Replace(ViewXML, XMLReadOnlyLocked, "")
    ViewXML = Replace(ViewXML, XMLReadOnlyUnlocked, "")
    If ActionBool Then ViewXML = Replace(ViewXML, XMLTimeEnd, XMLTimeEnd & XMLReadOnlyLocked)
    oView.XML = ViewXML

ViewLock_StateXML = True
End Function

'   Save (with trap in case something goes sideways)
'
Private Function ViewLock_StateSave( _
ByVal oView As Outlook.View, _
ByVal IncChangedCount As Boolean, _
ByVal SetLock As Boolean, _
ByVal SetXML As Boolean _
) As Boolean
Const ThisProc = "ViewLock_StateSave"
ViewLock_StateSave = True

    '   If WhatIf - pretend we made a change and done
    '
    If WhatIf Then
        If IncChangedCount Then Counts(Action) = Counts(Action) + 1
        Exit Function
    End If

    '   If called for - Update LockUserChanges
    '   If called for - Update the View.XML
    '
    If SetLock Then oView.LockUserChanges = ActionBool
    If SetXML Then If Not ViewLock_StateXML(oView) Then Stop: Exit Function

    '   Finally !
    '
    On Error Resume Next
        oView.Save
        If Err.Number = 0 Then
            If IncChangedCount Then Counts(Action) = Counts(Action) + 1
            Exit Function
        End If
    On Error GoTo 0
            
    Select Case ViewLock_MsgBox( _
            Proc:=ThisProc, _
            ErrNum:=Err.Number, _
            ErrDesc:=Err.Description, _
            ViewName:=oView.Name, _
            Folder:=oView.Parent.Parent, _
            Text:="Error saving View. Continue processing?", _
            Buttons:=vbYesNo, _
            Default:=vbDefaultButton2, _
            Icon:=vbCritical)
        Case vbYes
            Counts(Counts_ViewsSkipped_Error) = Counts(Counts_ViewsSkipped_Error) + 1
            Exit Function
        Case vbNo
            '   Continue
        Case Else
            '   Oops
            Stop: Exit Function
    End Select
    
ViewLock_StateSave = False
End Function

'   Handle Explorer.CurrentView
'
'   WTF?
'
'       Testing shows that I need to update LockUserChanges in all Explorers where
'       Explorer.CurrentView is using oView. (I assume Explorer.CurrentView is a
'       cached copy of the oView from Folder.Views).
'
'       But only one Explorer.CurrentView IS oView. All the others will match on
'       FolderPath and View.Name, but have a different Object reference.
'
'   https://learn.microsoft.com/en-us/office/vba/api/outlook.explorer.currentview
'
'       To obtain a View object for the view of the current Explorer, use
'       Explorer.CurrentView instead of the CurrentView property of the current Folder
'       object returned by Explorer.CurrentFolder.
'
'       You must save (Set) a reference to the View object returned by CurrentView before
'       you proceed to use it for any purpose.
'
Private Function ViewLock_StateExplorerCurrent(ByVal oView As Outlook.View) As Boolean
ViewLock_StateExplorerCurrent = False

    Dim ExplorersUpdated As Long    '   How many Explorers with this View have we updated?

    '   Walk all Explorers looking for a match to oView
    '
    Dim oExplorerView As Outlook.View
    Dim oExplorer As Outlook.Explorer
    For Each oExplorer In Application.Explorers: Do
    
        '   Ignore Explorers with no View (e.g. Outlook Today)
        '
        Dim GetViewError As Boolean
        On Error Resume Next
            Set oExplorerView = oExplorer.CurrentView
            GetViewError = (Err.Number <> 0)
        On Error GoTo 0
        If GetViewError Then Exit Do ' Next oExplorer
        
        '   Have to look for the View by View.Name (and FolderPath if not Shared)
        '   because if there are  multiple Explorers open to the same View, only
        '   one of them has the same Object reference as oView.

        '   If Name doesn't match - Next oExplorer
        '
        If oView.Name <> oExplorerView.Name Then Exit Do ' Next oExplorer
        
        '   If Not Shared
        '   - If FolderPath doesn't match - Next Explorer
        '
        If (ViewLock_ViewVType(oView) And VTypes_Shared) = 0 Then
            If oView.Parent.Parent.FolderPath <> oExplorerView.Parent.Parent.FolderPath Then Exit Do ' Next oExplorer
        End If
        
        '   We've found oView in at least one Explorer - will Return True
        '
        ViewLock_StateExplorerCurrent = True
        
        '   If first occurance
        '   - Run it through SharedSeen
        '   - If Not SharedSeen - Update Skipped Count
        '
        If ExplorersUpdated = 0 Then
            If Not ViewLock_StateSharedSeen(oView) Then
                If oExplorerView.LockUserChanges = ActionBool Then
                    Counts(Counts_ViewsSkipped_NoChange) = Counts(Counts_ViewsSkipped_NoChange) + 1
                End If
            End If
        End If
        
        '   If No Change - Next oExplorer
        '
        '   - Using LockUserChanges instead of XML because in this case the
        '   - XML is unreliable.
        '
        If oExplorerView.LockUserChanges = ActionBool Then
            ExplorersUpdated = ExplorersUpdated + 1
            Exit Do ' Next oExplorer
        End If
        
        '   Make the changes
        '   - Inc Chaged Count for the first occurance only.
        '   - Update only LockUserChanges, not the XML. Stupid will do that (hopefully).
        '
        If Not ViewLock_StateSave( _
            oView:=oExplorerView, _
            SetLock:=True, _
            SetXML:=False, _
            IncChangedCount:=(ExplorersUpdated = 0) _
            ) Then Stop: Exit Function
        
        ExplorersUpdated = ExplorersUpdated + 1
        
    Loop While False: Next oExplorer

End Function

'   Is the State of oView changing?
'
'   If Is Changing - Return True. Else - Update ViewsSkipped and Return False.
'
'   SPOS - Testing has shown that the State of oView.LockUserChanges is not
'   reliable for Views other than Explorer.CurrentView, so we check the XML.
'
Private Function ViewLock_StateChanging(ByVal oView As Outlook.View) As Boolean

    On Error Resume Next
    
        '   If the XML ReadOnly <> Current Action Boolean - ViewLock_StateChanging = True
        '
        ViewLock_StateChanging = ((InStr(1, oView.XML, XMLReadOnlyLocked) <> 0) <> ActionBool)
        
        '   If the Get oView.XML threw an Error
        '
        If Err.Number <> 0 Then
        
            '   If a Cached View that was renamed or deleted but Stupid is holding a copy
            '   - Inc the skipped count and Return False
            '
            If Err.Description = ErrViewNotFound Then
                Counts(Counts_ViewsSkipped_CacheGhost) = Counts(Counts_ViewsSkipped_CacheGhost) + 1
                ViewLock_StateChanging = False
                Exit Function
            End If
            
            Stop: Exit Function
            
        End If
        
    On Error GoTo 0
    
    '   If changing - Return True
    '   Else - Inc the skipped count and Return False
    '
    If ViewLock_StateChanging Then Exit Function
    Counts(Counts_ViewsSkipped_NoChange) = Counts(Counts_ViewsSkipped_NoChange) + 1

End Function

'   Does oView pass the current VTypes filter?
'
'   If oView passes the current VTypes filter - Return True
'   Else Inc Skipped Counter and Return False
'
Private Function ViewLock_StateVTypeFilter(ByVal oView As Outlook.View) As Boolean

    ViewLock_StateVTypeFilter = ((ViewLock_ViewVType(oView) And VTypes) <> 0)
    If ViewLock_StateVTypeFilter Then Exit Function
    
    Counts(Counts_ViewsSkipped_TypeFilter) = Counts(Counts_ViewsSkipped_TypeFilter) + 1

End Function

'   If oView is not a Shared View - Return False
'   If oView is a Shared View that I've already seen - Inc Counter and Return True
'   Else add it to the SharedSeen collection and Return False
'
Private Function ViewLock_StateSharedSeen(ByVal oView As Outlook.View) As Boolean
ViewLock_StateSharedSeen = False

    '   If not a Shared View - Return False
    '
    If (ViewLock_ViewVType(oView) And VTypes_Shared) = 0 Then Exit Function

    '   HaveSeen = Is oView in SharedSeen?
    '
    Dim SeenKey As String: SeenKey = oView.Name & vbFormFeed & ViewLock_FolderIPFRoot(oView.Parent.Parent)
    Dim HaveSeen As Boolean
    On Error Resume Next
        SharedSeen.Item SeenKey
        HaveSeen = (Err.Number = 0)
    On Error GoTo 0

    '   If I've seen this one before - Inc Counter and Return True
    '
    If HaveSeen Then
        Counts(Counts_ViewsSkipped_SharedSeen) = Counts(Counts_ViewsSkipped_SharedSeen) + 1
        ViewLock_StateSharedSeen = True
        Exit Function
    End If
    
    '   Else add it to the Collection and Return False
    '
    SharedSeen.Add "", SeenKey
 
End Function

' =====================================================================
'   Misc
' =====================================================================

'   Check/Set the current Environment.
'   Return False if not to my liking.
'
Private Function ViewLock_CurrentEnv() As Boolean
Const ThisProc = "ViewLock_CurrentEnv"
ViewLock_CurrentEnv = False

    Set CurrentView = Nothing
    Set CurrentFolder = Nothing
    
    On Error GoTo Error_Exit
        Set CurrentView = ActiveWindow.CurrentView
        Set CurrentFolder = CurrentView.Parent.Parent
    On Error GoTo 0
    
    ViewLock_CurrentEnv = True
    Exit Function

Error_Exit:

    MsgBox _
        "Err.Num = " & Err.Number & "  (0x" & Hex(Err.Number) & ")" & vbNewLine & _
        "Err.Desc = " & Err.Description & vbNewLine & vbNewLine & _
        "The Active Outlook Window must be an Outlook Folder Explorer.", vbOKOnly, "ViewLock"
    
End Function

'   Set the Status Display Text to Running Counts
'
Private Function ViewLock_FormStatusDisplayCounts()

    '   Get the Wide Scrop Show and Tell Counts block
    '   Pick off the first and last lines (Changed & Scanned Counts)
    '
    Dim ShowAndTellCounts() As String
    ShowAndTellCounts = Split(ViewLock_ShowAndTellCounts(), vbNewLine)
    
    '   Update the Form Status Display
    '
    ViewLock_FormStatusDisplay _
        "Scanning ..." & vbNewLine & _
        ShowAndTellCounts(0) & vbNewLine & _
        ShowAndTellCounts(UBound(ShowAndTellCounts))

End Function
'   Set the Status Display Text to default
'
Private Function ViewLock_FormStatusDisplayDefault()

    ViewLock_FormStatusDisplay _
        "Current Folder: '" & CurrentFolder.FolderPath & "'." & vbNewLine & _
        "Current View  : '" & _
            CurrentView.Name & "'. " & ViewLock_LockStateName(CurrentView) & ". " & _
            "(" & ViewLock_VTypesList(ViewLock_ViewVType(CurrentView)) & ")." & vbNewLine & _
        "Current Scope : " & ViewLock_ScopeString

End Function

'   If the Form is Open - Set the StatusDisplay Text
'
Private Function ViewLock_FormStatusDisplay(ByVal Text As String)

    If Form Is Nothing Then Exit Function
    Form.StatusDisplay = Text

End Function

'   Get the VTypes of a View
'
Private Function ViewLock_ViewVType(ByVal oView As Outlook.View) As Long
    
    Select Case oView.SaveOption
        Case olViewSaveOptionAllFoldersOfType
            ViewLock_ViewVType = VTypes_Shared
        Case olViewSaveOptionThisFolderEveryone
            ViewLock_ViewVType = VTypes_Public
        Case olViewSaveOptionThisFolderOnlyMe
            ViewLock_ViewVType = VTypes_Private
        Case Else
            '   Ooops
            Stop: Exit Function
    End Select

End Function

Private Function ViewLock_ScopeNames() As String

    Dim ScopeNames(Scopes_First To Scopes_Last)
    ScopeNames(Scopes_Stores) = "All Views on the System."
    ScopeNames(Scopes_Store) = "All Views in the Current Store (.pst)."
    ScopeNames(Scopes_Folders) = "All Views in the Current Folder and SubFolders."
    ScopeNames(Scopes_Folder) = "All Views in the Current Folder."
    ScopeNames(Scopes_View) = "Current View."
    
    ViewLock_ScopeNames = ScopeNames(Scope)
    
End Function

Private Function ViewLock_ScopeShortNames() As String

    Dim ScopeShortNames(Scopes_First To Scopes_Last) As String
    ScopeShortNames(Scopes_Stores) = "System"
    ScopeShortNames(Scopes_Store) = "Current Store (.pst)"
    ScopeShortNames(Scopes_Folders) = "Current Folder and All SubFolders"
    ScopeShortNames(Scopes_Folder) = "Current Folder"
    ScopeShortNames(Scopes_View) = "Current View"
    
    ViewLock_ScopeShortNames = ScopeShortNames(Scope)
    
End Function

Private Function ViewLock_ActionNames() As String

    Dim ActionNames(Actions_First To Actions_Last) As String
    ActionNames(Actions_Unlock) = "Unlock"
    ActionNames(Actions_Lock) = "LOCK"
    ActionNames(Actions_State) = "State"
    ActionNames(Actions_Save) = "Save"
    ActionNames(Actions_Status) = "Status"
    
    ViewLock_ActionNames = ActionNames(Action)
    
End Function

Private Function ViewLock_CountNames()

    ReDim CountNames(Counts_First To Counts_Last, CountNames_First To CountNames_Last)
    
    CountNames(Counts_ViewsUnlocked, CountNames_Type) = "View"
    CountNames(Counts_ViewsUnlocked, CountNames_Desc) = "Unlocked"
    CountNames(Counts_ViewsLocked, CountNames_Type) = "View"
    CountNames(Counts_ViewsLocked, CountNames_Desc) = "LOCKED"
    
    CountNames(Counts_ViewsSkipped_NoChange, CountNames_Type) = "View"
    CountNames(Counts_ViewsSkipped_NoChange, CountNames_Desc) = "Skipped - No Change"
    CountNames(Counts_ViewsSkipped_TypeFilter, CountNames_Type) = "View"
    CountNames(Counts_ViewsSkipped_TypeFilter, CountNames_Desc) = "Skipped - Type Filter"
    CountNames(Counts_ViewsSkipped_SharedSeen, CountNames_Type) = "View"
    CountNames(Counts_ViewsSkipped_SharedSeen, CountNames_Desc) = "Skipped - Shared Duplicate"
    CountNames(Counts_ViewsSkipped_Error, CountNames_Type) = "View"
    CountNames(Counts_ViewsSkipped_Error, CountNames_Desc) = "Skipped - Error Setting State"
    CountNames(Counts_ViewsSkipped_CacheGhost, CountNames_Type) = "View"
    CountNames(Counts_ViewsSkipped_CacheGhost, CountNames_Desc) = "Skipped - Cache Ghost"
    
    CountNames(Counts_FoldersSkipped_Pooled, CountNames_Type) = "Folder"
    CountNames(Counts_FoldersSkipped_Pooled, CountNames_Desc) = "Skipped - Hidden Pooled Search Folder"
    CountNames(Counts_FoldersSkipped_SharePoint, CountNames_Type) = "Folder"
    CountNames(Counts_FoldersSkipped_SharePoint, CountNames_Desc) = "Skipped - SharePoint Folder"
    CountNames(Counts_FoldersSkipped_Config, CountNames_Type) = "Folder"
    CountNames(Counts_FoldersSkipped_Config, CountNames_Desc) = "Skipped - Configuration Folder"
    
    CountNames(Counts_ViewCount, CountNames_Type) = "View"
    CountNames(Counts_ViewCount, CountNames_Desc) = "in"
    CountNames(Counts_FolderCount, CountNames_Type) = "Folder"
    CountNames(Counts_FolderCount, CountNames_Desc) = "in"
    CountNames(Counts_StoreCount, CountNames_Type) = "Store"
    CountNames(Counts_StoreCount, CountNames_Desc) = "Scanned"

End Function

'   Return a string (All, Shared, Public, Private) from VTypes
'
Private Function ViewLock_VTypesList(ByVal VTypes As Long) As String

    Select Case VTypes
    
        Case VTypes_None
            ViewLock_VTypesList = "No"
        Case VTypes_All
            ViewLock_VTypesList = "All"
        Case Else
            ViewLock_VTypesList = _
                IIf(VTypes And VTypes_Shared, "Shared, ", "") & _
                IIf(VTypes And VTypes_Public, "Public, ", "") & _
                IIf(VTypes And VTypes_Private, "Private, ", "")
            ViewLock_VTypesList = Left(ViewLock_VTypesList, Len(ViewLock_VTypesList) - 2)
            
    End Select
    
End Function

'   Return a string ("LOCKED" or "Unlocked") based on oView.LockUserChanges
'
Private Function ViewLock_LockStateName(ByVal oView As Outlook.View) As String

    ViewLock_LockStateName = IIf(oView.LockUserChanges, "LOCKED", "Unlocked")
    
End Function

'   Return a string based on the Scope, VTypes, Current Folder and View
'
Private Function ViewLock_ScopeString() As String

    Select Case Scope
        Case Scopes_Stores
            ViewLock_ScopeString = "System."
        Case Scopes_Store
            ViewLock_ScopeString = "Store (.pst) '" & CurrentFolder.Store & "'."
        Case Scopes_Folders
            ViewLock_ScopeString = "Folder '" & CurrentFolder.FolderPath & "' and Subfolders."
        Case Scopes_Folder
            ViewLock_ScopeString = "Folder '" & CurrentFolder.FolderPath & "'."
        Case Scopes_View
            ViewLock_ScopeString = "View '" & CurrentView.Name & "'."
        Case Else
            Stop: Exit Function
    End Select
    
    If Scope = Scopes_View Then
        ViewLock_ScopeString = ViewLock_ScopeString & " (" & ViewLock_VTypesList(VTypes) & ")."
    Else
        ViewLock_ScopeString = ViewLock_ScopeString & " " & ViewLock_VTypesList(VTypes) & " Views."
    End If

End Function

'   Get the Folder Type (IPFRoot) for a Folder
'
Private Function ViewLock_FolderIPFRoot(ByVal oFolder As Outlook.Folder) As String

    '   Get the Property Raw value
    '   Property does not exist or "" -> assume IPM.Note
    '
    Dim oPA As Outlook.PropertyAccessor
    Set oPA = oFolder.PropertyAccessor
    Dim PropValue As String

    On Error Resume Next
        PropValue = oPA.GetProperty(PR_CONTAINER_CLASS)
        Select Case Err.Number
            Case 0
                ' Continue
            Case PAPropertyNotFound
                PropValue = "IPF.Note"
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0
    If PropValue = "" Then PropValue = "IPF.Note"

    '   Must start with "IPF."
    '
    If Left(PropValue, 4) <> "IPF." Then Stop: Exit Function
    
    '   Get only the second piece - the IPFRoot
    '
    Dim IPFRoot As String
    IPFRoot = Mid(PropValue, 5)
    IPFRoot = Mid(IPFRoot, 1, InStr(1, IPFRoot & ".", ".") - 1)
    
    ViewLock_FolderIPFRoot = IPFRoot

End Function

'   Get a PropTag value from an Item
'
Private Function ViewLock_GetProperty(ByVal Item As Object, ByVal PropTag As String, ByRef Value As Variant) As Boolean
ViewLock_GetProperty = False

    Dim PA As Outlook.PropertyAccessor
    
    On Error GoTo ErrExit
    
        Set PA = Item.PropertyAccessor
        Value = PA.GetProperty(PropTag)

    On Error GoTo 0
    
ViewLock_GetProperty = True
ErrExit: End Function

'   Show a MsgBox and get any responce
'
Private Function ViewLock_MsgBox( _
    ByVal Proc As String, _
    Optional ByVal Folder As String, _
    Optional ByVal WhatIfMsg As Long = WhatIfMsgs_None, _
    Optional ByVal ErrNum As Long, _
    Optional ByVal ErrDesc As String, _
    Optional ByVal Warning As String, _
    Optional ByVal Counts As String, _
    Optional ByVal ViewName As String, _
    Optional ByVal Text As String, _
    Optional ByVal Buttons As Long = vbOKOnly, _
    Optional ByVal Default As Long = vbDefaultButton1, _
    Optional ByVal Icon As Long = vbExclamation _
    ) As Integer

    '   Make the box fixed width and as wide as possible.
    '
    Dim Title As String
    Title = ThisModule & Space(100)
    
    '   Build the Msg Header Lines
    '
    Dim MsgHdr(MsgHdrs_First To MsgHdrs_Last) As Variant
    
    '   Always
    '
    If Caller <> "" Then MsgHdr(MsgHdrs_Macro) = "Caller: '" & Caller & "'."
    If Proc <> "" Then MsgHdr(MsgHdrs_Proc) = "Proc: '" & Proc & "'."
    If ActionName <> "" Then MsgHdr(MsgHdrs_Action) = "Action: " & ActionName & "."
    If Folder <> "" Then
        MsgHdr(MsgHdrs_Folder) = vbNewLine & "Folder: '" & Folder & "'."
    Else
        MsgHdr(MsgHdrs_Folder) = vbNewLine & "Current Folder: '" & CurrentFolder.FolderPath & "'."
    End If
    If ScopeString <> "" Then MsgHdr(MsgHdrs_Scope) = vbNewLine & "Current Scope: " & ScopeString
    
    '   Pick the right WhatIf message
    '
    If WhatIf Then
    
        MsgHdr(MsgHdrs_WhatIf) = vbNewLine & "!!  WhatIf is TRUE. No changes "
        
        Select Case WhatIfMsg
            Case WhatIfMsgs_Before
                MsgHdr(MsgHdrs_WhatIf) = MsgHdr(MsgHdrs_WhatIf) & "will be made.  !!"
            Case WhatIfMsgs_After
                MsgHdr(MsgHdrs_WhatIf) = MsgHdr(MsgHdrs_WhatIf) & "were made.  !!"
            Case WhatIfMsgs_None
                MsgHdr(MsgHdrs_WhatIf) = ""
            Case Else
                Stop: Exit Function
        End Select
        
    End If
    
    If ErrNum <> 0 Then MsgHdr(MsgHdrs_ErrNum) = "Err.Number: " & ErrNum & " (0x" & Hex(ErrNum) & ")."
    If ErrDesc <> "" Then MsgHdr(MsgHdrs_ErrDesc) = vbNewLine & "Error: '" & ErrDesc & "'."
    If Warning <> "" Then MsgHdr(MsgHdrs_Warning) = vbNewLine & Warning
    If Counts <> "" Then MsgHdr(MsgHdrs_Counts) = vbNewLine & Counts
    If ViewName <> "" Then MsgHdr(MsgHdrs_ViewName) = vbNewLine & "View: '" & ViewName & "'."
    If Text <> "" Then MsgHdr(MsgHdrs_Text) = vbNewLine & Text
    
    '   Build the Msg Header String
    '
    Dim Header As String
    Dim HdrInx As Long
    For HdrInx = MsgHdrs_First To MsgHdrs_Last
        If MsgHdr(HdrInx) <> "" Then Header = Header & vbNewLine & MsgHdr(HdrInx)
    Next HdrInx
    Header = Mid(Header, 3)
    
    '   Show the box and return any responce
    '
    ViewLock_MsgBox = MsgBox(Header, Buttons + Default + Icon, Title)

End Function
