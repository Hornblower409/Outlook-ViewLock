![Ribbon Custom Group](https://github.com/user-attachments/assets/e8f26436-e578-4076-b3dd-c5ecb1ac61bf)
# Outlook-ViewLock
Lock Outlook Views, preventing any accidental changes from being saved when you close the Explorer.

## Purpose
To compensate for the lack of a "Save Changes?" step after modifying the Settings of an Outlook View.

ViewLock can lock a View, preventing any accidental changes from being saved when you close the Explorer, and allows you to Edit a View with the option to not save your changes.

When you make changes to an Outlook View there is no way to cancel or undo them. If you accidentally click on a column heading it permanently changes the Sort order for that View. If you have modified any of the Advanced View Settings (Columns, Group By, Sort, etc.) and don't like the new design, there's no way to undo your changes once you hit "OK" on the Settings dialog. There are ways to backup and restore a View, but they are manual and tedious.

## Install
For help on using the VBA Editor, running Macros, or adding Macros to your Quick Access Toolbar or Ribbon see the Slipstick Systems web site article: [How to use Outlook's VBA Editor](https://www.slipstick.com/developer/how-to-use-outlooks-vba-editor/)

This is a standalone Module with no external references and one Form. To install from the VBA Editor do:

- File -> Import: ViewLock.bas
- File -> Import: ViewLockForm.frm

The Module has five Macros:

- ViewLock_Lock
- ViewLock_Unlock
- ViewLock_State
- ViewLock_Save
- ViewLock_Form

The first four operate only on the current View. ViewLock_Form opens a User Form that allows you Lock/Unlock Views in the current Folder, the current Folder and all it's subfolders, the current Store (.pst file) or the entire system.

## Using ViewLock
### Step 1 - Lock All Views
Run the "ViewLock_Form" macro. Select the "System" Scope, "All" Type, check the "What If?" box, and click the "Lock" button. If everything runs OK, then uncheck the "What If?" box and click "Lock" again. This locks all the Views on your system.

Now, if you inadvertently make changes to a locked View, just close any Explorers using that View. When you reopen the View, it will have reverted to the unmodified version.

### Step 2 - Making Changes
When you need to edit a Locked View, there are two methods:

Lock, Edit, Save - Make sure the View is locked and leave it locked, make your changes, and then run Save. The disadvantage of this method is that if you close the Explorer without running Save your changes are lost.

Unlock, Edit, Save - Unlock the View, make your changes, and Save or Lock it. The disadvantage of this method is that you can't discard any changes because your changes become permanent when you close the Explorer.

## Side Effects on Open Views
Changing the state of an open View may sometimes hoark up it's appearance. Don't Panic. Just close and reopen the Explorer or switch to a different View and back again.

## Lock Pickers
I have found the following situations where the Locked state of a View is ignored. I'm sure there are more.

Standard Views - If you have modified and Locked any of the Outlook Standard Views (e.g. Compact, Single, Preview) then doing a "Reset" in the Advanced View Settings dialog ignores any locks.

Shared Views - These are Views that have "All Xxxx folders" in the "Can Be Used On" column of the Manage All Views dialog. (e.g. "All Mail and Post folders"). There is really only one copy of these Views on your system, even though they appear in multiple Folders. Making a change to the Locked state on any one of them changes all occurrences of that View for that Folder type system wide.

## Legal
Copyright (C) 2024 Lycon Of Texas

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License Version 3 as published by the Free Software Foundation.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.

## Screen Shots
![Store Scope Form ](https://github.com/user-attachments/assets/11728c3f-c2ec-41cd-b443-e9c2c5a3ee39)
![System Scope - Show And Tell](https://github.com/user-attachments/assets/3a36e21c-605f-44fc-9a9e-66f64286d5b2)
