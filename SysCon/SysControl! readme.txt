System Control 2.0
By Wesley Crossman

If you have any question or comments, please e-mail me at
wesley_crossman@yahoo.com. Have fun telling Windows what to do!

***** 'Selectable Items' Frame *****

** Start By: **
a: Clicking the Desktop, Start, or Taskbar shortcut button
-or-
b: Entering a HWnd directly
-or-
c: Selecting the 'HWnd Tree' and choosing a HWnd out of the listbox.
-or-
d: Holding down on the 'Drag the Mouse' button and dragging the mouse to any item on anything to select!

[Note: After you select the first target, you can use the listbox to select many new ones.]

** Options: **
[These are commands that affect the target only. Some of these options are in the 'Extended Options' button/popup.]
* 'Get/Set Misc': This will give you various bits of info on the target, show a small picture of it, and allow you to change its text.
* 'Kill': Destroy the item without closing the main prog. This is strongly discouraged and is not recommended. Use 'Close' instead if possible.
* 'Close': Post a message telling the target to close. Posting this to a child window will function similarly to 'Kill', but if you are posting it to a main window, it will close the program. It is highly recommended that you use this instead of 'Kill'.
* 'Flash Window': This will make the item's menubar & entry on the system tray flash 5 times.
* 'Lock': This will block programs from changing the visible elements of the target. (only one at a time)
* 'Unlock': This will free ANY lock.
* 'Enable': Allows user interaction
* 'Disable': Disallows user interaction (does not work on all targets)
* 'Send a Message': This allows you to send a message to the target. It is recommended you have documentation on the topic before experimenting.
* 'Edit the Style Yourself': This will let you change a property of a style bit. As with 'Send a Message', it is quite complicated for the beginner and is only recommended for use with the proper documentation.
* 'Minimize Target': Try it! You can minimize normal windows, buttons, textboxes, etc.! Its effects are generally identical to the 'Minimize' button on a toolbar.
* 'Normalize Target': The effect of this is identical to the 'Restore' button on a toolbar, except that you can use it on anything.
* 'Maximize Target': The effect of this is identical to the 'Maximize' button on a toolbar, except that you can use it on anything.
* 'Stay-On-Top On': Make the target stay on top.
* 'Stay-On-Top Off': Make the target not stay on top.
* 'New Parent': This allows you to transplant the target to any HWnd! You may want to set the Ctrl+Alt+A option to 'Undo', because sometimes this is necessary in cases of transplants covering up options. Remember, if you copy a target to another window besides SysControl, you might crash the unlucky recipient! Other programs are touchy!

** Features: **
[These are items that only supply information.]
* Light: Yellow on nothing selected, Green on existing window, Red on lost window (destroyed by some program)
* 'Parents going up' / 'All Children' / 'All HWnds' Listbox: This listbox serves three functions. It can list an object's children, an object's parents, or every HWnd in a directory-like structure. You can select HWnds by double-clicking on a list item. The different modes can be selected using the buttons under the listbox. You may note slight discrepancies in 'Parents going up' and the two other modes. This is because of the different techniques used to gather the data. 'Parents' uses the GetParent API, and the others use FindWindowEx recursively.
* 'Me' Button: This will fill in the 'New Parent' textbox with Console's HWnd. See "New Parent"
* 'Capture Picture' Button: Capture a picture of only the target. (It will actually capture a picture of the area of the object. If there is something over it, the program will capture that.)

***** 'Running Programs' Frame *****
[Note: Select an entry for much more info and access to the priority control.]
* 'Refresh' Button: Refresh the process list with current program data.
* 'Activate + Send Text' Button: Send text & activate the selected process listbox entry.
* 'Activate' Button: Attempt to activate the selected process listbox entry.
* 'Priority' Nested Listbox: This will allow you to change the priority of any program. However, some care is needed in this area; since very few programs run in realtime mode without freezing the entire system, and a few have some problems in high mode  (e.g. Winamp).

***** Options & Misc. Features *****

* 'Register as Screen Saver' Toggle Button: This will make the system unresponsive to Ctrl+Esc, Alt+Tab, Ctrl+Alt+Del, etc. It really has no actual function, except to demonstrate that API feature.
* 'Stay on Top' Toggle Button: This will make the window stay on top. Nothing special there. :-)
* 'Highlight HWnds' Button: This will suspend all screen activity (while working) and it will cover the entire screen with colors that will be different depending on the HWnd! It's great to analyze all kinds of programs! It will restore the screen after 4 seconds.
* 'On Ctrl+Alt+A' Frame: This sets the hotkey to do the specified function. (The hotkey is not 100% reliable. It was written to be safe. This means if it does not detect the hotkey, try, try again.)
* 'Screen Capture' Button: This will capture the screen onto a window with the ability to save it.
* 'System Info' Button: This gives you large amounts of information on the computer & OS the program is running on.