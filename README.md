# Enhanced AutoRecover Add-in for Excel
Automatic back-up of open Excel workbooks.

# What's the problem?
Excel crashed. This happens from time to time - after all, it's an impressive chunk of software with vast capabilities, constantly being pushed to its limits by "innovative" end-users, often run on under-powered hardware. I don't fault Excel for crashing, that's what auto-recover is for. Thank goodness for auto-recover, right?

So you launch Excel back up again, but wait...
The "Document Recovery" pane shows only "Your Valuable Work.xlsx [Original]" "Version created last time the user saved the file", but you've been making changes for hours!

Just kidding, that never happens! Really, it shows "Your Valuable Work.xlsx [Autosaved]" "Version created from the last Autosave", phew.
So you open up that Auto-saved version, but wait...
A wild pop-up appears: "We found a problem with some content in 'Your Valuable Work.xlsx'. Do you want us to try and recover as much as we can?" This is probably your own fault for killing Excel in Task Manager before it had a chance to finish what you told it to do.
How robust of you Excel, why \[Yes\], please do recover my data... 

"The file is corrupt and cannot be opened." Well, you did your best Excel, and surely there's a limit to how many back-to-back modal dialogs you're allowed to display in a row, so let's go see what online help is available. Hmm, if these problems keep happening to you, clearly the solution is to repair Excel, okay maybe uninstall then re-install Microsoft Office from scratch, okay next wipe your PC and reinstall Windows entirely.

Just kidding, Excel and its built-in auto-recover never fails.
But perhaps you just made some big changes, saved your file, and then realized that you've made an awful mistake and need to revert your changes. Silly human. But don't worry "Undo" to the rescue... wait, what do you mean "Can't Undo". Why does refreshing pivot tables or running a macro wipe out my undo history? (Because computers are hard).
This solution might help you out in that scenario as well.

Whatever, this whole idea is superfluous now that Excel has added the bold new "Auto-Save" toggle at the top-left of Excel. Wait... what do you mean this feature is only for people who have their files synchronized to OneDrive™?

# Current Features
- Sets up a timer to auto-save each workbook every 5 minutes.
  - Uses the "SaveCopyAs" interop method, which can save out a copy without affecting the current open file.
- Saves go into a folder called "AutoRecovery", in the same folder as the workbook. Copies are date-and-time stamped.
- Keeps only the 10 most recent saved copies.
- Hooks into all open workbooks when the add-in is started.
- Monitors for newly-opened workbooks while the add-in is running.
- Monitors workbook SheetChanged event to only create new save points if a change has been detected since the last save.
  - TODO: Test whether this works in various conditions, such as with Application.EnableEvents disabled.


# Planned Features
- Make the back-up interval and maximum number of back-ups configurable.
- Make the save location configurable.
- More intelligently create back-ups only when workbook contents have changed.
- More intelligently prune back-ups to include a broader set of checkpoints, rather than merely keeping the 10 most recent saves.

# Planned Robustness Improvements
- Deal with a likely myriad of issues surrounding workbooks opened from network drives, temporary folders, in read-only mode, in some sort of other protection mode, etc.
- Notification (toast) system if a workbook cannot be saved for some reason?
- Deal with workbooks that have side-effects when saving (e.g. external data is configured to be cleared on save?
  - This might not be an issue using the current "SaveCopyAs" approach.
- Deal with workbooks that take a long time to save (e.g. very large workbooks)
  - Perhaps allow configuration overrides on a per-workbook basis.
  - Perhaps have an alternative mode that will notify/remind the user to save periodically, rather than forcibly saving.
 
# Screen Shots
Not much to show right now, but here's what it looks to have the add-in up and running, making back-ups of your workbook:
![Image of Explorer](https://github.com/alainbryden/excel-enhanced-autorecover-addin/raw/master/images/explorer.png)
 
# Installation Instructions
- Currently, no "releases", so download the source code and build it yourself.
- You can open the add-in (as a normal workbook) on an as-needed basis, but this is one of those cases where you probably won't remember you need it until it's too late, so you may wish to set it up to auto-load.

Here's a handy one-pager for walking you through registering an add-in to auto-load at startup.
![Excel install instructions](https://github.com/alainbryden/excel-enhanced-autorecover-addin/raw/master/images/install.png)
1. Open Excel's Options Dialog
2. Click the "Add-ins" section
3. At the botom, ensure the "Manage" drop-down is set to "Excel Add-ins", and click "Go..."
4. Click "Browse" to navigate the add-in file.
5. Select either the 32 bit or 64 bit add-in based on the version of office you have installed.
   - Note that most people have Office 32-bit installed, even on a 64-bit computer, since this is the Microsoft recommendation.
   - In my case, I'm grabbing the add-in out of the "git/excel-enhanced-autorecover-addin/bin/Debug" build directory directly. You may wish to save out a stable build somewhere easier to track down.
6. Click "OK". If you see a prompt asking if you wish to make a copy to your "add-ins" folder, it's up to you. I choose "No" because having Excel make a separate copy of the add-in in a difficult-to-locate user settings folder makes updating the add-in a pain.
7. You should see the add-in appear in the available add-ins list now, checked to auto-load at startup.
