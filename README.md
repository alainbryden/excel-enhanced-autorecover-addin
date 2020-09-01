# Excel Enhanced Auto-recover Add-in
Manually manages auto-back-up and recovery of Excel documents being edited (because Excel AutoRecover is unreliable)

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
