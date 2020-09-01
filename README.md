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