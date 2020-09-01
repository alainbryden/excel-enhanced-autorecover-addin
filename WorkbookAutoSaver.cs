using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace EnhancedExcelAutoRecover
{
    public class WorkbookAutoSaver : IDisposable
    {
        int MaxBackups { get; } = 10;
        /// <summary>The auto-recover interval to use for all newly opened workbook.</summary>
        /// TODO: Make this settable, but on change we must cancel and create a new PeriodicTask.
        TimeSpan AutoSaveInterval { get; } = TimeSpan.FromMinutes(5);

        private DateTime lastSave = DateTime.MinValue;
        private DateTime lastChange = DateTime.MinValue;
        private readonly Workbook workbook;
        private string saveFolderPath;
        private string extension;
        private Task autoSaveTask;
        private CancellationTokenSource cancellation;
        private bool disposedValue;

        public WorkbookAutoSaver(Workbook wb)
        {
            workbook = wb;
            // Register an event handler to shut ourself down if the workbook closes
            workbook.BeforeClose += (ref bool Cancel) => this.Dispose();
            workbook.SheetChange += (object _sheet, Range _rng) =>
                lastChange = DateTime.UtcNow;
            StartAutoSaveTask();
            // Immediate auto-save of the workbook in it's original state, in case the user messes up the original.
            AutoSave(true);
        }

        public void AutoSave() => AutoSave(false);

        public void AutoSave(bool force)
        {
            // Step 1: Ensure the directory exists
            DirectoryInfo saveDirectory = Directory.CreateDirectory(saveFolderPath);

            // Step 2: Skip this save if there have been no changes since the last save
            if (!force && (lastSave >= lastChange || workbook.Saved))
                return;

            // Step 3: Create a New Save
            lastSave = DateTime.UtcNow;
            workbook.SaveCopyAs($"{saveFolderPath}{DateTime.Now:yyyy-MM-dd HH_mm_ss}{extension}");

            // Step 4: TODO: Post-save, delete backups of the workbook that are identical:
            // - Could populate a cache of back-ups on disk, their lengths MD5 sums for quick comparison

            // Step 5: Limit total number of back-ups. Delete the oldest auto-recovery files.
            foreach (FileInfo fi in saveDirectory.GetFiles().OrderByDescending(x => x.LastWriteTime).Skip(MaxBackups))
                fi.Delete();
        }

        public void StartAutoSaveTask()
        {
            string fullPath = workbook.FullName;
            extension = Path.GetExtension(fullPath);
            saveFolderPath = Path.GetDirectoryName(fullPath) + @"\AutoRecovery\" + Path.GetFileName(fullPath) + @"\";
            cancellation = new CancellationTokenSource();
            autoSaveTask = PeriodicTask.Run(AutoSave, AutoSaveInterval, cancellation.Token);
        }

        public void StopAutoSaveTask()
        {
            cancellation?.Cancel();
            if (autoSaveTask != null)
            {
                Task.WaitAny(autoSaveTask, Task.Delay(1000));
                autoSaveTask.Dispose();
                autoSaveTask = null;
            }
            cancellation.Dispose();
            cancellation = null;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                    StopAutoSaveTask();
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
