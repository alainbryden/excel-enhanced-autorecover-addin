using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace EnhancedExcelAutoRecover
{
    public class ExcelAddin : IExcelAddIn, IDisposable
    {
        private bool disposedValue;
        private readonly List<WorkbookAutoSaver> workbookAutoSavers = new List<WorkbookAutoSaver>();

        public void AutoOpen()
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;
            foreach (Workbook wb in excelApp.Workbooks)
                workbookAutoSavers.Add(new WorkbookAutoSaver(wb));
            excelApp.WorkbookOpen += (Workbook wb) => workbookAutoSavers.Add(new WorkbookAutoSaver(wb));
        }

        public void AutoClose()
        {
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                    workbookAutoSavers.ForEach(wb => wb.Dispose());
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
