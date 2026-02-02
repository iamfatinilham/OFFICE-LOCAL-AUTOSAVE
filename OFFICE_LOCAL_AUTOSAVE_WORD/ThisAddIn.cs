using System;
using System.Windows.Forms; // Needed for timer
using Word = Microsoft.Office.Interop.Word;

namespace OFFICE_LOCAL_AUTOSAVE_WORD
{
    public partial class ThisAddIn
    {
        private Timer saveTimer;

        private void ThisAddin_Startup(object sender, System.EventArgs e)
        {
            // Create a lightweight timer
            saveTimer = new Timer();
            saveTimer.Tick += new EventHandler(AutoSave_Tick);

            // Set initial speed from saved settings
            UpdateTimerInterval();

            saveTimer.Start();
        }

        // This function can be called from the Settings Form to change speed instantly

        public void UpdateTimerInterval()
        {
            int seconds = Properties.Settings.Default.IntervalSeconds;

            // Do not allow 0 or negative numbers
            if (seconds < 1) seconds = 30;

            saveTimer.Interval = seconds * 1000; // Convert to miliseconds
        }

        private void AutoSave_Tick(object sender, EventArgs e)
        {
            try
            {
                // Get the active document safely
                Word.Document activeDoc = this.Application.ActiveDocument;

                // LOGIC: Only save if it was saved atleast once and is not read-only
                if (activeDoc != null &&
                    !string.IsNullOrEmpty(activeDoc.Path) &&
                    !activeDoc.ReadOnly &&
                    !activeDoc.Saved)
                {
                    activeDoc.Save(); // Local save
                }
            }
            catch
            {
                // If Word is busy, do nothing.
                // It will catch it on the next loop
            }
        }
        private void ThisAddin_Shutdown(object sender, System.EventArgs e)
        {
            if (saveTimer != null)
            {
                saveTimer.Stop();
                saveTimer.Dispose();
            }
        }
        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddin_Startup);
            this.Shutdown += new System.EventHandler(ThisAddin_Shutdown);
        }
        #endregion
    }
}